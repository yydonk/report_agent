import os
import glob
import json
import argparse
from decimal import Decimal, ROUND_HALF_UP
from src.doc_agent.doc_loader import convert_doc_to_docx
from src.doc_agent.word_rules import analyze_docx
from src.doc_agent.classify import refine_with_llm
from src.doc_agent.normalize import normalize_spec
from src.img_agent.deepseek_client import DeepSeekClient
from src.config.env import get_env_config
import re


def f2(v):
    try:
        return str(Decimal(v).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP))
    except Exception:
        try:
            return str(Decimal(str(v)).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP))
        except Exception:
            return "0.00"


def build_format_schema(normalized_list):
    docs = []
    for n in normalized_list:
        fmt = n.get("format_spec", {})
        page = fmt.get("page", {})
        ps = fmt.get("paragraph_styles", [])
        heads = fmt.get("headings", [])
        tables = fmt.get("tables", [])
        caps = fmt.get("captions", [])
        # XML-extracted style definitions: {style_name: {fontSizePt, fontEastAsia, bold, lineSpacing, ...}}
        style_defs = fmt.get("style_definitions") or {}

        def _style_val(style_name, key, fallback=None):
            return (style_defs.get(style_name) or {}).get(key, fallback)

        def _para_style_entry(it):
            name = it.get("style")
            # Prefer East Asian font (Chinese) if present, else ASCII font
            font = _style_val(name, "fontEastAsia") or _style_val(name, "fontAscii")
            # Prefer paragraph-level line_spacing if set, else style definition value
            line_sp = it.get("line_spacing") or _style_val(name, "lineSpacing")
            # Indent: prefer style_def EMU values, fall back to paragraph_format values
            first_emu = _style_val(name, "firstLineIndentEmu") or it.get("first_line_indent")
            left_emu = _style_val(name, "leftIndentEmu") or it.get("left_indent")
            right_emu = _style_val(name, "rightIndentEmu") or it.get("right_indent")
            return {
                "name": name,
                "appliesTo": "paragraph",
                "fontFamily": font,
                "fontSizePt": _style_val(name, "fontSizePt"),
                "bold": _style_val(name, "bold"),
                "italic": _style_val(name, "italic"),
                "alignment": _style_val(name, "alignment") or it.get("alignment"),
                "indent": {
                    "firstLineEmu": first_emu,
                    "leftEmu": left_emu,
                    "rightEmu": right_emu,
                },
                "lineSpacing": line_sp,
                "spaceBeforePt": _style_val(name, "spaceBeforePt"),
                "spaceAfterPt": _style_val(name, "spaceAfterPt"),
                "numberingRule": None,
                "allowedValues": {},
            }

        doc_item = {
            "fileName": os.path.basename(n.get("meta", {}).get("template_path", "")),
            "pageLayout": {
                "orientation": page.get("orientation"),
                "sizeEmu": {"width": page.get("size_emu", {}).get("width"), "height": page.get("size_emu", {}).get("height")},
                "marginsEmu": page.get("margins_emu", {}),
            },
            "headingLevels": [{"level": h.get("level"), "style": h.get("style"), "name": h.get("text")} for h in heads],
            "paragraphStyles": [_para_style_entry(it) for it in ps],
            "figureTableRules": {
                "captionStyles": [{"style": c.get("style"), "hasNumbering": c.get("has_numbering")} for c in caps],
                "tableStyles": [{"style": t.get("style"), "count": t.get("count")} for t in tables],
            },
            "headerFooter": {"header": None, "footer": None},
            "toc": {"exists": False, "style": None, "levels": None},
            "citationStyle": {"type": "numericBracket", "example": "[1]"},
        }
        docs.append(doc_item)
    out = {"documents": docs, "summary": {"totalDocs": f2(len(docs))}}
    return out


_FMT_TOKENS = ["样式", "黑体", "宋体", "楷体", "行距", "段前", "段后", "缩进",
               "居中", "对齐", "加粗", "小四", "四号", "五号", "Heading",
               "不分页", "同页", "首行缩进"]


def _is_format_heading(title: str) -> bool:
    """Return True if the heading is purely a format instruction with no real chapter name."""
    t = (title or "").strip()
    if not t:
        return False
    if re.match(r"^\s*\d+\.\s*(插图|表格|重要内容强调)\s*$", t):
        return True
    # After cleaning, if nothing remains it was purely a format note
    from src.doc_agent.word_rules import _clean_heading_text
    cleaned = _clean_heading_text(t)
    return not cleaned


def _strip_format_tokens(text: str) -> str:
    if not text:
        return text
    s = text
    s = re.sub(r"（[^）]*[样式字体行距段前段后缩进]+[^）]*）", "", s)
    s = re.sub(r"\(\([^)]*\)\)", "", s)
    lines = []
    for line in s.splitlines():
        if any(tok in line for tok in _FMT_TOKENS):
            continue
        lines.append(line)
    return "\n".join(ln for ln in lines if ln.strip())


def build_content_schema(normalized_list):
    cfg = get_env_config()
    try:
        abs_min = int(cfg.get("REPORT_ABSTRACT_MIN_WORDS", "100"))
    except Exception:
        abs_min = 100
    try:
        abs_max = int(cfg.get("REPORT_ABSTRACT_MAX_WORDS", "200"))
    except Exception:
        abs_max = 200
    try:
        refs_min = int(cfg.get("REPORT_REFERENCES_MIN_COUNT", "1"))
    except Exception:
        refs_min = 1
    chapters = {}
    for n in normalized_list:
        c = n.get("content_spec", {})
        outline = c.get("heading_outline", [])
        blocks = c.get("heading_blocks", [])
        block_map = {}
        for b in blocks:
            t = (b.get("text") or "").strip()
            if not t:
                continue
            block_map[t] = b.get("content") or ""
        req = set(c.get("required_sections", []))
        for h in outline:
            title = (h.get("text") or "").strip()
            key = title.replace(" ", "")
            if not key:
                continue
            if _is_format_heading(title):
                continue
            if key not in chapters:
                is_req = any(k in title for k in req) or any(k in title for k in ["摘要", "引言", "实验目的", "实验原理", "实验步骤", "实验数据", "结果分析", "结论", "参考文献", "附录"])
                chapters[key] = {
                    "title": title,
                    "description": _strip_format_tokens(block_map.get(title) or "") or None,
                    "minWordCount": f2(0),
                    "maxWordCount": f2(0),
                    "figureTable": {"minFigures": f2(0), "maxFigures": f2(0), "minTables": f2(0), "maxTables": f2(0)},
                    "dataFormat": {"precision": None, "units": [], "examples": []},
                    "citationCount": {"min": f2(0), "max": f2(0)},
                    "qualityMetrics": [{"name": "relevance", "weight": f2(0.50)}, {"name": "clarity", "weight": f2(0.50)}],
                    "validationRules": [],
                    "required": True if is_req else False,
                }
            if "摘要" in title:
                chapters[key]["minWordCount"] = f2(abs_min)
                chapters[key]["maxWordCount"] = f2(abs_max)
            if "参考文献" in title:
                chapters[key]["citationCount"] = {"min": f2(refs_min), "max": f2(999)}
    return {"chapters": chapters}


def main():
    ap = argparse.ArgumentParser(prog="build-schemas", description="遍历解析 .doc/.docx 模板并生成两类 schema")
    ap.add_argument("--out-dir", default=".")
    ap.add_argument("--no-llm", action="store_true")
    args = ap.parse_args()
    cwd = os.getcwd()
    docs = glob.glob(os.path.join(cwd, "*.doc")) + glob.glob(os.path.join(cwd, "**", "*.doc"), recursive=True)
    docx_list = convert_doc_to_docx(docs) if docs else []
    if not docx_list:
        docx_list = glob.glob(os.path.join(cwd, "*.docx")) + glob.glob(os.path.join(cwd, "**", "*.docx"), recursive=True)
    # filter out temp/lock files and previously generated artifacts
    norm = []
    for p in docx_list:
        base = os.path.basename(p)
        if base.startswith("~$"):
            continue
        winp = p.replace("/", "\\").lower()
        if "\\artifacts\\" in winp:
            continue
        norm.append(p)
    docx_list = norm
    normalized_list = []
    if docx_list:
        client = None if args.no_llm else DeepSeekClient()
        for p in docx_list:
            base = analyze_docx(p)
            llm_obj = refine_with_llm(base, client) if client else {}
            normalized = normalize_spec(base, llm_obj, p)
            normalized_list.append(normalized)
    fmt = build_format_schema(normalized_list)
    cnt = build_content_schema(normalized_list)
    def conv(x):
        if isinstance(x, bool) or x is None:
            return x
        if isinstance(x, (int, float)):
            return f2(x)
        if isinstance(x, dict):
            return {k: conv(v) for k, v in x.items()}
        if isinstance(x, list):
            return [conv(i) for i in x]
        return x
    fmt2 = conv(fmt)
    cnt2 = conv(cnt)
    out_dir = args.out_dir
    os.makedirs(out_dir, exist_ok=True)
    fmt_path = os.path.join(out_dir, "report_format_schema.json")
    cnt_path = os.path.join(out_dir, "report_content_schema.json")
    with open(fmt_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(fmt2, ensure_ascii=False, separators=(",", ":")))
        f.write("\n")
    with open(cnt_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(cnt2, ensure_ascii=False, separators=(",", ":")))
        f.write("\n")
    print(fmt_path)
    print(cnt_path)


if __name__ == "__main__":
    main()
