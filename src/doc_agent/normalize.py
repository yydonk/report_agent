import datetime
from typing import Dict, Any, List


def _llm_to_content(llm: Dict[str, Any]) -> Dict[str, Any]:
    if not isinstance(llm, dict):
        return {}
    # Support both English and Chinese keys
    c = llm.get("content_spec") or llm.get("内容规范") or {}
    out = {
        "required_sections": [],
        "data_recording_rules": [],
        "analysis_methods": [],
        "conclusion_requirements": [],
        "placeholders": [],
        "notes": [],
    }
    if isinstance(c, list):
        # merge list of dicts
        merged: Dict[str, Any] = {}
        for it in c:
            if isinstance(it, dict):
                for k, v in it.items():
                    merged[k] = v
        c = merged
    if not isinstance(c, dict):
        return out
    # required sections
    rs = c.get("required_sections") or c.get("必需章节") or []
    if not rs:
        x = c.get("结构完整性")
        if isinstance(x, dict):
            rs = x.get("必需章节", [])
        elif isinstance(x, list):
            tmp: List[str] = []
            for it in x:
                if isinstance(it, str):
                    tmp.append(it)
                elif isinstance(it, dict):
                    vv = it.get("必需章节")
                    if isinstance(vv, list):
                        tmp.extend([str(i) for i in vv])
            rs = tmp
    if isinstance(rs, list):
        out["required_sections"] = [str(x) for x in rs]
    # data rules
    dr = c.get("data_recording_rules") or c.get("数据记录要求") or []
    if isinstance(dr, dict):
        # Flatten dictionary fields into lines
        lines: List[str] = []
        for k, v in dr.items():
            if isinstance(v, (str, int, float)):
                lines.append(f"{k}: {v}")
            elif isinstance(v, list):
                for it in v:
                    lines.append(f"{k}: {it}")
        dr = lines
    if isinstance(dr, list):
        out["data_recording_rules"] = [str(x) for x in dr]
    # analysis methods
    am = c.get("analysis_methods") or c.get("分析方法规范") or c.get("分析方法") or []
    if isinstance(am, dict):
        lines = []
        for k, v in am.items():
            if isinstance(v, (str, int, float)):
                lines.append(f"{k}: {v}")
            elif isinstance(v, list):
                for it in v:
                    lines.append(f"{k}: {it}")
        am = lines
    if isinstance(am, list):
        out["analysis_methods"] = [str(x) for x in am]
    # conclusion requirements
    cr = c.get("conclusion_requirements") or c.get("结论撰写标准") or c.get("结论要求") or []
    if isinstance(cr, dict):
        lines = []
        for k, v in cr.items():
            if isinstance(v, (str, int, float)):
                lines.append(f"{k}: {v}")
            elif isinstance(v, list):
                for it in v:
                    lines.append(f"{k}: {it}")
        cr = lines
    if isinstance(cr, list):
        out["conclusion_requirements"] = [str(x) for x in cr]
    # placeholders
    ph = c.get("placeholders") or c.get("占位符与提示语") or c.get("占位与提示") or []
    if isinstance(ph, dict):
        lines = []
        for k, v in ph.items():
            if isinstance(v, (str, int, float)):
                lines.append(f"{k}: {v}")
            elif isinstance(v, list):
                for it in v:
                    lines.append(f"{k}: {it}")
        ph = lines
    if isinstance(ph, list):
        out["placeholders"] = [str(x) for x in ph]
    # notes
    notes = c.get("notes") or c.get("备注") or []
    if isinstance(notes, list):
        out["notes"] = [str(x) for x in notes]
    return out


def normalize_spec(base_spec: Dict[str, Any], llm_obj: Dict[str, Any], template_path: str) -> Dict[str, Any]:
    fmt = base_spec.get("format_spec", {}) if isinstance(base_spec, dict) else {}
    cnt = base_spec.get("content_spec", {}) if isinstance(base_spec, dict) else {}
    sections = fmt.get("sections") or []
    page = {}
    if sections:
        s0 = sections[0]
        page = {
            "orientation": s0.get("orientation"),
            "size_emu": {"width": s0.get("page_width"), "height": s0.get("page_height")},
            "margins_emu": s0.get("margins", {}),
        }
    paragraph_styles = fmt.get("paragraph_styles") or []
    table_styles = fmt.get("table_styles") or []
    captions = fmt.get("captions") or []
    headings_outline = cnt.get("heading_outline") or []
    heading_blocks = cnt.get("heading_blocks") or []
    # content rules
    base_content_rules = {
        "required_sections": cnt.get("required_sections") or [],
        "data_recording_rules": cnt.get("data_recording") or [],
        "analysis_methods": cnt.get("analysis_methods") or [],
        "conclusion_requirements": cnt.get("conclusion_requirements") or [],
        "placeholders": cnt.get("placeholders") or [],
        "notes": [],
    }
    llm_content_rules = _llm_to_content(llm_obj)
    # prefer LLM if present, else base
    content_spec = {}
    for k in base_content_rules.keys():
        v = llm_content_rules.get(k)
        content_spec[k] = v if v else base_content_rules.get(k) or []
    content_spec["heading_outline"] = headings_outline
    content_spec["heading_blocks"] = heading_blocks
    format_spec = {
        "page": page,
        "paragraph_styles": paragraph_styles,
        "style_definitions": fmt.get("style_definitions") or {},
        "headings": [{"level": h.get("level"), "style": h.get("style"), "text": h.get("text")} for h in headings_outline],
        "tables": table_styles,
        "captions": captions,
    }
    return {
        "format_spec": format_spec,
        "content_spec": content_spec,
        "meta": {"template_path": template_path, "generated_at": datetime.datetime.utcnow().isoformat() + "Z"},
    }
