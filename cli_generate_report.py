import argparse
import json
import os
from src.img_agent.deepseek_client import DeepSeekClient
from src.report_agent.generator import generate_report
from src.report_agent.renderer import render_report, render_report_from_draft, read_draft_docx
from src.config.env import get_env_config


def main():
    ap = argparse.ArgumentParser(prog="gen-report", description="合并内容要求与实验内容，结合格式导出 Word 实验报告")
    ap.add_argument("--format", required=True)
    ap.add_argument("--content", required=True)
    ap.add_argument("--out-docx", required=True)
    # 内容来源：草稿 docx 或原始实验数据（二选一）
    source = ap.add_mutually_exclusive_group(required=True)
    source.add_argument("--draft", help="用户编辑完的草稿 docx 路径（直接套格式，跳过LLM生成）")
    source.add_argument("--experiment", help="实验内容 json/txt 路径（调用LLM生成内容）")
    ap.add_argument("--target-length", type=int)
    ap.add_argument("--api_key")
    ap.add_argument("--model")
    ap.add_argument("--base_url")
    args = ap.parse_args()

    with open(args.format, "r", encoding="utf-8") as f:
        fmt = json.load(f)
    with open(args.content, "r", encoding="utf-8") as f:
        cnt = json.load(f)

    os.makedirs(os.path.dirname(args.out_docx), exist_ok=True) if os.path.dirname(args.out_docx) else None

    if args.draft:
        # 直接从用户编辑的草稿套格式，保留所有内容（含图片）
        render_report_from_draft(args.draft, fmt, args.out_docx)
    else:
        exp_path = args.experiment
        if exp_path.lower().endswith(".json"):
            with open(exp_path, "r", encoding="utf-8") as f:
                exp_obj = json.load(f)
            exp_text = json.dumps(exp_obj, ensure_ascii=False)
        else:
            with open(exp_path, "r", encoding="utf-8") as f:
                exp_text = f.read()
        cfg = get_env_config()
        client = DeepSeekClient(
            api_key=args.api_key or cfg.get("DASHSCOPE_API_KEY"),
            base_url=args.base_url or cfg.get("DASHSCOPE_BASE_URL"),
            model=args.model or cfg.get("DASHSCOPE_MODEL") or "qwen-plus",
        )
        gen = generate_report(cnt, exp_text, client, target_length_chars=args.target_length)
        render_report(fmt, gen, args.out_docx)

    print(args.out_docx)


if __name__ == "__main__":
    main()
