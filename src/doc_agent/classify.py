import json
from typing import Dict, Any, List
from src.img_agent.deepseek_client import DeepSeekClient


def refine_with_llm(spec: Dict[str, Any], client: DeepSeekClient) -> Dict[str, Any]:
    fmt = spec.get("format_spec", {})
    cnt = spec.get("content_spec", {})
    text = spec.get("full_text_sample", "")
    prompt = {
        "instruction": "对给定的模板规则与文本样本进行复核，输出两个分类：格式规范与内容规范。每一类请给出清晰层级与要点，严格JSON。",
        "format_spec": fmt,
        "content_spec": cnt,
        "sample_text": text,
        "schema": {
            "format_spec": {
                "page": {"orientation": "", "margins": {}, "page_size": ""},
                "paragraph": [{"style": "", "line_spacing": "", "space_before": "", "space_after": ""}],
                "headings": [{"level": 1, "style": "", "name": ""}],
                "tables": [{"style": "", "count": 0}],
                "captions": [{"style": "", "has_numbering": True, "rules": ""}],
                "others": []
            },
            "content_spec": {
                "required_sections": [],
                "data_recording_rules": [],
                "analysis_methods": [],
                "conclusion_requirements": [],
                "placeholders": [],
                "notes": []
            }
        }
    }
    messages = [
        {"role": "system", "content": "你是一个严谨的文档模板规范分析助手。仅输出严格JSON。"},
        {"role": "user", "content": json.dumps(prompt, ensure_ascii=False)},
    ]
    out = client.chat(messages, temperature=0.2, max_tokens=2048)
    s = out.strip()
    i = s.find("{")
    j = s.rfind("}")
    if i != -1 and j != -1 and j >= i:
        s = s[i : j + 1]
    try:
        obj = json.loads(s)
        return obj
    except Exception:
        return {"format_spec": fmt, "content_spec": cnt}
