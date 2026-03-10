import json
from typing import Dict, Any, Optional
from src.img_agent.deepseek_client import DeepSeekClient


def generate_report(content_schema: Dict[str, Any], experiment_text: str, client: DeepSeekClient, target_length_chars: Optional[int] = None) -> Dict[str, Any]:
    chapters = content_schema.get("chapters", {})
    brief = {}
    for k, v in chapters.items():
        brief[k] = {
            "title": v.get("title"),
            "minWordCount": v.get("minWordCount"),
            "maxWordCount": v.get("maxWordCount"),
            "description": v.get("description"),
            "figureTable": v.get("figureTable"),
            "dataFormat": v.get("dataFormat"),
            "citationCount": v.get("citationCount"),
            "required": v.get("required"),
        }
    prompt = {
        "instruction": "根据给定的章节约束与实验内容撰写完整实验报告正文。严格以JSON输出，每章包含 title 与 content 字段，content 为纯文本。",
        "chapters": brief,
        "experiment_text": experiment_text,
        "rules": {
            "language": "zh",
            "format": "json",
            "constraints": ["满足每章字数下限", "覆盖描述中的要点", "章节顺序与标题保持一致"]
        }
    }
    if target_length_chars:
        prompt["rules"]["target_length_chars"] = target_length_chars
    messages = [
        {"role": "system", "content": "你是一个严谨的实验报告写作助手。仅输出严格JSON。"},
        {"role": "user", "content": json.dumps(prompt, ensure_ascii=False)},
    ]
    s = client.chat(messages, temperature=0.2, max_tokens=4096)
    i = s.find("{")
    j = s.rfind("}")
    if i != -1 and j != -1 and j >= i:
        s = s[i : j + 1]
    try:
        obj = json.loads(s)
        return obj
    except Exception:
        res = {}
        for k, v in chapters.items():
            res[k] = {"title": v.get("title"), "content": ""}
        return res
