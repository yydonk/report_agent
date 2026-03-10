import json
from typing import Dict, Any, List
from .schemas import StructuredExperiment, StepItem, TableData, OCRResult, empty_structured
from .deepseek_client import DeepSeekClient


def build_prompt_text(ocr: OCRResult) -> str:
    raw = ocr.raw_text.strip()
    head = "从以下OCR文本中抽取标准化的实验要素，输出严格的JSON。"
    fields = {
        "title": "实验标题",
        "objective": "实验目的",
        "theory": "实验原理",
        "apparatus": ["器材A", "器材B"],
        "steps": [{"description": "分步骤描述", "parameters": {}, "observation": "", "notes": ""}],
        "data": {"tables": [{"name": "", "columns": [], "rows": [], "units": {}}], "observations": [], "raw_text": ""},
        "analysis": "实验结果分析",
    }
    schema_str = json.dumps(fields, ensure_ascii=False, indent=2)
    example = "仅输出JSON，不要添加额外文字。"
    return f"{head}\nJSON字段:\n{schema_str}\n{example}\nOCR文本:\n{raw}"


def parse_with_deepseek(ocr: OCRResult, client: DeepSeekClient) -> StructuredExperiment:
    prompt = build_prompt_text(ocr)
    messages = [
        {"role": "system", "content": "你是一个科学实验助手。请基于文本抽取结构化要素，确保准确性，缺失项留空。只输出JSON。"},
        {"role": "user", "content": prompt},
    ]
    content = client.chat(messages, temperature=0.1, max_tokens=2048)
    text = content.strip()
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end >= start:
        text = text[start : end + 1]
    data = {}
    try:
        data = json.loads(text)
    except Exception:
        data = {}
    s = empty_structured()
    if isinstance(data, dict):
        s.title = str(data.get("title", "") or "")
        s.objective = str(data.get("objective", "") or "")
        s.theory = str(data.get("theory", "") or "")
        app = data.get("apparatus") or []
        if isinstance(app, list):
            s.apparatus = [str(x) for x in app]
        steps_raw = data.get("steps") or []
        steps: List[StepItem] = []
        if isinstance(steps_raw, list):
            for i, it in enumerate(steps_raw):
                if isinstance(it, dict):
                    st = StepItem(
                        step_id=str(it.get("step_id") or f"{i+1}"),
                        description=str(it.get("description") or ""),
                        parameters=it.get("parameters"),
                        observation=str(it.get("observation") or "") if it.get("observation") is not None else None,
                        notes=str(it.get("notes") or "") if it.get("notes") is not None else None,
                    )
                    steps.append(st)
                elif isinstance(it, str):
                    steps.append(StepItem(step_id=str(i + 1), description=it))
        s.steps = steps
        d = data.get("data") or {}
        if isinstance(d, dict):
            s.data = d
        s.analysis = str(data.get("analysis", "") or "")
    if not s.data:
        s.data = {"tables": [], "observations": [], "raw_text": ocr.raw_text}
    else:
        if "raw_text" not in s.data:
            s.data["raw_text"] = ocr.raw_text
    return s
