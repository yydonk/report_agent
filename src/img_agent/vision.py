import base64
import io
import json
import os
import requests
from typing import List, Dict, Any
from PIL import Image
from .schemas import StructuredExperiment, StepItem, empty_structured
from .deepseek_client import DeepSeekClient

# DashScope native vision endpoint (qwen-vl does NOT work via compatible-mode base64)
_DASHSCOPE_VL_URL = "https://dashscope.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation"


def _img_to_b64(path: str, max_w: int = 1600) -> str:
    """Return pure base64 string (no data-URL prefix)."""
    img = Image.open(path).convert("RGB")
    w, h = img.size
    if w > max_w:
        nh = int(h * (max_w / float(w)))
        img = img.resize((max_w, nh))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    return base64.b64encode(buf.getvalue()).decode("utf-8")


def _img_to_data_url(path: str, max_w: int = 1600) -> str:
    return f"data:image/jpeg;base64,{_img_to_b64(path, max_w)}"


def _call_dashscope_vl(image_paths: List[str], api_key: str, model: str, prompt: str) -> str:
    """Call DashScope native multimodal API. Returns raw text response."""
    user_content = []
    for p in image_paths:
        b64 = _img_to_b64(p)
        user_content.append({"image": f"data:image/jpeg;base64,{b64}"})
    user_content.append({"text": prompt})

    payload = {
        "model": model,
        "input": {
            "messages": [
                {"role": "system", "content": [{"text": "你是严谨的科学实验助手，只输出严格JSON。"}]},
                {"role": "user", "content": user_content},
            ]
        },
        "parameters": {"temperature": 0.1, "max_tokens": 8192},
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    resp = requests.post(_DASHSCOPE_VL_URL, headers=headers, json=payload, timeout=300)
    if not resp.ok:
        raise RuntimeError(f"DashScope VL API error {resp.status_code}: {resp.text[:300]}")
    data = resp.json()
    # Native API response: output.choices[0].message.content is a list of {text:...}
    try:
        content = data["output"]["choices"][0]["message"]["content"]
        if isinstance(content, list):
            return "".join(c.get("text", "") for c in content if isinstance(c, dict))
        return str(content)
    except Exception:
        raise RuntimeError(f"Unexpected DashScope VL response: {str(data)[:300]}")


def _build_vision_messages(image_paths: List[str]) -> List[Dict[str, Any]]:
    """Build OpenAI-compatible messages (kept for non-DashScope providers)."""
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
    content: List[Dict[str, Any]] = [
        {
            "type": "text",
            "text": "你将收到多张实验相关图片（包含文字、步骤、表格、公式与图表）。请严格依据图片内容，抽取标准化实验要素，输出严格的JSON，禁止添加额外说明。\nJSON字段定义如下：\n"
            + schema_str
            + "\n若信息缺失请留空或给出空数组，保持字段完整。",
        }
    ]
    for p in image_paths:
        data_url = _img_to_data_url(p)
        content.append({"type": "image_url", "image_url": {"url": data_url}})
    messages = [
        {"role": "system", "content": "你是严谨的科学实验助手，只输出严格JSON。"},
        {"role": "user", "content": content},
    ]
    return messages


def _build_vl_prompt() -> str:
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
    return (
        "你将收到多张实验相关图片（包含文字、步骤、表格、公式与图表）。"
        "请严格依据图片内容，抽取标准化实验要素，输出严格的JSON，禁止添加额外说明。\n"
        "JSON字段定义如下：\n" + schema_str +
        "\n若信息缺失请留空或给出空数组，保持字段完整。"
    )


def parse_with_deepseek_vision(image_paths: List[str], client: DeepSeekClient) -> StructuredExperiment:
    # Use DashScope native VL API when api_key is present (compatible-mode ignores base64 images)
    api_key = client.api_key
    model = client.model
    messages = _build_vision_messages(image_paths)
    text = client.chat(messages, temperature=0.1, max_tokens=8192).strip()
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end >= start:
        text = text[start : end + 1]
    data: Dict[str, Any] = {}
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
                    steps.append(
                        StepItem(
                            step_id=str(it.get("step_id") or f"{i+1}"),
                            description=str(it.get("description") or ""),
                            parameters=it.get("parameters"),
                            observation=str(it.get("observation") or "") if it.get("observation") is not None else None,
                            notes=str(it.get("notes") or "") if it.get("notes") is not None else None,
                        )
                    )
                elif isinstance(it, str):
                    steps.append(StepItem(step_id=str(i + 1), description=it))
        s.steps = steps
        d = data.get("data") or {}
        if isinstance(d, dict):
            s.data = d
        s.analysis = str(data.get("analysis", "") or "")
    if not s.data:
        s.data = {"tables": [], "observations": [], "raw_text": ""}
    elif "raw_text" not in s.data:
        s.data["raw_text"] = ""
    return s
