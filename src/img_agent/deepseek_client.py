import os
import json
from typing import List, Dict, Any, Optional
import requests
from src.config.env import get_env_config


class DeepSeekClient:
    def __init__(self, api_key: Optional[str] = None, base_url: Optional[str] = None, model: str = "deepseek-chat"):
        # Support multiple env var names to allow different providers without code changes
        cfg = get_env_config()
        self.api_key = (
            api_key
            or os.getenv("DEEPSEEK_API_KEY")
            or cfg.get("DASHSCOPE_API_KEY")
            or os.getenv("OPENAI_API_KEY")
        )
        self.base_url = (
            base_url
            or os.getenv("DEEPSEEK_BASE_URL")
            or cfg.get("DASHSCOPE_BASE_URL")
            or os.getenv("OPENAI_BASE_URL")
            or "https://api.deepseek.com/chat/completions"
        )
        # Explicit model argument takes priority; env vars are only defaults
        self.model = model or os.getenv("DEEPSEEK_MODEL") or cfg.get("DASHSCOPE_MODEL")

    def chat(self, messages: List[Dict[str, str]], temperature: float = 0.2, max_tokens: Optional[int] = None) -> str:
        if not self.api_key:
            raise RuntimeError("Missing DEEPSEEK_API_KEY")
        payload: Dict[str, Any] = {
            "model": self.model,
            "messages": messages,
            "temperature": temperature,
        }
        if max_tokens is not None:
            payload["max_tokens"] = max_tokens
        headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
        url = self.base_url
        resp = requests.post(url, headers=headers, data=json.dumps(payload), timeout=300)
        if not resp.ok:
            # Bubble up server error text for easier troubleshooting
            try:
                err = resp.json()
            except Exception:
                err = {"error": resp.text}
            raise RuntimeError(f"DeepSeek API error {resp.status_code}: {err}")
        data = resp.json()
        if "choices" in data and data["choices"]:
            content = data["choices"][0]["message"]["content"]
            # qwen-vl returns content as a list of blocks, extract text
            if isinstance(content, list):
                content = "".join(
                    block.get("text", "") for block in content if isinstance(block, dict)
                )
            return content
        if "output_text" in data:
            return data["output_text"]
        raise RuntimeError("Unexpected response from DeepSeek API")
