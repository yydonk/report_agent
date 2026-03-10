import os

_DEFAULTS = {
    "DASHSCOPE_BASE_URL": "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
    "DASHSCOPE_MODEL": "qwen-plus",
    "REPORT_ABSTRACT_MIN_WORDS": "100",
    "REPORT_ABSTRACT_MAX_WORDS": "200",
    "REPORT_REFERENCES_MIN_COUNT": "1",
    "REPORT_TARGET_CHARS": "3000",
}


def get_env_config():
    return {k: os.getenv(k, _DEFAULTS.get(k)) for k in [
        "DASHSCOPE_API_KEY",
        "DASHSCOPE_BASE_URL",
        "DASHSCOPE_MODEL",
        "REPORT_ABSTRACT_MIN_WORDS",
        "REPORT_ABSTRACT_MAX_WORDS",
        "REPORT_REFERENCES_MIN_COUNT",
        "REPORT_TARGET_CHARS",
    ]}
