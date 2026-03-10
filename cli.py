import argparse
import json
import os
from typing import List
from src.img_agent.deepseek_client import DeepSeekClient
from src.img_agent.parser import parse_with_deepseek
from src.img_agent.schemas import OCRResult
from src.img_agent.vision import parse_with_deepseek_vision


def run_vision(paths: List[str], output: str, api_key: str = None, model: str = None, base_url: str = None):
    client = DeepSeekClient(api_key=api_key, model=model or "deepseek-vl", base_url=base_url)
    s = parse_with_deepseek_vision(paths, client)
    os.makedirs(os.path.dirname(output), exist_ok=True) if os.path.dirname(output) else None
    with open(output, "w", encoding="utf-8") as f:
        f.write(s.to_json())
    print(output)


def run_demo(output: str):
    sample = """
实验标题：电阻串联与并联关系测定
实验目的：掌握欧姆定律应用，验证串并联等效电阻计算方法
实验原理：通过测量电压、电流，依据欧姆定律计算电阻；串联电阻相加，并联电阻倒数相加
实验器材：直流电源，电阻箱，万用表，导线，面包板
实验步骤：1. 连接串联电路，记录总电流与两端电压；2. 连接并联电路，记录各支路电流与总电流；3. 多组电压电流测量求等效电阻
实验数据：表1 电压-电流测量数据；图1 电路连接示意
实验结果分析：根据数据拟合R≈99.8Ω，与标称100Ω一致，误差0.2%
"""
    ocr = OCRResult(blocks=[], raw_text=sample)
    client = DeepSeekClient()
    s = parse_with_deepseek(ocr, client)
    os.makedirs(os.path.dirname(output), exist_ok=True) if os.path.dirname(output) else None
    with open(output, "w", encoding="utf-8") as f:
        f.write(s.to_json())
    print(output)


def main():
    parser = argparse.ArgumentParser(prog="img2struct", description="图片实验内容解析子agent")
    sub = parser.add_subparsers(dest="cmd")
    p1 = sub.add_parser("run")
    p1_group = p1.add_mutually_exclusive_group(required=True)
    p1_group.add_argument("--images", nargs="+")
    p1_group.add_argument("--dir")
    p1.add_argument("--output", required=True)
    p1.add_argument("--api_key", required=False)
    p1.add_argument("--model", required=False, default="deepseek-vl")
    p1.add_argument("--base_url", required=False)
    p2 = sub.add_parser("demo")
    p2.add_argument("--output", required=True)
    p2.add_argument("--api_key", required=False)
    args = parser.parse_args()
    if args.cmd == "run":
        paths = []
        if getattr(args, "dir", None):
            import glob
            root = args.dir
            patterns = ["*.png", "*.jpg", "*.jpeg", "*.bmp", "*.tif", "*.tiff"]
            for pat in patterns:
                paths.extend(glob.glob(os.path.join(root, pat)))
        else:
            paths = args.images
        run_vision(
            paths,
            args.output,
            api_key=getattr(args, "api_key", None),
            model=getattr(args, "model", None),
            base_url=getattr(args, "base_url", None),
        )
    elif args.cmd == "demo":
        if getattr(args, "api_key", None):
            os.environ["DEEPSEEK_API_KEY"] = args.api_key
        run_demo(args.output)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
