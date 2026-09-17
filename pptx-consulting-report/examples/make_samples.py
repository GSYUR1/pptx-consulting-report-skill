# -*- coding: utf-8 -*-
"""生成 README 展示样张：风格 A (ILLUMINERA 青) / 风格 B (IQVIA 深蓝) 各 3 页。

用法（在本仓库根目录执行）:
    python examples/make_samples.py
输出:
    samples/sample_styleA.pptx / sample_styleB.pptx（同目录附 QA 渲染截图）
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "scripts"))
from builders import (  # noqa: E402
    new_deck, save_deck, build_cover, build_compare, build_matrix,
)

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples")
os.makedirs(OUT, exist_ok=True)

# ---------- 演示内容（虚构数据，仅作版式展示） ----------
COVER = {
    "title": "中国创新药出海趋势研究",
    "subtitle": "2026 年市场格局与机会展望",
    "meta": ["示例报告 · 仅作版式展示", "2026 年 9 月"],
}
COMPARE = [
    ("国内市场", [
        "集采常态化后，仿制药利润空间持续压缩",
        "创新药医保谈判通过率提升，放量周期缩短",
        "院内市场仍是主要销售渠道",
    ]),
    ("海外市场", [
        "License-out 交易金额屡创新高",
        "欧美监管沟通路径日益成熟",
        "新兴市场本地化合作需求增长",
    ]),
]
MATRIX = {
    "headers": ["治疗领域 × 市场吸引力", "规模", "增速", "竞争强度", "准入门槛"],
    "rows": ["肿瘤", "代谢", "自免", "神经"],
    "cells": {0: {0: 3, 1: 2, 2: 1, 3: 2}, 1: {0: 2, 1: 3, 2: 2, 3: 1},
              2: {0: 2, 1: 1, 2: 3, 3: 2}, 3: {0: 1, 1: 2, 2: 2, 3: 3}},
    "labels": {0: {0: "高", 1: "高", 2: "高", 3: "中"}, 1: {0: "高", 1: "高", 2: "中", 3: "低"},
               2: {0: "中", 1: "低", 2: "高", 3: "中"}, 3: {0: "低", 1: "中", 2: "中", 3: "高"}},
}


def build(style, suffix):
    prs, manifest = new_deck(style)
    build_cover(prs, style=style, title=COVER["title"], subtitle=COVER["subtitle"],
                meta_lines=COVER["meta"], manifest=manifest)
    build_compare(prs, style=style, columns=COMPARE, manifest=manifest)
    build_matrix(prs, style=style, headers=MATRIX["headers"], row_labels=MATRIX["rows"],
                 cells=MATRIX["cells"], cell_labels=MATRIX["labels"], manifest=manifest)
    path = os.path.join(OUT, f"sample_{suffix}.pptx")
    save_deck(prs, path, manifest)
    print("saved:", path)
    return path


if __name__ == "__main__":
    build("A", "styleA")
    build("B", "styleB")
