# -*- coding: utf-8 -*-
"""翻译替换脚本：按映射表替换 PPTX 中的文本（含表格单元格），保留全部格式与坐标。
用于草稿模式保险路径——不动版式，只换文字。

用法: python translate_deck.py <输入.pptx> <映射.json> <输出.pptx>

映射.json 格式（精确匹配，逐 run 替换）:
{
  "患者理想的抗VEGF品牌": "Ideal Anti-VEGF Brand for Patients",
  "强调与Lucentis相比的改善": "Improvement over Lucentis emphasized"
}
支持 run 级碎片合并：先把同一段落内所有 run 文本拼接查找，
命中后将译文写回第一个 run、清空其余 run（格式跟随第一个 run）。
"""
import sys, json
from pptx import Presentation


def load_map(path):
    return json.load(open(path, encoding="utf-8"))


def translate_text_frame(tf, tmap):
    hits = []
    for para in tf.paragraphs:
        runs = para.runs
        if not runs:
            continue
        full = "".join(r.text for r in runs)
        if full in tmap:
            runs[0].text = tmap[full]
            for r in runs[1:]:
                r.text = ""
            hits.append(full)
        else:
            # run 级精确替换
            for r in runs:
                if r.text in tmap:
                    r.text = tmap[r.text]
                    hits.append(r.text)
    return hits


def main():
    src, map_path, dst = sys.argv[1], sys.argv[2], sys.argv[3]
    tmap = load_map(map_path)
    prs = Presentation(src)
    all_hits, missed = [], set(tmap.keys())
    for i, slide in enumerate(prs.slides):
        for sh in slide.shapes:
            if sh.has_text_frame:
                for h in translate_text_frame(sh.text_frame, tmap):
                    all_hits.append((i + 1, h))
                    missed.discard(h)
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    for cell in row.cells:
                        for h in translate_text_frame(cell.text_frame, tmap):
                            all_hits.append((i + 1, h))
                            missed.discard(h)
    prs.save(dst)
    print(f"已替换 {len(all_hits)} 处 -> {dst}")
    if missed:
        print("⚠ 映射表中未命中（请核对原文）:")
        for m in missed:
            print("  -", m)


if __name__ == "__main__":
    main()
