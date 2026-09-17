# -*- coding: utf-8 -*-
"""从 OfficeCLI dump JSON 提取指定页面的元素规格（坐标/字号/色值）。
用法: python extract_slide_spec.py <dump.json> <页码,逗号分隔>
输出: 每页的元素规格表（cm / pt / hex色值）
"""
import sys, json, re
from collections import defaultdict

EMU_PER_CM = 360000
PT_PER_CM = 28.3465


def to_cm(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return round(v / EMU_PER_CM, 2)
    s = str(v)
    m = re.match(r"([\d.]+)(emu|pt|cm)", s)
    if not m:
        return None
    num, unit = float(m.group(1)), m.group(2)
    if unit == "emu":
        return round(num / EMU_PER_CM, 2)
    if unit == "pt":
        return round(num / PT_PER_CM, 2)
    return round(num, 2)


def norm_color(c):
    if not c:
        return None
    c = str(c)
    m = re.search(r"([0-9A-Fa-f]{6})", c)
    return "#" + m.group(1).upper() if m else c


def main():
    dump_path = sys.argv[1]
    pages = set(int(p) for p in sys.argv[2].split(","))
    data = json.load(open(dump_path, encoding="utf-8"))

    shapes = defaultdict(dict)   # (slide, shape_id) -> props
    order = []

    def slide_of(path):
        m = re.match(r"/slide\[(\d+)\]", path or "")
        return int(m.group(1)) if m else None

    for op in data:
        cmd = op.get("command")
        if cmd == "add":
            parent = slide_of(op.get("parent", ""))
            if parent in pages and op.get("type") not in ("run", "paragraph"):
                sid = op["props"].get("id") or op["props"].get("name") or f"n{len(order)}"
                key = (parent, str(sid))
                shapes[key].update(op.get("props", {}))
                if key not in order:
                    order.append(key)
        elif cmd == "set":
            p = slide_of(op.get("path", ""))
            if p in pages:
                m = re.search(r"shape\[@id=(\d+)\]", op.get("path", ""))
                sid = m.group(1) if m else None
                if sid:
                    key = (p, sid)
                    shapes[key].update(op.get("props", {}))
                    if key not in order:
                        order.append(key)

    for (slide, sid) in order:
        pr = shapes[(slide, sid)]
        name = pr.get("name", "?")
        geo = {k: to_cm(pr.get(k)) for k in ("x", "y", "width", "height")}
        fs = pr.get("fontSize") or pr.get("font-size") or pr.get("size")
        fill = norm_color(pr.get("fill"))
        fontcolor = norm_color(pr.get("color") or pr.get("fontColor"))
        txt = str(pr.get("text", ""))[:36].replace("\n", " ")
        bits = [f'id={sid}', f'"{name}"',
                f"x={geo['x']} y={geo['y']} w={geo['width']} h={geo['height']}"]
        if fs: bits.append(f"字号={fs}")
        if fill: bits.append(f"填充={fill}")
        if fontcolor: bits.append(f"字色={fontcolor}")
        if pr.get("bold") == "true": bits.append("粗体")
        if txt: bits.append(f"「{txt}」")
        print(f"[p{slide}] " + " ".join(str(b) for b in bits))


if __name__ == "__main__":
    main()
