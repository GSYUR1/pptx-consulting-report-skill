# -*- coding: utf-8 -*-
"""质检脚本：渲染 PPTX 全部页面为截图，并调用 OfficeCLI issues 检测排版问题。
用法: python qa_render.py <成品.pptx> [截图输出目录]
流程（对应 SKILL.md 第 4 步）：
  1. officecli view <file> screenshot --page 1-N  （整册渲染为长图）
  2. 拆分长图为单页 PNG
  3. officecli view <file> issues --json          （自动检测溢出/越界）
需要 officecli 在 PATH 中。
"""
import sys, os, subprocess, glob
from PIL import Image

import shutil

# Windows 下 PATH 中的 officecli 实际是 officecli.cmd，subprocess 需要绝对路径
_cmd = shutil.which("officecli") or shutil.which("officecli.cmd")
OFFICECLI = _cmd if _cmd else "officecli"


def run(cmd):
    print("+", " ".join(cmd))
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.stdout.strip():
        print(r.stdout.strip()[-800:])
    if r.returncode != 0 and r.stderr.strip():
        print("STDERR:", r.stderr.strip()[-500:])
    return r


def main():
    src = sys.argv[1]
    outdir = sys.argv[2] if len(sys.argv) > 2 else "qa_pages"
    os.makedirs(outdir, exist_ok=True)

    # 1. 页数
    r = run([OFFICECLI, "view", src, "outline"])
    import re
    m = re.search(r"(\d+) slides", r.stdout)
    n = int(m.group(1)) if m else 1
    print(f"总页数: {n}")

    # 2. 整册渲染
    tall = os.path.join(outdir, "_tall.png")
    run([OFFICECLI, "view", src, "screenshot", "-o", tall, "--page", f"1-{n}"])

    # 3. 拆分
    im = Image.open(tall)
    w, h = im.size
    ph = h // n
    for i in range(n):
        im.crop((0, i * ph, w, (i + 1) * ph)).save(os.path.join(outdir, f"p{i+1:02d}.png"))
    os.remove(tall)
    print(f"截图已拆分至 {outdir}/p01..p{n:02d}.png")

    # 4. issues 检测
    r = run([OFFICECLI, "view", src, "issues", "--json"])
    if r.stdout.strip():
        print("=== issues 检测结果 ===")
        print(r.stdout.strip()[-1500:])
    else:
        print("issues 检测无输出（可能无问题或命令不支持）")


if __name__ == "__main__":
    main()
