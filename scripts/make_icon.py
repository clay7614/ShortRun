"""
PNG から ICO を生成する簡易スクリプト。

前提:
	- Pillow がインストールされていること (pip install Pillow)

使い方 (PowerShell):
	- 透過を維持:   python scripts/make_icon.py .\src.png .\assets\icon.ico
	- 非透過で白:   python scripts/make_icon.py .\src.png .\assets\icon.ico --no-alpha
	- 非透過で色:   python scripts/make_icon.py .\src.png .\assets\icon.ico --bg FFFFFF

オプション:
	--no-alpha           透過を無効にして白背景で塗りつぶす
	--bg RRGGBB|#RGB     指定色で背景を塗りつぶして透過を無効化（#付き/無し、3桁/6桁HEX対応）
"""

from __future__ import annotations

import sys
from pathlib import Path

try:
	from PIL import Image
except Exception as e:
	print("Pillow が必要です。pip install Pillow でインストールしてください。", file=sys.stderr)
	raise


def _parse_hex_color(s: str) -> tuple[int, int, int]:
	s = s.strip()
	if s.startswith("#"):
		s = s[1:]
	if len(s) == 3:
		s = "".join(ch * 2 for ch in s)
	if len(s) != 6:
		raise ValueError("color must be RGB hex like FFFFFF or #FFF")
	r = int(s[0:2], 16)
	g = int(s[2:4], 16)
	b = int(s[4:6], 16)
	return (r, g, b)


def make_ico(src_png: Path, dst_ico: Path, *, bg_color: tuple[int, int, int] | None = None) -> None:
	base_img = Image.open(src_png).convert("RGBA")
	# 透過を無効化する場合は背景に合成して RGB 化
	work = base_img
	if bg_color is not None:
		rgb_bg = Image.new("RGB", base_img.size, bg_color)
		# alpha をマスクにして貼り付け
		rgb_bg.paste(base_img, mask=base_img.split()[-1])
		work = rgb_bg

	sizes = [256]
	icon_base = work.resize((sizes[0], sizes[0]), Image.LANCZOS)
	dst_ico.parent.mkdir(parents=True, exist_ok=True)
	# base から他サイズも生成して埋め込む
	icon_base.save(dst_ico, format="ICO", sizes=[(s, s) for s in sizes])
	print(f"ICO generated: {dst_ico} (alpha={'off' if bg_color is not None else 'on'})")


def main(argv: list[str]) -> int:
	if len(argv) < 3:
		print("Usage: python scripts/make_icon.py <src_png> <dst_ico> [--no-alpha | --bg <hex>]", file=sys.stderr)
		return 2
	src = Path(argv[1])
	dst = Path(argv[2])
	if not src.is_file():
		print(f"Not found: {src}", file=sys.stderr)
		return 1

	bg_color: tuple[int, int, int] | None = None
	i = 3
	while i < len(argv):
		arg = argv[i]
		if arg == "--no-alpha":
			bg_color = (255, 255, 255)
			i += 1
			continue
		if arg == "--bg":
			if i + 1 >= len(argv):
				print("--bg option requires a color value (e.g., FFFFFF)", file=sys.stderr)
				return 2
			try:
				bg_color = _parse_hex_color(argv[i + 1])
			except Exception as e:
				print(f"Invalid color: {argv[i + 1]} ({e})", file=sys.stderr)
				return 2
			i += 2
			continue
		else:
			print(f"Unknown option: {arg}", file=sys.stderr)
			return 2

	make_ico(src, dst, bg_color=bg_color)
	return 0


if __name__ == "__main__":
	raise SystemExit(main(sys.argv))

