"""
生成应用图标 assets/app.ico（纯标准库实现，不依赖 Pillow）。

图标语义：蓝色圆角底 + 白色文档 + 涂黑的敏感信息条，一眼看懂"文档脱敏"。
ICO 内嵌 PNG（Vista 及以上原生支持），一次写入 16/32/48/64/128/256 六种尺寸。
"""

from __future__ import annotations

import struct
import zlib
from pathlib import Path

W = 256
BG_TOP = (37, 99, 175)
BG_BOTTOM = (23, 62, 122)
PAGE = (255, 255, 255)
INK = (203, 213, 225)
BAR = (17, 24, 39)
ACCENT = (239, 68, 68)


def _blend(dst, src, alpha):
    return tuple(round(d + (s - d) * alpha) for d, s in zip(dst, src))


def _rounded_alpha(x, y, x0, y0, x1, y1, r):
    """返回该像素落在圆角矩形内的覆盖率（简单抗锯齿）。"""
    cx = min(max(x, x0 + r), x1 - r)
    cy = min(max(y, y0 + r), y1 - r)
    dx, dy = x - cx, y - cy
    dist = (dx * dx + dy * dy) ** 0.5
    if x < x0 - 1 or x > x1 + 1 or y < y0 - 1 or y > y1 + 1:
        return 0.0
    edge = r - dist
    if dx == 0 and dy == 0:
        edge = 1.0
    return max(0.0, min(1.0, edge + 0.5))


def render() -> list[list[tuple[int, int, int, int]]]:
    px = [[(0, 0, 0, 0) for _ in range(W)] for _ in range(W)]

    # 1) 蓝色圆角底 + 竖向渐变
    for y in range(W):
        t = y / (W - 1)
        base = tuple(round(a + (b - a) * t) for a, b in zip(BG_TOP, BG_BOTTOM))
        for x in range(W):
            a = _rounded_alpha(x + 0.5, y + 0.5, 6, 6, W - 6, W - 6, 52)
            if a > 0:
                px[y][x] = (*base, round(255 * a))

    # 2) 白色文档
    px0, py0, px1, py1 = 66, 44, 190, 212
    for y in range(py0 - 2, py1 + 2):
        for x in range(px0 - 2, px1 + 2):
            if not (0 <= x < W and 0 <= y < W):
                continue
            a = _rounded_alpha(x + 0.5, y + 0.5, px0, py0, px1, py1, 10)
            if a > 0:
                dst = px[y][x][:3]
                px[y][x] = (*_blend(dst, PAGE, a), 255)

    # 3) 文本行：普通行浅灰，敏感行涂黑
    rows = [
        (66, 26, INK), (92, 26, INK),
        (118, 26, BAR),      # 被脱敏
        (144, 26, INK),
        (170, 26, BAR),      # 被脱敏
    ]
    for top, h, color in rows:
        left = px0 + 16
        right = px1 - (16 if color is INK else 34)
        for y in range(top, top + h):
            for x in range(left, right):
                if 0 <= x < W and 0 <= y < W:
                    a = _rounded_alpha(x + 0.5, y + 0.5, left, top, right, top + h, 5)
                    if a > 0:
                        px[y][x] = (*_blend(px[y][x][:3], color, a), 255)

    # 4) 右下角红点，暗示"已处理/警示"
    ccx, ccy, cr = 196, 196, 30
    for y in range(ccy - cr - 2, ccy + cr + 2):
        for x in range(ccx - cr - 2, ccx + cr + 2):
            if not (0 <= x < W and 0 <= y < W):
                continue
            d = (((x + 0.5 - ccx) ** 2 + (y + 0.5 - ccy) ** 2) ** 0.5)
            a = max(0.0, min(1.0, cr - d + 0.5))
            if a > 0:
                dst = px[y][x][:3] if px[y][x][3] else BG_BOTTOM
                px[y][x] = (*_blend(dst, ACCENT, a), max(px[y][x][3], round(255 * a)))
    return px


def _resize(src, size):
    """box filter 缩放，保证小尺寸不糊。"""
    out = []
    scale = W / size
    for j in range(size):
        row = []
        y0, y1 = int(j * scale), max(int((j + 1) * scale), int(j * scale) + 1)
        for i in range(size):
            x0, x1 = int(i * scale), max(int((i + 1) * scale), int(i * scale) + 1)
            r = g = b = a = n = 0
            for y in range(y0, min(y1, W)):
                for x in range(x0, min(x1, W)):
                    pr, pg, pb, pa = src[y][x]
                    r += pr * pa; g += pg * pa; b += pb * pa; a += pa; n += 1
            if n == 0 or a == 0:
                row.append((0, 0, 0, 0))
            else:
                row.append((r // a, g // a, b // a, a // n))
        out.append(row)
    return out


def _png(pixels, size) -> bytes:
    raw = bytearray()
    for row in pixels:
        raw.append(0)
        for r, g, b, a in row:
            raw += bytes((r, g, b, a))

    def chunk(tag: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(bytes(raw), 9)) + chunk(b"IEND", b""))


def _icns(images: dict[int, bytes]) -> bytes:
    """把尺寸 -> PNG 数据组装成 Apple .icns 文件（纯标准库）。"""
    # (类型码, 逻辑尺寸, 实际像素尺寸)
    entries = [
        (b"icp4", 16, 16),
        (b"icp5", 32, 32),
        (b"icp6", 64, 64),
        (b"ic07", 128, 128),
        (b"ic08", 256, 256),
        (b"ic09", 512, 512),
        (b"ic10", 1024, 1024),
        (b"ic11", 16, 32),   # 16@2x
        (b"ic12", 32, 64),   # 32@2x
        (b"ic13", 128, 256), # 128@2x
        (b"ic14", 256, 512), # 256@2x
    ]
    blobs = b""
    for type_code, _logical, actual in entries:
        data = images.get(actual)
        if not data:
            continue
        size = 8 + len(data)
        blobs += type_code + struct.pack(">I", size) + data

    total_size = 8 + len(blobs)
    return b"icns" + struct.pack(">I", total_size) + blobs


def main() -> None:
    base = render()
    ico_sizes = [16, 32, 48, 64, 128, 256]
    icns_actual_sizes = [16, 32, 64, 128, 256, 512, 1024]

    ico_images = [_png(base if s == W else _resize(base, s), s) for s in ico_sizes]
    icns_images = {s: _png(base if s == W else _resize(base, s), s) for s in icns_actual_sizes}

    assets = Path(__file__).resolve().parent.parent / "assets"
    assets.mkdir(parents=True, exist_ok=True)

    # ---- .ico (Windows) ----
    out_ico = assets / "app.ico"
    header = struct.pack("<HHH", 0, 1, len(ico_images))
    offset = 6 + 16 * len(ico_images)
    entries, blobs = b"", b""
    for size, data in zip(ico_sizes, ico_images):
        entries += struct.pack(
            "<BBBBHHII",
            0 if size >= 256 else size, 0 if size >= 256 else size,
            0, 0, 1, 32, len(data), offset,
        )
        blobs += data
        offset += len(data)
    out_ico.write_bytes(header + entries + blobs)
    print(f"已生成 {out_ico}  ({out_ico.stat().st_size / 1024:.1f} KB, {len(ico_sizes)} 种尺寸)")

    # ---- .icns (macOS) ----
    out_icns = assets / "app.icns"
    out_icns.write_bytes(_icns(icns_images))
    print(f"已生成 {out_icns}  ({out_icns.stat().st_size / 1024:.1f} KB)")


if __name__ == "__main__":
    main()
