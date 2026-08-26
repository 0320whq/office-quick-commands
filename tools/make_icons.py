"""Generate simple PNG icons (16/32/80) for the 快捷指令 Office Add-in.

Uses only the Python standard library (zlib) so no third-party packages are needed.
A blue square with a lighter inner square reads clearly at ribbon size.
"""
import os
import zlib
import struct

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "assets")


def png_chunk(ctype, data):
    chunk = ctype + data
    return struct.pack(">I", len(data)) + chunk + struct.pack(">I", zlib.crc32(chunk) & 0xFFFFFFFF)


def make_png(path, size, bg, fg):
    raw = bytearray()
    margin = max(1, size // 8)
    for y in range(size):
        raw.append(0)  # filter type 0 (None)
        for x in range(size):
            if x < margin or x >= size - margin or y < margin or y >= size - margin:
                c = bg
            else:
                c = fg
            raw += bytes((c[0], c[1], c[2], 255))
    compr = zlib.compress(bytes(raw), 9)
    ihdr = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)  # 8-bit, RGBA
    png = (
        b"\x89PNG\r\n\x1a\n"
        + png_chunk(b"IHDR", ihdr)
        + png_chunk(b"IDAT", compr)
        + png_chunk(b"IEND", b"")
    )
    with open(path, "wb") as f:
        f.write(png)
    print("wrote", path, size, "px")


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    bg = (37, 99, 235)    # accent blue
    fg = (147, 197, 253)  # light blue inner
    for size in (16, 32, 80):
        make_png(os.path.join(OUT_DIR, f"icon-{size}.png"), size, bg, fg)


if __name__ == "__main__":
    main()
