#!/usr/bin/env bash
# Иконки и картинки сайта из исходников в src/logo. Нужен ffmpeg и python.
#   bash src/make_icons.sh
# Исходники сгенерированы в Google Flow (Nano Banana 2.1):
#   src/logo/floppy-mascot-2048.jpeg     знак: дискета разводит руками
#   src/logo/estimator-mascot-2048.jpeg  иллюстрация: сметчица с дискетой
set -e
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
F="$ROOT/src/logo/floppy-mascot-2048.jpeg"
E="$ROOT/src/logo/estimator-mascot-2048.jpeg"
OUT="$ROOT/web"
T="$(mktemp -d)"
W='colorlevels=rimax=0.965:gimax=0.965:bimax=0.965'   # почти белый фон -> чисто белый
ff(){ ffmpeg -hide_banner -loglevel error -y "$@"; }
pal(){ # палитра 96 цветов: плоская картинка сжимается в разы
  ff -i "$1" -vf "palettegen=max_colors=96:reserve_transparent=0" "$T/pal.png"
  ff -i "$1" -i "$T/pal.png" -lavfi "paletteuse=dither=none" "$2"
}
mkdir -p "$OUT/img"

ff -i "$F" -vf "crop=1250:1250:398:401,$W,pad=1400:1400:75:75:white" "$T/full.png"          # дискета целиком, с полями
ff -i "$F" -vf "crop=812:862:618:526,$W,pad=980:980:(ow-iw)/2:(oh-ih)/2:white" "$T/body.png"  # только корпус, для 16-48 px

for s in 16 32 48; do ff -i "$T/body.png" -vf "scale=$s:$s:flags=lanczos" "$T/fav-$s.png"; done
cp "$T/fav-32.png" "$OUT/favicon-32x32.png"

ff -i "$T/full.png" -vf "scale=180:180:flags=lanczos" "$T/a.png"; pal "$T/a.png" "$OUT/apple-touch-icon.png"
ff -i "$T/full.png" -vf "scale=192:192:flags=lanczos" "$T/b.png"; pal "$T/b.png" "$OUT/icon-192.png"
ff -i "$T/full.png" -vf "scale=512:512:flags=lanczos" "$T/c.png"; pal "$T/c.png" "$OUT/icon-512.png"
cp "$OUT/icon-512.png" "$OUT/logo-512.png"
ff -i "$T/full.png" -vf "scale=330:330:flags=lanczos,pad=512:512:91:91:white" "$T/m.png"; pal "$T/m.png" "$OUT/icon-maskable-512.png"
ff -i "$F" -vf "crop=1250:1250:398:401,$W,scale=144:144:flags=lanczos" "$T/mark.png"; pal "$T/mark.png" "$OUT/img/logo-mark.png"
ff -i "$E" -vf "crop=1760:1760:143:143,$W,scale=840:840:flags=lanczos" "$T/est.png"; pal "$T/est.png" "$OUT/img/mascot-estimator.png"

# favicon.ico: 16, 32, 48 в одном файле (PNG внутри ICO)
python - "$T/fav-16.png" "$T/fav-32.png" "$T/fav-48.png" "$OUT/favicon.ico" <<'PY'
import struct, sys
*srcs, out = sys.argv[1:]
data = [open(p, "rb").read() for p in srcs]
head = struct.pack("<HHH", 0, 1, len(data))
off = 6 + 16 * len(data)
entries = b""
for s, d in zip([16, 32, 48], data):
    entries += struct.pack("<BBBBHHII", s, s, 0, 0, 1, 32, len(d), off)
    off += len(d)
open(out, "wb").write(head + entries + b"".join(data))
PY
rm -rf "$T"
echo "иконки готовы в web/"
