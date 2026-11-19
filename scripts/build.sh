#!/usr/bin/env bash
# Membangun ulang gambar, docx, dan pdf Milestone 1-2 ke folder out/.
# Langkah pdf memakai Microsoft Word (Windows) lewat finish.ps1.
set -e
cd "$(dirname "$0")"
mkdir -p ../out
python figs.py ../gambar
python m1.py ../gambar ../out/M1.docx
python m2.py ../gambar ../out/M2.docx
if command -v powershell >/dev/null 2>&1; then
  for m in M1 M2; do
    powershell -NoProfile -ExecutionPolicy Bypass -File finish.ps1 \
      -file "$(cygpath -w "$PWD/../out/$m.docx" 2>/dev/null || echo "$PWD/../out/$m.docx")" \
      -pdf  "$(cygpath -w "$PWD/../out/$m.pdf"  2>/dev/null || echo "$PWD/../out/$m.pdf")"
  done
fi
