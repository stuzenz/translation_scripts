#!/usr/bin/env nix-shell
#!nix-shell -i bash -p bash

for file in *.pptx; do
    base="${file%.pptx}"

    python pptx-translator-api.py "${file}" "${base}_zh-Hant_TW.pptx" --source-lang en --target-lang zh-Hant --gemini-model gemini-2.0-flash
    python pptx-translator-api.py "${file}" "${base}_zh-Hans_Ch.pptx" --source-lang en --target-lang zh-Hans --gemini-model gemini-2.0-flash


done