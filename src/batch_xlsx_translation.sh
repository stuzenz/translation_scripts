#!/usr/bin/env nix-shell
#!nix-shell -i bash -p bash

for file in *.xlsm; do
    base="${file%.xlsm}"

    python excel_translator_gemini3.py translate "$file" --target-lang ja

done