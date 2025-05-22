#!/usr/bin/env nix-shell
#!nix-shell -i bash -p bash

for file in *.vsdx; do
    base="${file%.vsdx}"

    python visio_translator10.py "$file" --target-lang ja --dual-language
    python visio_translator10.py "$file" --target-lang ja
    
done
