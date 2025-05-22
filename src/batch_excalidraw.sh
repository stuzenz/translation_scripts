
for file in *.excalidraw; do
    base="${file%.escalidraw}"
    mv "$file" "$base.json"
    python excalidraw_translate2.py "$file" --target-lang ja --output latest1.json
    mv "$base.json" "$base_ja.excalidraw"

done

# for file in *.json; do
#     base="${file%.json}"
#     # mv "$file" "$base.json"
#     python visio_translator10.py "$file" --target-lang ja --dual-language --output-pdf
    

# done