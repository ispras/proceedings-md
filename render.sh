base=$(dirname "$0")

pandoc \
      "$base/main.md" \
      -o "$base/main.docx" \
      --filter "$base/pandoc-filter.js"