base=$(dirname "$0")

pandoc \
      "$base/sample.md" \
      -o "$base/main.docx" \
      --filter "$base/pandoc-filter.js"