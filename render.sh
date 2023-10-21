base=$(dirname "$0")

input="$base/sample.md"
output="$base/main.docx"

pandoc \
      "$input" \
      -o "$output.tmp" \
      --filter "$base/pandoc-filter.js"

node main.js "$output.tmp" "$output"