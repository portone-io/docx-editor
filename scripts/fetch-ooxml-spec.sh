#!/usr/bin/env bash
# Downloads the ECMA-376 5th edition specification PDFs into spec/pdf/ (gitignored).
# The schemas are already committed under spec/schemas/; only the prose is fetched.
set -euo pipefail

package_dir="$(cd "$(dirname "$0")/.." && pwd)"
pdf_dir="$package_dir/spec/pdf"
mkdir -p "$pdf_dir"

fetch() {
  local url="$1" sha="$2" zip pdf
  zip="$pdf_dir/$(basename "$url")"
  if [ ! -f "$zip" ]; then
    echo "downloading $(basename "$url")"
    curl -fsSL -o "$zip" "$url"
  fi
  echo "$sha  $zip" | shasum -a 256 -c - >/dev/null
  (cd "$pdf_dir" && unzip -o -q "$zip" "*.pdf")
}

fetch "https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip" \
  "9d0bcad9cf06054785b03762fcfadbf6bab7e54a5f9d69434e34b7fd464d4129"
fetch "https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip" \
  "bd25da1109f73762356596918bf5ff8b74a1331642dba5f1c1d1dfc6bed34ecd"

echo "done:"
ls "$pdf_dir"/*.pdf
