#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUTPUT_FILE="${1:-$ROOT_DIR/outlook-mcp.mcpb}"

if [[ ! -f "$ROOT_DIR/manifest.json" ]]; then
  echo "manifest.json not found in project root: $ROOT_DIR" >&2
  exit 1
fi

if ! command -v mcpb >/dev/null 2>&1; then
  echo "mcpb CLI not found in PATH. Install it and try again." >&2
  exit 1
fi

node -e "
const fs = require('fs');
const path = require('path');
const manifestPath = path.join(process.argv[1], 'manifest.json');
const m = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
const secret = m?.server?.mcp_config?.env?.MS_CLIENT_SECRET || '';
if (secret && !/your-|\\$\\{|<|changeme/i.test(secret)) {
  console.error('WARNING: manifest.json contains a non-placeholder MS_CLIENT_SECRET value.');
  console.error('         Consider replacing hardcoded secrets before distributing the MCPB.');
}
" "$ROOT_DIR"

echo "Validating manifest..."
mcpb validate "$ROOT_DIR/manifest.json"

echo "Packing MCPB: $OUTPUT_FILE"
rm -f "$OUTPUT_FILE"
mcpb pack "$ROOT_DIR" "$OUTPUT_FILE"

echo "MCPB bundle created at: $OUTPUT_FILE"
