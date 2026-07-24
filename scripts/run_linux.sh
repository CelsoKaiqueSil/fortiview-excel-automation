#!/usr/bin/env bash
set -euo pipefail

repository="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
"${repository}/.venv/bin/python" -m fortiview_excel.pipeline \
  --input "${repository}/data/traffic_sample.csv" \
  --output "${repository}/output/traffic_weekly.xlsx"
