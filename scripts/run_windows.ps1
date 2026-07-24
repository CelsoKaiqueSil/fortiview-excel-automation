$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$python = Join-Path $repository '.venv\Scripts\python.exe'

& $python -m fortiview_excel.pipeline `
  --input (Join-Path $repository 'data\traffic_sample.csv') `
  --output (Join-Path $repository 'output\traffic_weekly.xlsx')
