param([string]$Config = "Debug", [switch]$SingleFile)
$ErrorActionPreference = "Stop"
$root  = Split-Path -Parent $PSScriptRoot
$bin   = Join-Path $root "CubeConnector\bin\$Config"
$stage = Join-Path $root "release\stage"
$out   = Join-Path $root "release"
Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $stage | Out-Null

Copy-Item (Join-Path $bin "CubeConnector-AddIn64-packed.xll") (Join-Path $stage "CubeConnector.xll")
if (-not $SingleFile) {
    Copy-Item (Join-Path $bin "WebView2Loader.dll") (Join-Path $stage "WebView2Loader.dll")
}

$zip = Join-Path $out "CubeConnector.zip"
Remove-Item $zip -Force -ErrorAction SilentlyContinue
Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $zip
Write-Host "Built $zip"

# Single-file mode: also emit the bare, directly-loadable .xll (no extraction needed).
$bareXll = Join-Path $out "CubeConnector.xll"
Remove-Item $bareXll -Force -ErrorAction SilentlyContinue
if ($SingleFile) {
    Copy-Item (Join-Path $stage "CubeConnector.xll") $bareXll
    Write-Host "Built $bareXll (single-file)"
}

Get-ChildItem $stage | Select-Object Name, Length | Format-Table | Out-String | Write-Host
