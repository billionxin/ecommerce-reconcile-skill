param(
    [string]$InputDir,
    [string]$AsOf = (Get-Date -Format 'yyyy-MM-dd'),
    [string]$OutputDir,
    [switch]$Demo,
    [switch]$Test,
    [switch]$OpenResult
)
$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$runtimeRoot = Join-Path ([Environment]::GetFolderPath('UserProfile')) '.cache/codex-runtimes/codex-primary-runtime/dependencies'
$pythonExe = Join-Path $runtimeRoot 'python/python.exe'
if (-not (Test-Path -LiteralPath $pythonExe)) {
    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if (-not $pythonCmd) { throw '找不到 Python。请在 Codex 中加载工作区依赖后运行。' }
    $pythonExe = $pythonCmd.Source
}
$scriptDir = Join-Path $projectRoot 'ecommerce-reconcile/scripts'
if ($Test) {
    & $pythonExe -X utf8 (Join-Path $scriptDir 'test_reconcile.py')
    exit $LASTEXITCODE
}
$moduleLink = Join-Path $projectRoot 'node_modules'
if (-not (Test-Path -LiteralPath (Join-Path $moduleLink '@oai/artifact-tool'))) {
    $bundledModules = Join-Path $runtimeRoot 'node/node_modules'
    if (-not (Test-Path -LiteralPath (Join-Path $bundledModules '@oai/artifact-tool'))) {
        throw '找不到 Codex 的 @oai/artifact-tool 依赖。请先在 Codex 中加载工作区依赖，不要猜测安装来源。'
    }
    if (Test-Path -LiteralPath $moduleLink) { throw 'node_modules 已存在但不包含所需依赖，请检查目录。' }
    New-Item -ItemType Junction -Path $moduleLink -Target $bundledModules | Out-Null
}
if ($Demo) {
    $InputDir = Join-Path $projectRoot '演示/待核对原始数据'
    $AsOf = '2026-09-19'
}
if (-not $InputDir) { throw '请通过 -InputDir 指定四表目录，或使用 -Demo 查看模拟数据。' }
if (-not $OutputDir) {
    $suffix = (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + ([guid]::NewGuid().ToString('N').Substring(0,6))
    $OutputDir = Join-Path $projectRoot ('运行结果/' + $suffix)
}
& $pythonExe -X utf8 (Join-Path $scriptDir 'reconcile.py') --input $InputDir --output $OutputDir --as-of $AsOf
$runCode = $LASTEXITCODE
if ($runCode -eq 0 -or $runCode -eq 2) {
    Write-Host ('已生成：' + (Join-Path $OutputDir '经营利润报表.xlsx'))
    if ($runCode -eq 2) { Write-Host '存在阻断问题；这是核对结果，请查看异常清单。' }
    if ($OpenResult) { Invoke-Item -LiteralPath (Join-Path $OutputDir '经营利润报表.xlsx') }
}
exit $runCode
