<#
.SYNOPSIS
  Excel PQ 批量自动刷新 - Step 1: Power Automate 环境检查
.DESCRIPTION
  检测 PAD 是否就绪。返回 JSON 报告对象，含 5 项检测结果 + 综合状态。
.NOTES
  必须 UTF-8 with BOM 保存（含中文路径安全）
  用法: powershell -NoProfile -ExecutionPolicy Bypass -File step1-check-pad.ps1 [-OutFile <path>] [-TargetFolder <path>]
  变更 (v1.1.0):
    - 增加 -OutFile 参数（默认落 %TEMP%），兼容"无 stdout 回显"的执行环境
    - 增加 executionPolicy / psVersion 诊断字段
    - status != READY 时自动填充 notes 处理建议
  变更 (v1.4.0):
    - 新增 Excel 进程体检：excelTotal / excelUntitled / lockFiles / needPrompt
    - 判据由"只要有 1 个无标题进程就打断用户"改为组合条件（过宽会打扰用户）：
        ① 无标题进程 >= 5
        ② 存在无标题进程 且 目标目录里有 ~$ 锁文件（这才真会踩坑）
        ③ 无标题进程 >= 3
      三条满足任一才置 needPrompt = true。传 -TargetFolder 才会检查锁文件。

#>

param(
    [string]$OutFile,
    [string]$TargetFolder
)

$ErrorActionPreference = 'Stop'
if (-not $OutFile) { $OutFile = Join-Path $env:TEMP 'step1-report.json' }

$report = [PSCustomObject]@{
    padInstalled    = $false
    padVersion      = 'N/A'
    agentInstalled  = $false
    servicesOk      = $false
    javaDirExists   = $false
    recent1722      = 0
    executionPolicy = 'N/A'
    psVersion       = 'N/A'
    status          = 'UNKNOWN'
    notes           = ''
}

# === 主程序检测 ===
$padExe = 'C:\Program Files (x86)\Power Automate Desktop\dotnet\PAD.Designer.exe'
if (Test-Path $padExe) {
    $report.padInstalled = $true
    $report.padVersion = (Get-Item $padExe).VersionInfo.FileVersion
}

# === agent 检测 ===
$report.agentInstalled = Test-Path 'C:\Program Files (x86)\Power Automate agent for virtual desktops\PAD.RDP.ControlAgent.exe'

# === 关键服务检测 ===
$required = 'UIFlowService','PADJavaSyncServiceRDP','UIFlowLogShipper','PADCrashMonitor'
$svcState = Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -in $required } | ForEach-Object { "$($_.Name)=$($_.Status)" }
$report.servicesOk = @($svcState | Where-Object { $_ -notmatch 'Running' }).Count -eq 0
$report | Add-Member -NotePropertyName servicesDetail -NotePropertyValue ($svcState -join '; ')

# === PAD_JAVA 目录检测 ===
$report.javaDirExists = Test-Path 'C:\Users\Public\Documents\Microsoft\Power Automate Desktop\PAD_JAVA'

# === 1722 错误检测 ===
$1722 = Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddDays(-1)} -ErrorAction SilentlyContinue | Where-Object { $_.Message -match 'PAD\.Java\.Sync\.Service\.Host.*PAD_JAVA' }
$report.recent1722 = @($1722).Count

# === 运行环境诊断 ===
$report.psVersion = $PSVersionTable.PSVersion.ToString()
try { $report.executionPolicy = (Get-ExecutionPolicy).ToString() } catch { }

# === Excel 进程体检 (v1.4.0) ===
# Quit() 后残留的无窗口 EXCEL.EXE 会被误认为"用户正开着的表"，
# 但只有 1~2 个时完全无害 —— 判据必须收紧，否则每次都要打扰用户。
$excelProcs = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
$excelUntitled = @($excelProcs | Where-Object { [string]::IsNullOrWhiteSpace($_.MainWindowTitle) })
$lockFiles = @()
if ($TargetFolder -and (Test-Path $TargetFolder)) {
    $lockFiles = @(Get-ChildItem -Path $TargetFolder -Filter '~$*' -File -ErrorAction SilentlyContinue)
}
$needPrompt = $false
$promptReason = ''
if ($excelUntitled.Count -ge 5) {
    $needPrompt = $true; $promptReason = "无窗口 Excel 进程达 $($excelUntitled.Count) 个，建议先清理"
} elseif ($excelUntitled.Count -ge 1 -and $lockFiles.Count -ge 1) {
    $needPrompt = $true; $promptReason = "目标目录存在 $($lockFiles.Count) 个锁文件（~$）且有无窗口 Excel 进程，文件可能被占用"
} elseif ($excelUntitled.Count -ge 3) {
    $needPrompt = $true; $promptReason = "无窗口 Excel 进程 $($excelUntitled.Count) 个，偏多"
}
$report | Add-Member -NotePropertyName excelTotal    -NotePropertyValue $excelProcs.Count
$report | Add-Member -NotePropertyName excelUntitled -NotePropertyValue $excelUntitled.Count
$report | Add-Member -NotePropertyName lockFiles     -NotePropertyValue @($lockFiles | ForEach-Object { $_.Name })
$report | Add-Member -NotePropertyName needPrompt    -NotePropertyValue $needPrompt
$report | Add-Member -NotePropertyName promptReason  -NotePropertyValue $promptReason

# === 综合判定 ===
if (-not $report.padInstalled)       { $report.status = 'MISSING_PAD';       $report.notes = '执行 winget install --id Microsoft.PowerAutomateDesktop --exact' }
elseif (-not $report.agentInstalled) { $report.status = 'MISSING_AGENT';     $report.notes = '运行 Setup.Microsoft.PowerAutomateAgent.exe -Install -ACCEPTEULA -Silent' }
elseif (-not $report.servicesOk)     { $report.status = 'SERVICES_DOWN';     $report.notes = '检查已停止服务的 ImagePath 是否存在，再 Start-Service 试启' }
elseif (-not $report.javaDirExists)  { $report.status = 'MISSING_PAD_JAVA';  $report.notes = '新建 PAD_JAVA 目录后 Restart-Service PADJavaSyncServiceRDP' }
elseif ($report.recent1722 -gt 0)    { $report.status = 'BLOCKED_1722';      $report.notes = '走 1722 排障剧本（卸载重装）' }
else                                  { $report.status = 'READY';            $report.notes = '环境就绪，可进入 Step 2' }

$json = $report | ConvertTo-Json -Depth 3
$json | Out-File $OutFile -Encoding utf8
Write-Output $json
Write-Host "报告已写入: $OutFile"
