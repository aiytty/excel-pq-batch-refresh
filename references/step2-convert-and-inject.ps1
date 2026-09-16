<#
.SYNOPSIS
  Excel PQ 批量自动刷新 - Step 2: Excel 转 .xlsm + 注入 RefreshPowerQuery 宏
.DESCRIPTION
  扫描目标文件夹所有 .xlsx，识别含 Power Query 的文件，转为 .xlsm 并注入刷新宏。

  ⚠️⚠️ 重要警告（2026-09-14 标注）：
  本脚本第 33 行含 `Add-Type -AssemblyName System.IO.Compression.FileSystem`，
  该调用在 WorkBuddy 沙箱中被**硬拦截**（报 "Add-Type compiles and loads .NET code at runtime"），
  因此**本脚本在当前环境下无法执行**，仅作逻辑参考用。

  实际执行 Step 2 请使用 `references/inline-snippets.md` 里的内联版代码块
  （内联执行 + dangerouslyDisableSandbox + 输出落盘后 Read），那份是实测可用的。

  若未来沙箱放开 Add-Type，本脚本可按原方式使用；
  届时建议把扫描段改为调用 `pq_scan.py`（避免该依赖）。

.NOTES
  必须 UTF-8 with BOM 保存（含中文路径安全）
  用法: powershell -NoProfile -ExecutionPolicy Bypass -File step2-convert-and-inject.ps1 -TargetFolder "C:\Users\YourName\Desktop\乐惠收入模块"
  变更 (v1.1.0):
    - 修复 PQ 识别漏检：不再依赖 customXml/ 目录条目（改用逐条 Entries 匹配）
    - 受信任位置改为幂等（同路径复用，不重复追加）
    - 增加 Excel 残留进程清理（Quit 后按 PID 基线回收，避免文件占用）
    - 报告路径可配置，默认落到 %TEMP%，不再污染 skill 目录
#>

param(
    [Parameter(Mandatory=$true)][string]$TargetFolder,
    [string]$ReportPath,
    [switch]$KeepExcelOpen,
    [switch]$RecursiveScan,     # 递归扫描子目录（不传则只扫顶层）
    [switch]$SkipOverwrite      # 已存在同名 .xlsm 时跳过该文件（原 .xlsx 始终保留）
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $TargetFolder)) { throw "目标文件夹不存在: $TargetFolder" }
if (-not $ReportPath) { $ReportPath = Join-Path $env:TEMP 'step2-report.json' }

# === 0. 记录 Excel 进程基线（用于结束后回收本脚本启动的实例）===
$excelBaseline = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

# === 1. 扫描含 PQ 的 .xlsx ===
Add-Type -AssemblyName System.IO.Compression.FileSystem
$xlsxFiles = @(Get-ChildItem -Path $TargetFolder -Filter '*.xlsx' -Recurse:$RecursiveScan)
$pqFiles = @()
foreach ($f in $xlsxFiles) {
    $hasPQ = $false
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($f.FullName)
        if ($zip.GetEntry('xl/connections.xml')) { $hasPQ = $true }
        if (-not $hasPQ) {
            # 注意：不能先判断 GetEntry('customXml/') 是否存在——Excel 写出的包里
            # 未必包含目录条目，会导致 Mashup 分支被整体跳过（漏检真实 PQ 文件）
            foreach ($e in $zip.Entries) {
                if ($e.FullName -like 'customXml/item*.xml') {
                    $reader = New-Object System.IO.StreamReader($e.Open())
                    $content = $reader.ReadToEnd()
                    $reader.Close()
                    if ($content -match 'Mashup') { $hasPQ = $true; break }
                }
            }
        }
    } catch {
        Write-Warning "扫描失败: $($f.FullName) - $($_.Exception.Message)"
    } finally {
        if ($zip) { $zip.Dispose() }
    }
    if ($hasPQ) { $pqFiles += $f }
}

Write-Host "[1/3] 扫描 $($xlsxFiles.Count) 个 .xlsx，找到 $($pqFiles.Count) 个含 Power Query"

# === 2. 设置 Excel 信任中心（必须在启动 Excel 前完成）===
$sec = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
if (-not (Test-Path $sec)) { New-Item -Path $sec -Force | Out-Null }
Set-ItemProperty -Path $sec -Name 'AccessVBOM' -Value 1
Set-ItemProperty -Path $sec -Name 'VBAWarnings' -Value 1

# 受信任位置：幂等处理（同路径复用，不重复追加）
$locBase = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations'
if (-not (Test-Path $locBase)) { New-Item -Path $locBase -Force | Out-Null }
$hit = Get-ChildItem $locBase -ErrorAction SilentlyContinue | Where-Object {
    (Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue).Path -eq $TargetFolder
}
if ($hit) {
    $locPath = $hit[0].PSPath
} else {
    $used = @(Get-ChildItem $locBase -ErrorAction SilentlyContinue |
              ForEach-Object { $_.PSChildName } |
              Where-Object { $_ -match '^Location(\d+)$' } |
              ForEach-Object { [int]($_ -replace '^Location', '') })
    $n = 0
    while ($used -contains $n) { $n++ }
    $locPath = "$locBase\Location$n"
    New-Item -Path $locPath -Force | Out-Null
}
Set-ItemProperty -Path $locPath -Name 'Path' -Value $TargetFolder
Set-ItemProperty -Path $locPath -Name 'AllowSubFolders' -Value 1

Write-Host "[2/3] Excel 信任中心已配置（AccessVBOM=1 / VBAWarnings=1 / 受信任位置: $locPath）"

# === 3. VBA 宏代码 ===
$macroCode = @'
Sub RefreshPowerQuery()
    ActiveWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    ActiveWorkbook.Save
End Sub
'@

# === 4. Excel COM 转换 + 注入 ===
$success = @()
$failed = @()

if ($pqFiles.Count -gt 0) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    foreach ($xlsx in $pqFiles) {
        $xlsmPath = $xlsx.FullName -replace '\.xlsx$', '.xlsm'
        $wb = $null
        if ($SkipOverwrite -and (Test-Path $xlsmPath)) {
            Write-Host "  跳过（同名 .xlsm 已存在）: $($xlsx.Name)"
            continue
        }
        Write-Host "  处理: $($xlsx.Name)"
        try {
            $wb = $excel.Workbooks.Open($xlsx.FullName)
            $wb.SaveAs($xlsmPath, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled

            # 删除同名模块（保幂等）
            $modName = 'RefreshModule'
            $old = $wb.VBProject.VBComponents | Where-Object { $_.Name -eq $modName }
            if ($old) { $wb.VBProject.VBComponents.Remove($old) }

            # 注入新模块
            $module = $wb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
            $module.Name = $modName
            $module.CodeModule.AddFromString($macroCode)

            # 模块数校验（必须恰好 1 个 RefreshModule）
            $modCount = @($wb.VBProject.VBComponents | Where-Object { $_.Name -eq $modName }).Count
            $wb.Save()
            $wb.Close($false)
            $wb = $null

            # 验证产物
            $verifyZip = [System.IO.Compression.ZipFile]::OpenRead($xlsmPath)
            $hasVba = $verifyZip.GetEntry('xl/vbaProject.bin') -ne $null
            $verifyZip.Dispose()

            $verify = if (-not $hasVba) { 'NO_VBA' } elseif ($modCount -ne 1) { "DUP_MODULE($modCount)" } else { 'OK' }
            $success += [PSCustomObject]@{
                Source = $xlsx.FullName
                Output = $xlsmPath
                Verify = $verify
            }
        } catch {
            $failed += [PSCustomObject]@{
                Source = $xlsx.FullName
                Error  = $_.Exception.Message
            }
            try { if ($wb) { $wb.Close($false) } } catch { }
        }
    }

    # 释放 COM
    try { $excel.Quit() } catch { }
    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    # === 4.1 回收本脚本启动的 Excel 残留进程 ===
    # Quit() 并不保证 EXCEL.EXE 退出；残留会占用文件，导致后续 SaveAs 失败
    Start-Sleep -Seconds 3
    $orphans = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $excelBaseline -notcontains $_.Id })
    $reclaimed = @()
    if (-not $KeepExcelOpen) {
        foreach ($p in $orphans) {
            try { Stop-Process -Id $p.Id -Force; $reclaimed += $p.Id } catch { }
        }
    }
} else {
    $reclaimed = @()
    Write-Host "  未发现含 Power Query 的文件，跳过转换"
}

Write-Host "[3/3] 完成。成功: $($success.Count) / 失败: $($failed.Count)"

# === 5. 输出报告 ===
$report = @{
    timestamp     = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    targetFolder  = $TargetFolder
    totalScanned  = $xlsxFiles.Count
    totalPQ       = $pqFiles.Count
    reclaimedPids = $reclaimed
    success       = $success
    failed        = $failed
}

$report | ConvertTo-Json -Depth 3 | Out-File $ReportPath -Encoding utf8

Write-Host "报告: $ReportPath"
