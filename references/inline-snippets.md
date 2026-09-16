# 可复制代码块（inline snippets）

> **本文件从 SKILL.md 拆出**（v1.4.0）。SKILL.md 正文只留关键约束与陷阱清单，**真正执行时从这里复制整段代码**。
>
> ⚠️ 执行方式优先级：**① 直接调用 `references/*.ps1`（实测可用）→ ② 若被拦则用本文件的内联版**。
> 2026-09-14 实测修正：此前记录的"`.ps1` 一律被静默拦截"**不成立**——纯 cmdlet 脚本（如 `step1-check-pad.ps1`）可直接 `& '路径' -参数` 调用成功。
> 但**含 COM / Add-Type 的逻辑仍必须内联执行**（`dangerouslyDisableSandbox` + 输出 `Out-File` 落盘后 Read）。

---

## Step 1 环境检查（内联兜底版）

优先直接调用 `references/step1-check-pad.ps1`；仅当文件执行被拦时用本段。

```powershell
$target = '<target_folder>'      # 传了才会检查目标目录的锁文件
$out = "$env:TEMP\step1-report.json"
$r = [ordered]@{}

$padExe = 'C:\Program Files (x86)\Power Automate Desktop\dotnet\PAD.Designer.exe'
$r.padInstalled = Test-Path $padExe
if ($r.padInstalled) { $r.padVersion = (Get-Item $padExe).VersionInfo.FileVersion } else { $r.padVersion = 'N/A' }
$r.agentInstalled = Test-Path 'C:\Program Files (x86)\Power Automate agent for virtual desktops\PAD.RDP.ControlAgent.exe'

$required = 'UIFlowService','PADJavaSyncServiceRDP','UIFlowLogShipper','PADCrashMonitor'
$svc = @(Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -in $required })
$r.servicesDetail = (($svc | ForEach-Object { "$($_.Name)=$($_.Status)" }) -join '; ')
$r.servicesOk = @($svc | Where-Object { $_.Status -ne 'Running' }).Count -eq 0

$r.javaDirExists = Test-Path 'C:\Users\Public\Documents\Microsoft\Power Automate Desktop\PAD_JAVA'
$r.recent1722 = @(Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddDays(-1)} -ErrorAction SilentlyContinue |
                  Where-Object { $_.Message -match 'PAD\.Java\.Sync\.Service\.Host.*PAD_JAVA' }).Count

# === Excel 进程体检（v1.4.0 收紧判据）===
$ep = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
$eu = @($ep | Where-Object { [string]::IsNullOrWhiteSpace($_.MainWindowTitle) })
$lf = @()
if ($target -and (Test-Path $target)) { $lf = @(Get-ChildItem $target -Filter '~$*' -File -ErrorAction SilentlyContinue) }
$r.excelTotal    = $ep.Count
$r.excelUntitled = $eu.Count
$r.lockFiles     = @($lf | ForEach-Object { $_.Name })
# 满足任一才提示用户：①无标题 >=5 ②有锁文件且有无标题 ③无标题 >=3
$r.needPrompt    = ($eu.Count -ge 5) -or (($eu.Count -ge 1) -and ($lf.Count -ge 1)) -or ($eu.Count -ge 3)

if (-not $r.padInstalled)       { $r.status = 'MISSING_PAD' }
elseif (-not $r.agentInstalled) { $r.status = 'MISSING_AGENT' }
elseif (-not $r.servicesOk)     { $r.status = 'SERVICES_DOWN' }
elseif (-not $r.javaDirExists)  { $r.status = 'MISSING_PAD_JAVA' }
elseif ($r.recent1722 -gt 0)    { $r.status = 'BLOCKED_1722' }
else                            { $r.status = 'READY' }

[PSCustomObject]$r | ConvertTo-Json -Depth 3 | Out-File $out -Encoding utf8
```

---

## 2.1 扫描 + 预筛选（用固定脚本，别临时写）

```bash
"C:/Users/Administrator/.workbuddy/binaries/python/versions/3.13.12/python.exe" \
  "C:/Users/Administrator/.workbuddy/skills/Excel PQ 批量刷新/references/pq_scan.py" \
  --target "<target_folder>" \
  --out "%TEMP%/pq-scan.json"
```

- 递归扫描加 `--recursive`；不加则只扫顶层
- 结果 JSON 落盘后 Read 读取，**不要依赖 stdout**（虽然本脚本会打印摘要，但以文件为准）
- 关键字段：`pqFiles`（含 PQ 的源文件数）、`existingXlsm`、`willProcess`、`willSkip`
- **`willSkip == pqFiles` 且用户选了"不覆盖"时，本次会 0 个文件可处理 → 先告知用户，别闷头跑**

> 为什么不用 PowerShell 扫描：需要 `Add-Type -AssemblyName System.IO.Compression.FileSystem`，该调用在 WorkBuddy 沙箱中被硬拦（报 `Add-Type compiles and loads .NET code at runtime`）。

---

## 2.2 设置 Excel 信任中心（⚠️ 必须在启动 Excel 之前完成）

```powershell
$target = '<target_folder>'
$sec = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
if (-not (Test-Path $sec)) { New-Item -Path $sec -Force | Out-Null }
Set-ItemProperty -Path $sec -Name 'AccessVBOM'  -Value 1 -Type DWord   # 信任 VBA 工程对象模型（注入宏前置）
Set-ItemProperty -Path $sec -Name 'VBAWarnings' -Value 1 -Type DWord   # 启用所有宏

# 受信任位置：幂等 —— 同路径复用，不重复追加（否则会不断产生 Location8/Location9…）
$locBase = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations'
if (-not (Test-Path $locBase)) { New-Item -Path $locBase -Force | Out-Null }
$hit = Get-ChildItem $locBase -ErrorAction SilentlyContinue | Where-Object {
    (Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue).Path -eq $target
}
if ($hit) {
    $locPath = $hit[0].PSPath
    $trustedAction = 'REUSED: ' + $hit[0].PSChildName
} else {
    $used = @(Get-ChildItem $locBase -ErrorAction SilentlyContinue |
              ForEach-Object { $_.PSChildName } |
              Where-Object { $_ -match '^Location(\d+)$' } |
              ForEach-Object { [int]($_ -replace '^Location','') })
    $n = 0
    while ($used -contains $n) { $n++ }
    $locPath = "$locBase\Location$n"
    New-Item -Path $locPath -Force | Out-Null
    $trustedAction = "NEW: Location$n"
}
Set-ItemProperty -Path $locPath -Name 'Path' -Value $target
Set-ItemProperty -Path $locPath -Name 'AllowSubFolders' -Value 1
"$trustedAction | trustedTotal=" + @(Get-ChildItem $locBase -ErrorAction SilentlyContinue).Count |
    Out-File "$env:TEMP\step2-registry.txt" -Encoding utf8
```

---

## 2.3 VBA 宏代码（常量）

```vba
Sub RefreshPowerQuery()
    ActiveWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    ActiveWorkbook.Save
End Sub
```

---

## 2.4 转换 + 注入 + 内联验证（Excel COM 单实例循环）

> 本段已把**模块数验证**内联进循环（关表之前顺手读模块名），
> 因此不需要再另开一个 Excel 实例做 2.6 的独立复核。

```powershell
$target     = '<target_folder>'
$reportPath = "$env:TEMP\step2-report.json"
$modName    = 'RefreshModule'          # [由用户回答填充，默认 RefreshModule]
$waitSec    = 15                        # [由用户回答填充，默认 15]
$recursive  = $false                    # [由用户回答填充]
$schedule   = $null                     # [由用户回答填充，无则 null]

$macroCode = @'
Sub RefreshPowerQuery()
    ActiveWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    ActiveWorkbook.Save
End Sub
'@

# ⚠️ 先记录 Excel 进程基线，收尾时只回收本次启动的实例（必须在 New-Object 之前）
$baseline = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

$success = @()
$failed  = @()
$files = @(Get-ChildItem -Path $target -Filter '*.xlsx' -File -Recurse:$recursive |
           Sort-Object Name | Where-Object { $_.Name -notlike '~$*' })

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    foreach ($xlsx in $files) {
        $xlsmPath = $xlsx.FullName -replace '\.xlsx$', '.xlsm'
        $wb = $null
        try {
            $wb = $excel.Workbooks.Open($xlsx.FullName)
            $wb.SaveAs($xlsmPath, 52)          # 52 = xlOpenXMLWorkbookMacroEnabled

            # 删除同名模块（保幂等）
            $names = @($wb.VBProject.VBComponents | ForEach-Object { $_.Name })
            if ($names -contains $modName) {
                $wb.VBProject.VBComponents.Remove($wb.VBProject.VBComponents.Item($modName))
            }

            # 注入新模块
            $module = $wb.VBProject.VBComponents.Add(1)   # 1 = vbext_ct_StdModule
            $module.Name = $modName
            $module.CodeModule.AddFromString($macroCode)
            $wb.Save()

            # === 内联验证（关表之前读模块名，省一次 Excel 启动）===
            $names2 = @($wb.VBProject.VBComponents | ForEach-Object { $_.Name })
            $dup = @($names2 | Where-Object { $_ -eq $modName }).Count
            if ($dup -eq 1) { $verify = 'OK' }
            elseif ($dup -eq 0) { $verify = 'NO_MODULE' }
            else { $verify = "DUP_MODULE($dup)" }

            $wb.Close()
            $wb = $null

            $success += [PSCustomObject]@{
                Source  = $xlsx.FullName
                Output  = $xlsmPath
                Modules = ($names2 -join ', ')
                Verify  = $verify
            }
        } catch {
            $failed += [PSCustomObject]@{ Source = $xlsx.FullName; Error = $_.Exception.Message }
            if ($wb) { try { $wb.Close($false) } catch { } }
        }
    }
} catch {
    $failed += [PSCustomObject]@{ Source = 'COM_INIT'; Error = $_.Exception.Message }
} finally {
    if ($excel) {
        try { $excel.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# ⚠️ Quit() 不保证 EXCEL.EXE 退出（实测连跑 2 次残留 3 个）→ 必须显式回收本次启动的
Start-Sleep -Seconds 3
$reclaimed = @()
Get-Process EXCEL -ErrorAction SilentlyContinue |
    Where-Object { $baseline -notcontains $_.Id } |
    ForEach-Object { $reclaimed += $_.Id; try { Stop-Process -Id $_.Id -Force } catch { } }

# === 报告落盘（不要写进 skill 目录）===
$report = [PSCustomObject]@{
    timestamp     = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    targetFolder  = $target
    excelBaseline = $baseline
    totalScanned  = $files.Count
    reclaimedPids = $reclaimed
    success       = $success
    failed        = $failed
}
$report | ConvertTo-Json -Depth 4 | Out-File $reportPath -Encoding utf8

# === 进度续接存档（OPT-014：中断后能接着来）===
$progressPath = "$env:TEMP\excel-pq-progress.json"
[PSCustomObject]@{
    targetFolder = $target
    stepsDone    = @('step1','step2')
    step2Report  = $reportPath
    macroModule  = $modName
    macroName    = 'RefreshPowerQuery'
    waitSeconds  = $waitSec
    recursive    = $recursive
    schedule     = $schedule
    updatedAt    = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
} | ConvertTo-Json -Depth 3 | Out-File $progressPath -Encoding utf8

"step2 done -> $reportPath | success=$($success.Count) failed=$($failed.Count) reclaimed=$($reclaimed -join ',')"
```

---

## 2.5 产物验证（轻量，可选）

```powershell
foreach ($item in $success) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($item.Output)
    $hasVba = $zip.GetEntry('xl/vbaProject.bin') -ne $null
    $zip.Dispose()
    if (-not $hasVba) { $item | Add-Member -NotePropertyName 'verify2' -NotePropertyValue 'NO_VBA' }
}
```

---

## 2.6 独立复核（**降级为备用手段**）

`xl/vbaProject.bin` 存在只能证明"工程存在"，不能证明模块名/宏名对。
正常情况下 **2.4 的内联验证已经给出结论**，不需要跑本段。
**只有**在怀疑 2.4 结果时（例如 `Verify` 异常、或事后单独查一个 .xlsm）才用：

```powershell
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
Get-ChildItem '<target_folder>' -Filter '*.xlsm' -File | ForEach-Object {
    $wb = $xl.Workbooks.Open($_.FullName)
    $comps = @($wb.VBProject.VBComponents | ForEach-Object { $_.Name })
    Write-Output ("{0}: modules=[{1}]" -f $_.Name, ($comps -join ', '))
    $wb.Close($false)
}
$xl.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
```

期望：`RefreshModule` **恰好出现一次**（重复运行后仍是一次，才说明幂等生效）。

---

## Step 2 输出格式

报告默认写 `%TEMP%\step2-report.json`：

```json
{
  "timestamp": "2026-09-14 15:17:13",
  "targetFolder": "C:\\...\\销售差异模块",
  "excelBaseline": [41564],
  "totalScanned": 2,
  "reclaimedPids": [21004],
  "success": [{"Source": "C:...", "Output": "C:...", "Modules": "ThisWorkbook, Sheet1, RefreshModule", "Verify": "OK"}],
  "failed": []
}
```

`Verify` 取值：`OK` / `NO_MODULE`（模块没注进去）/ `DUP_MODULE(n)`（RefreshModule 重复 n 个，幂等失效）。
