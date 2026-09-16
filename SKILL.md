---
name: excel-pq-batch-refresh
description: 批量刷新指定文件夹内所有 Excel Power Query 的全流程工作流（环境检查 + .xlsm 化 + VBA 宏注入 + PAD 流程搭建指引 + 定时调度）。触发场景：用户要求批量刷新 Power Query、定时刷新 Excel、PQ 自动刷新、用 Power Automate 刷新 Excel、批量刷新 Excel 数据、自动刷新 .xlsx 数据源、让表格自己定时更新。
version: 1.4.2
created: 2026-09-10
# 注：以下字段（display_name / icon / visibility / triggers / inputs / outputs 等）
# 是 WorkBuddy 平台规范所需的元数据，用于界面显示与导入，
# 不要按通用 Agent Skills 规范「只留 name + description」删减。
display_name: "Excel PQ 批量刷新"
display_name_en: "Excel PQ Batch Refresh"
description_zh: "批量刷新文件夹内所有 Excel Power Query（含环境检查、.xlsm 化、VBA 宏注入、PAD 流程搭建指引、定时调度）"
description_en: "Batch-refresh every Excel Power Query in a folder — env check, .xlsm prep, VBA macro injection, PAD flow guide, scheduling."
visibility: "public"
icon: "https://openplatform-cdn.codebuddy.cn/public/skills/icons/excel-pq-batch-refresh.svg"
allowed-tools: Bash(powershell:*)
license: Personal
disable: false
triggers:
  - 批量刷新 Power Query
  - 定时刷新 Excel
  - PQ 自动刷新
  - Power Automate 刷新 Excel
  - 批量刷新 Excel 数据
  - 自动刷新 .xlsx 数据源
inputs:
  target_folder: "[必填] 含 .xlsx 的文件夹绝对路径（由用户粘贴提供）"
  schedule: "[可选] 定时规则，如 '每周一/五 09:00'，留空表示手动触发"
outputs:
  report: "已处理文件清单 + 失败列表 + 产物路径"
---

# Excel Power Query 批量自动刷新工作流

**当前版本**：v1.4.0　|　改动历史与实机测试记录：`references/CHANGELOG.md`

---

## ⛔ 安全边界（最高优先级，违反即停止）

> **只处理用户明确提供的那个文件夹，其他文件绝对不要动。**

1. **只认用户给的路径**：本 skill 只读写用户在 Step 0 里明确提供的 `target_folder`。除此之外，**任何文件、任何文件夹，一律不读取、不扫描、不列出、不复制、不打开、不修改**。
2. **禁止主动扫描用户目录找候选**（v1.4.0 新增）：问用户 `target_folder` 时，**直接请用户把路径粘贴过来**。**不要**为了"给几个候选路径"去 `ls` 用户的桌面、文档、下载等目录——即使只是只读，也会让人不适（2026-09-14 用户明确反馈）。
3. **禁止遍历目录树**：不得对任何盘符或共享盘执行「递归扫描整盘 / 列出全部文件 / 全盘搜索」这类操作。
4. **共享盘默认不碰**：`\\192.168.1.250\财务部` 等公司共享盘（账务、工资、报销、报表等），**除非用户在 Step 0 里把它作为 `target_folder` 明确给出，否则完全不访问**。
5. **越界即停**：任何步骤若需要访问 `target_folder` 以外的位置，**先停下来说明原因并询问用户**，得到明确同意后才做。

## 适用场景

- 需要把一个文件夹里**所有含 Power Query 的 Excel** 定时/手动批量刷新
- 典型场景：财务报表（日结/周结）、销售数据、自动化日报/周报
- 用户机器：Windows 10/11 + Microsoft 365 或 Office 2019+

## 工作流总览

```
Step 0: 开场问候 → 前置参数询问（先问候，再问参数）
    ↓
Step 1: 检查 Power Automate 就绪状态   ←── 若 Step 0 第 2 问答"不要"，直接跳到 Step 2
   └ 1.5: （本次装了 PAD 才做）给桌面放一个 PAD 快捷方式
    ↓
Step 2: Excel 批量转 .xlsm + 注入 RefreshPowerQuery 宏（含内联验证 + 进度存档）
    ↓
Step 3: 引导用户在 PAD 设计器搭建自动刷新流程 → 跑完做结果自检
    ↓
Step 4 (可选): 配置定时调度
```

> **跳过 Step 1 的分支**（v1.4.0 补充）：Step 0 第 2 问若用户答"不要"，**直接进 Step 2**；
> 若 Step 2 过程中出现 COM 异常或 Excel 报错，再回头补跑 Step 1 定位环境问题。

---

## 附件文件索引（用到哪个读哪个）

> 本 skill 采用渐进式披露：SKILL.md 只留主干与约束，细节都在 `references/` 下。
> **不要一上来把 references 全读了**——按下表按需读取。

| 文件 | 内容 | 什么时候读 |
|---|---|---|
| `references/pq_scan.py` | 只读扫描器（找含 PQ 的文件 + 预判跳过） | **每次都要用**（Step 0 扫描、Step 2.1） |
| `references/inline-snippets.md` | Step 1 / Step 2 的可复制代码块 | **执行 Step 1、Step 2 时** |
| `references/step1-check-pad.ps1` | Step 1 环境检查脚本（可直接调用） | 执行 Step 1 时 |
| `references/step3-pad-canvas-guide.md` | Step 3 小白逐击教程 + 3.9 运行结果自检 | **引导用户搭画布时**（Step 3） |
| `references/troubleshooting.md` | 完整失败诊断树 | **出问题时** |
| `references/CHANGELOG.md` | 版本历史 + 实机测试记录 | **需要追溯改动来源时**（日常不用读） |
| `references/step2-convert-and-inject.ps1` | Step 2 脚本版实现。⚠️ **含 `Add-Type`，当前沙箱下必被拦截**，仅作逻辑参考；实际执行请用 `inline-snippets.md` | 需要对照脚本逻辑时 |
| `references/vba-refresh-power-query.bas` | VBA 宏源码（手工导入用） | 需要在 VBE 里手工导入宏时 |

---

## Step 0 - 前置参数询问（必走）

> **硬性约定**：agent 加载本 skill 后，先问候、再问参数，拿到回答才跑 Step 1。
> 交互由 agent 用 AskUserQuestion 工具发起（每次 ≤4 个问题）。

### 提问硬规矩（v1.4.0 新增，必须遵守）

| 规矩 | 说明 |
|---|---|
| **不主动扫描找候选** | 问路径就直接请用户粘贴。**不要**去 `ls` 用户桌面/文档列候选 |
| **不出现行话** | 问题文案里**不得单独出现** `.xlsm` / `覆盖` / `子文件夹` / `递归` / `绝对路径` 这类词。必须出现时，**后面跟一句大白话解释**（例："以前弄过一次的（文件名后面是 .xlsm 结尾）"） |
| **说人话** | 面向初中生。其他禁用词："顶层""RRULE""宏模块""布尔""枚举" |
| **一屏内** | 问题 + 选项描述都要短，别让用户在选项里读小作文 |

### 开场问候（加载 skill 后第一句话就要说）

> **时机**：做任何检查、提任何问题之前先说。不要省略，也不要和参数表揉在一起。

**固定文案**（照抄）：

```
🎉 您好！我是您的 Excel 自动刷新小助手～

接下来我要带您做一件事：
把您文件夹里的 Excel 数据，变成能「自己刷新自己」的状态。
设置好之后，到点它自动就刷新完了，您完全不用管！😄

很简单的，全程我带着您走，
您只要回答我几个小问题就行 👇
```

**语气要求**：欢快热情、第一句必须是"您好"、emoji 克制（3~5 个）、长度 3~5 行、用词初中生能懂。

**反面示例**（不要这样）：

```
❌ 您好。本 skill 将执行 Excel Power Query 批量刷新工作流。
   请提供 target_folder 参数，并确认是否启用递归扫描。
```

### 第 1 轮（执行任何脚本前必问）

| # | 问题（照抄给用户看） | 真正含义 | 默认 | 必填 |
|---|---|---|---|---|
| 1 | 请把要处理的文件夹地址发给我 👉 打开那个文件夹，点一下最上面的地址栏，复制，然后粘贴给我 | `target_folder` | 无 | ✅ |
| 2 | 要不要先检查一下 Power Automate 有没有装好？ | 是否跑 Step 1 | `要` | 否 |
| 3 | 这个文件夹里面如果**还套着一层小文件夹**，小文件夹里的表要不要也算上？ | 是否递归扫子目录 | `要（推荐）` | 否 |

> 第 1 问的措辞要点：**给操作动作，不给术语**。说"点一下最上面的地址栏复制"，不说"请提供绝对路径"。
> 也不要替用户先去找——见"安全边界"第 2 条。
>
> **第 1 问的时机要点（2026-09-15 用户反馈，OPT-021）**：
> - **本会话里已经确认过路径的，直接沿用，不要再问**（用户明确说"直接用前面给你的路径"）
> - 路径未知时，就**只问这一个 plain 问题**，**禁止**附带"候选文件夹"选项——不扫描、不列桌面/文档里的疑似目录让用户挑（2026-09-15 实测：列了 3 个候选，用户还是得手打完整路径，纯属多余）
> - 提问次数能少则少：路径已知的场景下，Step 0 的必要提问可能只剩 0~1 个

### 扫描一次（第 1 轮之后、第 2 轮之前）

拿到 `target_folder` 后，先跑一次只读扫描，**用结果决定第 2 轮问什么**：

```bash
"C:/Users/Administrator/.workbuddy/binaries/python/versions/3.13.12/python.exe" \
  "C:/Users/Administrator/.workbuddy/skills/Excel PQ 批量刷新/references/pq_scan.py" \
  --target "<target_folder>" --out "%TEMP%/pq-scan.json"     # 含小文件夹时加 --recursive
```

读 `%TEMP%\pq-scan.json`，按下表决定：

| 扫描结果 | 怎么做 |
|---|---|
| `existingXlsm` 为空 | **跳过第 4 问**，直接说一句"这次都是新表，没有要覆盖的东西" |
| `willSkip == pqFiles` 且 `pqFiles > 0` | ⚠️ **提前警告**："按'保留旧的'这个选择，这次 **0 个文件**需要处理（旧的都还在）。你是想重新弄一遍，还是就这样？" |
| `pqFiles == 0` | 告知"这个文件夹里没找到带 Power Query 的表"，请用户确认文件夹是否选对，**不要继续** |
| 其他 | 正常问第 4 问 |

### 第 2 轮（处理 Excel 文件之前问）

| # | 问题（照抄给用户看） | 真正含义 | 默认 | 必填 |
|---|---|---|---|---|
| 1 | （**仅扫描发现有旧产物时才问**）这里面有几张表是**以前弄过一次的**（文件名后面是 .xlsm 结尾）。要不要重新弄一遍？选"不要"就是：原来弄好的不动，只弄没弄过的 | 是否覆盖同名 .xlsm | `不要` | 否 |
| 2 | 装宏用的那段代码，放在哪个"格子"里、叫什么名字？ | 宏模块名 / 宏函数名 | `RefreshModule` / `RefreshPowerQuery` | 否 |
| 3 | 刷新完等几秒再保存？ | wait_seconds | `15` | 否 |
| 4 | 要不要让它定时自动跑？ | 是否配 schedule | `不用` | 否 |

### 第 3 轮（仅用户选择"定时自动跑"后再问）

- 问："具体周几、几点几分自动跑？"（答如"每周一和周五上午 9 点"）
- agent 内部转成 RRULE，**不要给用户看这个术语**

### 进度续接（v1.4.0 新增）

> **Step 0 的第一件事，除了问候，是先读 `%TEMP%\excel-pq-progress.json`。**

- **文件不存在** → 按正常流程问参数
- **文件存在且在 7 天内** → 开场问候后补一句：

  > "上次您做到 Step 2 了 —— `[targetFolder]` 已经转好 `[n]` 个 .xlsm，还剩「在 Power Automate 里搭流程」没搭完。要接着来吗？"

  用户确认后**直接续 Step 3**，跳过已知参数的追问（路径、宏名、等待秒数都在存档里）。
- **文件存在但超过 7 天** → 视为过期，忽略它，按新任务走

> 存档由 Step 2 收尾时写入（见 `inline-snippets.md` 的 2.4 段）。这样一来，用户上午跑一半、下午回来，agent 也能接上。

### 用户说"按默认"时的处理

- 跳过对应轮次的提问，直接用默认值跑
- 但**"文件夹在哪里"永远必问**（没有默认值）

---

## 执行环境前提（先读，能省掉 80% 的试错）

| 前提 | 说明 | 对策 |
|---|---|---|
| 执行策略 | 本机 `Get-ExecutionPolicy` = **Restricted** | 先 `Set-ExecutionPolicy -Scope Process Bypass -Force`，或用 `-ExecutionPolicy Bypass -File` |
| 无 stdout 回显 | WorkBuddy 的 PowerShell 工具**不返回命令 stdout** | 所有输出必须 `Out-File -Encoding utf8` 落盘，再用 Read 读取。**不要用 `>`**（PS 5.1 会写成 UTF-16） |
| `.ps1` 能否执行 | ✅ **2026-09-14 实测修正**：纯 cmdlet 脚本（`step1-check-pad.ps1`）**可以直接 `& '路径' -参数` 调用成功**。此前记录的"`.ps1` 一律被静默拦截"不成立 | 优先直接调脚本；被拦时用 `inline-snippets.md` 的内联版 |
| COM / Add-Type | 沙箱内 `New-Object -ComObject`、`Add-Type` 被安全策略拦截 | **含 COM / Add-Type 的逻辑必须内联执行**（`dangerouslyDisableSandbox` + 落盘后 Read）；扫描则改用 Python `pq_scan.py`（零依赖、无 Add-Type） |
| 中文路径 | PS 5.1 读 UTF-8 无 BOM 的 `.ps1` 会把中文路径读乱 | `references/` 下所有 `.ps1` 必须 **UTF-8 with BOM**；改完务必复查前 3 字节是否为 `EF BB BF` |
| 脚本传参 | 通过 `powershell.exe -Command "...中文路径..."` 可能丢参数 | 优先在当前会话 `& "脚本路径" -Param ...` 直接调用 |
| 开工前查 Excel 进程 | 残留的无窗口 `EXCEL.EXE` 会被误认为"用户正开着的表"，还可能锁文件 | 跑 `step1-check-pad.ps1`（含进程体检）。**提示判据已收紧**：①无标题进程 ≥5 ②有 `~$` 锁文件且有无标题进程 ③无标题进程 ≥3 —— 满足任一才提示用户。**1 个无害进程不要再打断用户**（2026-09-14 修正） |

---

## Step 1 - 检查 Power Automate 是否就绪

### 检测项清单

| 项 | 检查位置 | 期望值 |
|---|---|---|
| PAD 主程序 | `C:\Program Files (x86)\Power Automate Desktop\dotnet\PAD.Designer.exe` | 存在 |
| PAD 版本 | 同上 `.VersionInfo.FileVersion` | >= 2.71 |
| agent 主目录 | `C:\Program Files (x86)\Power Automate agent for virtual desktops\` | 存在 |
| 关键服务 | UIFlowService / PADJavaSyncServiceRDP / UIFlowLogShipper / PADCrashMonitor | Running |
| PAD_JAVA 目录 | `C:\Users\Public\Documents\Microsoft\Power Automate Desktop\PAD_JAVA` | 存在 |
| 1722 状态 | Application 日志最近 24h 无 `PadJavaSyncServiceRDP` 启动错误 | 干净 |
| Excel 进程 | `Get-Process EXCEL` 的数量与无标题数 + 目标目录锁文件 | 见"执行环境前提"判据 |

### 调用方式

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
& "C:\Users\Administrator\.workbuddy\skills\Excel PQ 批量刷新\references\step1-check-pad.ps1" `
    -OutFile "$env:TEMP\step1-report.json" `
    -TargetFolder "<target_folder>"      # 传了才会检查目标目录的锁文件
```

然后用 Read 读取 `%TEMP%\step1-report.json`（不要依赖 stdout 回显）。

> 若脚本执行被拦，改用 `inline-snippets.md` 的「Step 1 环境检查（内联兜底版）」。

### 失败分支处理

| status | 处理方式 |
|---|---|
| `MISSING_PAD` | `winget install --id Microsoft.PowerAutomateDesktop --exact --accept-package-agreements --accept-source-agreements`；**装完走 1.5 建桌面快捷方式**（问用户放哪） |
| `MISSING_AGENT` | 从 `https://go.microsoft.com/fwlink/?linkid=2188766` 下载 `Setup.Microsoft.PowerAutomateAgent.exe`，运行 `-Install -ACCEPTEULA -Silent` |
| `SERVICES_DOWN` | 看报告里的 `servicesDetail` 找出哪个服务停了；**先查该服务的 ImagePath 是否存在**（决定是文件缺失还是被安全软件拦截），再按服务名 `Start-Service` 试启。注意：`PADJavaSyncServiceRDP` 停止**不一定**影响桌面流，可先询问用户是否要继续 |
| `MISSING_PAD_JAVA` | `New-Item -Path 'C:\Users\Public\Documents\Microsoft\Power Automate Desktop\PAD_JAVA' -ItemType Directory -Force`，然后 `Restart-Service PADJavaSyncServiceRDP` |
| `BLOCKED_1722` | 卸载重装流程（用 `/SKIPSTARTINGPOWERAUTOMATESERVICE` 装 PAD → `sc config UIFlowService obj= LocalSystem`），排障剧本见工作区 `memory/2026-08-31.md` |
| `READY` | 继续 Step 2 |

### 1.5 装完 PAD 后顺手做：给桌面放一个 Power Automate 快捷方式

> **仅在本次实际装了 PAD 时做**。**先问用户放哪**（我的桌面 / 公共桌面），别默认。

```powershell
$src = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Power Automate\Power Automate.lnk'
$dst = Join-Path $env:USERPROFILE 'Desktop\Power Automate.lnk'   # 公共桌面：'C:\Users\Public\Desktop'
if (Test-Path $src) { Copy-Item $src $dst -Force; $r = "COPIED: $dst" } else { $r = 'SRC_MISSING' }
$r | Out-File "$env:TEMP\step1-shortcut.txt" -Encoding utf8
```

**三条易错点**：
- ⚠️ **不要**用 `New-Object -ComObject WScript.Shell` 建快捷方式——本机安全策略拦 COM（复制官方 `.lnk` 可绕开）
- ⚠️ `.lnk` 是**二进制**，别当文本读；取路径要用字节 + ISO-8859-1 解码
- ⚠️ 官方快捷方式指向 `PAD.Console.Host.exe`（控制台宿主），**不是** `PAD.Designer.exe`

**官方快捷方式缺失时**：首选让用户从开始菜单把「Power Automate」拖到桌面；兜底再手建。

---

## Step 2 - Excel 批量转 .xlsm + 注入刷新宏

> **完整可复制代码在 `references/inline-snippets.md`**（2.1 扫描 → 2.2 注册表 → 2.4 转换注入 → 进度存档）。
> 本节只讲**关键决策与陷阱**。

### 关键决策

- 只处理含 Power Query 的文件（其他 .xlsx 跳过），避免污染
- 原 `.xlsx` 永远保留（不覆盖），产物 `.xlsm` 与原文件同目录
- 重复运行要**幂等**（先删同名模块再注入）

### 子步骤概要

| 步骤 | 做什么 | 落盘产物 |
|---|---|---|
| 2.1 | 跑 `pq_scan.py` 扫描 + 预筛选 | `%TEMP%\pq-scan.json` |
| 2.2 | 设 Excel 信任中心（`AccessVBOM=1` / `VBAWarnings=1` / 受信任位置**幂等**） | `%TEMP%\step2-registry.txt` |
| 2.3 | VBA 宏代码常量（`RefreshAll` → `CalculateUntilAsyncQueriesDone` → `Save`） | — |
| 2.4 | Excel COM 单实例循环：Open → `SaveAs(path, 52)` → 删同名模块 → 注入新模块 → **内联验证模块数** → Close | `%TEMP%\step2-report.json` |
| 2.4b | **写进度存档** `%TEMP%\excel-pq-progress.json` | 供下次续接 |
| 2.5 | （可选）轻量校验 `vbaProject.bin` 是否存在 | — |
| 2.6 | （备用）另开 Excel 独立复核模块 | — |

### Step 2 易错点（必须遵守）

- ⚠️ 注册表设置必须在 `New-Object -ComObject Excel.Application` **之前**完成（COM 实例化时读信任设置）
- ⚠️ 原 `.xlsx` 永远保留，不覆盖（让用户可回退）
- ⚠️ 模块名与宏名必须一致（默认 `RefreshModule` / `RefreshPowerQuery`）
- ⚠️ `SaveAs` 第二个参数 `52` = xlOpenXMLWorkbookMacroEnabled，**不要写成别的**
- ⚠️ 收尾**必须回收 Excel 残留进程**（`Quit()` 不够，实测连跑 2 次残留 3 个）；只回收**本次启动的**（用开工前的 PID 基线比对，别误杀用户自己开着的）
- ⚠️ 报告与存档**不要写进 skill 的 `references/` 目录**（会污染技能包），一律写 `%TEMP%`
- ⚠️ 含中文路径的 `.ps1` 必须 UTF-8 with BOM
- ⚠️ 受信任位置必须**幂等**（同路径复用），否则会不断追加 `Location8/Location9…` 污染注册表
- ⚠️ **不要**依赖 `GetEntry('customXml/')` 目录条目识别 PQ（会漏检），遍历 `namelist()` 逐条匹配 `customXml/item*.xml`（`pq_scan.py` 已内置正确逻辑）

### Step 2 输出

`%TEMP%\step2-report.json`：

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

`Verify` 取值：`OK` / `NO_MODULE`（模块没注进去）/ `DUP_MODULE(n)`（重复 n 个，幂等失效）。

---

## Step 3 - 引导用户搭建 PAD 自动刷新工作流

> **本节的完整内容在 `references/step3-pad-canvas-guide.md`——开始引导前读它。**
> 这一步**不出代码**，PAD 设计器是 GUI 操作，需要用户亲手搭。

**引导要点（先记住这 4 条，细节读指南）**：

1. **一次只给 1-2 个小节**，不要整页甩给用户。每个动作讲清"在哪搜 → 拖到哪 → 每个框填什么"
2. **动手前先讲"搭前必读 3 条"**：不要删了重拖动作（会触发变量后缀漂移）／「启动 Excel」必须选"启动新实例"／运行前关闭所有目标 Excel
3. **搭完先要截图，agent 逐动作复核变量引用后再让用户点运行**（人在回路，避免带病运行）
4. **跑完之后要做结果自检**（指南 3.9）——看 PAD 有没有报错、看 `.xlsm` 修改时间戳有没有变新、抽一张表打开看数据、**检查 Excel 进程有没有残留**。**"跑绿了"不等于"刷新成功了"**
5. **⚠️ 运行中途不要点 ⏹ 停止**：停止只终止流程逻辑，**不会关闭它已经开出来的 Excel**，必然留下僵尸进程（会锁住文件）。真觉得卡住了，先由 agent 判断（单文件 60~90 秒属正常）

**8 个动作速览**：① 获取文件夹中的文件（筛 `*.xlsm`）→ ② For each → ③ 启动 Excel（**新实例**）→ ④ 运行宏 `RefreshPowerQuery` → ⑤ 等待 15 秒 → ⑥ 保存 Excel → ⑦ 关闭 Excel（不保存）→ ⑧ **终止进程 `EXCEL`**（兜底清理残留壳，PAD 关闭动作关不干净进程，2026-09-15 实测）。

---

## Step 4 - 定时调度（可选）

**仅当 Step 0 里 schedule 不为空时配置**（挂在 **Main** 上，挂错对象会导致只有单个模块被调度）：

1. PAD 左侧流属性 → 计划运行
2. 勾选"启用计划运行"
3. 添加时间点（例：每周一 + 每周五 09:00 → `FREQ=WEEKLY;BYDAY=MO,FR;BYHOUR=9;BYMINUTE=0`）
4. ⚠️ 电脑必须**保持开机、不进入休眠**（PAD 计划运行依赖本地调度）

详见 `references/step3-pad-canvas-guide.md` 的 3.7。

---

## 失败诊断（最常踩的 5 条）

| 现象 | 根因 | 处理 |
|---|---|---|
| 注入宏时 Excel 弹"VBA 工程对象模型访问被拒绝" | `AccessVBOM` 没设，或 Excel 已在运行 | 关 Excel → 重跑 2.2 → 重试 |
| 循环只刷第一个文件 / 后面动作报"变量未定义" | 变量后缀漂移（`Files2`/`ExcelInstance2`） | 按 step3 指南 3.3 统一变量引用；**不要删了重拖** |
| Power Query 数据没更新 | 等待秒数太短（第一大原因） | 调到 30 秒重跑 |
| 第二次运行报"文件被占用" | 上次残留 `EXCEL.EXE` | 回收残留进程，或手动结束 |
| 命令 exit 0 但看不到输出 | 执行环境不回显 stdout | `Out-File -Encoding utf8` 落盘后 Read |

> **完整诊断树（含 17 条现象）见 `references/troubleshooting.md`。**

---

## Examples（典型调用场景）

### Example 1: 端到端（最完整）
**用户请求**："把 `D:\财务报表` 里的 Excel 都变成能自己刷新的"

**AI 执行**：
1. 说开场问候 → 第 1 轮问 3 个问题（**直接请用户粘贴路径**，不扫描其目录）
2. 跑 `pq_scan.py` → 读 `%TEMP%\pq-scan.json`
3. 按扫描结果问第 2 轮（有旧产物才问"要不要重新弄一遍"；宏名/等待秒数/是否定时）
4. Step 1 环境检查（若第 2 问答"要"）
5. Step 2 转换注入 → 读 `%TEMP%\step2-report.json` → 写进度存档
6. 读 `step3-pad-canvas-guide.md`，一小节一小节带用户搭画布 → 要截图复核 → 跑结果自检

### Example 2: 周一/周五定时刷新
同 Example 1，Step 4 配 `FREQ=WEEKLY;BYDAY=MO,FR;BYHOUR=9;BYMINUTE=0`，并提示保持开机不休眠。

### Example 3: 仅环境检查
**用户请求**："检查我这台电脑 Power Automate 装好了没"
→ 只跑 `step1-check-pad.ps1`，读报告，按 `status` 给建议（见 Step 1 失败分支表）。

### Example 4: 仅转换、不搭 PAD 流
**用户请求**："把目标文件夹里的表都转成能刷新宏的状态，流程我自己搭"
→ 跑 Step 0 + Step 2 即可，跳过 Step 3/4，但**要提示用户**：还没搭 PAD 流程，光有宏不会自动刷新。

### Example 5: 中断后回来续做
**用户请求**：（隔了几小时）"继续吧"
→ 先读 `%TEMP%\excel-pq-progress.json` → 若在 7 天内，直接说"上次做到 Step 2，还剩搭画布，接着来吗？" → 确认后进 Step 3。

---

## 关联资源

> 以下文档**不在 skill 目录内**，位于工作区 `自动化工作流`，用绝对路径访问。

- 完整操作手册：`C:\Users\Administrator\WorkBuddy\自动化工作流\.workbuddy\memory\PowerAutomate批量刷新PowerQuery操作手册.md`
- Step 2 思路文档：`C:\Users\Administrator\WorkBuddy\自动化工作流\.workbuddy\memory\Excel批量转xlsm注入宏自动化思路.md`
- PAD 卸载重装排障剧本（1722）：`C:\Users\Administrator\WorkBuddy\自动化工作流\.workbuddy\memory\2026-08-31.md`
- **本 skill 的迭代台账**（优化意见与裁决记录）：`C:\Users\Administrator\WorkBuddy\自动化工作流\.learnings\SKILL-EVOLUTION.md`
