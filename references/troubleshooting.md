# 失败诊断树（troubleshooting）

> **本文件从 SKILL.md 拆出**（v1.4.0），并合并了原 `references/README.md` 的排查表。
> SKILL.md 里只留最常踩的 5 条，**完整排查表在本文件**——出问题时读。

---

## 按现象查

| 现象 | 根因 | 处理 |
|---|---|---|
| Step 2 注入宏时 Excel 弹"VBA 工程对象模型访问被拒绝" | `AccessVBOM` 没设，或 Excel 已在运行 | 关闭 Excel → 重跑 Step 2.2（务必在启动 Excel **之前**设好注册表）→ 重试 |
| Step 2 跑完但 .xlsm 里没有宏 | 重复运行没保幂等 + COM 残留 | 关 Excel → 检查 `EXCEL.EXE` 残留进程 → 结束 → 重跑 |
| Step 3 跑 PAD 报 1722 | agent / Java sync 异常 | 回 Step 1 走 `BLOCKED_1722` 分支 |
| Step 3 跑 PAD 报"找不到 RefreshPowerQuery 宏" | Step 2 没成功注入 / 宏名拼错 | 用 VBA 编辑器打开 .xlsm 看模块名；或回 Step 2.6 用 COM 复核 |
| 循环只刷第一个文件 | `ExcelInstance` 没在 For each 内部新建 | 检查"启动 Excel"是否在 For each 内部；**双击改位置/改引用，不要删除重拖**（会触发变量后缀漂移） |
| 循环一次都没执行 / 后面动作引用变量报"未定义" | 变量后缀漂移（动作 1/3 重拖过，产出 `Files2`/`ExcelInstance2`，下游仍引用旧名） | 按 step3 指南 3.3 自检清单统一变量引用；或在右侧"变量"面板重命名（会自动同步引用） |
| 运行 Excel 报"找不到实例 / 实例无效" | "启动 Excel"选了"使用现有的 Excel 进程"但目标没开 | 双击动作改为"启动新实例" |
| Power Query 数据没更新 | 等待秒数太短（第一大原因） | 调到 30 秒，重跑；详见 step3 指南 3.9 |
| 运行后 Excel 报"文件被独占打开" | PAD 关 Excel 前用户自己打开了 | 确保循环内最后一步是关闭 Excel |
| 报"在此系统上禁止运行脚本"（UnauthorizedAccess） | 执行策略为 Restricted | `Set-ExecutionPolicy -Scope Process Bypass -Force`，或用 `-ExecutionPolicy Bypass -File` 调起 |
| 明明有 PQ 的文件却没被处理 | 识别逻辑依赖 `customXml/` 目录条目 | 改为遍历 `$zip.Entries` / Python `namelist()` 匹配 `customXml/item*.xml`（见 Step 2.1；用 `pq_scan.py` 可避免此坑） |
| `New-Object -ComObject` 直接被拒 / `Add-Type` 被拒 | 沙箱安全策略拦截 COM 与动态编译 | 以沙箱外权限执行并取得用户确认；扫描改用 Python `pq_scan.py` |
| 第二次运行报"文件被占用"或 SaveAs 失败 | 上次残留 `EXCEL.EXE` | 回收残留进程（见 Step 2.4 收尾段），或先手动结束 `EXCEL.EXE` |
| 注册表 Trusted Locations 出现多条同路径 | 受信任位置非幂等 | 清理冗余项，改用 Step 2.2 的幂等写法 |
| 命令 exit 0 但看不到任何输出 | 执行环境不回显 stdout | `Out-File -Encoding utf8` 落盘后 Read（不要用 `>`，PS 5.1 会写成 UTF-16） |
| 中文路径乱码 / 脚本读到错误路径 | `.ps1` 缺 UTF-8 BOM | 重新以 **UTF-8 with BOM** 保存（VS Code：编码 → Save with Encoding → UTF-8 with BOM；Notepad++：编码 → UTF-8-BOM） |
| SaveAs 后看不到 .xlsm | 资源管理器没刷新 | 重新读取目录，或按 F5 |

---

## 新增 PQ 文件后不生效

**现象**：往目标文件夹里放了一个新的带 Power Query 的 .xlsx，但 PAD 跑完它没被刷新。

**原因**：PAD 流程只处理 `.xlsm` 文件（动作 1 的筛选是 `*.xlsm`）。新建的 `.xlsx` 还没经过 Step 2 转换。

**处理**：对新增文件单独跑一次 Step 2（`pq_scan.py` 会告诉你哪些源文件还没有对应产物），转换完 PAD 下次运行就会带上它。

---

## 跨机器使用注意（环境假设）

- Windows 10 / 11 + Microsoft 365 或 Office 2019+
- 已安装 Power Automate Desktop **2.71+**
- 已安装 Power Automate agent for virtual desktops **2.67+**
- Office 注册表分支按 **16.0** 定位（`HKCU:\Software\Microsoft\Office\16.0\Excel\...`）；若机器上是 Office 2016 以下，需自行确认版本号分支
- 若 PAD / `UIFlowService` 启动崩溃（1722），走 Step 1 的 `BLOCKED_1722` 分支，排障剧本见工作区 `memory/2026-08-31.md`
