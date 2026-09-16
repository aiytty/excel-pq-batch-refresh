# Excel PQ 批量刷新 (excel-pq-batch-refresh)

![version](https://img.shields.io/badge/version-1.4.3-blue)

![platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey)

![requires](https://img.shields.io/badge/requires-Power%20Automate%20Desktop%20%2B%20Excel-orange)

![license](https://img.shields.io/badge/license-Personal-green)

一个 [WorkBuddy](https://www.workbuddy.cn) Skill：**把一个文件夹里所有含 Power Query 的 Excel，变成"能自己刷新自己"的状态**——批量转 `.xlsm`、注入刷新宏、带小白一步步搭 Power Automate Desktop 刷新流程、可选定时调度。

> 写给谁用：每天要手动打开 N 张表 → 点刷新 → 保存的财务 / 运营同学。  
> 特别照顾**零编程基础**用户：所有提问不说行话，所有操作给截图。

---

## 它解决什么问题

| 手动现状                          | 用这个 Skill 之后                  |
| ----------------------------- | ----------------------------- |
| 每天打开 10 张表，逐个"全部刷新 → 保存 → 关闭" | 一次配置，之后点一次 ▶（或到点自动）全部刷完       |
| 表是 `.xlsx`，根本存不进宏             | 自动批量转 `.xlsm`，原文件永远保留可回退      |
| 不会写 VBA、不敢碰注册表                | 宏自动注入，信任中心自动配好（幂等，不污染）        |
| Power Automate 界面看不懂          | 逐击教程 + 实机配置截图，照着红框填就行         |
| 跑完不知道到底刷新成功没有                 | 4 步自检：红叉 / 时间戳 / 抽表看数据 / 进程残留 |



---

## 工作流程

```
Step 0  开场问候 → 问参数（路径必问，其余有默认值；支持中断后进度续接）
   ↓
Step 1  体检 Power Automate 就绪状态（主程序 / agent / 4 个服务 / 1722 事件）
   ↓
Step 2  Excel 批量转 .xlsm + 注入 RefreshPowerQuery 宏（COM 单实例循环，含内联验证）
   ↓
Step 3  引导用户在 PAD 设计器搭 8 动作刷新画布 → 跑完做结果自检
   ↓
Step 4  （可选）配置定时调度，如每周一/五 09:00
```

### PAD 画布的 8 个动作（Step 3 的产物）

```
① 获取文件夹中的文件（筛 *.xlsm）
② For each
③ 启动 Excel（必须"启动新实例"）
④ 运行 Excel 宏 RefreshPowerQuery
⑤ 等待 15 秒（PQ 异步刷新，宁多勿少）
⑥ 保存 Excel
⑦ 关闭 Excel（不保存——前面已单独保存）
⑧ 终止进程 EXCEL（兜底：PAD 关闭动作关不干净进程，实测必加）
```

> ⑧ 是实机跑出来的教训：不加它，每轮运行会残留无窗口的 Excel "僵尸壳"，锁住文件导致下次运行报"文件被占用"。

---

## 核心特性

- **安全边界硬约束**：只读写用户明确提供的那个文件夹，其余一律不碰；禁止主动扫描用户目录找"候选路径"；公司共享盘默认不访问
- **幂等可重跑**：同名宏模块先删后注入；受信任位置复用不重复追加；原 `.xlsx` 永不覆盖
- **进度续接**：Step 2 收尾写进度存档，中断后回来说一句"继续"，从上次的步骤接着来
- **说人话**：提问文案禁用 `.xlsm` / 递归 / 绝对路径等术语，面向初中生水平（用户实测反馈驱动）
- **实机验证**：每条关键结论都有实测记录，踩过的坑（变量后缀漂移、僵尸进程、stdout 无回显、UTF-8 BOM……）都写进了文档和诊断树

## 仓库结构

```
excel-pq-batch-refresh/
├── SKILL.md                          # 主文件：工作流主干 + 安全边界 + 陷阱清单
├── workbuddy.json                    # WorkBuddy 平台元数据
├── _icon.svg                         # 64×64 图标
├── _skillhub_meta.json               # SkillHub 元数据
└── references/                       # 按需加载的细节文档与脚本
    ├── pq_scan.py                    #   零依赖只读扫描器（找含 PQ 的文件 + 预判跳过）
    ├── inline-snippets.md            #   Step 1/2 可直接复制的代码块（执行用这个）
    ├── step1-check-pad.ps1           #   PAD 环境体检脚本（可直接调用）
    ├── step2-convert-and-inject.ps1  #   Step 2 脚本版（逻辑参考，沙箱内被拦）
    ├── vba-refresh-power-query.bas   #   VBA 宏源码（VBE 手工导入用）
    ├── step3-pad-canvas-guide.md     #   PAD 画布逐击教程 + 配置截图 + 结果自检
    ├── assets/                       #   动作配置实机截图（照图填红框）
    ├── troubleshooting.md            #   失败诊断树（17 条现象）
    └── CHANGELOG.md                  #   版本历史 + 实机测试记录
```

## 环境要求

| 项      | 要求                                               |
| ------ | ------------------------------------------------ |
| 系统     | Windows 10 / 11                                  |
| Office | Microsoft 365 或 Office 2019+（需含 Power Query）     |
| 自动化    | Power Automate Desktop（免费版即可，Step 1 会体检，缺了给安装命令） |
| AI 端   | WorkBuddy（加载 Skill 并驱动全流程）                       |

## 安装使用

### 方式一：导入 WorkBuddy（推荐）

1. 下载本仓库（Code → Download ZIP）
2. WorkBuddy → 专家·技能·连接器 → 技能 → 导入，选择 ZIP
3. 对 WorkBuddy 说：**"批量刷新 Power Query，文件夹是 `[你的文件夹路径]`"**，跟着引导走完 Step 0~4

### 方式二：只用脚本（不用 WorkBuddy）

`references/` 下的脚本可以独立使用：

```powershell
# 环境体检
Set-ExecutionPolicy -Scope Process Bypass -Force
& ".\references\step1-check-pad.ps1" -OutFile "$env:TEMP\step1-report.json"

# 扫描目标文件夹里含 Power Query 的文件（需 Python 3.x，零第三方依赖）
python .\references\pq_scan.py --target "D:\你的文件夹" --recursive --out "$env:TEMP\pq-scan.json"
```

宏转换与注入涉及 Excel COM，完整可复制代码见 `references/inline-snippets.md`（2.4 节）。

## 常见问题（Top 5，完整 17 条见 troubleshooting.md）

| 现象                            | 根因                           | 处理                        |
| ----------------------------- | ---------------------------- | ------------------------- |
| 注入宏时 Excel 弹"VBA 工程对象模型访问被拒绝" | `AccessVBOM` 没设，或 Excel 已在运行 | 关 Excel → 重设注册表 → 重试      |
| 循环只刷第一个文件 / 报"变量未定义"          | 变量后缀漂移（`Files2`）             | 别删了重拖动作，双击改参数统一变量名        |
| PQ 数据没更新                      | 等待秒数太短（第一大原因）                | 等待调到 30 秒重跑               |
| 第二次运行报"文件被占用"                 | 上次残留 `EXCEL.EXE`             | 结束残留进程；确认画布有第 ⑧ 个动作       |
| 运行中途按了 ⏹，后面文件没刷               | 停止只停流程、不关 Excel              | 让它跑完（单文件 60~90 秒属正常），别中途停 |

## 版本

当前 **v1.4.3**。完整版本历史与实机测试记录见 [`references/CHANGELOG.md`](references/CHANGELOG.md)。

迭代方式：每次实跑暴露的问题都立案编号（OPT-xxx）、用户确认后统一落地、升版本号记账——目前累计落地 **23 条**实战优化。

## 安全与隐私

- Skill 本体只处理**用户当次明确指定的文件夹**，不读取、不扫描其他任何位置
- 所有报告 / 存档写入系统临时目录（`%TEMP%`），不污染技能包与业务目录
- 不含任何遥测、网络上传逻辑；脚本可全文审阅

## License

Personal —— 供个人学习与内部使用。欢迎提 Issue 交流踩坑经验。
