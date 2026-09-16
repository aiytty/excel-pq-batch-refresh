#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel PQ 批量刷新 - Step 2.1 只读扫描器

用途
    扫描目标文件夹，找出所有含 Power Query 的 Excel 文件，
    并预判哪些文件会因「已存在同名 .xlsm」而被跳过。

只读保证
    不修改、不移动、不打开任何文件。仅以 zipfile 方式读取 .xlsx 的包结构
    （xl/connections.xml 与 customXml/item*.xml），整个过程不启动 Excel。

为什么用 Python 而不是 PowerShell
    PowerShell 版扫描需要 `Add-Type -AssemblyName System.IO.Compression.FileSystem`，
    该调用在 WorkBuddy 沙箱中被硬拦截（报 "Add-Type compiles and loads .NET code at runtime"）。
    Python 标准库 zipfile 无此限制，且更快更稳，零第三方依赖。

用法
    python pq_scan.py --target "C:\\path\\to\\folder" [--recursive] [--out report.json]

    --target     必填，目标文件夹绝对路径（中文路径可用）
    --recursive  可选，连子文件夹一起扫；不传只扫顶层
    --out        可选，报告落盘路径；默认 %TEMP%\\pq-scan.json

输出字段
    target        扫描的目标文件夹
    recursive     是否递归
    totalFiles    扫到的 Excel 文件总数（.xlsx/.xlsm/.xls 全部计入）
    pqFiles       含 Power Query 的**源文件（.xlsx）**数量——这才是 Step 2 要处理的集合
    existingXlsm  已存在的 .xlsm 文件名列表（配合「要不要重新弄一遍」的提问使用）
    willProcess   本次预计会处理的文件数（源文件里没有同名 .xlsm 的）
    willSkip      本次预计会跳过的文件数（源文件已有同名 .xlsm，且用户选了「不覆盖」）
    willSkipNames 会被跳过的源文件名
    files         逐文件明细：name / ext / kind / sizeKB / modified / hasPQ / reason
                  kind = source（源文件）| output（已生成的 .xlsm 产物）

    reason 取值：connections.xml | customXml/Mashup | no-PQ | xls-unsupported | ZIP_ERR: ...
"""

import argparse
import datetime
import json
import os
import sys
import zipfile

EXCEL_EXTS = (".xlsx", ".xlsm", ".xls")


def has_pq(path):
    """判断一个 Excel 文件是否含 Power Query。

    返回 (bool, reason)。

    判定顺序（两者命中其一即为含 PQ）：
      1. 包里存在 xl/connections.xml —— 绝大多数 PQ 文件走这条
      2. customXml/item*.xml 里出现 "Mashup" 字样

    注意：第 2 步必须逐条遍历 namelist()，
    不能先判断 customXml/ 目录条目是否存在——Excel 写出的包未必包含目录条目，
    那样会让整个 Mashup 分支被跳过，导致真实 PQ 文件漏检（v1.1.0 修过的坑）。
    """
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            if "xl/connections.xml" in names:
                return True, "connections.xml"
            for n in names:
                if n.startswith("customXml/item") and n.endswith(".xml"):
                    try:
                        if "Mashup" in z.read(n).decode("utf-8", "ignore"):
                            return True, "customXml/Mashup"
                    except Exception:
                        continue
            return False, "no-PQ"
    except Exception as e:
        return False, "ZIP_ERR: %s" % e


def collect(target, recursive):
    """收集目标文件夹下的 Excel 文件。不递归时只看顶层。"""
    out = []
    if recursive:
        for root, _dirs, names in os.walk(target):
            for n in names:
                if n.lower().endswith(EXCEL_EXTS) and not n.startswith("~$"):
                    out.append(os.path.join(root, n))
    else:
        try:
            for n in os.listdir(target):
                p = os.path.join(target, n)
                if os.path.isfile(p) and n.lower().endswith(EXCEL_EXTS) and not n.startswith("~$"):
                    out.append(p)
        except Exception as e:
            print("无法读取目标文件夹: %s" % e, file=sys.stderr)
    return sorted(out)


def main():
    ap = argparse.ArgumentParser(description="Excel PQ 批量刷新 - 只读扫描器")
    ap.add_argument("--target", required=True, help="目标文件夹绝对路径")
    ap.add_argument("--recursive", action="store_true", help="连子文件夹一起扫")
    ap.add_argument("--out", default=None, help="报告落盘路径，默认 %%TEMP%%\\pq-scan.json")
    args = ap.parse_args()

    target = os.path.abspath(args.target)
    if not os.path.isdir(target):
        print("目标文件夹不存在: %s" % target, file=sys.stderr)
        return 2

    out_path = args.out
    if not out_path:
        out_path = os.path.join(os.environ.get("TEMP", "."), "pq-scan.json")

    files = collect(target, args.recursive)

    details = []
    pq_paths = []
    existing_xlsm = []

    for p in files:
        name = os.path.basename(p)
        ext = os.path.splitext(name)[1].lower()
        try:
            st = os.stat(p)
            size_kb = round(st.st_size / 1024.0, 1)
            modified = datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M")
        except Exception:
            size_kb, modified = 0, "N/A"

        if ext == ".xls":
            pq, reason = False, "xls-unsupported"
        else:
            pq, reason = has_pq(p)

        # kind: source = 待处理的源文件（.xlsx/.xls）；output = 已生成的产物（.xlsm）
        # 只有 source 才参与「会不会被处理」的预判——产物本身不是处理对象
        kind = "output" if ext == ".xlsm" else "source"
        if pq and kind == "source":
            pq_paths.append(p)
        if kind == "output":
            existing_xlsm.append(name)

        details.append({
            "name": name,
            "ext": ext,
            "kind": kind,
            "sizeKB": size_kb,
            "modified": modified,
            "hasPQ": pq,
            "reason": reason,
        })

    # 预判：含 PQ 的 .xlsx，若同名 .xlsm 已存在，则在「不覆盖」策略下会被跳过
    existing_set = set(n.lower() for n in existing_xlsm)
    will_skip, will_process = [], []
    for p in pq_paths:
        name = os.path.basename(p)
        stem = os.path.splitext(name)[0].lower()
        if stem + ".xlsm" in existing_set:
            will_skip.append(name)
        else:
            will_process.append(name)

    report = {
        "target": target,
        "recursive": bool(args.recursive),
        "scannedAt": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "totalFiles": len(files),
        "pqFiles": len(pq_paths),
        "existingXlsm": existing_xlsm,
        "willProcess": len(will_process),
        "willSkip": len(will_skip),
        "willSkipNames": will_skip,
        "files": details,
    }

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    print("scan written to %s" % out_path)
    n_source = sum(1 for d in details if d["kind"] == "source")
    n_output = len(details) - n_source
    print("totalFiles=%d (源文件 %d / 已生成 %d) pqSourceFiles=%d willProcess=%d willSkip=%d"
          % (report["totalFiles"], n_source, n_output,
             report["pqFiles"], report["willProcess"], report["willSkip"]))
    for d in details:
        print("  [%-6s] %-42s %8sKB  PQ=%s (%s)"
              % (d["kind"], d["name"], d["sizeKB"], d["hasPQ"], d["reason"]))
    if existing_xlsm:
        print("  已存在的 .xlsm: %s" % ", ".join(existing_xlsm))
    return 0


if __name__ == "__main__":
    sys.exit(main())
