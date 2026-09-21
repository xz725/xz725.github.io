#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_data.py — 从 raw_data/about.xlsx 生成 _data/about.json

用法（仓库根目录）：
    uv run --with openpyxl .agents/skills/update-website/scripts/build_data.py

数据源为 xlsx 中第 2 个 sheet 到最后一个 sheet，与网站 tab 一一对应。
输出为 JSON（Jekyll 原生支持 _data/*.json），全量覆盖写入。
"""
import json
import re
import sys
from pathlib import Path

import openpyxl

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parents[3]  # .agents/skills/update-website/scripts -> repo root
XLSX = REPO_ROOT / "raw_data" / "about.xlsx"
OUT = REPO_ROOT / "_data" / "about.json"

CJK_RE = re.compile(r"[一-鿿]")


def clean(v):
    """单元格归一化：None/占位符 -> ''，数字去尾零，字符串去空白。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return "" if s in {"—", "-", "--"} else s


def has_cjk(s):
    return bool(CJK_RE.search(s or ""))


def rows_of(ws):
    """跳过第 1 行（合并标题行）与第 2 行（表头），返回数据行。"""
    rows = []
    for r in ws.iter_rows(min_row=3, values_only=True):
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        rows.append([clean(v) for v in r])
    return rows


def split_citation(a, b):
    """期刊/会议 sheet 最后两列在数据行中与表头互换：按是否含中文判断。
    返回 (citation_en, note_zh)。"""
    if has_cjk(a) and not has_cjk(b):
        return b, a
    return a, b


def extract_patent_title(s):
    """美国专利"英文名称"列常含发明人前缀，取引号内文本；无引号则取整串。"""
    s = s.strip()
    for q_open, q_close in [("“", "”"), ('"', '"')]:
        i = s.find(q_open)
        if i != -1:
            t = s[i + 1:]
            j = t.find(q_close)
            if j != -1:
                t = t[:j]
            return t.strip()
    return s


def build(wb):
    data = {}

    # 工作经历：序号|开始|结束|单位中|单位英|部门|职位中|职位英|状态|中文描述|英文描述|备注
    data["experience"] = [
        {
            "start": r[1], "end": r[2],
            "org_zh": r[3], "org_en": r[4],
            "dept": r[5],
            "role_zh": r[6], "role_en": r[7],
            "status": r[8],
            "desc_zh": r[9], "desc_en": r[10],
            "note": r[11] if len(r) > 11 else "",
        }
        for r in rows_of(wb["工作经历"]) if r[0]
    ]

    # 学术兼职与职称：序号|类别|开始|结束|机构中|机构英|职务中|职务英|备注
    data["appointments"] = [
        {
            "category": r[1], "start": r[2], "end": r[3],
            "org_zh": r[4], "org_en": r[5],
            "role_zh": r[6], "role_en": r[7],
            "note": r[8] if len(r) > 8 else "",
        }
        for r in rows_of(wb["学术兼职与职称"]) if r[0]
    ]

    # 教育经历：序号|开始|结束|学校中|学校英|国家|学位|专业|导师|备注
    data["education"] = [
        {
            "start": r[1], "end": r[2],
            "school_zh": r[3], "school_en": r[4],
            "country": r[5], "degree": r[6], "major": r[7],
            "advisor": r[8], "note": r[9] if len(r) > 9 else "",
        }
        for r in rows_of(wb["教育经历"]) if r[0]
    ]

    # 荣誉奖项：序号|年份|类别|名次|名称中|名称英|颁发方|备注
    data["honors"] = [
        {
            "year": r[1], "category": r[2], "rank": r[3],
            "name_zh": r[4], "name_en": r[5],
            "issuer": r[6], "note": r[7] if len(r) > 7 else "",
        }
        for r in rows_of(wb["荣誉奖项"]) if r[0]
    ]

    # 期刊文章：编号|年份|...|引用/说明(数据行中两列互换)
    journals = []
    for r in rows_of(wb["期刊文章"]):
        if not r[0]:
            continue
        citation, note_zh = split_citation(r[12] if len(r) > 12 else "",
                                           r[13] if len(r) > 13 else "")
        journals.append({"num": r[0], "year": r[1], "citation": citation,
                         "note_zh": note_zh})
    data["journals"] = journals

    # 会议文章：编号|年份|...|引用/说明(数据行中两列互换)；编号 'T' 为博士论文
    conferences, thesis = [], []
    for r in rows_of(wb["会议文章"]):
        if not r[0]:
            continue
        citation, note_zh = split_citation(r[8] if len(r) > 8 else "",
                                           r[9] if len(r) > 9 else "")
        item = {"num": r[0], "year": r[1], "citation": citation,
                "note_zh": note_zh}
        (thesis if r[0].upper() == "T" else conferences).append(item)
    data["conferences"] = conferences
    data["thesis"] = thesis

    # 美国专利：编号|专利号|英文名称|中文名称|发明人|授权年份|国家|法律状态|备注
    data["patents_intl"] = [
        {
            "num": r[0], "pid": r[1],
            "title_en": extract_patent_title(r[2]), "title_zh": r[3],
            "inventors": r[4], "year": r[5],
            "country": r[6], "status": r[7],
            "note": r[8] if len(r) > 8 else "",
        }
        for r in rows_of(wb["美国专利"]) if r[0]
    ]

    # 中国专利：编号|专利号|名称|发明人|授权年|授权月|授权日|专利类型|法律状态|备注
    data["patents_cn"] = [
        {
            "num": r[0], "pid": r[1], "name": r[2], "inventors": r[3],
            "year": r[4], "month": r[5], "day": r[6],
            "ptype": r[7], "status": r[8],
            "note": r[9] if len(r) > 9 else "",
        }
        for r in rows_of(wb["中国专利"]) if r[0]
    ]

    return data


def main():
    if not XLSX.exists():
        sys.exit(f"未找到数据文件: {XLSX}")
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    data = build(wb)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
        f.write("\n")
    print(f"已生成 {OUT}")
    for k, v in data.items():
        print(f"  {k}: {len(v)} 条")


if __name__ == "__main__":
    main()
