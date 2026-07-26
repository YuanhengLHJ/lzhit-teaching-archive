#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
教学专家分析报告生成脚本（lzhit-teaching-archive）

基于成绩数据与试卷分析，按高校教学专家视角生成《教学专家分析报告.md》。
分析框架见 references/expert-analysis.md。

可作为模块导入（generate 函数），也可独立运行。

用法:
    python expert_report.py --data-file grades.csv \\
        --course-name 高等数学 --semester 2025-2026-2 \\
        --course-type T1 --teacher 张老师 --class 计科2301 \\
        --output ./教学专家分析报告.md

依赖: 无（纯统计计算，可选 numpy 用于偏度/峰度）
"""

import argparse
import datetime as _dt
import statistics
import sys
from pathlib import Path

AI_NOTE = "本材料由AI辅助生成，仅供参考，请认真核对后由相关负责人签字确认。"


# --------------------------------------------------------------------------- #
# 数据读取
# --------------------------------------------------------------------------- #
def read_data(path: str) -> list:
    if not path:
        return []
    p = Path(path)
    if not p.exists():
        return []
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        return _read_xlsx(p)
    import csv
    with open(p, "r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def _read_xlsx(path: Path) -> list:
    try:
        import openpyxl
    except ImportError:
        return []
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(h).strip() if h is not None else f"col{i}" for i, h in enumerate(rows[0])]
    return [dict(zip(header, r)) for r in rows[1:]]


# --------------------------------------------------------------------------- #
# 统计计算
# --------------------------------------------------------------------------- #
def extract_scores(rows: list, col: str = "总评") -> list:
    scores = []
    for r in rows:
        v = r.get(col)
        if v is None or v == "":
            continue
        try:
            scores.append(float(v))
        except (ValueError, TypeError):
            continue
    return scores


def descriptive_stats(scores: list) -> dict:
    if not scores:
        return {"n": 0}
    n = len(scores)
    mean = statistics.mean(scores)
    std = statistics.stdev(scores) if n > 1 else 0
    median = statistics.median(scores)
    passed = sum(1 for s in scores if s >= 60)
    excellent = sum(1 for s in scores if s >= 85)
    failed = sum(1 for s in scores if s < 60)
    return {
        "n": n,
        "mean": round(mean, 2),
        "std": round(std, 2),
        "median": round(median, 2),
        "max": max(scores),
        "min": min(scores),
        "pass_rate": round(passed / n * 100, 1),
        "excellent_rate": round(excellent / n * 100, 1),
        "fail_rate": round(failed / n * 100, 1),
    }


def distribution_shape(scores: list) -> dict:
    """判断分布形态。"""
    if len(scores) < 5:
        return {"shape": "样本不足", "skewness": None, "kurtosis": None}
    n = len(scores)
    mean = statistics.mean(scores)
    std = statistics.stdev(scores) if n > 1 else 1
    if std == 0:
        return {"shape": "无变异", "skewness": 0, "kurtosis": None}
    # 偏度
    skew = sum((x - mean) ** 3 for x in scores) / (n * std ** 3)
    # 峰度（简化）
    kurt = sum((x - mean) ** 4 for x in scores) / (n * std ** 4) - 3
    # 形态判断
    if abs(skew) < 0.5:
        shape = "近似正态分布"
    elif skew > 0.5:
        shape = "正偏（左偏，低分较多）"
    else:
        shape = "负偏（右偏，高分较多）"
    # 双峰检测（简化：看直方图分布）
    bins = [0] * 10
    for s in scores:
        idx = min(int(s // 10), 9)
        bins[idx] += 1
    peaks = sum(1 for i in range(1, 9) if bins[i] > bins[i - 1] and bins[i] > bins[i + 1])
    if peaks >= 2 and max(bins) > n * 0.2:
        shape = "双峰或多峰分布（学生水平分化）"
    return {"shape": shape, "skewness": round(skew, 2), "kurtosis": round(kurt, 2), "histogram": bins}


def abnormal_scores(scores: list) -> dict:
    high = [s for s in scores if s >= 98]
    low = [s for s in scores if s <= 20]
    return {"abnormal_high": high, "abnormal_low": low}


# --------------------------------------------------------------------------- #
# 教学质量判断
# --------------------------------------------------------------------------- #
def judge_teaching(stats: dict, shape: dict) -> str:
    if stats.get("n", 0) == 0:
        return "数据不足，无法判断。"
    parts = []
    pr = stats["pass_rate"]
    sh = shape["shape"]
    if pr >= 90 and "正态" in sh:
        parts.append("教学效果良好：及格率高且成绩分布近似正态，题目区分度合理。")
    elif pr < 60:
        parts.append("教学效果需关注：及格率偏低，建议排查原因（题目偏难/教学薄弱环节/学生基础）。")
    elif "双峰" in sh:
        parts.append("学生水平分化明显（双峰分布），建议分层教学或针对薄弱群体加强辅导。")
    elif "正偏" in sh:
        parts.append("成绩偏低分集中，可能题目偏难或教学效果不佳，建议复盘重点章节教学。")
    elif "负偏" in sh:
        parts.append("成绩偏高分集中，题目区分度可能不足，建议增加区分度高的题目。")
    else:
        parts.append("教学效果总体正常，建议持续优化。")
    return " ".join(parts)


def weak_points(rows: list) -> list:
    """识别薄弱知识点（基于各题得分率，若有题号列）。"""
    # 简化：若数据含题号列（Q1/Q2/题1/题2...），计算各题得分率
    if not rows:
        return []
    sample = rows[0]
    q_cols = [k for k in sample if any(p in str(k) for p in ("题", "Q", "q")) and "题号" not in str(k)]
    weak = []
    for col in q_cols:
        vals = []
        for r in rows:
            v = r.get(col)
            try:
                vals.append(float(v))
            except (ValueError, TypeError):
                continue
        if vals:
            rate = sum(vals) / len(vals)
            if rate < 60:
                weak.append({"题目": col, "平均得分": round(rate, 1), "建议": "得分率偏低，对应知识点需加强"})
    return weak


def suggestions(stats: dict, shape: dict, weak: list) -> list:
    s = []
    if weak:
        s.append("【短期】针对薄弱知识点（" + "、".join(w["题目"] for w in weak) + "）增加练习与讲解。")
    if stats.get("fail_rate", 0) > 20:
        s.append("【短期】不及格率较高，建议组织辅导答疑，关注后进生。")
    if "双峰" in shape.get("shape", ""):
        s.append("【长期】学生水平分化，建议探索分层教学或差异化作业。")
    if "负偏" in shape.get("shape", ""):
        s.append("【长期】题目区分度不足，建议优化题型搭配，增加中等难度题。")
    s.append("【长期】持续收集学生反馈，迭代教学设计。")
    return s


# --------------------------------------------------------------------------- #
# 报告生成
# --------------------------------------------------------------------------- #
def generate(course, data_rows, course_type_name="理论课（考试）", exam_data=None):
    """生成专家分析报告，返回 (report_text, stats_dict)。"""
    scores = extract_scores(data_rows, "总评") or extract_scores(data_rows, "期末")
    stats = descriptive_stats(scores)
    shape = distribution_shape(scores)
    abn = abnormal_scores(scores)
    weak = weak_points(data_rows)
    teaching_judge = judge_teaching(stats, shape)
    sugg = suggestions(stats, shape, weak)

    lines = ["# 教学专家分析报告", ""]
    lines.append("> 本报告由 lzhit-teaching-archive 技能以高校教学专家视角生成，"
                 "基于课程归档材料与成绩数据，仅供参考。")
    lines.append("")
    lines.append("## 课程信息")
    for k, v in course.items():
        lines.append(f"- {k}: {v}")
    lines.append(f"- 课程类型: {course_type_name}")
    lines.append(f"- 学生人数: {stats.get('n', 0)}")
    lines.append(f"- 生成时间: {_dt.datetime.now().isoformat(timespec='seconds')}")
    lines.append("")

    lines.append("## 一、题目设计合理性分析")
    lines.append("> 注：题目设计分析依赖试卷题目数据。若未提供试卷，本节基于成绩分布间接推断。")
    lines.append("")
    if exam_data:
        lines.append("### 1.1 覆盖度")
        lines.append("- 待结合试卷题目与大纲知识点分析。")
        lines.append("### 1.2 难度分布")
        lines.append("- 待结合试卷题目分析。")
    else:
        lines.append("### 1.1–1.4 题目设计")
        lines.append("- 未提供试卷题目数据，无法直接分析覆盖度与难度分布。")
        lines.append(f"- 间接推断：成绩分布形态为「{shape['shape']}」，"
                     f"平均分 {stats.get('mean', 'N/A')}，可结合经验判断题目难度。")
    lines.append("")

    lines.append("## 二、考试成绩分布分析")
    if stats.get("n", 0) == 0:
        lines.append("- 数据不足，无法分析。")
    else:
        lines.append("### 2.1 描述性统计")
        lines.append(f"- 学生人数: {stats['n']}")
        lines.append(f"- 平均分: {stats['mean']}")
        lines.append(f"- 标准差: {stats['std']}")
        lines.append(f"- 中位数: {stats['median']}")
        lines.append(f"- 最高分: {stats['max']}")
        lines.append(f"- 最低分: {stats['min']}")
        lines.append(f"- 及格率: {stats['pass_rate']}%")
        lines.append(f"- 优秀率: {stats['excellent_rate']}%")
        lines.append(f"- 不及格率: {stats['fail_rate']}%")
        lines.append("")
        lines.append("### 2.2 分布形态")
        lines.append(f"- 形态: {shape['shape']}")
        if shape.get("skewness") is not None:
            lines.append(f"- 偏度: {shape['skewness']}")
            lines.append(f"- 峰度: {shape.get('kurtosis', 'N/A')}")
        lines.append("")
        lines.append("### 2.3 异常分数")
        lines.append(f"- 异常高分(≥98): {len(abn['abnormal_high'])} 个 {abn['abnormal_high'][:5]}")
        lines.append(f"- 异常低分(≤20): {len(abn['abnormal_low'])} 个 {abn['abnormal_low'][:5]}")
    lines.append("")

    lines.append("## 三、教学质量判断与建议")
    lines.append("### 3.1 教学效果")
    lines.append(f"- {teaching_judge}")
    lines.append("")
    lines.append("### 3.2 薄弱环节")
    if weak:
        for w in weak:
            lines.append(f"- {w['题目']}：平均得分 {w['平均得分']}，{w['建议']}")
    else:
        lines.append("- 未识别到明显薄弱环节（或未提供题目级数据）。")
    lines.append("")
    lines.append("### 3.3 改进建议")
    for sg in sugg:
        lines.append(f"- {sg}")
    lines.append("")

    lines.append("## 四、总结")
    if stats.get("n", 0) > 0:
        lines.append(f"本课程共 {stats['n']} 名学生参评，平均分 {stats['mean']}，"
                     f"及格率 {stats['pass_rate']}%，成绩分布呈「{shape['shape']}」。"
                     f"{teaching_judge}")
    else:
        lines.append("数据不足，建议补充成绩数据后重新分析。")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append(AI_NOTE)
    return "\n".join(lines), {"stats": stats, "shape": shape, "weak": weak}


# --------------------------------------------------------------------------- #
# 命令行入口
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="生成教学专家分析报告")
    ap.add_argument("--data-file", required=True, help="成绩数据文件（CSV/xlsx）")
    ap.add_argument("--course-name", default="", help="课程名")
    ap.add_argument("--semester", default="", help="学期")
    ap.add_argument("--course-type", default="T1", help="课程类型代码")
    ap.add_argument("--teacher", default="", help="任课教师")
    ap.add_argument("--class", dest="klass", default="", help="班级")
    ap.add_argument("--output", help="报告输出路径")
    args = ap.parse_args()

    data_rows = read_data(args.data_file)
    type_names = {"T1": "理论课（考试）", "T2": "理论课（考查）", "T3": "课程设计",
                  "T4": "实训", "T5": "实习"}
    course = {"课程名": args.course_name, "学期": args.semester,
              "任课教师": args.teacher, "班级": args.klass}

    report_text, _ = generate(course, data_rows, type_names.get(args.course_type, "理论课（考试）"))

    out = Path(args.output) if args.output else Path("教学专家分析报告.md")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(report_text, encoding="utf-8")
    print(f"✅ 专家分析报告已生成: {out}")


if __name__ == "__main__":
    main()
