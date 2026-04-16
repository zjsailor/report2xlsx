# -*- coding: utf-8 -*-
"""
PDF转MD转换器 - 用于PGT-M胚胎植入前遗传学检测报告
将PDF格式的报告转换为Markdown格式，便于后续解析

依赖安装:
    pip install pdfplumber

用法:
    python pdf_to_md.py -i <pdf_folder> [-o <output_folder>]

示例:
    python pdf_to_md.py -i D:\\md2excel -o D:\\md2excel\\md_output
    python pdf_to_md.py -i D:\\md2excel  # 使用默认输出文件夹
"""

import os
import sys
import re
import pdfplumber
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")


def clean_text(text):
    """清理文本，移除多余空白"""
    if not text:
        return ""
    text = re.sub(r"\s+", " ", str(text))
    return text.strip()


def is_table_row(line):
    """判断是否为表格行"""
    if "|" not in line:
        return False
    parts = [p.strip() for p in line.split("|")]
    parts = [p for p in parts if p]
    # 表格行通常有多个部分
    return len(parts) >= 3


def extract_page_content(page):
    """从单页提取内容"""
    content = {"text_lines": [], "tables": []}

    # 提取文本
    text = page.extract_text()
    if text:
        for line in text.split("\n"):
            line = clean_text(line)
            if line:
                content["text_lines"].append(line)

    # 提取表格
    tables = page.extract_tables()
    for table in tables:
        if table and len(table) > 0:
            md_table = []
            for row in table:
                if row:
                    cleaned_row = [clean_text(cell) if cell else "" for cell in row]
                    # 过滤空行
                    if any(cell for cell in cleaned_row):
                        md_table.append(cleaned_row)
            if md_table:
                content["tables"].append(md_table)

    return content


def format_as_md_table(table):
    """将表格格式化为MD格式"""
    if not table:
        return []

    md_lines = []
    for i, row in enumerate(table):
        md_row = "|" + "|".join(str(cell) for cell in row) + "|"
        md_lines.append(md_row)

        # 添加分隔行（第二行）
        if i == 0:
            col_count = len(row)
            separator = "|" + "|".join(["---"] * col_count) + "|"
            md_lines.append(separator)

    return md_lines


def pdf_to_md(pdf_path, output_path):
    """将单个PDF转换为MD"""

    md_lines = []

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        patient_name = ""

        for page_num, page in enumerate(pdf.pages, 1):
            content = extract_page_content(page)

            # 提取患者姓名（通常在第一页）
            if page_num == 1:
                for line in content["text_lines"]:
                    if "受检者姓名" in line:
                        match = re.search(r"受检者姓名[：:]\s*(\S+)", line)
                        if match:
                            patient_name = match.group(1)
                        break

            # 处理文本行
            in_detection_content = False
            in_attachment = False

            for line in content["text_lines"]:
                # 跳过页码和版本信息
                if re.match(r"^第\s*\d+\s*页", line):
                    continue
                if re.match(r"^版本号", line):
                    continue
                if re.match(r"^官网：", line):
                    continue
                if re.match(r"^地址：", line):
                    continue
                if re.match(r"^上海亿康", line):
                    continue

                # 处理标题
                if "胚胎植入前遗传学检测报告" in line:
                    md_lines.append(f"# {line}")
                elif re.match(r"^附件[一二三四五六]", line):
                    md_lines.append(f"\n## {line}")
                elif "检测局限性说明" in line:
                    md_lines.append(f"\n### {line}")
                elif "结果说明" in line and len(line) < 20:
                    md_lines.append(f"\n#### {line}")
                elif "目标变异检测结果" in line:
                    md_lines.append(f"\n##### {line}")
                    in_detection_content = True
                elif "SNP可用位点统计" in line:
                    md_lines.append(f"\n##### {line}")
                elif "SNP单体型分型图谱" in line:
                    md_lines.append(f"\n##### {line}")
                elif "位点验证图谱" in line:
                    md_lines.append(f"\n##### {line}")
                elif "检测结果注释" in line:
                    md_lines.append(f"\n### {line}")

                # 处理基本信息行 - 关键字段提取
                elif "女方姓名" in line or "男方姓名" in line:
                    # 保持原始格式，便于后续解析
                    md_lines.append(f"|{line}|")
                elif "年 龄" in line or "年龄" in line:
                    # 年龄行
                    md_lines.append(f"|{line}|")
                elif "收样日期" in line:
                    md_lines.append(f"|{line}|")
                elif "送检编号" in line:
                    md_lines.append(f"|{line}|")
                elif "送检条码" in line:
                    md_lines.append(f"|{line}|")

                # 处理检测内容中的变异信息
                elif "基因名称" in line and "变异位置" in line:
                    md_lines.append(f"|{line}|")
                elif "疾病名称" in line:
                    md_lines.append(f"|{line}|")

                # 处理胚胎检测结果表头
                elif any(
                    x in line
                    for x in ["样本名称", "形态学", "CNV检测结果", "异倍体", "携带状态"]
                ):
                    md_lines.append(f"|{line}|")
                elif "评级" in line and "结果解释" in line:
                    md_lines.append(f"|{line}|")

                # 跳过无意义的分隔线
                elif line == "---" or line == "——" or line == "——" * 10:
                    continue

                # 普通文本
                else:
                    # 检查是否是表格行（用|分隔）
                    if "|" in line:
                        parts = [p.strip() for p in line.split("|")]
                        parts = [p for p in parts if p]
                        if len(parts) >= 2:
                            md_lines.append("|" + "|".join(parts) + "|")
                            continue
                    md_lines.append(line)

            # 处理表格 - 添加到MD
            for table in content["tables"]:
                md_table = format_as_md_table(table)
                if md_table:
                    md_lines.append("")
                    md_lines.extend(md_table)
                    md_lines.append("")

    # 写入文件
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))

    return patient_name, len(md_lines)


def process_folder(pdf_folder, output_folder):
    """处理文件夹中的所有PDF文件"""

    pdf_folder = Path(pdf_folder)
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)

    pdf_files = list(pdf_folder.glob("*.pdf"))

    if not pdf_files:
        print(f"在 {pdf_folder} 中未找到PDF文件")
        return

    # 过滤掉非PDF文件
    pdf_files = [f for f in pdf_files if f.suffix.lower() == ".pdf"]

    print(f"找到 {len(pdf_files)} 个PDF文件")
    print(f"输出到: {output_folder}")
    print("-" * 60)

    success_count = 0
    error_count = 0
    results = []

    for pdf_file in sorted(pdf_files):
        try:
            output_file = output_folder / (pdf_file.stem + ".md")
            patient_name, line_count = pdf_to_md(pdf_file, output_file)
            print(f"[OK] {pdf_file.name}")
            if patient_name:
                print(f"    患者: {patient_name}, 行数: {line_count}")
            else:
                print(f"    行数: {line_count}")
            success_count += 1
            results.append((pdf_file.name, patient_name, "success"))
        except Exception as e:
            print(f"[FAIL] {pdf_file.name}: {e}")
            error_count += 1
            results.append((pdf_file.name, "", str(e)))

    print("-" * 60)
    print(f"完成: {success_count} 成功, {error_count} 失败")

    return results


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="PDF转MD转换器 - 将PGT-M胚胎检测报告PDF转换为Markdown格式"
    )
    parser.add_argument(
        "-i",
        "--input",
        default=r"D:\md2excel",
        help="PDF文件所在文件夹 (默认: D:\\md2excel)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=r"D:\md2excel\md_output",
        help="MD输出文件夹 (默认: D:\\md2excel\\md_output)",
    )
    args = parser.parse_args()

    print(f"PDF文件夹: {args.input}")
    print(f"MD输出文件夹: {args.output}")
    print()

    process_folder(args.input, args.output)
