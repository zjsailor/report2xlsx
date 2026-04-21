# Report2xlsx

PGT-M 胚胎植入前遗传学检测报告数据提取工具

## 功能概述

将 PGT-M 胚胎检测报告从 PDF 转换为 Markdown 格式，再解析为 Excel 表格。

## 目录结构

```
report2xlsx/
├── pdf_to_md.py           # PDF 转 MD 转换器
├── parse_reports.py        # MD 解析写入 Excel
├── fix_all_dual_gene.py    # 修复双基因报告的 U/V 列
├── markdown/               # MD 文件输出目录
├── Info100smp.xlsx         # Excel 输出
└── README.md               # 说明文档
```

## 环境要求

- Python 3.8+
- 依赖包：`pdfplumber`, `openpyxl`

## 安装依赖

```bash
pip install pdfplumber openpyxl
```

## 工作流程

### Step 1: PDF 转 MD

```bash
python pdf_to_md.py -i <pdf文件夹> -o <md输出文件夹>

# 示例
python pdf_to_md.py -i D:/md2excel/xh -o D:/md2excel/markdown
```

### Step 2: MD 转 Excel

```bash
python parse_reports.py -i <md文件夹> -o <输出excel>

# 示例
python parse_reports.py -i D:/md2excel/markdown -o D:/md2excel/Info100smp.xlsx
```

### Step 3: 修复双基因报告 U/V 列

```bash
python fix_all_dual_gene.py -i <md文件夹> -o <输出excel>

# 示例
python fix_all_dual_gene.py -i D:/md2excel/markdown -o D:/md2excel/Info100smp.xlsx
```

**注意**：此脚本用于修复双基因突变报告（如 ABCD1+GJB2）的 U列 和 V列。即使 Excel 中已有值，脚本也会用 MD 文件中的正确数据覆盖修正。

## 输出字段说明（22列）

| 列号 | 字段名 | 说明 |
|------|--------|------|
| 1 | 文件名 | PDF 文件名 |
| 2 | 送检编号 | 报告送检编号 |
| 3 | 送检条码 | 报告送检条码 |
| 4 | 收样日期 | 样本接收日期 |
| 5 | 女方姓名 | 患者姓名 |
| 6 | 女方年龄 | 年龄 |
| 7 | 男方姓名 | 配偶姓名 |
| 8 | 男方年龄 | 配偶年龄 |
| 9 | 疾病 | 遗传疾病名称 |
| 10 | 基因 | 致病基因（多个用逗号分隔） |
| 11 | 突变1 | 第一个突变位点 |
| 12 | 突变2 | 第二个突变位点（如有） |
| 13 | 胚胎编号 | 胚胎样本编号 |
| 14 | 形态学评级 | 胚胎形态学评分 |
| 15 | CNV检测结果 | 拷贝数变异检测结果 |
| 16 | CNV结果解释 | CNV 结果说明 |
| 17 | 异倍体检测结果 | 染色体非整倍体检测 |
| 18 | 携带状态 | 致病基因携带状态 |
| 19 | 目标变异1检测结果 | 第一个基因的突变检测结果 |
| 20 | 目标变异2相关信息 | 第二个基因的突变相关信息 |
| 21 | 目标变异1/SNP分型一致性 | U列 - 第一个基因的SNP分型一致性 |
| 22 | 目标变异2/SNP分型一致性 | V列 - 第二个基因的SNP分型一致性 |

## SNP分型一致性判断规则

| MD文件中显示 | Excel填入 |
|-------------|-----------|
| 一致 | 一致 |
| 不一致（位点扩增ADO） | 不一致（位点扩增ADO） |
| 不一致（检测异常） | 不一致（检测异常） |
| - (横杠/空白) | 不一致 |

## 注意事项

- 跳过预实验报告（包含 `_PGTMF_` 的文件）
- 样本名称格式多样，自动识别胚胎编号
- 双基因突变报告需运行 Step 3 修复 U/V 列

## License

MIT