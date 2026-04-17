# Report2xlsx

PGT-M 胚胎植入前遗传学检测报告数据提取工具

## 功能概述

将 PGT-M 胚胎检测报告从 PDF 转换为 Markdown 格式，再解析为 Excel 表格。

## 目录结构

```
report2xlsx/
├── pdf_to_md.py       # PDF 转 MD 转换器
├── parse_reports.py   # MD 解析写入 Excel
├── md_output/         # MD 文件输出目录
├── Info.xlsx          # 最终 Excel 输出
└── README.md          # 说明文档
```

## 环境要求

- Python 3.8+
- 依赖包：`pdfplumber`, `openpyxl`

## 安装依赖

```bash
pip install pdfplumber openpyxl
```

## 使用方法

### 1. PDF 转 MD

```bash
python pdf_to_md.py -i <pdf文件夹> [-o <输出文件夹>]

# 示例
python pdf_to_md.py -i D:/md2excel -o D:/md2excel/md_output
```

### 2. MD 解析写入 Excel

```bash
python parse_reports.py -i <md文件夹> [-o <输出excel>]

# 示例
python parse_reports.py -i D:/md2excel/md_output -o D:/md2excel/Info.xlsx
```

## 输出字段说明

| 列号 | 字段名 | 说明 |
|------|---------|------|
| 1 | 文件名 | PDF 文件名 |
| 2 | 送检编号 | 报告送检编号 |
| 3 | 送检条码 | 报告送检条码 |
| 4 | 收样日期 | 样本接收日期 |
| 5 | 女方姓名 | 患者姓名 |
| 6 | 女方年龄 | 年龄 |
| 7 | 男方姓名 | 配偶姓名 |
| 8 | 男方年龄 | 配偶年龄 |
| 9 | 疾病 | 遗传疾病名称 |
| 10 | 基因 | 致病基因 |
| 11 | 突变1 | 第一个突变位点 |
| 12 | 突变2 | 第二个突变位点（如有） |
| 13 | 胚胎编号 | 胚胎样本编号 |
| 14 | 形态学评级 | 胚胎形态学评分 |
| 15 | CNV检测结果 | 拷贝数变异检测结果 |
| 16 | CNV结果解释 | CNV 结果说明 |
| 17 | 异倍体检测结果 | 染色体非整倍体检测 |
| 18 | 携带状态 | 致病基因携带状态 |
| 19 | 目标变异检测结果 | S列 - 第一个突变检测结果 |
| 20 | 变异2相关信息 | T列 - AR双突变时第二个结果 |
| 21 | 目标变异/SNP分型一致性 | U列 - SNP分型一致性结果 |

## 注意事项

- 跳过预实验报告（包含 `_PGTMF_` 的文件）
- 样本名称格式多样，自动识别胚胎编号
- wangting等特殊格式报告已做兼容处理

## License

MIT
