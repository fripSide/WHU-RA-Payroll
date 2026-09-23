# 学生劳务费发放助手

上传 Excel 名单 → 填金额 → 填信息 → 导出 Word / PDF / Excel。

## 启动

```bash
./run.sh
```

浏览器会自动打开 http://127.0.0.1:8848 （端口被占用时自动换下一个）。

## 使用步骤

1. **上传 Excel** — 需包含"学号"、"姓名"、"金额"三列，列名自动识别，顺序无关。带金额列的话会一起读进来。
2. **设置金额** — 表格里直接填，学号、姓名、学院也可改。
3. **填写信息** — 先选项目，项目类型、编号、单位随之自动带出；兼职时段默认当前月份。
4. **导出** — 同时生成 Word、PDF、Excel，存到 `~/Downloads`。

## 导出文件

文件名格式 `{名单名}_{月份}_劳务费明细`。名单名取自上传的 Excel 文件名，同一个项目发多次靠它区分批次；月份取自"兼职时段"。

| 文件 | 用途 |
|------|------|
| `1-basic_9月_劳务费明细.docx` | Word 明细表，需要手改时改这份 |
| `1-basic_9月_劳务费明细.pdf` | 报销提交用 |
| `1-basic_9月_劳务费名单.xls` | 学号、姓名、金额，用于上传到系统中生成发放列表 |

PDF 由 `src/pdf_gen.py` 基于 reportlab 生成，不需要装 Office。它是从自动生成的那份 Word 转出来的，**手改 Word 之后要重新生成对应 PDF**。

非科研模板会在合计行上一行自动写入发放事由，格式为 `发放{第一个人姓名}等{总人数}人，参与组内科研项目的科研补助。明细如上表。`

## 目录结构

```
run.sh              启动入口
src/                代码
  server.py           HTTP 服务（上传 / 导出 / 下载）
  store.py            路径常量、文件名清洗、底板选择
  docx_gen.py         生成 Word 明细表
  pdf_gen.py          由 docx 生成 PDF（reportlab）
  xls_gen.py          读写 .xls / .xlsx 名单
  bootstrap.py        依赖引导
data/               项目库 projects.json、人员与历史发放 people.json
templates/          两份官方 Word 底板 + 上传表底板
web/index.html      界面（单文件）
vendor/             随仓库自带的依赖
docs/开发记录.md     需求与历史踩坑记录
```

数据目录和导出目录可用环境变量改：`BAOXIAO_DATA_DIR`、`BAOXIAO_OUTPUT_DIR`。

## 环境

Python 3.7+，Windows / macOS / Linux 均可。依赖（python-docx、reportlab、xlrd、xlwt、openpyxl）已随仓库放在 `vendor/` 下，启动时 `src/bootstrap.py` 自动补齐缺失的包。想装到自己环境：`pip install -r requirements.txt`。

## 常见问题

**Excel 列名怎么写？** 学号（学号 / 学生学号 / studentid）、姓名（姓名 / 学生姓名 / name）、学院（学院 / 所在学院 / 院系）、金额（金额 / 津贴 / 助研津贴，可选）。没有学院列时自动填"国家网络安全学院"。

**能导入之前导出的 Excel 吗？** 能，导出的 Excel 带金额，再次导入会自动填上。

**PDF 报错说含图片？** `pdf_gen.py` 只支持文字和表格底板。自定义底板里有图片要先移除，或者用 Word 另存为 PDF。内置的两个模板没这个问题。
