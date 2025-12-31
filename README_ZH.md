# mog-xltpl

基于 Excel COM 的 Windows 专用分支，生成并保存 .xlsx/.xlsm 模板，完整保留 VBA、图片和复杂格式。

[English](README_EN.md) | [日本語](README_JA.md)

> **重要**：仅支持 Windows + 已安装的 Microsoft Excel。**不支持 .xls (BIFF8)**，请先转换为 .xlsx/.xlsm/.xltx/.xltm，CLI 传入 .xls 会直接报错退出。

## 与原版的差异
- 保留 VBA 宏（.xlsm/.xltm）
- 模板中的图片、图形、形状完整保留
- 条件格式、数据验证等复杂格式不会丢失
- `{% img %}` 标签可无占位符插入图片（支持空格或逗号分隔）
- 使用 Excel COM 保存，确保命名空间/格式不被 openpyxl 擦除

## 环境与安装
- Windows + 已安装 Excel
- Python 3.8+
- pywin32 >= 311
- 支持格式：.xlsx/.xlsm/.xltx/.xltm（不支持 .xls）

```shell
pip install mog-xltpl
```

开发版：
```shell
git clone https://github.com/[YOUR-USERNAME]/mog-xltpl.git
cd mog-xltpl
pip install -e .[test]
```

使用 uv：
```shell
uv venv
uv pip install -e .[test]
uv run pytest
```

## CLI 用法
```shell
uv run xltpl template.xlsx output.xlsx vars.yaml
uv run xltpl template.xlsx output.xlsx vars.yaml --highlight-output
uv run xltpl template.xlsx output.xlsx vars.yaml --highlight-output --highlight-color FFFF9999
```

## YAML 变量示例
```yaml
vars:
  doc_type: "invoice"
  date: "2025-12-30"
  name: "张三"
  items:
    - name: "产品A"
      price: 1000
    - name: "产品B"
      price: 2000
```

## 图片插入
过滤器：
```jinja2
{{ image_path | img(120, 140) }}
{{ image_path | img(width=120, height=140) }}
```
标签（无需占位符）：
```jinja2
{% img image_path 120 140 %}
{% img image_path, 120, 140 %}
```

## 常用过滤器
- `sha256`: `{{ file | sha256 }}`
- `mtime`: `{{ file | mtime('%Y-%m-%d') }}`
- `to_fullwidth`: 半角数字和 `-` 转全角

## 简单 Python 示例
```python
from xltpl.writerx import BookWriter

writer = BookWriter("tpl.xlsx")
payloads = [{"name": "Hello Wizard", "items": ["1"] * 8}]
writer.render_book(payloads)
writer.save("result.xlsx")
```

## 为什么用 COM 保存
- openpyxl 保存可能丢失图片/命名空间
- COM 保存可保留 VBA、图片和格式命名空间

## 支持与限制
- 支持：.xlsx/.xlsm/.xltx/.xltm
- 不支持：.xls（请先转换）
- 功能：合并单元格、图片 (`{% img %}`)、数据验证、自动筛选、非字符串单元值 (`{% xv %}`)

## 许可与致谢
- MIT License
- 原项目：Zhang Yu / https://github.com/zhangyu836/xltpl
