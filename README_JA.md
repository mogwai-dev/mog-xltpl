git clone https://github.com/[YOUR-USERNAME]/mog-xltpl.git
# mog-xltpl

Excel テンプレートから xls/x ファイルを生成する Windows 専用フォーク版です（VBA画像書式を完全に保存）。

[English](README.md) | [中文](README_ZH.md)

> ** 重要：Windows専用版**
> [xltpl](https://github.com/zhangyu836/xltpl) のフォークで、Excel COM API を使い VBA/画像/書式を完全保持します。
> Linux/Mac では動作しません。クロスプラットフォームが必要な場合はオリジナル版をご利用ください。

## オリジナル版との主な違い
- VBA マクロを完全保存（xlsm/xltm）
- テンプレートの画像図形オブジェクトを保持
- 条件付き書式データ検証など複雑な書式を保持
- `{% img %}` タグ：プレースホルダーなし挿入、スペース区切り/カンマ区切り両対応
- Windows 専用（pywin32 + Excel COM）、Excel インストール必須

## インストール
要件: Windows / Microsoft Excel / Python 3.8+ / pywin32>=311

```
pip install mog-xltpl
```

開発版:
```
git clone https://github.com/[YOUR-USERNAME]/mog-xltpl.git
cd mog-xltpl
pip install -e .[test]
```

uv を使う場合:
```
uv venv
uv pip install -e .[test]
uv run pytest
```

## 使い方（CLI）
```
uv run xltpl template.xlsx output.xlsx vars.yaml
uv run xltpl template.xlsx output.xlsx vars.yaml --highlight-output
uv run xltpl template.xlsx output.xlsx vars.yaml --highlight-output --highlight-color FFFF9999
```

### Taskfile 連携例
```
# Taskfile.yml
version: '3'

vars:
  DOC_TYPE: invoice
  DATE: "2025-12-30"
  NAME: "山田太郎"

tasks:
  render:
    cmds:
      - xltpl templates/{{.DOC_TYPE}}.xlsx output/result.xlsx vars.yaml
```

### vars（YAML）の例
```
vars:
  doc_type: "invoice"
  date: "2025-12-30"
  name: "山田太郎"
  items:
    - name: "商品A"
      price: 1000
    - name: "商品B"
      price: 2000
```

## 画像挿入
フィルター:
```
{{ image_path | img(120, 140) }}
{{ image_path | img(width=120, height=140) }}
```
タグ（プレースホルダー不要）:
```
{% img image_path 120 140 %}   {# スペース区切り #}
{% img image_path, 120, 140 %} {# カンマ区切り #}
```

## よく使うフィルター
- `sha256`: `{{ file | sha256 }}`
- `mtime`: `{{ file | mtime('%Y-%m-%d') }}`
- `to_fullwidth`: 半角数字と `-` を全角に

## 簡単な Python 例
```
from xltpl.writerx import BookWriter

writer = BookWriter("tpl.xlsx")
payloads = [{"name": "Hello Wizard", "items": ["1"] * 8}]
writer.render_book(payloads)
writer.save("result.xlsx")
```

## なぜ COM 保存か
- openpyxl 保存では画像図形一部名前空間が失われることがある
- COM 保存なら VBA画像書式名前空間を保持できる

## ライセンス / クレジット
- MIT License
- Original: Zhang Yu / https://github.com/zhangyu836/xltpl

**現在の方法（pywin32 + COM API）：**
1. openpyxl でテンプレートを読み込み（読み取り専用、Jinja2 レンダリング用）
2. Excel COM API でテンプレートのコピーを開く
3. レンダリングされたデータからセル値のみを更新
4. COM API で保存 → すべての画像、図形、フォーマットが保持される

### 再現性のテスト

画像やフォーマットが保持されているか検証：

```bash
# 画像を含むテストテンプレートを作成
# static_image.xlsm を例として使用

# レンダリングを実行
xltpl static_image.xlsm static_image_out.xlsm static_image.yaml

# ファイルの整合性を検証
python -c "
import zipfile
z = zipfile.ZipFile('static_image_out.xlsm')
images = [n for n in z.namelist() if 'image' in n.lower() or 'drawing' in n.lower()]
print(f'保持された画像/図形: {len(images)}')
for img in images:
    print(f'  {img}')
"

# テンプレートと比較
python -c "
import zipfile
t = zipfile.ZipFile('static_image.xlsm')
o = zipfile.ZipFile('static_image_out.xlsm')
t_imgs = set([n for n in t.namelist() if 'drawing' in n or 'media' in n])
o_imgs = set([n for n in o.namelist() if 'drawing' in n or 'media' in n])
print(f'テンプレート: {len(t_imgs)} 個の画像関連ファイル')
print(f'出力: {len(o_imgs)} 個の画像関連ファイル')
print(f'一致: {t_imgs == o_imgs}')
"

# 名前空間が保持されているか検証
python -c "
import zipfile
o = zipfile.ZipFile('static_image_out.xlsm')
sheet_xml = o.read('xl/worksheets/sheet1.xml').decode('utf-8')
preserved = all([
    'xmlns:mc' in sheet_xml,
    'mc:Ignorable' in sheet_xml,
    'xr:uid' in sheet_xml
])
print(f'Excel 名前空間が保持されています: {preserved}')
"
```

### 並列実行の安全性

ツールはスレッドロックを使用して、並行する Excel COM 操作を防ぎます：
- 複数の xltpl プロセスを並列実行可能（Taskfile 経由など）
- 各プロセスは Excel COM を使用する前にロックを取得
- COM の競合を防ぎ、安定性を確保

Taskfile での並列実行例：

```yaml
tasks:
  process-all:
    deps:
      - task: process-file-1  # 並列実行
      - task: process-file-2  # 並列実行
    cmds:
      - echo "すべてのファイルを処理しました"
```

### 結合セルの扱い

COM API アプローチを使用すると、**テンプレート内の結合セルは自動的に保持されます**。ツールはセルの値のみを更新し、結合状態、スタイル、その他の属性は変更しません。つまり：
- ✅ 結合セルは結合されたまま保持
- ✅ セルのスタイルと罫線はそのまま
- ✅ セルの内容のみが更新される

## ライセンス
MIT License
