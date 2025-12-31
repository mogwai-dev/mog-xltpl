# Fork Notice / フォーク通知

## English

This project (mog-xltpl) is a **Windows-only fork** of [xltpl](https://github.com/zhangyu836/xltpl) by Zhang Yu.

### Why a Fork?

The original xltpl is a cross-platform library that uses openpyxl for saving Excel files. While this works well for basic templates, it **loses important content** when saving:

- ❌ VBA macros are removed from xlsm/xltm files
- ❌ Template images and shapes disappear
- ❌ Complex formatting is not preserved
- ❌ Conditional formatting and data validation are lost

This fork addresses these issues by:

- ✅ Using **Excel COM API** for saving (Windows-only)
- ✅ **Preserving VBA macros** completely
- ✅ **Keeping all template images** and shapes
- ✅ **Maintaining complex formatting** and data validation
- ✅ Adding **{% img %} tag** support without placeholders (space-separated syntax)

### Key Differences from Original

| Feature | Original xltpl | mog-xltpl |
|---------|---------------|----------------|
| Platform | Cross-platform (Windows/Mac/Linux) | **Windows only** |
| Save method | openpyxl | **Excel COM API** |
| VBA preservation | ❌ Lost | ✅ Preserved |
| Template images | ❌ Lost | ✅ Preserved |
| Complex formatting | ❌ Lost | ✅ Preserved |
| {% img %} syntax | Comma-separated | **Space + comma** |
| Performance | Fast (~0.1s/file) | Slower (~3s/file) |
| Excel installation | Not required | **Required** |

### When to Use Which Version

**Use original xltpl if:**
- You need cross-platform support (Linux/Mac)
- You don't have VBA or complex formatting in templates
- You need fast performance

**Use mog-xltpl if:**
- You're on Windows with Excel installed
- You need to preserve VBA macros
- You need to keep template images and shapes
- You need complex formatting preserved

### Technical Implementation

This fork (mog-xltpl) uses a **two-stage approach**:

1. **Stage 1 (openpyxl)**: Load template and perform Jinja2 rendering
2. **Stage 2 (COM)**: Open original template via COM, apply rendered values, save via COM

This preserves everything in the original template while still using openpyxl's powerful rendering capabilities.

### Credits

Original author: **Zhang Yu** (zhangyu836@gmail.com)  
Original repository: https://github.com/zhangyu836/xltpl

This fork (mog-xltpl) is maintained separately with Windows-specific enhancements.

### License

MIT License (same as original)

---

## 日本語

このプロジェクトは Zhang Yu による [xltpl](https://github.com/zhangyu836/xltpl) の **Windows専用フォーク版**です。

### なぜフォークしたのか？

オリジナルのxltplはクロスプラットフォームのライブラリで、Excelファイルの保存にopenpyxlを使用しています。これは基本的なテンプレートでは問題なく動作しますが、保存時に**重要なコンテンツが失われます**：

- ❌ VBAマクロがxlsm/xltmファイルから削除される
- ❌ テンプレートの画像や図形が消える
- ❌ 複雑な書式が保持されない
- ❌ 条件付き書式やデータ検証が失われる

このフォークは以下の方法でこれらの問題を解決します：

- ✅ 保存に**Excel COM API**を使用（Windows専用）
- ✅ **VBAマクロを完全に保存**
- ✅ **すべてのテンプレート画像**と図形を保持
- ✅ **複雑な書式**とデータ検証を維持
- ✅ プレースホルダーなしで**{% img %}タグ**をサポート（スペース区切り構文）

### オリジナルとの主な違い

| 機能 | オリジナル xltpl | xltpl-preserve |
|------|-----------------|----------------|
| プラットフォーム | クロスプラットフォーム（Windows/Mac/Linux） | **Windows専用** |
| 保存方式 | openpyxl | **Excel COM API** |
| VBA保存 | ❌ 失われる | ✅ 保存される |
| テンプレート画像 | ❌ 失われる | ✅ 保存される |
| 複雑な書式 | ❌ 失われる | ✅ 保存される |
| {% img %}構文 | カンマ区切り | **スペース+カンマ** |
| パフォーマンス | 高速（~0.1秒/ファイル） | 低速（~3秒/ファイル） |
| Excelインストール | 不要 | **必須** |

### どちらのバージョンを使うべきか

**オリジナルxltplを使うべき場合：**
- クロスプラットフォームサポートが必要（Linux/Mac）
- テンプレートにVBAや複雑な書式がない
- 高速なパフォーマンスが必要

**xltpl-preserveを使うべき場合：**
- WindowsでExcelがインストールされている
- VBAマクロを保存する必要がある
- テンプレート画像や図形を保持する必要がある
- 複雑な書式を保持する必要がある

### 技術的な実装

このフォークは**2段階アプローチ**を使用します：

1. **ステージ1（openpyxl）**：テンプレートを読み込み、Jinja2レンダリングを実行
2. **ステージ2（COM）**：COM経由で元のテンプレートを開き、レンダリング結果を適用、COM経由で保存

これにより、openpyxlの強力なレンダリング機能を使いながら、元のテンプレートのすべてを保持します。

### クレジット

オリジナル作者：**Zhang Yu** (zhangyu836@gmail.com)  
オリジナルリポジトリ：https://github.com/zhangyu836/xltpl

このフォークはWindows固有の拡張を加えて独立して維持されています。

### ライセンス

MIT License（オリジナルと同じ）

---

## 中文

此项目是 Zhang Yu 开发的 [xltpl](https://github.com/zhangyu836/xltpl) 的 **Windows 专用分支版本**。

### 为什么创建分支？

原版 xltpl 是一个跨平台库，使用 openpyxl 保存 Excel 文件。虽然对于基本模板效果很好，但保存时会**丢失重要内容**：

- ❌ VBA 宏会从 xlsm/xltm 文件中删除
- ❌ 模板图片和图形会消失
- ❌ 复杂格式无法保留
- ❌ 条件格式和数据验证会丢失

此分支通过以下方式解决这些问题：

- ✅ 使用 **Excel COM API** 保存（仅限 Windows）
- ✅ **完全保留 VBA 宏**
- ✅ **保留所有模板图片**和图形
- ✅ **保持复杂格式**和数据验证
- ✅ 添加无占位符的 **{% img %} 标签**支持（空格分隔语法）

### 与原版的主要区别

| 功能 | 原版 xltpl | xltpl-preserve |
|------|-----------|----------------|
| 平台 | 跨平台（Windows/Mac/Linux） | **仅 Windows** |
| 保存方式 | openpyxl | **Excel COM API** |
| VBA 保留 | ❌ 丢失 | ✅ 保留 |
| 模板图片 | ❌ 丢失 | ✅ 保留 |
| 复杂格式 | ❌ 丢失 | ✅ 保留 |
| {% img %} 语法 | 逗号分隔 | **空格+逗号** |
| 性能 | 快速（~0.1秒/文件） | 较慢（~3秒/文件） |
| Excel 安装 | 不需要 | **必需** |

### 应该使用哪个版本

**使用原版 xltpl 如果：**
- 需要跨平台支持（Linux/Mac）
- 模板中没有 VBA 或复杂格式
- 需要快速性能

**使用 xltpl-preserve 如果：**
- 使用 Windows 且已安装 Excel
- 需要保留 VBA 宏
- 需要保留模板图片和图形
- 需要保留复杂格式

### 技术实现

此分支使用**两阶段方法**：

1. **阶段1（openpyxl）**：加载模板并执行 Jinja2 渲染
2. **阶段2（COM）**：通过 COM 打开原始模板，应用渲染结果，通过 COM 保存

这样可以在使用 openpyxl 强大的渲染能力的同时，保留原始模板中的所有内容。

### 致谢

原作者：**Zhang Yu** (zhangyu836@gmail.com)  
原始仓库：https://github.com/zhangyu836/xltpl

此分支独立维护，添加了 Windows 特定增强功能。

### 许可证

MIT License（与原版相同）
