# HDX Word 提取器

从 HDX 中读取所选范围，直接生成合并后的 Word 手册。无需 ICS，不需要先运行 HTML 导出脚本，不生成中间 HTML 文件夹。原 HDX、示例 Word 和已生成的 HTML 离线包保持不变。

这是第一版 Word 转换工具：按示例的封面、版权、目录、章节正文、页眉页脚结构组织内容，支持可编辑表格、合并单元格、原始图片及内部书签。不是 ICS 排版引擎的复制品，不承诺字体、页数和逐页位置完全相同。

需要每份小于 10 MB 时，请阅读同目录的 `Word分卷使用说明.md`。新版 `--max-mb 10` 在生成过程中分卷，不生成整本 Word；增加 `--merge-word-attachments` 可将目录中的 DOC/DOCX 正文编入分卷。以下单文件说明适用于未启用分卷的旧模式。

## 首次准备

需要 Python 3.10 或更新版本，以及 `python-docx`、`lxml`。本机开发时检查到当前 `python` 已有这两个依赖，可以直接执行。

若另一台电脑提示缺少依赖，在项目目录执行：

```powershell
python -m pip install -r .\hdx_export_tools\requirements-word.txt
```

如果命令 `python` 不可用但 `py` 可用，下面所有命令可以把 `python` 换成 `py`。

## 第一次试跑

打开 PowerShell，进入项目：

```powershell
cd "D:\华为文档提取器"
```

先检查范围（不生成 Word）：

```powershell
python .\hdx_export_tools\export_word.py --chapter "用户手册 > iFC用户手册" --dry-run
```

应显示：29 个目录节点、29 个唯一正文。

正式生成：

```powershell
python .\hdx_export_tools\export_word.py --chapter "用户手册 > iFC用户手册" --out ".\Word导出\iFC用户手册.docx"
```

默认会在项目目录与第一层子目录中查找唯一的 `.hdx` 文件。找到多个文件时要求你明确指定，不会自行挑选。

也可以写出完整参数：

```powershell
python .\hdx_export_tools\export_word.py --hdx ".\HUAWEI IMS IP多媒体子系统解决方案 V500R019C50 产品文档 06\IMS_V500R019C50_06_zh_CZM0530J.hdx" --chapter "用户手册 > iFC用户手册" --out ".\Word导出\iFC用户手册.docx"
```

## 生成后检查

输出包含：

- `Word导出\iFC用户手册.docx`：正文、内嵌图片、目录和内部链接。
- `Word导出\iFC用户手册.report.json`：每页文本对照、表格和图片数量、章节外引用、错误和警告、实际耗时。
- 如果选定内容含可下载附件，会另外生成 `iFC用户手册_附件` 文件夹；请与 Word 文件一起移动。

### 目录项本身就是 Word/PDF 等文件

部分 HDX 的目录节点直接指向 `.doc` 等原始文件，并非 HTML。这类节点不做文字解码：工具保留完整原文件，在汇总 Word 的对应标题下提供相对链接，并校验附件的 SHA-256。

例如本包“用户手册”的 548 项实际为 **547 个 HTML 页面 + 1 个 DOC 文件**。默认单文件模式将该 DOC 留作附件。新版分卷模式加 `--merge-word-attachments` 后，会使用本机 Microsoft Word 转换旧 DOC，再将其正文按原目录位置编入分卷；具体依赖和验证边界见 `Word分卷使用说明.md`。

用 Microsoft Word 打开文件：

1. 按 `Ctrl+A`，再按 `F9`（部分键盘需 `Fn+F9`），选择“更新整个目录”。也可在目录上右键“更新域”。
2. 在 Word 左侧开启“视图 → 导航窗格”，确认章节层级。
3. 查看中文、复杂表格、大图和长命令块；试用目录和正文内部链接。Word 通常需要 `Ctrl+单击` 才能跟随链接。
4. 保存文档，以保存排版后计算的目录页码。

目录初始已包含可点击的条目，但页码显示 `—`。准确页码必须由 Word 的排版引擎计算，Python 不伪造页数；某些阅读器不会自动更新目录和页码，需要用 Word 更新。

程序显示“错误：0”说明自动检查通过，不代表所有页面视觉排版都已验收。报告中的 `visual_review` 会明确显示未进行人工版式检查。

出现转换错误时仍会保留 Word 供排查，并返回非零退出码；不要把这种结果当成最终通过。把终端输出和 `.report.json` 发回来即可定位。被选范围以外的引用会指向文末说明清单，不会偷偷扩大范围导出。

## 换范围或重复导出

列出前两层目录及节点数：

```powershell
python .\hdx_export_tools\export_word.py --list --list-depth 2
```

`--chapter` 使用完整中文层级，层级间用 `>`。该选项可以重复使用，将多个互不重叠的章节合并为同一 Word；不要同时选父章节和子章节。

再次导出时，优先更换输出名。如确定覆盖原 Word，加 `--overwrite`，并先关闭 Word 里的同名文件。

超过 200 个目录节点时会停止，提示你先确认范围。确认后自行加 `--allow-large` 才会执行。这个阈值是防误操作保护，并不是能力上限或耗时预测。

## 可选 PowerShell 快捷入口

```powershell
.\hdx_export_tools\导出Word.ps1 -DryRun
.\hdx_export_tools\导出Word.ps1
```

默认仍为 iFC 手册。支持 `-Chapter`、`-Hdx`、`-Out`、`-Overwrite` 和 `-AllowLarge`。
如果系统不允许运行 `.ps1`，直接使用上面的 Python 命令，无需修改系统策略。

## 当前验证范围与限制

- 已提供小型合成测试，覆盖中文转码、内部书签、同名文件、横纵合并表格、内嵌图片、附件、缺图报错、覆盖保护和范围预检查。
- 在真实 HDX 中选择“iFC触发原理”“配置iFC基础数据”“配置实例”进行内存转换核对，正文文本、表格和图片转换自动检查通过；没有替用户生成整份 iFC Word。
- 全部 29 项的实际 Word 导出与视觉验收留给首次试跑。没有逐页复刻或承诺与 ICS 示例分页相同。
- 自动检查比较去除网页导航后的正文文本；网页的目录菜单、上一节/父主题导航不作为 Word 正文重复输出。
- 动态网页交互、视频、公式等特殊元素可能需要另外适配，会在报告中提示；不执行包内 JavaScript，不访问外部网络。
- 图像保留源图像数据并按页面宽度缩放，不依赖外部图片文件。
- 超长命令、跨页合并表格、字体差异等仍需要 Word 中人工检查；在样例通过前不要直接扩大为整包导出。

开发验证命令：

```powershell
python .\hdx_export_tools\test_export_word.py
```
