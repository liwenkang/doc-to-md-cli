# doc_to_md_cli 使用说明

`doc_to_md_cli.py` 使用本机已安装的 Microsoft Word，将 `.doc` / `.docx` 批量转换为 Markdown 文档，并导出文档中的图片为 PNG，并在生成的 Markdown 中引用。

## 功能概览

- 支持单个文件或整个目录的 `.doc` / `.docx` 转换为 `.md`
- 自动识别标题、正文和表格，生成对应的 Markdown 结构
- 导出文档中的内嵌图片为 PNG，保存在单独的 `_images` 目录下，并在 Markdown 中引用
- 批量模式下复用同一个 Word 进程，并对常见 COM / RPC 异常进行重试与降级处理
- 提供安静模式和详细模式，方便在 CI 或本地调试时使用

## 环境要求

- 操作系统：Windows
- 已安装 Microsoft Word（支持通过 COM 调用）
- Python：3.8+（本仓库实际使用的是 Python 3.14）
- 依赖库：
  - `pywin32`（提供 `win32com.client` / `pywintypes`）
  - ImageMagick（命令行工具 `magick`，用于将 EMF 转为 PNG）

### 安装依赖

在当前 Python 环境中安装 `pywin32`：

```bash
pip install pywin32
```

安装 ImageMagick（若尚未安装）：

- 推荐方式：从官网下载安装包：https://imagemagick.org
- 安装时勾选“Add application directory to your system path”或确保在命令行中可以直接使用 `magick` 命令。

## 基本用法

在仓库根目录下执行脚本（示例使用 `python`，也可以使用具体解释器路径）：

### 单文件转换

```bash
python scripts/doc_to_md_cli.py "docs/spec-template.docx"
```

默认会在同级目录生成同名的 `.md` 文件，例如：

- 输入：`docs/spec-template.docx`
- 输出：`docs/spec-template.md`

可以通过 `-o/--output` 指定输出路径：

```bash
python scripts/doc_to_md_cli.py "docs/spec-template.docx" -o "docs/spec-template-converted.md"
```

### 目录批量转换

将某个目录下所有 `.doc` / `.docx` 文件批量转换：

```bash
python scripts/doc_to_md_cli.py docs
```

递归扫描子目录：

```bash
python scripts/doc_to_md_cli.py docs -r
```

脚本会：

- 为每个 Word 文件在同级目录生成同名 `.md` 文件
- 导出图片到对应的 `<md 文件名>_images/` 目录
- 输出批量转换统计信息（成功/失败个数）

## 命令行参数

```text
python scripts/doc_to_md_cli.py INPUT [options]
```

- `INPUT`（必须）：
  - 单个 `.doc` / `.docx` 文件路径，或
  - 某个目录路径（批量模式）

### 通用选项

- `-o, --output PATH`

  - 仅在单文件模式下有效
  - 指定输出 Markdown 文件路径；若省略，则与输入同名，仅扩展名改为 `.md`

- `--no-optimize`

  - 默认情况下，脚本会对生成的 Markdown 做一些轻量的清理：
    - 去掉常见控制字符
    - 合并多余的空行
    - 修正形如 `##标题` 为 `## 标题`
  - 加上此参数后，跳过这些清理步骤，保留最原始的导出结果

- `--quiet`

  - 安静模式：
    - 只输出最终统计信息和错误信息
    - 不输出每个文件的详细处理过程

- `--verbose`
  - 详细模式：
    - 输出每个文件的源/目标路径
    - 输出转换成功提示

> 注意：`--quiet` 与 `--verbose` 不能同时使用。

- `-r, --recursive`
  - 当 `INPUT` 是目录时，递归扫描子目录中的 `.doc` / `.docx` 文件进行批量转换

## 输出结构说明

以输入文件 `docs/spec-template.docx` 为例，转换后典型输出为：

- Markdown 文件：
  - `docs/spec-template.md`
- 图片目录：
  - `docs/spec-template_images/`
  - 内含导出的 PNG 图片：`img_1.png`, `img_2.png`, ...

Markdown 中的图片引用形式类似：

```markdown
![image1](spec-template_images/img_1.png)
```

在转换过程中：

- 脚本优先尝试使用 Word 的 `SaveAsPicture` 直接导出 PNG
- 若当前 Word 版本不支持或发生错误，则会导出 EMF 并调用 `magick` 将其转为 PNG
- 转换成功后会删除中间产生的 `.emf` 文件，最终仅保留 PNG

## 错误处理与重试机制

由于 Word COM 在高频调用或长时间批量转换时容易出现 RPC/指针类异常，脚本内置了一些防护逻辑：

- 段落/表格级降级：
  - 单个段落或整张表格处理失败时，只跳过该段落/表格，不影响整篇文档
- 文件级重试：
  - 单文件模式下，失败后会自动再重试一次
  - 目录批量模式下，每个文件最多尝试两次
- Word 进程级重启：
  - 在批量模式中，如捕获到典型的 RPC 服务器不可用、远程调用失败、无效指针等错误码
  - 会关闭当前 Word 进程，重新启动一个新的 Word 实例，然后对当前文件重试一次

这些机制可以显著提升在大批量转换场景下的稳定性。

## 使用建议与注意事项

- 仅支持安装了 Microsoft Word 的 Windows 环境
- 不建议并发（多进程/多线程）同时运行多个本脚本实例，以避免 Word COM 竞争和异常
- 批量转换时，最好确保相关 Word 文档未在前台被编辑或弹出对话框，避免影响 COM 操作
- 如果发现大量文件失败，可以先对某个单独文件加 `--verbose` 检查具体错误，再决定是否需要调整文档或环境配置

如需扩展功能（例如：自定义输出目录、忽略图片导出、支持更多 Word 对象类型等），可以在 `scripts/doc_to_md_cli.py` 的基础上继续演进。

## 常见问题（FAQ）

**Q1：运行时报错未找到 `win32com.client` 或 `pywintypes`，怎么办？**
请确认当前使用的 Python 环境已安装 `pywin32`：

```bash
pip install pywin32
```

并确保运行脚本时使用的解释器与安装 `pywin32` 的环境一致（例如同一个虚拟环境）。

**Q2：提示找不到 `magick` 命令，或者图片没有导出为 PNG？**
脚本在 Word 原生不支持 `SaveAsPicture` 时，会调用 ImageMagick 的 `magick` 命令将 EMF 转为 PNG：

- 请确认已正确安装 ImageMagick，并且命令行可以直接执行 `magick -version`
- 如果 `magick` 不可用，脚本会回退为使用 EMF 文件，Markdown 中的图片链接可能是 `.emf` 扩展名
- 建议安装并配置好 ImageMagick，以获得统一的 PNG 输出效果

**Q3：可以在 macOS 或 Linux 上使用这个脚本吗？**
目前不支持。脚本依赖 Windows 下的 Microsoft Word COM 自动化接口，因此只能在安装了 Word 的 Windows 环境中运行。

**Q4：为什么有些图片或图形在 Markdown 中没有显示？**
当前实现主要处理 Word 文档中的 `InlineShapes`（内嵌图片）。对于某些以浮动方式插入的 Shapes/图形（例如“嵌入型”之外的排版方式），可能不会被识别和导出。遇到这种情况时：

- 可以尝试在 Word 中将图片的布局调整为“与文字排列为嵌入型”，然后重新转换；或
- 需要额外扩展脚本逻辑，支持更多 Word 对象类型。

**Q5：可以并发（多进程/多终端）同时跑多个批量任务吗？**
不推荐。Word COM 在多进程/多线程并发调用时容易出现竞争、RPC 错误和随机崩溃。建议：

- 在单机上一次只运行一个批量任务；
- 如需并行处理大量文档，可以在多台机器上各自单进程运行。

**Q6：为什么生成的 Markdown 排版和原始 Word 不完全一致？**
本工具的目标是“高保真但不完全等价”的 Markdown：

- 标题、段落、表格和图片会尽量保留结构和内容；
- 复杂版式（多栏、特殊样式、某些 SmartArt/对象等）不会 1:1 还原；
- 如需手工整理，可在 Markdown 基础上做二次编辑。
