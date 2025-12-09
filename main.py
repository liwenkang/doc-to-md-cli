#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用本机已安装的 Microsoft Word，将 .doc / .docx 转为 Markdown。

用法示例：
    python scripts/doc_to_md_cli.py "docs/00 - 编码规则维护.docx"
    python scripts/doc_to_md_cli.py "docs/00 - 编码规则维护.docx" -o "docs/00-编码规则维护.md"
"""

import os
import sys
import re
import subprocess
from typing import Any, Optional, List

try:
    import win32com.client  # type: ignore
except ImportError:
    win32com = None  # 延后在运行时给出更友好的报错

try:
    import pywintypes  # type: ignore
except ImportError:  # pragma: no cover - 极端情况
    pywintypes = None  # type: ignore


def simple_optimize_markdown(text: str) -> str:
    """对生成的 Markdown 做一些简单清理，尽量不破坏原始结构。"""
    # 去掉明显的控制字符（常见于 Word 导出）
    text = text.replace("\x07", "")

    # 将 3 个及以上连续空行压缩为 2 行
    text = re.sub(r"\n{3,}", "\n\n", text)

    # 标题后确保有空格，例如 "##标题" -> "## 标题"
    text = re.sub(r"^(#{1,6})([^#\s])", r"\1 \2", text, flags=re.MULTILINE)

    return text


def _convert_with_word_instance(word: Any, doc_path: str, md_path: str, optimize: bool = True, verbosity: int = 1) -> None:
    """在给定的 Word 实例上执行单个文档的转换逻辑。出现异常时直接抛出。"""
    doc_path = os.path.abspath(doc_path)
    md_path = os.path.abspath(md_path)

    if verbosity >= 2:
        print(f"源文件: {doc_path}")
        print(f"目标文件: {md_path}")

    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"文件不存在: {doc_path}")

    doc = word.Documents.Open(
        doc_path,
        ReadOnly=True,
        AddToRecentFiles=False,
        ConfirmConversions=False,
    )
    try:
        markdown_lines: List[str] = []

        # 图片导出目录（与 md 同级，按文件名区分）
        image_dir = os.path.join(
            os.path.dirname(md_path),
            f"{os.path.splitext(os.path.basename(md_path))[0]}_images",
        )
        os.makedirs(image_dir, exist_ok=True)
        image_index = 1

        # 1. 段落（标题、正文 + 内嵌图片）
        for para in doc.Paragraphs:
            try:
                # 先拿到文本和图片，避免仅有图片的段落被整体跳过
                text = para.Range.Text.strip()

                try:
                    inline_shapes = para.Range.InlineShapes
                    count = inline_shapes.Count
                except Exception:
                    inline_shapes = None
                    count = 0

                # 文本部分
                if text:
                    try:
                        style_name = str(para.Style.NameLocal)
                    except Exception:
                        # 降级：无法获取样式时，当作普通段落处理
                        style_name = ""

                    # 判断是否为标题（兼容中文“标题 1”、英文“Heading 1”等）
                    is_heading = False
                    level = 1
                    if "标题" in style_name or "Heading" in style_name:
                        is_heading = True
                        for i in range(1, 7):
                            if str(i) in style_name:
                                level = i
                                break

                    if is_heading:
                        markdown_lines.append(f"{'#' * level} {text}\n\n")
                    else:
                        # 普通段落，尽量保留原始文本
                        markdown_lines.append(f"{text}\n\n")

                # 图片部分（即便没有文本，也要处理图片）
                if inline_shapes is not None and count:
                    for i in range(1, count + 1):
                        try:
                            ishape = inline_shapes.Item(i)

                            # 优先尝试使用 Word 的 SaveAsPicture 导出为 PNG
                            img_base = f"img_{image_index}"
                            png_path = os.path.join(image_dir, img_base + ".png")
                            img_path = png_path

                            try:
                                ishape.SaveAsPicture(png_path)  # 部分版本 Word 支持
                            except Exception:
                                # 兼容路径：使用 EMF 数据导出，再调用 ImageMagick 转 PNG
                                emf_path = os.path.join(image_dir, img_base + ".emf")
                                bits = ishape.Range.EnhMetaFileBits
                                with open(emf_path, "wb") as img_f:
                                    img_f.write(bits)

                                # 调用 ImageMagick 将 EMF 转为 PNG
                                try:
                                    subprocess.run(
                                        ["magick", emf_path, png_path],
                                        check=True,
                                        stdout=subprocess.DEVNULL,
                                        stderr=subprocess.DEVNULL,
                                    )
                                    if os.path.exists(png_path):
                                        # 转换成功，删除 EMF，只保留 PNG
                                        try:
                                            os.remove(emf_path)
                                        except Exception:
                                            pass
                                        img_path = png_path
                                    else:
                                        # 转换失败时，退回到 EMF
                                        img_path = emf_path
                                except Exception:
                                    # magick 命令不可用或转换失败，退回到 EMF
                                    img_path = emf_path

                            rel_path = os.path.relpath(
                                img_path, start=os.path.dirname(md_path)
                            ).replace("\\", "/")
                            markdown_lines.append(
                                f"![image{image_index}]({rel_path})\n\n"
                            )
                            image_index += 1
                        except Exception as e_img:
                            print(
                                f"警告：导出图片失败，已跳过。{e_img}",
                                file=sys.stderr,
                            )
            except Exception as e:
                # 降级：单个段落处理失败时，跳过该段落，避免整篇失败
                print(f"警告：处理段落时出错，已跳过。{e}", file=sys.stderr)
                continue

        # 2. 表格
        for table in doc.Tables:
            try:
                rows = table.Rows.Count
                cols = table.Columns.Count
                if rows == 0 or cols == 0:
                    continue

                markdown_lines.append("\n")

                for r in range(1, rows + 1):
                    cells = []
                    for c in range(1, cols + 1):
                        try:
                            cell = table.Cell(r, c)
                            cell_text = cell.Range.Text
                            # 去掉单元格结尾自带的特殊字符
                            cell_text = cell_text.replace("\r\x07", "").replace("\r", " ").replace("\n", " ")
                            cells.append(cell_text.strip())
                        except Exception:
                            cells.append("")

                    markdown_lines.append("| " + " | ".join(cells) + " |\n")

                    # 第一行后加表头分隔线
                    if r == 1:
                        markdown_lines.append("| " + " | ".join(["---"] * cols) + " |\n")

                markdown_lines.append("\n")
            except Exception as e:
                # 降级：整张表处理失败时，跳过该表，避免中断其他内容
                print(f"警告：处理表格时出错，已跳过。{e}", file=sys.stderr)
                continue

        markdown = "".join(markdown_lines)
        if optimize:
            markdown = simple_optimize_markdown(markdown)

        os.makedirs(os.path.dirname(md_path) or ".", exist_ok=True)
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(markdown)

        if verbosity >= 2:
            print("转换成功。")
    finally:
        # 只关闭文档，不关闭 Word 实例
        doc.Close(False)


def convert_doc_to_markdown(doc_path: str, md_path: str, optimize: bool = True, verbosity: int = 1) -> bool:
    """使用 Word COM 将 Word 文档转换为 Markdown 近似格式（单文件模式）。"""
    if win32com is None:
        print("错误：未安装 pywin32，无法调用 Word。请先安装 pywin32。", file=sys.stderr)
        return False

    word = None
    try:
        # 启动独立的 Word 实例
        word = win32com.client.Dispatch("Word.Application")  # type: ignore
        word.Visible = False
        word.DisplayAlerts = 0

        _convert_with_word_instance(word, doc_path, md_path, optimize=optimize, verbosity=verbosity)
        return True
    except Exception as e:
        print(f"转换失败: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return False
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def _default_output_path(input_path: str) -> str:
    base, _ = os.path.splitext(os.path.abspath(input_path))
    return base + ".md"


def _iter_doc_files(root: str, recursive: bool) -> List[str]:
    """枚举目录下所有 .doc / .docx 文件。"""
    exts = {".doc", ".docx"}
    root = os.path.abspath(root)

    results: list[str] = []
    if not recursive:
        for name in os.listdir(root):
            path = os.path.join(root, name)
            if os.path.isfile(path) and os.path.splitext(name)[1].lower() in exts:
                results.append(path)
    else:
        for dirpath, _dirnames, filenames in os.walk(root):
            for name in filenames:
                if os.path.splitext(name)[1].lower() in exts:
                    results.append(os.path.join(dirpath, name))
    return results


def _is_rpc_or_pointer_error(exc: Exception) -> bool:
    """判断是否为典型的 RPC / 指针类 COM 异常，需要通过重启 Word 来恢复。"""
    # COM 层错误码
    rpc_codes = {
        -2147023174,  # RPC 服务器不可用
        -2147023170,  # 远程过程调用失败
        -2147467261,  # 无效指针
    }

    if pywintypes is not None and isinstance(exc, pywintypes.com_error):  # type: ignore
        try:
            code = exc.args[0]
            if isinstance(code, int) and code in rpc_codes:
                return True
        except Exception:
            pass

    # win32com 封装的枚举异常
    if isinstance(exc, TypeError) and "does not support enumeration" in str(exc):
        return True

    return False


def main(argv: Optional[List[str]] = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(
        description="将 .doc/.docx 转为 Markdown（依赖本机 Microsoft Word），支持单文件或目录批量转换",
    )
    parser.add_argument("input", help="输入路径：单个 .doc/.docx 文件或目录")
    parser.add_argument(
        "-o",
        "--output",
        help="输出 Markdown 文件路径（默认与输入同名，扩展名改为 .md）",
    )
    parser.add_argument(
        "--no-optimize",
        action="store_true",
        help="不做简单的 Markdown 清理（保留最原始导出结果）",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="安静模式：仅输出统计和错误，不输出每个文件的处理信息",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="详细模式：输出每个文件的源/目标路径及转换成功提示",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="当 input 为目录时，递归扫描子目录中的 .doc/.docx 文件",
    )

    args = parser.parse_args(argv)

    if args.quiet and args.verbose:
        print("错误：--quiet 与 --verbose 不能同时使用", file=sys.stderr)
        return 1

    if args.quiet:
        verbosity = 0
    elif args.verbose:
        verbosity = 2
    else:
        verbosity = 1

    input_path = os.path.abspath(args.input)

    # 目录：批量转换
    if os.path.isdir(input_path):
        files = _iter_doc_files(input_path, recursive=args.recursive)
        if not files:
            print(f"目录下未找到 .doc/.docx 文件：{input_path}")
            return 0

        total = len(files)
        ok = 0
        fail = 0
        if verbosity >= 1:
            print(f"开始批量转换，文件数：{total}，目录：{input_path}")

        if win32com is None:
            print("错误：未安装 pywin32，无法调用 Word。请先安装 pywin32。", file=sys.stderr)
            return 1

        word = None
        try:
            # 批量场景：复用同一个 Word 实例，减少启动/销毁开销
            word = win32com.client.Dispatch("Word.Application")  # type: ignore
            word.Visible = False
            word.DisplayAlerts = 0

            for idx, path in enumerate(files):
                out_path = _default_output_path(path)
                if verbosity >= 1:
                    print("-" * 60)
                    print(f"[{idx + 1}/{total}] 处理：{path}")

                success = False
                last_exc: Optional[Exception] = None

                for attempt in range(2):  # 最多两次：首次 + 一次重启重试
                    try:
                        _convert_with_word_instance(
                            word,
                            path,
                            out_path,
                            optimize=not args.no_optimize,
                            verbosity=verbosity,
                        )
                        success = True
                        break
                    except Exception as e:
                        last_exc = e
                        print(f"转换失败（第 {attempt + 1} 次）: {e}", file=sys.stderr)
                        import traceback
                        traceback.print_exc()

                        if attempt == 0 and _is_rpc_or_pointer_error(e):
                            # 典型 RPC/指针错误：重启 Word 后再试一次
                            if verbosity >= 1:
                                print("检测到 RPC/指针类错误，准备重启 Word 实例后重试...", file=sys.stderr)
                            try:
                                word.Quit()
                            except Exception:
                                pass
                            # 重新创建 Word 实例
                            word = win32com.client.Dispatch("Word.Application")  # type: ignore
                            word.Visible = False
                            word.DisplayAlerts = 0
                        else:
                            # 非 RPC 类错误，或者已重试过一次，直接跳出
                            break

                if success:
                    ok += 1
                else:
                    fail += 1
                    if last_exc is not None:
                        print(f"最终失败文件：{path}，错误：{last_exc}", file=sys.stderr)

        finally:
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass

        if verbosity >= 1:
            print("=" * 60)
        print(f"批量转换完成：成功 {ok} 个，失败 {fail} 个，共 {total} 个。")
        return 0 if fail == 0 else 1

    # 单文件：保持原有行为，但增加一次重试
    output_path = args.output or _default_output_path(input_path)
    success = convert_doc_to_markdown(
        input_path,
        output_path,
        optimize=not args.no_optimize,
        verbosity=verbosity,
    )
    if not success:
        print("首次转换失败，准备重试一次...", file=sys.stderr)
        success = convert_doc_to_markdown(
            input_path,
            output_path,
            optimize=not args.no_optimize,
            verbosity=verbosity,
        )

    return 0 if success else 1


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
