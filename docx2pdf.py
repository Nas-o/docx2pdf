#!/usr/bin/env python3
"""docx2pdf -- convert Word documents to PDF on Linux.

Command line front-end for :mod:`d2p`.  Examples::

    ./docx2pdf report.docx                  # -> report.pdf
    ./docx2pdf -d out/ *.docx               # batch into a directory
    ./docx2pdf -r ./documents              # recurse a folder tree
    ./docx2pdf --engine builtin a.docx     # force the pure-Python engine
    ./docx2pdf --doctor                    # what is available on this machine

Exit status: ``0`` everything converted, ``1`` at least one failure,
``2`` nothing could be attempted (bad usage / no input).
"""

from __future__ import annotations

import argparse
import glob
import logging
import multiprocessing
import os
import sys
import time
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed
from typing import List, Optional, Sequence, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from d2p import __version__  # noqa: E402
from d2p.engines import (DEFAULT_TIMEOUT, LIBREOFFICE_EXTENSIONS,
                         ConversionError, ConversionResult, LibreOfficeEngine, doctor,
                         engines_status, select_engine)  # noqa: E402

SUPPORTED_EXTENSIONS = LIBREOFFICE_EXTENSIONS

USAGE_EPILOG = """\
支持的输入格式 / accepted inputs:
  .docx .docm .dotx .dotm        (内置引擎与 LibreOffice)
  .doc .odt .rtf .wps .txt ...   (需要 LibreOffice，安装方式见 README)

示例 / examples:
  docx2pdf 报告.docx                    生成 报告.pdf
  docx2pdf -d pdf/ 报告.docx 合同.docx   输出到 pdf/ 目录
  docx2pdf -r 资料/ -d pdf/ -j 4         递归转换并并行处理
  docx2pdf --engine builtin *.docx       强制使用内置引擎
  docx2pdf --doctor                      检查环境与字体
"""


# --------------------------------------------------------------------------- #
# input collection
# --------------------------------------------------------------------------- #


def _looks_like_glob(value: str) -> bool:
    return any(ch in value for ch in "*?[")


def _scan_directory(path: str, recursive: bool) -> List[str]:
    found: List[str] = []
    if recursive:
        for root, dirnames, filenames in os.walk(path):
            dirnames[:] = [d for d in dirnames if not d.startswith(".")]
            for name in filenames:
                if os.path.splitext(name)[1].lower() in SUPPORTED_EXTENSIONS:
                    found.append(os.path.join(root, name))
    else:
        try:
            with os.scandir(path) as entries:
                for entry in entries:
                    if entry.is_file() and os.path.splitext(entry.name)[1].lower() in SUPPORTED_EXTENSIONS:
                        found.append(entry.path)
        except OSError:
            pass
    return found


def collect_inputs(values: Sequence[str], recursive: bool,
                   strict: bool = True) -> Tuple[List[str], List[str]]:
    """Expand files, directories and glob patterns into a file list."""
    files: List[str] = []
    problems: List[str] = []
    for value in values:
        if os.path.isdir(value):
            matches = _scan_directory(value, recursive)
            if not matches:
                problems.append(f"目录中没有可转换的文档 / no documents in directory: {value}")
            files.extend(matches)
            continue
        if os.path.isfile(value):
            files.append(value)
            continue
        if _looks_like_glob(value):
            matches = sorted(p for p in glob.glob(value, recursive=True) if os.path.isfile(p))
            if not matches:
                problems.append(f"通配符没有匹配到文件 / pattern matched nothing: {value}")
            files.extend(matches)
            continue
        problems.append(f"文件不存在 / no such file: {value}")

    deduped: List[str] = []
    seen = set()
    for path in files:
        absolute = os.path.abspath(path)
        if absolute in seen:
            continue
        seen.add(absolute)
        deduped.append(path)
    if strict:
        for path in deduped:
            ext = os.path.splitext(path)[1].lower()
            if ext not in SUPPORTED_EXTENSIONS:
                problems.append(f"跳过不支持的文件类型 / unsupported file type: {path}")
    deduped = [p for p in deduped
               if os.path.splitext(p)[1].lower() in SUPPORTED_EXTENSIONS]
    return deduped, problems


def plan_outputs(files: Sequence[str], output: Optional[str], outdir: Optional[str],
                 recursive_roots: Sequence[str]) -> List[Tuple[str, str]]:
    """Map each source file onto its target PDF path."""
    pairs: List[Tuple[str, str]] = []
    single = len(files) == 1

    def relative_base(path: str) -> str:
        best = ""
        absolute = os.path.abspath(path)
        for root in recursive_roots:
            root_abs = os.path.abspath(root)
            if absolute.startswith(root_abs + os.sep) and len(root_abs) > len(best):
                best = root_abs
        return best

    for path in files:
        stem, _ = os.path.splitext(path)
        if output and single and not os.path.isdir(output):
            target = output
        else:
            destination_root = outdir or (output if output else None)
            if destination_root is None:
                target = stem + ".pdf"
            else:
                base = relative_base(path)
                if base:
                    rel = os.path.relpath(path, base)
                    target = os.path.join(destination_root, os.path.splitext(rel)[0] + ".pdf")
                else:
                    target = os.path.join(destination_root, os.path.basename(stem) + ".pdf")
        pairs.append((path, target))
    return pairs


# --------------------------------------------------------------------------- #
# conversion
# --------------------------------------------------------------------------- #


def convert_one(source: str, target: str, engine_name: str, font_dirs: Optional[Sequence[str]],
                timeout: float, overwrite: bool, dry_run: bool = False) -> ConversionResult:
    result = ConversionResult(source=source, target=target)
    if not overwrite and os.path.exists(target):
        result.skipped = True
        result.ok = True
        result.engine = "-"
        return result
    if dry_run:
        result.ok = True
        result.engine = "dry-run"
        return result

    started = time.perf_counter()
    try:
        engine = select_engine(engine_name, os.path.splitext(source)[1], font_dirs)
        result.engine = engine.name
        converted = engine.convert(source, target, timeout=timeout)
        result.ok = True
        result.pages = converted.pages
        result.warnings = converted.warnings
    except ConversionError as exc:
        result.ok = False
        result.error = str(exc)
        try:
            if os.path.exists(target) and os.path.getsize(target) == 0:
                os.remove(target)
        except OSError:
            pass
    except Exception as exc:  # unexpected: keep the batch going
        result.ok = False
        result.error = f"{type(exc).__name__}: {exc}"
        if os.environ.get("D2P_TRACEBACK"):
            import traceback

            traceback.print_exc()
        else:
            logging.getLogger("d2p").debug("conversion of %s failed", source, exc_info=True)
        try:
            if os.path.exists(target):
                os.remove(target)
        except OSError:
            pass
    result.seconds = time.perf_counter() - started
    return result


def _make_pool(use_processes: bool, jobs: int):
    """Create a worker pool, falling back to serial when it is unavailable.

    Some containers/sandboxes forbid the POSIX semaphores multiprocessing needs;
    in that case ``(None, "serial")`` is returned and the caller runs serially
    (which is also faster than threads for the CPU-bound built-in engine).
    """
    if not use_processes:
        try:
            return ThreadPoolExecutor(max_workers=jobs), "threads"
        except Exception as exc:
            print(f"提示: 无法创建线程池（{exc}），改为串行 / falling back to serial",
                  file=sys.stderr)
            return None, "serial"
    try:
        context = multiprocessing.get_context("fork") if hasattr(os, "fork") else None
        return ProcessPoolExecutor(max_workers=jobs, mp_context=context), "processes"
    except Exception as exc:
        print(f"提示: 当前环境无法创建多进程（{type(exc).__name__}: {exc}），"
              f"改为串行转换 / multiprocessing unavailable, running serially",
              file=sys.stderr)
        return None, "serial"


def _format_result(result: ConversionResult, index: int, total: int, quiet: bool) -> List[str]:
    lines: List[str] = []
    name = os.path.basename(result.source)
    if result.skipped:
        message = f"[{index}/{total}] 跳过 {name}（目标已存在，使用 --overwrite 覆盖）" if not quiet else ""
        if message:
            lines.append(message)
        return lines
    if result.ok:
        stamp = f"{result.seconds:.2f}s"
        pages = f", {result.pages}页" if result.pages else ""
        lines.append(f"[{index}/{total}] OK   {name} -> {result.target} "
                     f"({result.engine}, {stamp}{pages})")
        for warning in result.warnings[:5]:
            lines.append(f"        ! {warning}")
        if len(result.warnings) > 5:
            lines.append(f"        ! ... 另外 {len(result.warnings) - 5} 条提示")
    else:
        lines.append(f"[{index}/{total}] 失败 {name}: {result.error}")
    return lines


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docx2pdf",
        description="Linux 下把 Word 文档转换为 PDF（内置引擎 + 可选 LibreOffice）"
                    " / convert Word documents to PDF on Linux",
        epilog=USAGE_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("inputs", nargs="*", help="Word 文档、目录或通配符 / files, directories or globs")
    parser.add_argument("-o", "--output", help="输出文件（单个输入）或输出目录 / output file or directory")
    parser.add_argument("-d", "--outdir", help="输出目录 / output directory")
    parser.add_argument("-r", "--recursive", dest="recursive", action="store_true", default=None,
                        help="递归处理目录（默认开启）/ recurse into directories (default)")
    parser.add_argument("--no-recursive", dest="recursive", action="store_false",
                        help="只处理目录第一层 / only the top level of directories")
    parser.add_argument("--engine", choices=["auto", "builtin", "libreoffice"], default="auto",
                        help="转换引擎，默认 auto（优先 LibreOffice）/ engine (default: auto)")
    parser.add_argument("--overwrite", dest="overwrite", action="store_true", default=True,
                        help="覆盖已存在的 PDF（默认）/ overwrite existing PDFs (default)")
    parser.add_argument("--no-clobber", dest="overwrite", action="store_false",
                        help="跳过已存在的 PDF / skip existing PDFs")
    parser.add_argument("-j", "--jobs", type=int, default=0,
                        help="并行任务数，默认 min(CPU,4)，1 表示串行 / parallel jobs")
    parser.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT,
                        help=f"单个文件超时秒数（LibreOffice），默认 {DEFAULT_TIMEOUT:.0f}")
    parser.add_argument("--font-dir", action="append", default=[], metavar="DIR",
                        help="额外的字体目录（可重复）/ extra font directory")
    parser.add_argument("-n", "--dry-run", action="store_true", help="只显示将执行的操作 / list only")
    parser.add_argument("-q", "--quiet", action="store_true", help="仅在出错时输出 / quiet")
    parser.add_argument("-v", "--verbose", action="store_true", help="输出更多信息 / verbose")
    parser.add_argument("--doctor", action="store_true", help="检查环境与字体 / environment check")
    parser.add_argument("--engines", action="store_true", help="列出可用引擎 / list engines")
    parser.add_argument("--fonts", action="store_true", help="列出检测到的字体 / list fonts")
    parser.add_argument("--version", action="version", version=f"d2p {__version__}")
    return parser


def _print_engines() -> int:
    print(f"d2p {__version__} 可用引擎 / available engines")
    for status in engines_status():
        mark = "OK " if status["available"] else "-- "
        print(f"[{mark}] {status['name']:<12} {status['detail']}")
        print(f"        支持 / supports: {' '.join(status['extensions'])}")
    return 0


def _print_fonts(font_dirs: Sequence[str]) -> int:
    from d2p.fonts import get_font_book

    print(get_font_book(font_dirs or None).explain())
    return 0


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.WARNING,
        format="%(levelname)s %(name)s: %(message)s",
    )
    font_dirs = args.font_dir or None
    recursive = True if args.recursive is None else args.recursive

    if args.doctor:
        print(doctor(font_dirs))
        return 0
    if args.engines:
        return _print_engines()
    if args.fonts:
        return _print_fonts(font_dirs or [])

    if not args.inputs:
        parser.print_help()
        return 2

    files, problems = collect_inputs(args.inputs, recursive)
    for problem in problems:
        print(f"警告: {problem}", file=sys.stderr)
    if not files:
        print("没有找到可转换的文档 / no convertible documents found", file=sys.stderr)
        return 2

    roots = [value for value in args.inputs if os.path.isdir(value)]
    if args.output and args.outdir:
        print("不能同时使用 -o 和 -d / use either -o or -d", file=sys.stderr)
        return 2
    if args.output and len(files) > 1 and not os.path.isdir(args.output):
        os.makedirs(args.output, exist_ok=True)

    pairs = plan_outputs(files, args.output, args.outdir, roots)
    total = len(pairs)
    jobs = args.jobs if args.jobs and args.jobs > 0 else min(os.cpu_count() or 1, 4)
    jobs = max(1, min(jobs, total))

    # The built-in engine is pure Python (CPU bound), so threads would only
    # fight over the GIL -- processes are faster.  LibreOffice conversion is
    # mostly waiting on a subprocess, where threads are ideal.
    use_processes = args.engine == "builtin" or (
        args.engine == "auto" and not LibreOfficeEngine().available())
    if use_processes and jobs > 1:
        from d2p.fonts import get_font_book

        get_font_book(font_dirs)      # warm caches so forked workers inherit them

    results: List[ConversionResult] = []

    def run_serial() -> None:
        for index, (source, target) in enumerate(pairs, start=1):
            result = convert_one(source, target, args.engine, font_dirs, args.timeout,
                                 args.overwrite, args.dry_run)
            results.append(result)
            for line in _format_result(result, index, total, args.quiet):
                print(line, flush=True)

    pool = None
    mode = "serial"
    if not args.dry_run and jobs > 1:
        pool, mode = _make_pool(use_processes, jobs)

    if not args.quiet:
        detail = f", {mode}" if mode != "serial" else ""
        print(f"准备转换 {total} 个文档 / converting {total} document(s) "
              f"[engine={args.engine}, jobs={jobs}{detail}]")

    if pool is None:
        run_serial()
    else:
        handled: set = set()
        try:
            with pool:
                futures = {
                    pool.submit(convert_one, source, target, args.engine, font_dirs,
                                args.timeout, args.overwrite, False): (index, source, target)
                    for index, (source, target) in enumerate(pairs, start=1)
                }
                for future in as_completed(futures):
                    index, source, _ = futures[future]
                    try:
                        result = future.result()
                    except Exception as exc:  # worker died: report, keep going
                        result = ConversionResult(source=source, ok=False,
                                                  error=f"{type(exc).__name__}: {exc}")
                    results.append(result)
                    handled.add(index)
                    for line in _format_result(result, index, total, args.quiet):
                        print(line, flush=True)
        except Exception as exc:  # pool itself broke: finish the rest serially
            print(f"提示: 并行执行失败（{type(exc).__name__}: {exc}），"
                  f"剩余文件改为串行 / parallel run failed, finishing serially",
                  file=sys.stderr)
            for index, (source, target) in enumerate(pairs, start=1):
                if index in handled:
                    continue
                result = convert_one(source, target, args.engine, font_dirs, args.timeout,
                                     args.overwrite, args.dry_run)
                results.append(result)
                for line in _format_result(result, index, total, args.quiet):
                    print(line, flush=True)

    ok = sum(1 for r in results if r.ok and not r.skipped)
    skipped = sum(1 for r in results if r.skipped)
    failed = [r for r in results if not r.ok]
    elapsed = sum(r.seconds for r in results)

    if not args.quiet or failed or skipped:
        print("-" * 52)
        print(f"完成: 成功 {ok}，跳过 {skipped}，失败 {len(failed)}，总耗时 {elapsed:.2f}s"
              f" / done: {ok} ok, {skipped} skipped, {len(failed)} failed")
    for result in failed:
        print(f"  x {result.source}: {result.error}", file=sys.stderr)

    if args.verbose:
        for result in results:
            if result.ok and not result.skipped:
                print(f"  - {result.source} -> {result.target}")
    return 1 if failed else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n中断 / interrupted", file=sys.stderr)
        sys.exit(130)
