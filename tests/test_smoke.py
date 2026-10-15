#!/usr/bin/env python3
"""End-to-end smoke tests for docx2pdf (no pytest required).

    python3 tests/test_smoke.py            # run everything
    python3 tests/test_smoke.py -v         # also print each check

Checks document conversion quality (text, CJK, tables, images, lists,
headers/footers) and command line behaviour (batch, recursion, clobbering,
dry-run, error exits).  Exits non-zero if any check fails.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CLI = os.path.join(ROOT, "docx2pdf.py")
sys.path.insert(0, HERE)
sys.path.insert(0, ROOT)

import make_samples  # noqa: E402

VERBOSE = "-v" in sys.argv or "--verbose" in sys.argv

_results = []


def check(name: str, condition: bool, detail: str = "") -> bool:
    _results.append((name, bool(condition), detail))
    if VERBOSE or not condition:
        mark = "PASS" if condition else "FAIL"
        print(f"  [{mark}] {name}" + (f" -- {detail}" if detail and not condition else ""))
    return bool(condition)


def run_cli(*args: str, expect: int = 0):
    proc = subprocess.run([sys.executable, CLI, *args], capture_output=True, text=True)
    if proc.returncode != expect:
        check(f"exit code for {' '.join(args)}", False,
              f"expected {expect}, got {proc.returncode}; stderr={proc.stderr.strip()[:400]}")
    return proc


def pdf_text(path: str) -> str:
    import pymupdf

    with pymupdf.open(path) as doc:
        return "\n".join(page.get_text() for page in doc)


def pdf_flat_text(path: str) -> str:
    """Whitespace-normalised text, for checks that may wrap across lines."""
    return " ".join(pdf_text(path).split())


def pdf_pages(path: str) -> int:
    import pymupdf

    with pymupdf.open(path) as doc:
        return doc.page_count


def _make_docm(source: str, target: str) -> None:
    """Turn a .docx into a macro-enabled .docm package (content type only)."""
    import zipfile

    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.decode("utf-8").replace(
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                    "application/vnd.ms-word.document.macroEnabled.main+xml",
                ).encode("utf-8")
            zout.writestr(item, data)


def main() -> int:
    try:
        import pymupdf  # noqa: F401
        import reportlab  # noqa: F401
        import docx  # noqa: F401
    except Exception as exc:
        print(f"缺少依赖 / missing dependency: {exc}")
        return 2

    work = tempfile.mkdtemp(prefix="d2p-test-")
    try:
        print("== 准备样本 / generating samples ==")
        samples = os.path.join(work, "samples")
        make_samples.main([samples])
        cn = os.path.join(samples, "sample_cn.docx")
        en = os.path.join(samples, "sample_en.docx")

        print("== 基本转换 / basic conversion ==")
        out = os.path.join(work, "out")
        proc = run_cli("-d", out, cn, en)
        cn_pdf = os.path.join(out, "sample_cn.pdf")
        en_pdf = os.path.join(out, "sample_en.pdf")
        check("两个 PDF 均生成 / both PDFs exist",
              os.path.exists(cn_pdf) and os.path.exists(en_pdf))
        check("PDF 非空 / PDFs are not empty",
              os.path.getsize(cn_pdf) > 2000 and os.path.getsize(en_pdf) > 2000,
              f"{os.path.getsize(cn_pdf)} / {os.path.getsize(en_pdf)}")
        check("中文样本分页正确 / CN sample is 2 pages", pdf_pages(cn_pdf) == 2,
              f"got {pdf_pages(cn_pdf)}")

        text = pdf_text(cn_pdf)
        for needle in ("季度业务报告", "核心指标", "Key Metrics", "供应链成本下降",
                       "第一阶段", "合计 Total", "图 1", "Revenue", "Mixed English"):
            check(f"文本包含 / text contains {needle!r}", needle in text)
        check("列表符号 / bullet marker present", "•" in text)
        check("阿拉伯数字编号 / decimal list numbering", "1. 完成第四季度预算评审。" in text)
        check("中文编号 / chinese list numbering", "一、第一阶段：渠道梳理。" in text)
        check("页眉 / header rendered", "季度业务报告 · 内部资料" in text)
        check("页脚页码 / footer page field", "第 1 页，共 2 页" in text)
        check("符号字符 / symbol glyphs", "→" in text and "★" in text and "①" in text)
        check("引用段落两端对齐 / justified paragraph indented",
              "这是引用段落，用于检查缩进与对齐效果。" in text)

        import pymupdf

        with pymupdf.open(cn_pdf) as doc:
            images = len(doc[0].get_images())
        check("图片已嵌入 / image embedded", images >= 1, f"images={images}")

        en_text = pdf_flat_text(en_pdf)
        for needle in ("Product Requirements Document", "Background", "Latency", "98%"):
            check(f"英文样本包含 / EN text contains {needle!r}", needle in en_text)

        print("== 宏文档 / macro-enabled document ==")
        docm = os.path.join(work, "macro.docm")
        _make_docm(cn, docm)
        docm_out = os.path.join(work, "macro.pdf")
        run_cli("-o", docm_out, docm)
        check("docm 可转换 / .docm converts",
              os.path.exists(docm_out) and "季度业务报告" in pdf_text(docm_out))

        print("== 输出路径 / output paths ==")
        single = os.path.join(work, "single.pdf")
        run_cli("-o", single, cn)
        check("-o 指定文件 / -o writes a file", os.path.exists(single))

        nested = os.path.join(work, "tree", "a", "b")
        os.makedirs(nested, exist_ok=True)
        shutil.copy(cn, os.path.join(nested, "deep.docx"))
        shutil.copy(cn, os.path.join(work, "tree", "top.docx"))
        tree_out = os.path.join(work, "tree-out")
        run_cli("-r", "-d", tree_out, os.path.join(work, "tree"))
        check("递归保持目录结构 / recursion keeps structure",
              os.path.exists(os.path.join(tree_out, "a", "b", "deep.pdf")) and
              os.path.exists(os.path.join(tree_out, "top.pdf")))
        top_out = os.path.join(work, "top-out")
        run_cli("--no-recursive", "-d", top_out, os.path.join(work, "tree"))
        check("--no-recursive 只处理第一层 / only top level",
              os.path.exists(os.path.join(top_out, "top.pdf")) and
              not os.path.exists(os.path.join(top_out, "a")))

        print("== 行为 / behaviour ==")
        before = os.path.getmtime(single)
        proc = run_cli("--no-clobber", "-o", single, cn)
        check("--no-clobber 跳过 / skips existing", "跳过" in proc.stdout,
              proc.stdout.strip()[-200:])
        check("--no-clobber 不重写 / file untouched", os.path.getmtime(single) == before)

        dry = os.path.join(work, "dry")
        run_cli("--dry-run", "-d", dry, cn)
        check("--dry-run 不写文件 / dry-run writes nothing", not os.path.exists(dry))

        run_cli("--doctor")
        run_cli("--engines")
        run_cli("--fonts")

        proc = run_cli("不存在的文件.docx", expect=2)
        check("缺失输入返回 2 / missing input exits 2", proc.returncode == 2)

        junk = os.path.join(work, "junk.xyz")
        with open(junk, "w", encoding="utf-8") as handle:
            handle.write("hello")
        proc = run_cli(junk, expect=2)
        check("不支持的类型返回 2 / unsupported input exits 2", proc.returncode == 2)

        print("== 引擎 / engines ==")
        from d2p.engines import LibreOfficeEngine

        if LibreOfficeEngine().available():
            lo_out = os.path.join(work, "lo.pdf")
            run_cli("--engine", "libreoffice", "-o", lo_out, cn)
            check("LibreOffice 引擎可用 / libreoffice engine works",
                  os.path.exists(lo_out) and pdf_pages(lo_out) >= 1)
        else:
            proc = run_cli("--engine", "libreoffice", "-o", os.path.join(work, "x.pdf"), cn,
                           expect=1)
            check("缺少 LibreOffice 时给出提示 / helpful error when LO is missing",
                  "LibreOffice" in (proc.stdout + proc.stderr))
            print("  (跳过 LibreOffice 转换: 未安装 / skipped: not installed)")

        print("== Python API ==")
        import d2p

        api_target = os.path.join(work, "api.pdf")
        api_result = d2p.convert_file(cn, api_target)
        check("convert_file 可用 / convert_file works",
              api_result.ok and os.path.exists(api_target),
              api_result.error)
        check("available_engines 报告 builtin / lists builtin",
              "builtin" in d2p.available_engines())
        broken = d2p.convert_file(os.path.join(work, "missing.docx"),
                                  os.path.join(work, "missing.pdf"))
        check("API 以结果返回错误 / API reports errors without raising",
              (not broken.ok) and bool(broken.error))
        batch = d2p.convert_many([cn, en], outdir=os.path.join(work, "api-batch"))
        check("convert_many 批量 / convert_many batch",
              len(batch) == 2 and all(r.ok for r in batch))

        print("== 内置引擎细节 / built-in details ==")
        from d2p.docxmodel import read_docx

        model = read_docx(cn)
        check("表格解析 / table parsed",
              any(type(b).__name__ == "Table" for b in model.blocks))
        check("页面设置为 A4 / A4 page size",
              abs(model.section.page_width - 595.28) < 3)
        check("无未支持图形警告 / no unsupported-object warnings",
              not any("EMF" in w for w in model.warnings), str(model.warnings))

    finally:
        shutil.rmtree(work, ignore_errors=True)

    failed = [name for name, ok, _ in _results if not ok]
    print("-" * 52)
    print(f"通过 / passed: {len(_results) - len(failed)}/{len(_results)}")
    if failed:
        print("失败 / failed:")
        for name in failed:
            print(f"  - {name}")
        return 1
    print("全部通过 / all checks passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
