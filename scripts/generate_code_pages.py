from __future__ import annotations

import argparse
import json
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


LINES_PER_PAGE = 56
PAGES_PER_DOC = 30
FULL_DOC_THRESHOLD = 60
CODE_FONT_SIZE = 7.5
CODE_LINE_SPACING_PT = 10

INCLUDE_EXTENSIONS = {
    ".java",
    ".kt",
    ".py",
    ".js",
    ".jsx",
    ".ts",
    ".tsx",
    ".css",
    ".scss",
    ".less",
    ".html",
    ".vue",
    ".xml",
    ".properties",
    ".sql",
    ".json",
    ".yaml",
    ".yml",
}

EXCLUDED_DIRS = {
    ".git",
    ".idea",
    ".vscode",
    ".next",
    ".nuxt",
    ".cache",
    ".turbo",
    ".parcel-cache",
    ".pytest_cache",
    "__pycache__",
    "node_modules",
    "dist",
    "build",
    "out",
    "target",
    "coverage",
    "vendor",
    "tmp",
    "temp",
    "docs",
    "screenshots",
    "软件著作权申请资料",
}

PRIORITY_PARTS = [
    ("java-server", "src"),
    ("src",),
    ("app",),
    ("server",),
    ("backend",),
    ("frontend",),
    ("java-server", "resources"),
    ("resources",),
    ("public",),
]


@dataclass(frozen=True)
class SourceFile:
    path: Path
    relative_path: str
    category: str


def set_run_font(run, size: float, bold: bool = False, ascii_font: str = "Courier New", east_asia: str = "SimSun") -> None:
    run.font.name = ascii_font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), east_asia)
    run.font.size = Pt(size)
    run.bold = bold


def set_paragraph_body(paragraph) -> None:
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.left_indent = Pt(0)
    fmt.right_indent = Pt(0)
    fmt.first_line_indent = Pt(0)
    fmt.line_spacing = Pt(CODE_LINE_SPACING_PT)
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")


def count_lines(path: Path) -> int:
    return len(read_text(path).splitlines())


def infer_category(relative_path: str) -> str:
    lower = relative_path.lower()
    if lower.endswith((".java", ".kt")):
        return "后端源码"
    if lower.endswith((".ts", ".tsx", ".js", ".jsx", ".vue")):
        return "前端脚本"
    if lower.endswith((".css", ".scss", ".less")):
        return "界面样式"
    if lower.endswith(".html"):
        return "页面模板"
    if lower.endswith((".xml", ".properties", ".yaml", ".yml", ".json", ".sql")):
        return "配置与数据脚本"
    return "项目源码"


def is_excluded(path: Path, project_root: Path) -> bool:
    relative_parts = path.relative_to(project_root).parts
    return any(part in EXCLUDED_DIRS for part in relative_parts)


def priority_rank(relative_parts: tuple[str, ...]) -> int:
    parts_lower = tuple(part.lower() for part in relative_parts)
    for index, prefix in enumerate(PRIORITY_PARTS):
        if parts_lower[: len(prefix)] == prefix:
            return index
    return len(PRIORITY_PARTS)


def collect_source_files(project_root: Path) -> list[SourceFile]:
    files: list[SourceFile] = []
    for path in project_root.rglob("*"):
        if not path.is_file():
            continue
        if path.suffix.lower() not in INCLUDE_EXTENSIONS:
            continue
        if is_excluded(path, project_root):
            continue
        relative = path.relative_to(project_root).as_posix()
        files.append(SourceFile(path=path, relative_path=relative, category=infer_category(relative)))

    files.sort(key=lambda item: (priority_rank(tuple(item.path.relative_to(project_root).parts)), item.relative_path.lower()))
    return files


def build_listing_lines(files: Iterable[SourceFile]) -> list[str]:
    lines: list[str] = []
    for source in files:
        lines.append(f"// ===== File: {source.relative_path} =====")
        lines.append(f"// ===== Category: {source.category} =====")
        lines.extend(read_text(source.path).splitlines())
        lines.append("")
    return lines


def page_lines(lines: list[str], lines_per_page: int) -> list[list[str]]:
    total_pages = (len(lines) + lines_per_page - 1) // lines_per_page
    padded = lines + [""] * (total_pages * lines_per_page - len(lines))
    return [padded[index * lines_per_page : (index + 1) * lines_per_page] for index in range(total_pages)]


def add_code_page(doc: Document, page_content: list[str], add_break_after: bool) -> None:
    paragraph = doc.add_paragraph()
    set_paragraph_body(paragraph)
    run = paragraph.add_run("\n".join(page_content))
    set_run_font(run, CODE_FONT_SIZE, ascii_font="Courier New", east_asia="SimSun")
    if add_break_after:
        doc.add_page_break()


def init_doc(title: str, system_name: str, version: str) -> Document:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.25)
    section.bottom_margin = Cm(1.25)
    section.left_margin = Cm(1.2)
    section.right_margin = Cm(1.2)
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "SimSun"
    normal_style._element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")
    normal_style.font.size = Pt(12)

    heading = doc.add_paragraph()
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = heading.add_run(title)
    set_run_font(run, 16, bold=True, ascii_font="SimSun", east_asia="SimSun")

    for line in [
        f"软件名称：{system_name}",
        f"版本号：{version}",
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(line)
        set_run_font(run, 12, ascii_font="SimSun", east_asia="SimSun")

    doc.add_page_break()
    return doc


def build_combined_doc(system_name: str, version: str, front_page_numbers: list[int], front_pages: list[list[str]], back_page_numbers: list[int], back_pages: list[list[str]], output_path: Path) -> None:
    doc = init_doc(f"{system_name} - 代码材料（前30页+后30页）", system_name, version)

    seen_pages = set(front_page_numbers)
    unique_back = [(page_number, back_pages[index]) for index, page_number in enumerate(back_page_numbers) if page_number not in seen_pages]

    for index, _page_number in enumerate(front_page_numbers):
        is_last_front = index == len(front_page_numbers) - 1
        add_break = not is_last_front or bool(unique_back)
        add_code_page(doc, front_pages[index], add_break)

    if unique_back:
        for index, (_page_number, page_content) in enumerate(unique_back):
            add_code_page(doc, page_content, index != len(unique_back) - 1)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)


def build_full_doc(system_name: str, version: str, all_pages: list[list[str]], output_path: Path) -> None:
    doc = init_doc(f"{system_name} - 全量代码材料", system_name, version)
    for index, page_content in enumerate(all_pages, start=1):
        add_code_page(doc, page_content, index != len(all_pages))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)


def convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> bool:
    escaped_docx = str(docx_path).replace("'", "''")
    escaped_pdf = str(pdf_path).replace("'", "''")
    script = (
        "$ErrorActionPreference='Stop';"
        "$word=New-Object -ComObject Word.Application;"
        "$word.Visible=$false;"
        f"$doc=$word.Documents.Open('{escaped_docx}');"
        f"$doc.SaveAs([ref]'{escaped_pdf}', [ref]17);"
        "$doc.Close();"
        "$word.Quit();"
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", script],
            check=True,
            capture_output=True,
            text=True,
        )
        return True
    except subprocess.CalledProcessError:
        return False


def load_manifest(path: Path | None) -> dict:
    if path and path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return {}


def build_application_info(system_name: str, version: str, owner: str, source_files: list[SourceFile], total_lines: int, total_pages: int, languages: list[str]) -> str:
    backend_lines = sum(count_lines(item.path) for item in source_files if item.path.suffix.lower() in {".java", ".kt"})
    frontend_lines = total_lines - backend_lines
    extras = []
    if total_pages > FULL_DOC_THRESHOLD:
        extras.append("代码超过60页，已补充全量代码Word文档。")
        extras.append("已输出前后60页合并版PDF代码文档。")
    return "\n".join(
        [
            f"软件全称：{system_name}",
            f"版本号：{version}",
            f"著作权人：{owner}",
            f"编程语言：{'、'.join(languages) if languages else '未识别'}",
            f"源程序量：{total_lines} 行（完整项目物理总行数）",
            f"其中后端源码：{backend_lines} 行",
            f"其中前端及其他源码：{frontend_lines} 行",
            f"源码文件数：{len(source_files)}",
            f"代码材料总页数：{total_pages} 页",
            f"每页代码行数目标：54-58 行（当前默认 {LINES_PER_PAGE} 行）",
            *extras,
        ]
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate one combined front-30-plus-back-30 software copyright code document.")
    parser.add_argument("--project", required=True, help="Project root path")
    parser.add_argument("--manifest", help="Manifest JSON path")
    parser.add_argument("--system-name", help="Software name shown in the documents")
    parser.add_argument("--version", help="Software version")
    parser.add_argument("--owner", help="Owner name for the application info text")
    parser.add_argument("--output-root", help="Output directory, defaults to <project>/软件著作权申请资料/正式资料")
    parser.add_argument("--lines-per-page", type=int, default=LINES_PER_PAGE, help="Lines per page, default 50")
    parser.add_argument("--pages-per-doc", type=int, default=PAGES_PER_DOC, help="Pages per document, default 30")
    parser.add_argument("--skip-app-info", action="store_true", help="Do not emit 申请表信息.txt")
    parser.add_argument("--skip-pdf", action="store_true", help="Do not emit PDF for the combined code document")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    project_root = Path(args.project).resolve()
    manifest = load_manifest(Path(args.manifest).resolve() if args.manifest else (project_root / "softcopyright-manifest.json"))

    system_name = args.system_name or manifest.get("system_name") or project_root.name
    version = args.version or manifest.get("version") or "V1.0"
    owner = args.owner or manifest.get("owner") or "项目开发团队"
    output_root = Path(args.output_root).resolve() if args.output_root else project_root / "软件著作权申请资料" / "正式资料"

    source_files = collect_source_files(project_root)
    if not source_files:
        raise SystemExit("No source files matched the built-in rules.")

    listing_lines = build_listing_lines(source_files)
    paged = page_lines(listing_lines, args.lines_per_page)
    total_pages = len(paged)
    total_lines = sum(count_lines(item.path) for item in source_files)

    front_page_numbers = list(range(1, min(args.pages_per_doc, total_pages) + 1))
    back_start = max(1, total_pages - args.pages_per_doc + 1)
    back_page_numbers = list(range(back_start, total_pages + 1))
    front_pages = [paged[number - 1] for number in front_page_numbers]
    back_pages = [paged[number - 1] for number in back_page_numbers]

    combined_docx = output_root / f"{system_name}-代码(前30页+后30页).docx"
    combined_pdf = output_root / f"{system_name}-代码(前30页+后30页).pdf"
    build_combined_doc(system_name, version, front_page_numbers, front_pages, back_page_numbers, back_pages, combined_docx)

    full_docx = None
    if total_pages > FULL_DOC_THRESHOLD:
        full_docx = output_root / f"{system_name}-代码(全部).docx"
        build_full_doc(system_name, version, paged, full_docx)

    pdf_created = False
    if not args.skip_pdf:
        pdf_created = convert_docx_to_pdf(combined_docx, combined_pdf)

    languages = manifest.get("environments", {}).get("languages", [])
    app_info_path = output_root / "申请表信息.txt"
    if not args.skip_app_info:
        app_info_path.write_text(
            build_application_info(system_name, version, owner, source_files, total_lines, total_pages, languages),
            encoding="utf-8",
        )

    print(f"COMBINED_DOCX={combined_docx}")
    print(f"COMBINED_PDF={combined_pdf if pdf_created else 'NOT_CREATED'}")
    print(f"FULL_DOCX={full_docx if full_docx else 'NOT_REQUIRED'}")
    if not args.skip_app_info:
        print(f"APP_INFO={app_info_path}")
    print(f"SOURCE_FILES={len(source_files)}")
    print(f"TOTAL_LINES={total_lines}")
    print(f"TOTAL_PAGES={total_pages}")
    print(f"LINES_PER_PAGE={args.lines_per_page}")


if __name__ == "__main__":
    main()
