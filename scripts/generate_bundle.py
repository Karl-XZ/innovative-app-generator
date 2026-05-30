from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from datetime import date
from pathlib import Path

from generate_code_pages import (
    build_tech_features_text,
    collect_source_files,
    infer_environment,
    infer_feature_categories,
    infer_main_functions,
    infer_purpose_and_industry,
)


def run(command: list[str]) -> None:
    subprocess.run(command, check=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate the complete software copyright bundle from a project.")
    parser.add_argument("--project", required=True, help="Project root path")
    parser.add_argument("--manifest", help="Optional manifest JSON path; missing fields will be auto-filled")
    parser.add_argument("--output-root", help="Output directory, defaults to <project>/软件著作权申请资料/正式资料")
    return parser.parse_args()


def json_load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def json_dump(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def extract_docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        try:
            raw = archive.read("word/document.xml")
        except KeyError:
            return ""
    return raw.decode("utf-8", errors="replace")


def collect_rendered_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".json", ".txt", ".html", ".md"}:
        return path.read_text(encoding="utf-8", errors="replace")
    if suffix == ".docx":
        return extract_docx_text(path)
    return ""


def find_garble_markers(text: str, *, allow_code_question_runs: bool = False) -> list[str]:
    markers: list[str] = []
    if re.search(r"\?{2,}", text) and not allow_code_question_runs:
        markers.append("question-runs")
    if "\ufffd" in text:
        markers.append("replacement-char")
    suspicious_tokens = [
        "截图文件：",
        "请结合代码段中的条件判断",
        "同样的算法口径与代码对应关系",
    ]
    for token in suspicious_tokens:
        if token in text:
            markers.append(token)
    return markers


def validate_rendered_artifact(path: Path) -> None:
    text = collect_rendered_text(path)
    if not text:
        return
    allow_code_question_runs = path.suffix.lower() == ".docx" and "代码" in path.name
    markers = find_garble_markers(text, allow_code_question_runs=allow_code_question_runs)
    if markers:
        raise RuntimeError(f"Artifact validation failed for {path}: {', '.join(markers)}")


def clean_filename_stem(name: str) -> str:
    stem = Path(name).stem
    stem = re.sub(r"^[0-9]+[-_ ]*", "", stem)
    stem = stem.replace("_", " ").replace("-", " ").strip()
    return stem or "界面"


def title_from_name(name: str) -> str:
    mapping = {
        "home": "首页",
        "overview": "综合看板",
        "dashboard": "综合看板",
        "patient": "患者管理",
        "patients": "患者管理",
        "appointment": "预约挂号",
        "appointments": "预约挂号",
        "consultation": "在线问诊",
        "schedule": "医生排班",
        "record": "电子病历",
        "records": "电子病历",
        "prescription": "处方管理",
        "prescriptions": "处方管理",
        "billing": "支付结算",
        "analytics": "数据统计",
        "setting": "系统设置",
        "settings": "系统设置",
        "followup": "随访管理",
        "followups": "随访管理",
        "operation": "运营调度",
        "operations": "运营调度",
        "triage": "智能分诊",
    }
    lowered = clean_filename_stem(name).lower()
    for key, value in mapping.items():
        if key in lowered:
            return value
    return clean_filename_stem(name)


def discover_screenshots(project_root: Path) -> list[dict]:
    preferred_dirs = [
        project_root / "docs" / "screenshots",
        project_root / "docs" / "softcopyright-assets",
    ]
    results: list[dict] = []
    seen: set[str] = set()
    for root in preferred_dirs:
        if not root.exists():
            continue
        for path in sorted(root.rglob("*")):
            if path.suffix.lower() not in {".png", ".jpg", ".jpeg", ".webp"}:
                continue
            relative = path.relative_to(project_root).as_posix()
            lowered = relative.lower()
            if any(token in lowered for token in ("review-", "page-", "section-", "thumb", "thumbnail")):
                continue
            title = title_from_name(path.stem)
            if title in seen:
                continue
            seen.add(title)
            results.append({"file": relative, "title": title})
    return results


def guess_route(title: str) -> str:
    routes = {
        "首页": "/",
        "综合看板": "/",
        "患者管理": "/patients",
        "预约挂号": "/appointments",
        "在线问诊": "/consultation",
        "医生排班": "/schedule",
        "电子病历": "/records",
        "处方管理": "/prescriptions",
        "支付结算": "/billing",
        "数据统计": "/analytics",
        "系统设置": "/settings",
        "智能分诊": "/triage",
        "随访管理": "/followup",
        "运营调度": "/operations",
    }
    return routes.get(title, f"/{title}")


def discover_code_refs(project_root: Path, source_files: list) -> list[dict]:
    refs: list[dict] = []
    patterns = [
        re.compile(r"^\s*(?:public|private|protected)?\s*(?:static\s+)?[\w<>\[\], ?]+\s+([A-Za-z_]\w*)\s*\([^;]*\)\s*\{"),
        re.compile(r"^\s*(?:const|let|var)\s+([A-Za-z_]\w*)\s*=\s*(?:async\s*)?\([^)]*\)\s*=>"),
        re.compile(r"^\s*function\s+([A-Za-z_]\w*)\s*\("),
    ]
    for source in source_files:
        if len(refs) >= 10:
            break
        lines = source.path.read_text(encoding="utf-8", errors="replace").splitlines()
        for index, line in enumerate(lines, start=1):
            matched_name = ""
            for pattern in patterns:
                match = pattern.search(line)
                if match:
                    matched_name = match.group(1)
                    break
            if not matched_name:
                continue
            refs.append(
                {
                    "file": source.relative_path,
                    "title": f"{matched_name}核心逻辑",
                    "function_name": matched_name,
                    "start_line": index,
                    "end_line": min(len(lines), index + 24),
                    "explanation": f"{matched_name}承担该模块的关键业务判断、数据整理或结果构造。",
                    "input": "输入当前模块所需的业务数据、状态参数或界面提交信息。",
                    "output": "输出模块结果对象、渲染数据或接口响应内容。",
                    "variables": [],
                    "line_explanations": [],
                }
            )
            break
    return refs


def infer_user_groups(system_name: str, industry: str) -> list[str]:
    lowered = f"{system_name}{industry}".lower()
    if "医院" in lowered or "医疗" in lowered:
        return ["管理人员", "医生", "护士", "运营人员", "患者"]
    if "电梯" in lowered:
        return ["运维人员", "值班人员", "管理人员", "监控人员"]
    return ["管理员", "业务人员", "终端用户"]


def validate_manifest_for_manual(manifest: dict) -> None:
    modules = manifest.get("modules") or []
    if not modules:
        raise RuntimeError("Manifest validation failed: modules 不能为空，不能再由生成器自动填充模板化模块说明。")

    required_module_fields = ["summary", "algorithm", "input_output", "data_flow", "special_handling", "steps", "code_refs"]
    for index, module in enumerate(modules, start=1):
        missing = [field for field in required_module_fields if not module.get(field)]
        if missing:
            raise RuntimeError(
                f"Manifest validation failed: 第 {index} 个模块 {module.get('title', '<未命名模块>')} 缺少字段 {', '.join(missing)}。"
            )

    if not manifest.get("exports"):
        raise RuntimeError("Manifest validation failed: exports 不能为空，不能使用通用导出占位文案。")

    if not manifest.get("faq"):
        raise RuntimeError("Manifest validation failed: faq 不能为空，不能使用通用故障排除模板。")

    tests = manifest.get("tests") or {}
    if not tests.get("built_in") or not tests.get("cases"):
        raise RuntimeError("Manifest validation failed: tests.built_in 与 tests.cases 必须填写真实测试内容。")

    maintenance = manifest.get("maintenance") or {}
    if not maintenance.get("guide") or not maintenance.get("items"):
        raise RuntimeError("Manifest validation failed: maintenance.guide 与 maintenance.items 必须填写真实维护内容。")


def build_synthesized_manifest(project_root: Path, manifest_path: Path) -> dict:
    base = {}
    if manifest_path.exists():
        try:
            base = json_load(manifest_path)
        except Exception:
            base = {}
    source_files = collect_source_files(project_root)
    system_name = base.get("system_name") or project_root.name
    version = base.get("version") or "V1.0"
    env = infer_environment(project_root, base, source_files)
    purpose, industry = infer_purpose_and_industry(system_name, base, project_root)
    screenshots = discover_screenshots(project_root)
    code_refs = discover_code_refs(project_root, source_files)
    modules = list(base.get("modules") or [])
    synthesized = {
        "system_name": system_name,
        "version": version,
        "owner": base.get("owner") or "项目开发团队",
        "complete_date": base.get("complete_date") or date.today().isoformat(),
        "purpose": purpose,
        "industry": industry,
        "user_groups": base.get("user_groups") or infer_user_groups(system_name, industry),
        "architecture": base.get("architecture") or "前后端分离 B/S 架构，前端负责界面与交互，后端负责业务规则、数据组织与导出接口。",
        "environments": env,
        "technical_feature_categories": base.get("technical_feature_categories")
        or [item for item in infer_feature_categories(base, system_name, industry).split("、") if item],
        "technical_features": base.get("technical_features")
        or [part.strip() for part in build_tech_features_text(base).split("；") if part.strip()],
        "main_functions": base.get("main_functions") or infer_main_functions(project_root, base),
        "modules": modules,
        "exports": base.get("exports") or [],
        "faq": base.get("faq") or [],
        "tests": base.get("tests") or {},
        "maintenance": base.get("maintenance") or {},
    }
    return synthesized


def ensure_manifest(project_root: Path, manifest_path: Path) -> Path:
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest = build_synthesized_manifest(project_root, manifest_path)
    json_dump(manifest_path, manifest)
    return manifest_path


def main() -> None:
    args = parse_args()
    project_root = Path(args.project).resolve()
    manifest = Path(args.manifest).resolve() if args.manifest else project_root / "softcopyright-manifest.json"
    output_root = Path(args.output_root).resolve() if args.output_root else project_root / "软件著作权申请资料" / "正式资料"
    scripts_dir = Path(__file__).resolve().parent
    output_root.mkdir(parents=True, exist_ok=True)
    staging_root = Path(tempfile.mkdtemp(prefix="innovative-app-bundle-"))

    manifest = ensure_manifest(project_root, manifest)
    validate_manifest_for_manual(json_load(manifest))
    validate_rendered_artifact(manifest)

    run(
        [
            sys.executable,
            str(scripts_dir / "generate_html_manual_from_manifest.py"),
            "--project",
            str(project_root),
            "--manifest",
            str(manifest),
            "--output-root",
            str(staging_root),
        ]
    )
    run(
        [
            sys.executable,
            str(scripts_dir / "generate_code_pages.py"),
            "--project",
            str(project_root),
            "--manifest",
            str(manifest),
            "--output-root",
            str(output_root),
        ]
    )

    manifest_data = json_load(manifest)
    system_name = manifest_data.get("system_name") or project_root.name

    manual_docx = staging_root / f"{system_name}_\u8f6f\u4ef6\u8457\u4f5c\u6743\u9274\u5b9a\u6750\u6599.docx"
    manual_pdf = staging_root / f"{system_name}_\u8f6f\u4ef6\u8457\u4f5c\u6743\u9274\u5b9a\u6750\u6599.pdf"

    final_map = {
        manual_docx: output_root / f"{system_name}\u624b\u518c.docx",
        manual_pdf: output_root / f"{system_name}\u624b\u518c.pdf",
    }
    for src, dst in final_map.items():
        if src.exists():
            shutil.copy2(src, dst)
            validate_rendered_artifact(dst)

    txt_output = output_root / f"{system_name}.txt"
    code_docx_output = output_root / f"{system_name}代码（完整版）.docx"
    for artifact in [txt_output, code_docx_output]:
        if artifact.exists():
            validate_rendered_artifact(artifact)

    print(f"MANIFEST={manifest}")
    print(f"OUTPUT_ROOT={output_root}")


if __name__ == "__main__":
    main()
