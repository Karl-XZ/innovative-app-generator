from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def run(command: list[str]) -> None:
    subprocess.run(command, check=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate the complete software copyright bundle from a project and manifest.")
    parser.add_argument("--project", required=True, help="Project root path")
    parser.add_argument("--manifest", help="Manifest JSON path, defaults to <project>/softcopyright-manifest.json")
    parser.add_argument("--output-root", help="Output directory, defaults to <project>/软件著作权申请资料/正式资料")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    project_root = Path(args.project).resolve()
    manifest = Path(args.manifest).resolve() if args.manifest else project_root / "softcopyright-manifest.json"
    output_root = Path(args.output_root).resolve() if args.output_root else project_root / "软件著作权申请资料" / "正式资料"
    scripts_dir = Path(__file__).resolve().parent

    run([sys.executable, str(scripts_dir / "generate_html_manual_from_manifest.py"), "--project", str(project_root), "--manifest", str(manifest), "--output-root", str(output_root)])
    run([sys.executable, str(scripts_dir / "generate_code_pages.py"), "--project", str(project_root), "--manifest", str(manifest), "--output-root", str(output_root)])

    print(f"OUTPUT_ROOT={output_root}")


if __name__ == "__main__":
    main()
