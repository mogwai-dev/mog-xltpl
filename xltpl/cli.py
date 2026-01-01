import argparse
import sys
import re
from pathlib import Path

try:
    import yaml
except ImportError:
    yaml = None  # type: ignore

from jinja2 import Template

from .writer import BookWriter
from .writerx import BookWriter as BookWriterx


_TASKFILE_PLACEHOLDER_RE = re.compile(r"\{\{\s*\.(\w+)\s*\}\}")
_MAX_VAR_RENDER_PASSES = 5  # avoid infinite recursion when vars reference each other cyclically


def _normalize_taskfile_placeholders(obj):
    """Convert Taskfile-style {{ .VAR }} to Jinja2-style {{ VAR }} recursively."""
    if isinstance(obj, str):
        return _TASKFILE_PLACEHOLDER_RE.sub(r"{{ \1 }}", obj)
    if isinstance(obj, list):
        return [_normalize_taskfile_placeholders(v) for v in obj]
    if isinstance(obj, dict):
        return {k: _normalize_taskfile_placeholders(v) for k, v in obj.items()}
    return obj


def _merge_var_list(var_list):
    """Allow Taskfile-style list form: vars: - KEY: value."""
    merged = {}
    for item in var_list:
        if not isinstance(item, dict) or len(item) != 1:
            raise SystemExit("Each item in 'vars' list must be a single-key mapping")
        merged.update(item)
    return merged


def _render_inline_templates(obj, context: dict):
    """Render strings that still contain Jinja-style placeholders using the given context."""
    if isinstance(obj, str):
        if "{{" in obj or "{%" in obj:
            return Template(obj).render(context)
        return obj
    if isinstance(obj, list):
        return [_render_inline_templates(v, context) for v in obj]
    if isinstance(obj, dict):
        return {k: _render_inline_templates(v, context) for k, v in obj.items()}
    return obj


def _resolve_vars(payloads: dict):
    """Render templated values using the vars themselves, with a safety cap on passes."""
    current = payloads
    for _ in range(_MAX_VAR_RENDER_PASSES):
        rendered = _render_inline_templates(current, current)
        if rendered == current:
            break
        current = rendered
    return current


def load_yaml(yaml_path: Path):
    if yaml is None:
        raise SystemExit("PyYAML not installed. Install with: pip install pyyaml")
    with yaml_path.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def choose_writer(template_path: Path, highlight: bool = False, highlight_color: str | None = None):
    ext = template_path.suffix.lower()
    # Treat macro-enabled and template xlsx as xlsx family
    xlsx_family = {".xlsx", ".xlsm", ".xltx", ".xltm"}
    if ext in xlsx_family:
        return BookWriterx(str(template_path), highlight=highlight, highlight_color=highlight_color)
    # Fallback to legacy .xls writer
    return BookWriter(str(template_path), highlight=highlight, highlight_color=highlight_color)


def render_from_yaml(template_path: Path, output_path: Path, yaml_path: Path, highlight_output: Path | None = None, highlight_color: str | None = None):
    """Render Excel from template using vars from YAML config.

    Optionally emits a highlighted variant when highlight_output is given.
    """
    data = load_yaml(yaml_path)
    
    # Get vars (Taskfile-style)
    payloads = data.get("vars")
    if payloads is None:
        raise SystemExit("YAML must contain 'vars' section")
    if isinstance(payloads, list):
        payloads = _merge_var_list(payloads)
    if not isinstance(payloads, dict):
        raise SystemExit("'vars' must be a dict or a list of single-key dicts")
    payloads = _normalize_taskfile_placeholders(payloads)
    payloads = _resolve_vars(payloads)
    
    if not template_path.exists():
        raise SystemExit(f"Template file not found: {template_path}")
    
    # Create output directory if needed
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    def _render_sheet_name(name_tpl: str, ctx: dict) -> str:
        """Render template sheet name with Jinja using vars context."""
        rendered = Template(name_tpl).render(ctx)
        rendered = rendered.strip()
        if not rendered:
            raise SystemExit("Rendered sheet name is empty. Check your template sheet name.")
        if len(rendered) > 31 or any(ch in rendered for ch in "[]:*?/\\"):
            raise SystemExit(f"Invalid sheet name '{rendered}'. Excel sheet names must be <=31 chars and not contain []:*?/\\")
        return rendered

    def _build_payloads(writer_obj):
        """Build payloads for each sheet, rendering template sheet names with vars."""
        all_payloads_local = []
        for sheet_state in writer_obj.sheet_resource_map.sheet_resources:
            sheet_payload = dict(payloads)
            sheet_payload['tpl_index'] = sheet_state.index
            # Always render template sheet name with Jinja (e.g., {{ sheet_name }} -> hogehoge)
            sheet_payload['sheet_name'] = _render_sheet_name(sheet_state.name, payloads)
            all_payloads_local.append(sheet_payload)
        # Check for duplicate rendered sheet names
        names = [p.get('sheet_name') for p in all_payloads_local]
        dupes = {n for n in names if names.count(n) > 1}
        if dupes:
            dupes_list = ", ".join(sorted(dupes))
            raise SystemExit(f"Duplicate sheet_name detected after rendering: {dupes_list}. Ensure template sheet names render to unique values.")
        return all_payloads_local

    # Render Excel - template sheet names are rendered with Jinja
    writer = choose_writer(template_path, highlight=False, highlight_color=highlight_color)
    all_payloads = _build_payloads(writer)
    writer.render_book(all_payloads)
    writer.save(str(output_path))
    print(f"Rendered: {output_path}")

    if highlight_output:
        # Wait for COM cleanup before starting second save
        import time
        import gc
        gc.collect()
        time.sleep(2.0)
        
        highlight_output.parent.mkdir(parents=True, exist_ok=True)
        writer_hl = choose_writer(template_path, highlight=True, highlight_color=highlight_color)
        all_payloads_hl = _build_payloads(writer_hl)
        writer_hl.render_book(all_payloads_hl)
        writer_hl.save(str(highlight_output))
        print(f"Rendered (highlight): {highlight_output}")


def build_parser():
    p = argparse.ArgumentParser(description="Render Excel from template and variables (YAML)")
    p.add_argument("template", help="path to Excel template file")
    p.add_argument("output", help="path to output Excel file")
    p.add_argument("yaml", help="path to YAML config file (contains vars)")
    p.add_argument("--version", action="version", version="%(prog)s 1.0.0")
    p.add_argument("--highlight-output", action="store_true", help="also emit a highlighted copy using default name <output>_highlight", dest="highlight_output")
    p.add_argument("--highlight-color", help="highlight color (ARGB, e.g., FFFF9999)")
    return p


def main(argv: list[str] | None = None):
    args = build_parser().parse_args(argv)
    
    template_path = Path(args.template).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()
    yaml_path = Path(args.yaml).expanduser().resolve()
    highlight_output: Path | None
    if args.highlight_output:
        highlight_output = output_path.with_name(f"{output_path.stem}_highlight{output_path.suffix}")
    else:
        highlight_output = None
    highlight_color = args.highlight_color
    
    if not template_path.exists():
        raise SystemExit(f"Template file not found: {template_path}")
    if not yaml_path.exists():
        raise SystemExit(f"YAML file not found: {yaml_path}")
    
    # Reject .xls files (only .xlsx is supported)
    if template_path.suffix.lower() == '.xls':
        raise SystemExit(f"Error: .xls files are not supported. Please convert '{template_path.name}' to .xlsx format.")
    
    render_from_yaml(template_path, output_path, yaml_path, highlight_output=highlight_output, highlight_color=highlight_color)


if __name__ == "__main__":  # pragma: no cover
    main()
