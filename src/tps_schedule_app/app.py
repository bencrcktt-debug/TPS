from __future__ import annotations

from pathlib import Path
import argparse
import sys

from .docx_schedule import (
    build_day_label_map,
    load_template_metadata,
    sanitize_filename,
    write_schedule_docx,
)
from .excel_schedule import load_workbook_schedule


def generate_documents(
    excel_path: str | Path,
    template_path: str | Path,
    output_dir: str | Path,
    *,
    sheet_label_overrides: dict[str, str] | None = None,
    filename_pattern: str | None = None,
) -> list[Path]:
    workbook = load_workbook_schedule(excel_path)
    template = load_template_metadata(template_path)
    day_label_map = build_day_label_map(
        workbook.sheet_names,
        template.day_labels,
        overrides=sheet_label_overrides,
    )
    filename_pattern = filename_pattern or template.filename_pattern
    if "{name}" not in filename_pattern:
        raise ValueError("The filename pattern must include a {name} placeholder.")

    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)

    created_files = []
    for person in workbook.people:
        output_name = sanitize_filename(filename_pattern.format(name=person.name))
        output_path = destination / output_name
        write_schedule_docx(
            person=person,
            template_path=template_path,
            output_path=output_path,
            title_prefix=template.title_prefix,
            day_label_map=day_label_map,
        )
        created_files.append(output_path)

    return created_files


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    if args.gui or (argv is None and len(sys.argv) == 1):
        _launch_gui()
        return 0

    if not args.excel or not args.template or not args.output_dir:
        parser.error("--excel, --template, and --output-dir are required when running without --gui.")

    created_files = generate_documents(
        excel_path=args.excel,
        template_path=args.template,
        output_dir=args.output_dir,
        sheet_label_overrides=_parse_sheet_label_overrides(args.sheet_label),
        filename_pattern=args.filename_pattern,
    )

    print(f"Generated {len(created_files)} schedule document(s) in {Path(args.output_dir).resolve()}")
    return 0


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert a TPS master Excel schedule into individual DOCX schedules.",
    )
    parser.add_argument("--excel", type=Path, help="Path to the master Excel workbook.")
    parser.add_argument("--template", type=Path, help="Path to the sample DOCX template.")
    parser.add_argument("--output-dir", type=Path, help="Folder where generated DOCX files will be written.")
    parser.add_argument(
        "--sheet-label",
        action="append",
        default=[],
        metavar="SHEET=LABEL",
        help="Override a day heading, for example: --sheet-label WED=Wednesday, April 8th",
    )
    parser.add_argument(
        "--filename-pattern",
        help="Output filename pattern. Include {name}, for example: '{name} TPS 2026.docx'",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch the desktop interface instead of running from CLI arguments.",
    )
    return parser


def _parse_sheet_label_overrides(entries: list[str]) -> dict[str, str]:
    overrides = {}
    for entry in entries:
        sheet_name, separator, label = entry.partition("=")
        if not separator or not sheet_name.strip() or not label.strip():
            raise ValueError(f"Invalid --sheet-label value: {entry!r}")
        overrides[sheet_name.strip()] = label.strip()
    return overrides


def _launch_gui() -> None:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except Exception as exc:  # pragma: no cover
        raise RuntimeError("Tkinter is required to launch the desktop interface.") from exc

    root = tk.Tk()
    root.title("TPS Schedule Generator")
    root.resizable(False, False)

    excel_var = tk.StringVar()
    template_var = tk.StringVar()
    output_var = tk.StringVar()
    status_var = tk.StringVar(value="Select the source workbook, the sample DOCX, and an output folder.")

    def browse_excel() -> None:
        selected = filedialog.askopenfilename(
            title="Choose the master Excel workbook",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if selected:
            excel_var.set(selected)

    def browse_template() -> None:
        selected = filedialog.askopenfilename(
            title="Choose the sample DOCX template",
            filetypes=[("Word document", "*.docx")],
        )
        if selected:
            template_var.set(selected)

    def browse_output() -> None:
        selected = filedialog.askdirectory(title="Choose the output folder")
        if selected:
            output_var.set(selected)

    def generate() -> None:
        try:
            if not excel_var.get() or not template_var.get() or not output_var.get():
                raise ValueError("All three paths are required.")
            created = generate_documents(
                excel_path=excel_var.get(),
                template_path=template_var.get(),
                output_dir=output_var.get(),
            )
        except Exception as exc:
            status_var.set(str(exc))
            messagebox.showerror("Generation failed", str(exc))
            return

        summary = f"Generated {len(created)} schedule document(s) in {Path(output_var.get()).resolve()}"
        status_var.set(summary)
        messagebox.showinfo("Generation complete", summary)

    frame = tk.Frame(root, padx=12, pady=12)
    frame.grid(row=0, column=0, sticky="nsew")

    tk.Label(frame, text="Excel workbook").grid(row=0, column=0, sticky="w", pady=(0, 6))
    tk.Entry(frame, textvariable=excel_var, width=58).grid(row=1, column=0, padx=(0, 8))
    tk.Button(frame, text="Browse", width=10, command=browse_excel).grid(row=1, column=1)

    tk.Label(frame, text="Sample DOCX").grid(row=2, column=0, sticky="w", pady=(12, 6))
    tk.Entry(frame, textvariable=template_var, width=58).grid(row=3, column=0, padx=(0, 8))
    tk.Button(frame, text="Browse", width=10, command=browse_template).grid(row=3, column=1)

    tk.Label(frame, text="Output folder").grid(row=4, column=0, sticky="w", pady=(12, 6))
    tk.Entry(frame, textvariable=output_var, width=58).grid(row=5, column=0, padx=(0, 8))
    tk.Button(frame, text="Browse", width=10, command=browse_output).grid(row=5, column=1)

    tk.Button(frame, text="Generate schedules", command=generate, width=24).grid(
        row=6,
        column=0,
        columnspan=2,
        pady=(16, 8),
    )
    tk.Label(frame, textvariable=status_var, wraplength=500, justify="left").grid(
        row=7,
        column=0,
        columnspan=2,
        sticky="w",
    )

    root.mainloop()
