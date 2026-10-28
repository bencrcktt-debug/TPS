# TPS Schedule Generator

This app converts the TPS master Excel staff schedule into one DOCX schedule per person, using an existing sample DOCX as the formatting template.

## What it does

- Reads the master `.xlsx` workbook directly from its Office XML files.
- Extracts each person's schedule from every day sheet.
- Infers event ranges from the colored schedule blocks in the workbook.
- Writes one `.docx` file per person while preserving the sample document container.

## Run it

### Desktop app

```sh
python main.py
```

This opens a small Tkinter window where you choose:

- the master Excel workbook
- the sample DOCX file
- the output folder

### Command line

```sh
python main.py --excel "C:\\path\\to\\master.xlsx" --template "C:\\path\\to\\sample.docx" --output-dir "C:\\path\\to\\output"
```

Optional overrides:

```sh
python main.py ^
  --excel "C:\\path\\to\\master.xlsx" ^
  --template "C:\\path\\to\\sample.docx" ^
  --output-dir "C:\\path\\to\\output" ^
  --sheet-label "WED=Wednesday, April 8th" ^
  --sheet-label "THURS=Thursday, April 9th" ^
  --sheet-label "FRI=Friday, April 10th" ^
  --filename-pattern "{name} TPS 2026.docx"
```

## Notes

- No external Python packages are required.
- The sample DOCX is used to infer the title prefix, day labels, and default filename pattern.
- Single-slot items are written as a single time. Multi-slot colored blocks are written as time ranges.

## Tests

```sh
python -m unittest discover -s tests -v
```
