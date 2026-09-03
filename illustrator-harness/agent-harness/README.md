# CLI-Anything Illustrator

`cli-anything-illustrator` controls Adobe Illustrator through its Windows COM
automation interface. It provides commands for documents, layers, vector
shapes, text, and export operations.

## Requirements

- Windows 10 or 11
- Adobe Illustrator 2023 or newer
- Python 3.10 or newer

## Install

From this directory, run:

```powershell
py -m pip install -e .
cli-anything-illustrator --help
```

The package installs `click` and `pywin32` automatically.

## Quick start

```powershell
# Create an unsaved 500 x 500 pt document, then save it.
cli-anything-illustrator project new --width 500 --height 500
cli-anything-illustrator project save logo.ai

# Work with the active document.
cli-anything-illustrator text add "Brand" --x 100 --y 100 --font Arial --font-size 72
cli-anything-illustrator shape rect --x 50 --y 50 --w 200 --h 100
cli-anything-illustrator export svg output.svg
```

Use `--json` before the command group to produce machine-readable output:

```powershell
cli-anything-illustrator --json project info
```

Running `cli-anything-illustrator` without a command starts the interactive
REPL. Illustrator itself is only contacted when a command needs the COM
backend, so `--help` remains available for installation diagnostics.

## Tests

The default tests do not launch Illustrator:

```powershell
py -m pytest cli_anything/illustrator/tests -v
```
