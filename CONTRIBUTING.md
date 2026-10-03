# Contributing

Thanks for helping improve Google Photos Takeout Organizer. The most useful contributions are reproducible bug reports, sample filename patterns, metadata edge cases, and focused fixes.

## Before Opening an Issue

- Search existing issues for the same problem.
- Include your operating system, Python version, and project version/commit if known.
- Describe the selected source/temp/output layout and the relevant activity-log message.
- For date bugs, provide the filename pattern and the expected date source. Do not attach personal photos, videos, Takeout ZIPs, sidecars, or report workbooks.
- When possible, reproduce the issue with a small synthetic media file or a redacted metadata snippet.

## Development Setup

Use Python 3.10 or later from the repository root:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements-dev.txt
```

Run tests before submitting:

```powershell
python -m unittest discover -s tests -v
python -m coverage run --source=takeout_organizer,date_helpers -m unittest discover -s tests
python -m coverage report -m
```

The project uses `unittest`; add regression tests in `tests/test_takeout_organizer.py`. Keep tests independent of private media by generating fixtures under `TemporaryDirectory` or mocking ffprobe.

## Pull Request Guidelines

- Keep changes focused and explain the user-visible behavior being changed.
- Add or update tests for date precedence, filename parsing, archive handling, workbook values, or GUI behavior affected by the change.
- Update `README.md` when supported formats, installation, behavior, or output changes.
- Do not add private media, ZIP archives, temporary extraction folders, logs containing personal paths, generated reports, or credentials.
- Include test results and note any manual GUI checks that were not performed.

## Filename and Metadata Requests

When suggesting a new filename format, include a few representative names, the expected parsed date, and any ambiguous cases. Date sources must preserve the documented precedence and reject years outside 1950–2100.