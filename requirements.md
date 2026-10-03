# Requirements

## Purpose

Provide a desktop Python application for organizing Google Photos Takeout ZIP exports into a year/month library while preserving album groupings.

## User workflow

1. Launch the application with `python main.py`.
2. Select the folder containing one or more Takeout `.zip` files, the temporary extraction folder, and the output folder.
3. Start the operation only after all three locations are selected.

## Functional requirements

- Search the selected ZIP folder recursively for `.zip` files.
- Extract every archive into a unique run folder under the selected temporary folder. Never delete or modify source archives; keep archive contents isolated to prevent same-path overwrites.
- Reject unsafe archive paths that could extract outside the temporary folder.
- Replace Windows-invalid characters in ZIP path components during extraction and log the original and sanitized names.
- If an individual ZIP member has a CRC/read error, log it as an error, remove any partial extracted file, and continue with remaining members and archives.
- Recognize common photo and video extensions, including the formats used by Google Photos Takeout.
- Accept only dates with years from 1950 through 2100, inclusive, from every source.
- Determine media date in this order: supported filename timestamp, embedded media EXIF date, matching video `.thm`/`.THM` EXIF date, video container `creation_time`, Google Takeout `creationTime.timestamp`, then `photoTakenTime.timestamp`. Do not use filesystem modification time as a date source.
- Support the existing camera/WhatsApp patterns plus dotted WhatsApp timestamps, `DDMMYYYY` with optional trailing camera digits, and unpadded `VideoYYYYMDD` names.
- Use `ffprobe` when available on `PATH` to inspect video container/stream `creation_time` tags; continue to Takeout JSON dates when ffprobe is unavailable or the video has no such tag.
- Support Takeout sidecar JSON files named `<media filename>.json`, supplemental metadata sidecars, and `<media stem>.json`.
- Find THM sidecars by replacing the video suffix with `.thm` or `.THM`, matching the existing `sync_image_dates.py` behavior.
- Copy dated media to `<output>/<year>/<year> <Month>/<filename>`.
- For media inside an album folder, also copy it to `<output>/<year>/<album name>/<filename>`. Album names containing a year, such as `2026 - trip`, are preserved.
- Store any media whose path or filename identifies it as a collage under `<output>/<year>/collage/` as well as its month folder.
- Copy media with no usable filename, EXIF, or Takeout date to `<output>/unknown date/`.
- At an output destination, skip a file when an existing file has the same name and byte size; mark it as pre-existing in the audit. If the name exists with a different size, preserve both using a numbered suffix.
- Do not copy JSON or THM sidecars as media. Report archives, media found, month copies, album copies, and unknown-date copies.
- Write `media_organization_report.xlsx` in the output folder with one audit row per media file and columns for serial number, source path/name, filename date, media EXIF date, THM date, media-created date, Google Photos date, original archive location, resulting output locations, and whether an identical name/size existed.
- Keep the GUI responsive during extraction and organization, and report completion or errors.
- Show an activity log in the GUI with the ZIP currently extracting and each media file currently being processed.
- Render error log entries distinctly and include the skipped extraction error count in the completion summary.

## Non-functional requirements

- Windows desktop compatibility; use Python 3.10 or newer, Tkinter, Pillow, and openpyxl. `ffprobe` is optional but enables embedded video creation dates.
- Use a temporary run directory so repeat runs do not overwrite prior extracted content.
- Keep the organizer self-contained; do not modify the existing date-sync script.
- Provide automated unit and integration tests using Python's standard `unittest` framework.
- Provide a development dependency set with line-coverage measurement for the test suite.
- Provide GitHub-facing project documentation and a Windows CI test workflow.
- Publish under the name "Google Photos Takeout Organizer" with the MIT License.