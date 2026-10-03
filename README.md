# Google Photos Takeout Organizer

> A local Windows app that organizes Google Photos Takeout ZIPs into a date-sorted photo and video library, preserving albums and producing an Excel audit report.

Google Photos Takeout Organizer is a Windows desktop app for turning Google Photos Takeout ZIP archives into an organized photo and video library. It keeps album copies, uses the best available capture date, records every decision in an Excel workbook, and leaves the source ZIP files untouched.

This is an independent community project and is not affiliated with or endorsed by Google LLC. Suggested GitHub repository name: `google-photos-takeout-organizer`.

## What It Does

- Finds `.zip` files recursively in the folder you select.
- Extracts each archive into its own temporary directory so files from different ZIPs cannot overwrite one another.
- Organizes dated photos and videos by year and month, with additional copies for albums and collages.
- Uses filename dates first, then embedded EXIF, matching video THM EXIF, video creation metadata, and Google Photos JSON dates.
- Puts files without a supported date in `unknown date` instead of guessing from filesystem dates.
- Skips an output file if the destination already contains the same filename and byte size. A same-name file with a different size is retained and the incoming copy receives a numbered suffix.
- Writes `media_organization_report.xlsx` with source dates, original locations, output locations, and duplicate decisions.
- Shows extraction, processing, sanitization, and recoverable errors in the GUI activity log.

The app processes files locally. It does not upload your archives or media. `ffprobe`, when available, is run locally to inspect video metadata.

## Requirements

- Windows 10 or later
- Python 3.10 or later
- Tkinter, normally included with the standard Windows Python installer
- Pillow and openpyxl (installed below)
- Optional: FFmpeg's `ffprobe` on `PATH` to read embedded video `creation_time` tags

## Install

Open PowerShell in the repository folder and run:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

If PowerShell blocks virtual-environment activation, run the venv interpreter directly instead:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py
```

To enable the test-coverage command, install the development dependencies:

```powershell
python -m pip install -r requirements-dev.txt
```

### Optional Video Metadata

Install FFmpeg and add its `bin` directory to `PATH` so `ffprobe.exe` can be found. The app continues to work without it; video files then fall through to Google Photos JSON metadata when filename, EXIF, and THM dates are unavailable.

## Run

From the repository folder:

```powershell
python main.py
```

In the window, choose all three locations before starting:

1. **ZIP folder**: the folder containing Takeout ZIP files. Subfolders are searched too.
2. **Temporary folder**: a location with enough free space for the uncompressed archives. GoogLi creates a unique `takeout_import_*` folder and deletes that run's extracted data after processing. Other files in the selected temporary folder are left untouched.
3. **Output folder**: the organized library and Excel report destination.

The activity log reports which ZIP/member is being extracted and which media file is being processed. Recoverable errors, such as a damaged ZIP member or a renamed Windows-invalid archive path, are marked `[ERROR]` or logged with the original and sanitized paths; processing continues where possible.

## Output Layout

```text
output/
|-- 2021/
|   |-- 2021 April/
|   |-- Family album/
|   `-- collage/
|-- 2022/
|   |-- 2022 October/
|   `-- 2022 - trip/
|-- unknown date/
`-- media_organization_report.xlsx
```

Dated media is copied to `YEAR/YEAR Month/`. Album media also gets a copy directly under its year folder. Collage media also gets a copy under `YEAR/collage/`. Media with no supported date goes to the top-level `unknown date/` folder. Sidecars are used for metadata, not copied as media. Source ZIPs are never deleted or modified.

## Date Sources

Dates are checked in this order. The first usable date wins:

1. Date/time parsed from the media filename
2. Media EXIF: `DateTimeOriginal` (tag 36867), then `DateTime` (tag 306)
3. Matching video's `.thm` or `.THM` sidecar EXIF
4. Video container or stream `creation_time`, read with `ffprobe` when available
5. Google Photos JSON `creationTime.timestamp`
6. Google Photos JSON `photoTakenTime.timestamp`

Only years **1950 through 2100 inclusive** are accepted. Filesystem modification time is not used. A date-only filename is interpreted as midnight. Google JSON timestamps are Unix timestamps; EXIF and filename dates have no timezone unless encoded in the filename.

## Supported Filename Patterns

The parser recognizes these patterns in the filename stem. Examples are illustrative; invalid dates and years outside the accepted range are ignored.

| Pattern | Example | Parsed date |
| --- | --- | --- |
| Dotted date and time | `2007-01-01 00.13.46 2007 01 01` | 2007-01-01 00:13:46 |
| WhatsApp dotted time | `WhatsApp Image 2021-04-29 at 18.46.14 (3)` | 2021-04-29 18:46:14 |
| VLC snapshot | `vlcsnap-2011-05-13-15h20m43s172` | 2011-05-13 15:20:43 |
| Windows Camera | `WIN_20200330_09_31_37_Pro` | 2020-03-30 09:31:37 |
| Compact date/time with separator | `20230315_182317`, `20230315-182317` | 2023-03-15 18:23:17 |
| Compact date/time | `20230315182317` | 2023-03-15 18:23:17 |
| Delimited screenshot | `2023-03-15_18-23-17` | 2023-03-15 18:23:17 |
| Windows Phone date | `WP_20140407_005` | 2014-04-07 00:00:00 |
| WhatsApp date | `IMG-20211011-WA0008` | 2021-10-11 00:00:00 |
| Generic compact date | `20211011` | 2021-10-11 00:00:00 |
| Day-month-year with optional trailing digits | `27112009 27 11(nov) 2009`, `01012010024` | 2009-11-27, 2010-01-01 |
| Video date with unpadded month | `Video2007730423257` | 2007-07-30 00:00:00 |

The report includes separate columns for the filename, media EXIF, THM EXIF, video media-created, and Google Photos dates, even when a higher-priority source determines the output folder.

## Supported Media

Images: `.jpg`, `.jpeg`, `.png`, `.gif`, `.tif`, `.tiff`, `.webp`, `.heic`, `.heif`, `.bmp`, `.dng`, `.cr2`, `.nef`, `.arw`, `.orf`, `.rw2`.

Videos: `.mp4`, `.mov`, `.m4v`, `.avi`, `.mkv`, `.3gp`, `.3g2`, `.wmv`, `.mpg`, `.mpeg`, `.mts`, `.m2ts`, `.mod`, `.webm`.

## Excel Audit

The output workbook has one row per media file and includes:

- Serial number and extracted media path/name
- Date found in the filename, media EXIF, THM EXIF, video metadata, and Google Photos JSON
- Original location in the ZIP and the actual output path(s)
- Whether a same-name, same-size destination file already existed

When the same filename and size exists at a destination, the app skips that copy and records the existing destination. Filename and size are the duplicate test; file contents are not hashed.

## Tests

Run the unit and synthetic ZIP integration tests from the repository folder:

```powershell
python -m unittest discover -s tests -v
```

With the development requirements installed, measure line coverage:

```powershell
python -m coverage run --source=takeout_organizer,date_helpers -m unittest discover -s tests
python -m coverage report -m
```

All test archives and media are generated in temporary folders. Tests do not use a personal photo collection.

## Contributing

Bug reports, filename examples, and pull requests are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md) for how to provide useful, privacy-safe reports and run checks locally.

## License

This project is released under the [MIT License](LICENSE). Copyright (c) 2026 Vaishvik Satyam.

## Upload to GitHub

1. Create an empty GitHub repository. Do not add a README or `.gitignore` there because this folder already contains them.
2. In PowerShell opened at this project folder, review the repository and stage its files:

	```powershell
	# Run from the project folder
	git status --short
	git add -A
	git status --short
	git diff --cached --stat
	```

	Check the staged list carefully. `.gitignore` excludes common media, archives, logs, temporary extraction data, and generated reports, but always confirm that no private files are staged.
3. Commit and connect the empty GitHub repository, replacing `USERNAME` and `REPOSITORY` with your GitHub values:

	```powershell
	git commit -m "Prepare Google Photos Takeout Organizer for GitHub"
	git remote add origin https://github.com/USERNAME/REPOSITORY.git
	git push -u origin main
	```

