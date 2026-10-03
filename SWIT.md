# Software Integration Test (SWIT)

## Automated integration test

The end-to-end tests in `tests/test_takeout_organizer.py` build synthetic Google Photos ZIPs with image and video files, THM EXIF, JSON sidecars, a year-bearing album name, and an intentionally corrupted CRC member. They verify year bounds and the supplied filename patterns, ffprobe metadata parsing, collage placement, unknown-date output, duplicate skipping by filename and size, Excel audit contents, Windows-invalid path sanitization, CRC error logging and continuation, omission of sidecars from media output, and preservation of the original ZIPs. Run them with:

```powershell
python -m unittest discover -s tests -v
```

## GUI smoke test

1. Start `python main.py` and confirm the three folder selectors and activity log are visible.
2. Select a folder containing a small synthetic or disposable Takeout ZIP, a new temporary folder, and a new output folder.
3. Click **Organize Takeout**. Confirm the window remains responsive and reports completion.
4. Confirm the activity log names each ZIP and extracted member, then names each media file as it is processed. Confirm the temp folder contains one run folder per operation and one isolated directory per ZIP.
5. Confirm dated files are copied under `YEAR/YEAR Month`, album files under `YEAR/album name`, undated files under `unknown date`, and `media_organization_report.xlsx` appears in the output; confirm the source ZIP is unchanged.
6. Run a second time with the same output and confirm same-name/same-size files are skipped and noted in the report, while same-name/different-size files receive numbered names.
7. Try an archive containing a folder name with a Windows-invalid character such as `:`; confirm extraction logs the sanitized name and continues.
8. Try an empty ZIP folder and an invalid ZIP; confirm the GUI reports a useful error without deleting source files.

Run the smoke test against disposable sample data, not the only copy of a personal archive. For large exports, confirm the selected temporary and output volumes have enough free space.