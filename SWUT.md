# Software Unit Test (SWUT)

Run the unit suite from the repository root:

```powershell
python -m unittest discover -s tests -v
```

The tests cover 1950–2100 year bounds, supplied filename patterns, filename/EXIF/THM/video-created/Takeout date precedence, collage routing, date-folder versus album detection, ZIP path traversal rejection, Windows path sanitization, CRC corruption recovery, recursive archive discovery, isolated extraction, invalid/empty archives, filename-and-size duplicate behavior, Excel audit contents, and synthetic Takeout archives. All test archives and media are generated in temporary directories; the personal photo library is not read or changed.

To measure line coverage when `coverage` is installed:

```powershell
coverage run -m unittest discover -s tests
coverage report -m
```

The suite uses the Python standard library plus the application's Pillow and openpyxl dependencies.