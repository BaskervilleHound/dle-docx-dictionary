# DLEX Portable Windows Release

## Download

Use `DLEX-portable.exe` from the GitHub Release assets.

This is a portable Windows build: no installer and no Python setup required.
Double-click the exe to open the graphical interface.

## Notes

- The app includes Chromium for DLE lookups, so the download is large.
- Windows SmartScreen may warn on first launch because the exe is unsigned.
- Keep an internet connection available while looking up words.
- The default output file is `diccionario.docx`.

## Command Line

The same executable also supports the command-line workflow:

```powershell
.\DLEX-portable.exe amor
.\DLEX-portable.exe --words-file words.txt --output diccionario.docx
```

## Verification

SHA256:

```text
DLEX-portable.exe  0D0A52F672C2DB306D9F4581D8516D9326B30C47C46003BE7F29A1D955F2BA53
```
