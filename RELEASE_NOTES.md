# DLEX Portable Windows Release

## Download

Download `DLEX-portable.exe` from this GitHub Release. It is a portable Windows
application; no Python installation or installer is required. Double-click the
executable to open the graphical interface.

## What's new

- Uses the rae-api.com API for dictionary lookups.
- Automatically paces requests according to API quota headers, with a
  conservative free-tier pace until the API reports a limit.
- Supports an optional API key and an optional extra safety delay.
- Includes the same command-line workflow as `dlex.py`.

To use an API key, place `API.txt` next to `DLEX-portable.exe`. The key is
provided by the API provider and uses the quota assigned to that account.
Never share or commit the key file.

## Notes

- An internet connection is required for lookups.
- Windows SmartScreen may warn on first launch because the executable is
  unsigned.
- The default output file is `diccionario.docx`.

## Command line

```powershell
.\DLEX-portable.exe amor
.\DLEX-portable.exe --words-file words.txt --output diccionario.docx
```

## Verification

SHA256:

```text
See SHA256SUMS.txt in the repository.
```
