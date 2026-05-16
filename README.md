# DLE DOCX Dictionary Builder

Small command-line tool that looks up Spanish words in the online DLE dictionary
and writes formatted entries into a Word document.

<div align="center">
  <img width="611" height="419" alt="image" src="https://github.com/user-attachments/assets/6638adf6-6bca-4e06-bebb-d012a41f45f5"/>
  <p>Example output</p>
</div>

## Features

- Look up one word or a list of words from `words.txt`.
- Append entries to `diccionario.docx` or another DOCX path.
- Keep entries alphabetized under centered letter headings.
- Skip words that are already present in the DOCX.
- Print words not found in DLE at the end of each run.
- Use the optional graphical interface on Windows.

## Requirements

- Python 3.10+
- Playwright
- python-docx
- CustomTkinter

Install Python dependencies:

```powershell
pip install -r requirements.txt
```

Install the Playwright browser once:

```powershell
playwright install chromium
```

## Usage

## Portable Windows Release

For the GitHub Release, download `DLEX-portable.exe` and double-click it to
open the graphical interface. It is portable and does not require installing
Python.

The older folder build in `dist/DLEX/` only works when `DLEX.exe` stays beside
its `_internal` folder. Do not distribute that single `DLEX.exe` by itself.

Build the single-file release artifact locally with:

```powershell
python -m PyInstaller .\DLEX-portable.spec
```

Look up one word and save it to `diccionario.docx`:

```powershell
python dlex.py amor
```

Read all words from `words.txt`:

```powershell
python dlex.py --words-file
```

Use a custom word list and output file:

```powershell
python dlex.py --words-file words.txt --output diccionario.docx
```

Open the graphical interface:

```powershell
python dlex.py --gui
```

`words.txt` can contain comma-separated or newline-separated words:

```text
palabra, diccionario, entrada
definición
```

## Generated Files

The repository does not include the generated dictionary files. Running the
script creates or updates `diccionario.docx` by default:

```powershell
python dlex.py --words-file
```

You can also choose a different output path with `--output`.

## Output Format

Generated DOCX files use this dictionary style:

- centered bold letter headings, 14 pt
- justified dictionary entries, 10 pt
- etymology in parentheses, 9 pt
- numbered senses inline with `||`

Generated `.docx` and `.pdf` files are ignored by default in `.gitignore`.
Remove those ignore rules if you decide to publish generated dictionaries too.
