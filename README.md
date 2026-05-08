# DLE DOCX Dictionary Builder

Small command-line tool that looks up Spanish words in the online DLE dictionary
and writes formatted entries into a Word document.

## Features

- Look up one word or a list of words from `words.txt`.
- Append entries to `diccionario.docx` or another DOCX path.
- Keep entries alphabetized under centered letter headings.
- Skip words that are already present in the DOCX.
- Print words not found in DLE at the end of each run.

## Requirements

- Python 3.10+
- Playwright
- python-docx

Install Python dependencies:

```powershell
pip install -r requirements.txt
```

Install the Playwright browser once:

```powershell
playwright install chromium
```

## Usage

Look up one word and save it to `diccionario.docx`:

```powershell
python test.py amor
```

Read all words from `words.txt`:

```powershell
python test.py --words-file
```

Use a custom word list and output file:

```powershell
python test.py --words-file words.txt --output diccionario.docx
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
python test.py --words-file
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
