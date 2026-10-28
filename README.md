# DLE DOCX Dictionary Builder

Small tool that looks up Spanish words through the rae-api.com API and writes
formatted entries into a Word document.

<div align="center">
  <img width="671" height="568" alt="image" src="https://github.com/user-attachments/assets/f03ed508-4933-42cb-a29a-cf79b1f61493" />
  <p>Portable .exe interface</p>
</div>

<div align="center">
  <img width="611" height="419" alt="image" src="https://github.com/user-attachments/assets/6638adf6-6bca-4e06-bebb-d012a41f45f5"/>
  <p>Example output</p>
</div>

## Features

- Look up one word or a list of words from `words.txt` through rae-api.com.
- Append entries to `diccionario.docx` or another DOCX path.
- Keep entries alphabetized under centered letter headings.
- Skip words that are already present in the DOCX.
- Respect API request limits, log quota information returned by the API, and stop clearly when a quota is exhausted.
- Optionally add a safety delay between requests; the API limit is respected automatically.
- Print words not found in DLE at the end of each run.
- Use the optional graphical interface on Windows.

## Requirements and API key

- Python 3.10+
- python-docx
- CustomTkinter (for the GUI)
- Internet access to rae-api.com

Install Python dependencies:

```powershell
pip install -r requirements.txt
```

An API key is optional. To use one, keep `API.txt` next to `dlex.py`; for the
portable executable, place it next to `DLEX-portable.exe`. The file may contain
a raw key or a labeled line such as `RAE API KEY: <key>`. `API.txt` is ignored
by Git. Never commit or share it. A key uses the quota assigned by the API
provider to its account; a key alone does not guarantee a higher limit.

## API limits and usage

The API documents these request quotas; they may change, so the response headers
reported during a run are authoritative:

| Tier | Requests per minute | Requests per day |
| --- | ---: | ---: |
| Free | 10 | 100 |
| Developer | 60 | 5,000 |
| Extended | 300 | 50,000 |

When the API provides a tier limit, DLEX paces later requests to fit that
per-minute allowance. Until it has received limit headers, it conservatively
waits at least 6 seconds between requests (the free-tier rate). It then adjusts
its pacing based on the limit headers. The optional `--delay-between-words`
value adds a safety delay; it can slow requests down but cannot make them
faster than the API limit.

DLEX logs the tier and remaining per-minute/daily quota when those headers are
provided. If a quota is exhausted or the API returns HTTP 429, the batch stops
instead of retrying automatically. A 429 response's `Retry-After` value is
shown when available. The API key is sent only in the `X-API-Key` request header
and is never written to logs or the request URL.

API reference: <https://rae-api.com/docs/api-reference/>

## Portable Windows release

For the GitHub Release, download `DLEX-portable.exe` and double-click it to
open the graphical interface. It does not require installing Python.

Build the single-file release artifact locally with:

```powershell
python -m PyInstaller .\DLEX-portable.spec
```

## Usage

Look up one word through the API and save it to `diccionario.docx`:

```powershell
python dlex.py amor
```

Read all words from `words.txt`:

```powershell
python dlex.py --words-file
```

Set an optional additional safety delay between words (default: 0 seconds; accepted range: 0-120):

```powershell
python dlex.py --words-file --delay-between-words 5
```

In the graphical interface, this control is under **Show advanced settings**.

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

## Generated files

The repository does not include generated dictionary files. Running the script
creates or updates `diccionario.docx` by default. You can choose a different
output path with `--output`.

## Output format

Generated DOCX files use this dictionary style:

- centered bold letter headings, 14 pt
- justified dictionary entries, 10 pt
- etymology in parentheses, 9 pt
- numbered senses inline with `||`, with each sense number shown once
- synonym and antonym lists omitted from definition text

Generated `.docx` and `.pdf` files are ignored by default in `.gitignore`.
Remove those ignore rules if you decide to publish generated dictionaries too.
