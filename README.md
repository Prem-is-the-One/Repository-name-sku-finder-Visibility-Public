# Excel Live Search Website Starter

This starter converts an Excel worksheet into `data/data.json`, extracts supported
embedded worksheet images into `data/images/`, and serves a fast browser search UI.

## 1. Create your local settings

Copy:

`settings.example.json`

to:

`settings.json`

Then edit:

- `excel_file`
- `sheet_name`
- `header_row`
- `auto_git_push`

## 2. Create a Python virtual environment

Windows PowerShell:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

If PowerShell blocks activation, use Command Prompt:

```bat
.venv\Scripts\activate.bat
pip install -r requirements.txt
```

## 3. Export the Excel workbook

```powershell
python tools/export_excel.py
```

## 4. Run locally

```powershell
python -m http.server 5500
```

Open:

`http://localhost:5500`

Do not open `index.html` directly with `file://`, because browser security can block
the JSON fetch.

## 5. Watch Excel automatically

In a second VS Code terminal:

```powershell
python tools/watch_excel.py
```

Every Excel save rebuilds the website data. If `auto_git_push` is `true`, it also
runs Git add/commit/push.

## 6. GitHub Pages

Publish the repository to GitHub, then in the repository:

Settings -> Pages -> Build and deployment -> Deploy from a branch -> `main` / `(root)`

For GitHub Free, a Pages repository normally needs to be public. Anything published
on the site should be treated as public data.

## Notes about images

The exporter supports standard anchored worksheet images and includes a best-effort
extractor for modern Excel "Place in Cell" images stored through Rich Data.

Excel has several internal image formats. If your particular workbook exports text
but not its images, the site still works; the image extractor can then be adapted
to that workbook's exact Excel structure.
