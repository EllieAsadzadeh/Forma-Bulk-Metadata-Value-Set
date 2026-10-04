# Forma metadata push

This script copies custom-attribute metadata from an Autodesk Files log Excel export into Autodesk Construction Cloud / Forma Data Management.

It matches each spreadsheet row to the file with the same name in the same folder, then writes attributes onto **that row’s file version only** (for example V1). It does not update every version of the file.

## What it does

1. Reads the `Files` sheet of the Files log report.
2. Finds each file in Forma by folder path and file name.
3. Resolves the version from the **Version number** column (`V1` → version 1).
4. Pushes every Forma custom-attribute column from that row onto that version.
5. Treats a blank Excel cell as a clear: the matching Forma value is removed.

The Files log **Description** column is written to the Forma **File Description** attribute. If Description is blank, File Description is cleared in Forma.

Other custom attributes are matched by column name, including:

- Reference
- Phase
- Suitability Code
- Baseline Version
- Status
- Design Stage (Milestone)
- Design Status
- Baseline Changes

## Requirements

- Python 3.10 or later
- An [Autodesk Platform Services](https://aps.autodesk.com/) app with `data:read` and `data:write`
- That app added to the ACC account as a custom integration, and to the target project with permission to edit documents
- Packages:

```text
pip install openpyxl requests python-dotenv
```

## Setup

1. Copy `.env.example` to `.env` next to `push_metadata_to_acc.py`.
2. Put your APS application id and secret in `.env`.
3. Set `APS_REGION` in `.env` to the data region of the target ACC account.
4. Place the Files log workbook in this folder and set `EXCEL_PATH` in `push_metadata_to_acc.py` to that file name.
5. On the `Files` sheet, set **Project ID** to the ACC project GUID (no `b.` prefix). Use the same ID on every row. Leave **Item ID (URN)** and **Version ID (URN)** blank on the first run; the script fills them in.

## Run

From this folder:

```text
python push_metadata_to_acc.py
```

Close the Excel file before you run, or the script cannot save the resolved Item / Version IDs back to the workbook. The Forma updates still go through even if the save fails.

A successful line looks like:

```text
[ok] row 2: example.dwg V1 -> pushed (1 set, 8 cleared)
```

`set` means a non-blank Excel value was written. `cleared` means a blank Excel cell removed the Forma value.

## Version matching

| Spreadsheet | Forma target |
| --- | --- |
| Name + Folder name and path | The file in that folder |
| Version number (`V1`, `V2`, …) | That specific version URN (`?version=1`, `?version=2`, …) |

If Version number is missing, the row is skipped. The script will not fall back to the latest version.

## What to put on GitHub

Include:

- `push_metadata_to_acc.py`
- `README.md`
- `.gitignore`
- `.env.example`

Do **not** include:

- `.env` (live credentials)
- The Files log `.xlsx` if it contains data you do not want public

### Should `.gitignore` go in the zip?

Yes. Put `.gitignore` in the zip.

GitHub does not invent ignore rules for you. The file in this repo tells git to skip `.env`, so the APS secret is not committed when you unzip, `git init`, and push. Without it, it is easy to upload credentials by mistake.

`.gitignore` itself is not a secret. It is a small text file that should be committed.

## Security

- Keep `.env` on your machine only.
- If an APS secret was ever committed or zipped, rotate it in the APS portal and update local `.env`.
- The APS app can read and write document metadata in every project it is added to. Limit project membership to what this job needs.
