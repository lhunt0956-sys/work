# work

Tools for generating **工作任务联系单** (Work Task Contact Form) Word documents
with [python-docx](https://python-docx.readthedocs.io/).

## Requirements

- Python 3.12+
- Dependencies pinned in [`requirements.txt`](requirements.txt)

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

> On Debian/Ubuntu, creating a virtual environment requires the
> `python3-venv` system package (e.g. `sudo apt-get install python3.12-venv`).

## Usage

Generate the example form:

```bash
python generate_form.py
# -> Wrote .../output/工作任务联系单.docx
```

Override any field from the command line:

```bash
python generate_form.py \
  --编号 RW-2024-042 \
  --项目名称 "龙江二期项目" \
  --output output/my_form.docx
```

Generated documents are written under `output/` (git-ignored).

## Project layout

| Path | Purpose |
| --- | --- |
| `generate_form.py` | Reference application that builds the DOCX form. |
| `requirements.txt` | Pinned Python dependencies. |
| `output/` | Destination for generated `.docx` files (git-ignored). |
