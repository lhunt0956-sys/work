#!/usr/bin/env bash
# Idempotent dependency setup for the work (docx generation) project.
# Creates a project-local virtual environment and installs pinned deps.
set -euo pipefail

cd "$(dirname "$0")/.."

python3 -m venv .venv
./.venv/bin/python -m pip install --upgrade pip
./.venv/bin/python -m pip install -r requirements.txt

echo "Environment ready. Run: ./.venv/bin/python generate_form.py"
