#!/usr/bin/env bash
set -e
if [ -d ".venv" ]; then
  . .venv/bin/activate
fi
python3 "Image2ATF_V2.8.py"
