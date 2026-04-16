#!/bin/bash
cd "$(dirname "$0")"
source .venv/bin/activate
python email_processor.py >> logs/email_processor.log 2>&1
