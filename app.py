"""Streamlit entrypoint for SheetPilot Agent v2.0.10.

Streamlit reruns this file after every widget interaction.  Executing the UI
script explicitly keeps each rerun independent from Python's import cache.
"""

from pathlib import Path
from runpy import run_path


run_path(str(Path(__file__).with_name("app_chat.py")), run_name="__main__")
