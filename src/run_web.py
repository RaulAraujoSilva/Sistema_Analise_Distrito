# -*- coding: utf-8 -*-
"""Entry point for the web interface. Run: python src/run_web.py"""
import sys
from pathlib import Path

# Ensure src/ is on the path so config and other modules can be imported
SRC_DIR = Path(__file__).resolve().parent
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

import uvicorn

if __name__ == "__main__":
    uvicorn.run(
        "web.app:app",
        host="127.0.0.1",
        port=8000,
        reload=True,
        reload_dirs=[str(SRC_DIR / "web")],
    )
