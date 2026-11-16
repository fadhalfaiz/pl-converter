from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / ".runtime" / "python"))

from streamlit.web import cli

sys.argv = [
    "streamlit", "run", str(ROOT / "app.py"),
    "--server.address=127.0.0.1", "--server.port=8501",
    "--server.headless=true", "--browser.gatherUsageStats=false",
    "--global.developmentMode=false",
]
cli.main()
