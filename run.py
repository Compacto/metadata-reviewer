"""Inicialização local da aplicação Streamlit a partir do PyCharm."""
from pathlib import Path
import os
import sys

sys.dont_write_bytecode = True
os.environ["PYTHONDONTWRITEBYTECODE"] = "1"

if __name__ == "__main__":
    from streamlit.web import cli
    project = Path(__file__).resolve().parent
    os.chdir(project)
    sys.argv = ["streamlit", "run", str(project / "streamlit_app.py")]
    raise SystemExit(cli.main())
