import runpy
from pathlib import Path

# Streamlit Cloud entrypoint: redirige automáticamente a app.py
current_dir = Path(__file__).resolve().parent
app_path = current_dir / "app.py"

runpy.run_path(str(app_path), run_name="__main__")
