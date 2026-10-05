from pathlib import Path
from datetime import datetime
import json
import re


PROFILE_DIR = Path("data") / "perfiles_base"


def clean_profile_name(name: str) -> str:
    clean = str(name).strip()
    clean = re.sub(r"[^a-zA-Z0-9_-]+", "_", clean)
    clean = clean.strip("_")

    if not clean:
        raise ValueError("El nombre del perfil no puede estar vacio")

    return clean


def save_base_profile(profile_name: str, config: dict) -> Path:
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)

    clean_name = clean_profile_name(profile_name)
    output_path = PROFILE_DIR / f"{clean_name}.json"

    data = {
        "tipo": "BASE_EXCEL_CONFIG",
        "version": 1,
        "nombre_perfil": profile_name,
        "guardado_en": datetime.now().isoformat(timespec="seconds"),
        "config": config,
    }

    with output_path.open("w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)

    return output_path


def load_base_profile(file_path: str | Path) -> dict:
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"No existe el perfil base: {path}")

    with path.open("r", encoding="utf-8") as file:
        data = json.load(file)

    if data.get("tipo") != "BASE_EXCEL_CONFIG":
        raise ValueError("El archivo seleccionado no es una configuracion de base")

    if "config" not in data:
        raise ValueError("El perfil base no tiene config")

    return data
