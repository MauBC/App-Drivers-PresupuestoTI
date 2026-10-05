from pathlib import Path
from datetime import datetime
import json
import re


PROFILE_DIR = Path("data") / "perfiles_drivers" / "reglas_ceco"


def clean_profile_name(name: str) -> str:
    clean = str(name).strip()
    clean = re.sub(r"[^a-zA-Z0-9_-]+", "_", clean)
    clean = clean.strip("_")

    if not clean:
        raise ValueError("El nombre del perfil no puede estar vacio")

    return clean


def save_ceco_rules_profile(profile_name: str, rules: list[dict]) -> Path:
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)

    clean_name = clean_profile_name(profile_name)
    output_path = PROFILE_DIR / f"{clean_name}.json"

    data = {
        "tipo": "CECO_RULES_PROFILE",
        "version": 1,
        "nombre_perfil": profile_name,
        "guardado_en": datetime.now().isoformat(timespec="seconds"),
        "rules": rules,
    }

    with output_path.open("w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)

    return output_path


def load_ceco_rules_profile(file_path: str | Path) -> dict:
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"No existe el perfil de reglas CECO: {path}")

    with path.open("r", encoding="utf-8") as file:
        data = json.load(file)

    if data.get("tipo") != "CECO_RULES_PROFILE":
        raise ValueError("El archivo seleccionado no es un perfil de reglas CECO")

    if "rules" not in data:
        raise ValueError("El perfil no tiene reglas CECO")

    return data
