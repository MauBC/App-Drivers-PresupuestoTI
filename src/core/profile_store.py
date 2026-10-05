from pathlib import Path
import json
import re


PROFILE_DIR = Path("data") / "perfiles_custom"


def clean_profile_name(name: str) -> str:
    clean = str(name).strip()
    clean = re.sub(r"[^a-zA-Z0-9_-]+", "_", clean)
    clean = clean.strip("_")

    if not clean:
        raise ValueError("El nombre del perfil no puede estar vacio")

    return clean


def save_profile(profile_name: str, profile_data: dict) -> Path:
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)

    clean_name = clean_profile_name(profile_name)
    output_path = PROFILE_DIR / f"{clean_name}.json"

    with output_path.open("w", encoding="utf-8") as file:
        json.dump(profile_data, file, ensure_ascii=False, indent=2)

    return output_path
