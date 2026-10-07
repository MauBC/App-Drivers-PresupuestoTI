"""Copias locales de resultados; no ejecuta de nuevo el cálculo."""
import json
import os
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4


def recovery_directory():
    return Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir())) / "AppDriversPresupuestoTI" / "recovery"


def _write_json(path, value):
    # Registros inmutables: nunca se reemplaza un evento anterior.
    temporary = path.with_name(".event_" + uuid4().hex + ".part")
    with temporary.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2)
        stream.flush()
        os.fsync(stream.fileno())
    os.rename(temporary, path)


def prepare(source, destination, expected):
    from src.core.excel_safety import file_hash
    folder = recovery_directory() / uuid4().hex
    folder.mkdir(parents=True)
    payload = folder / ("resultado" + Path(source).suffix)
    shutil.copyfile(source, payload)
    with payload.open("r+b") as stream:
        os.fsync(stream.fileno())
    request = {"id": folder.name, "destination": str(Path(destination).resolve()),
               "expected": expected, "sha256": file_hash(payload),
               "created": datetime.now(timezone.utc).isoformat(), "payload": payload.name}
    _write_json(folder / "request.json", request)
    return folder.name


def _folder(identifier):
    if len(identifier) != 32 or any(c not in "0123456789abcdef" for c in identifier):
        raise ValueError("Identificador de recuperación inválido")
    return recovery_directory() / identifier


def complete(identifier, destination):
    path = _folder(identifier) / "completed.json"
    if not path.exists():
        _write_json(path, {"destination": str(destination), "saved_locally": True})
    # Conservar el recibo, retirar datos del Excel después de confirmar el guardado.
    folder = path.parent
    request = json.loads((folder / "request.json").read_text(encoding="utf-8"))
    if request["payload"] not in {"resultado.xlsx", "resultado.xlsm"}:
        raise ValueError("Nombre de copia de recuperación inválido")
    try:
        (folder / request["payload"]).unlink(missing_ok=True)
    except OSError:
        from src.core.excel_safety import logger
        logger.exception("Guardado confirmado; no se pudo limpiar la copia local: %s", folder)


def pending():
    records = []
    for path in sorted(recovery_directory().glob("*/request.json")):
        if (path.parent / "completed.json").exists():
            continue
        try:
            record = json.loads(path.read_text(encoding="utf-8"))
            record["id"] = path.parent.name
            records.append(record)
        except (OSError, ValueError):
            from src.core.excel_safety import logger
            logger.exception("Registro de recuperación ilegible: %s", path)
    return records


def restore(identifier, alternate=None):
    from src.core.excel_safety import atomic_write, file_hash, current_hash, destination_guard
    folder = _folder(identifier)
    # También serializa dos intentos de recuperar el mismo registro.
    with destination_guard(folder):
        request = json.loads((folder / "request.json").read_text(encoding="utf-8"))
        receipt = folder / "completed.json"
        if receipt.exists():
            return Path(json.loads(receipt.read_text(encoding="utf-8"))["destination"])
        payload = folder / request["payload"]
        if payload.parent != folder or file_hash(payload) != request["sha256"]:
            raise ValueError("La copia de recuperación está dañada. No se publicó el archivo.")
        # Si el proceso murió después de publicar pero antes del recibo, no repetir.
        attempts = [request]
        for event in sorted(folder.glob("attempt_*.json")):
            attempts.append(json.loads(event.read_text(encoding="utf-8")))
        for attempt in attempts:
            target = Path(attempt["destination"])
            try:
                saved_hash = current_hash(target)
            except OSError:
                # Un destino bloqueado no impide elegir una copia en otro lugar.
                continue
            if saved_hash == request["sha256"]:
                complete(identifier, target)
                return target
        target = Path(alternate).resolve() if alternate else Path(request["destination"])
        if alternate and target == Path(request["destination"]):
            raise ValueError("Elija otro nombre para guardar una copia separada.")
        if target.suffix.lower() != payload.suffix.lower():
            raise ValueError(f"Conserve la extensión {payload.suffix} del libro recuperado.")
        expected = None if alternate else request["expected"]
        _write_json(folder / f"attempt_{uuid4().hex}.json", {"destination": str(target)})
        result = atomic_write(target, lambda temporary: shutil.copyfile(payload, temporary),
                              expected=expected, recoverable=False)
        complete(identifier, result)
        return result
