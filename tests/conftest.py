from pathlib import Path

import pytest
from openpyxl import Workbook


@pytest.fixture(autouse=True)
def isolated_recovery(tmp_path, monkeypatch, request):
    # Los tests nunca crean pendientes en los datos reales del usuario.
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "localappdata"))
    yield
    if "tk_root" in request.fixturenames:
        # Tk debe liberar los widgets del test desde el hilo de la interfaz,
        # antes de que el siguiente worker active el recolector de Python.
        import gc
        gc.collect()


@pytest.fixture(scope="session")
def tk_root():
    import customtkinter as ctk
    root = ctk.CTk()
    root.withdraw()
    yield root
    root.destroy()


@pytest.fixture
def excel_file(tmp_path):
    def create(name, headers, rows, header_row=1):
        path = tmp_path / name
        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.title = "Datos"
            for _ in range(header_row - 1):
                sheet.append(["Titulo"])
            sheet.append(headers)
            for row in rows:
                sheet.append(row)
            other = workbook.create_sheet("Conservar")
            other["A1"] = "=1+2"
            other.sheet_state = "hidden"
            workbook.save(path)
        finally:
            workbook.close()
        return path
    return create


def source_config(path: Path, headers, header_row=1):
    columns = [
        {"alias": header, "header_detectado": header,
         "indice_0": i, "indice_1": i + 1, "mostrar_salida": True}
        for i, header in enumerate(headers)
    ]
    return {
        "archivo": {"ruta": str(path), "nombre": path.name},
        "hoja": {"nombre_detectado": "Datos"},
        "fila_header": header_row,
        "columnas": columns,
        "columna_busqueda": columns[0],
    }


@pytest.fixture
def search_config(excel_file):
    template = excel_file("template.xlsx", ["Clave", "Detalle"], [
        [" José-Pérez ", "uno"], ["Ausente", "dos"], [None, "tres"],
        [None, None],
    ], header_row=2)
    base = excel_file("maestro.xlsx", ["Clave", "Importe"], [
        ["JOSE PEREZ", 10], ["José Pérez", 99], [None, 50],
    ])
    return {
        "template": source_config(template, ["Clave", "Detalle"], 2),
        "base": source_config(base, ["Clave", "Importe"]),
        "match": {"tipo_busqueda": "exacta", "score_exacto": 95, "score_aproximado": 70},
        "salida": {"modo": "excel_nuevo", "nombre_hoja_resultado": "Cruce",
                   "incluir_columnas_tecnicas": True},
    }
