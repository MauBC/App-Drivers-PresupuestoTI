import customtkinter as ctk
import pytest
from openpyxl import load_workbook

from conftest import source_config
from src.core.excel_safety import check_writable
from src.gui.widgets import excel_config_panel as gui


@pytest.fixture
def panel(monkeypatch, tk_root):
    errors = []
    monkeypatch.setattr(gui.messagebox, "showerror", lambda title, message: errors.append(message))
    widget = gui.ExcelConfigPanel(tk_root, title="Maestro", role="base")
    try:
        yield widget, errors
    finally:
        widget.close_session()
        widget.destroy()


def test_configuration_reads_then_releases_and_refreshes(panel, excel_file, monkeypatch):
    widget, errors = panel
    source = excel_file("maestro.xlsx", ["Clave", "Importe"], [[1, 10]])
    monkeypatch.setattr(gui.filedialog, "askopenfilename", lambda **kwargs: str(source))
    widget.select_file()
    assert widget.sheet_names == ["Datos", "Conservar"]
    assert widget.excel_session is None
    check_writable(source)
    widget.sheet_entry.delete(0, "end")
    widget.sheet_entry.insert(0, "Datos")
    widget.header_entry.delete(0, "end")
    widget.header_entry.insert(0, "1")
    widget.validate_sheet()
    widget.load_headers()
    assert widget.headers == ["Clave", "Importe"]
    check_writable(source)
    workbook = load_workbook(source)
    try:
        workbook["Datos"]["B1"] = "Monto"
        workbook.save(source)
    finally:
        workbook.close()
    # Sigue funcionando incluso tras el close_session que usan los ejecutores.
    widget.close_session()
    widget.load_headers()
    assert widget.headers == ["Clave", "Monto"]
    check_writable(source)
    assert not errors


def test_profile_restores_config_without_open_workbook(panel, excel_file):
    widget, errors = panel
    source = excel_file("maestro.xlsx", ["Clave", "Importe"], [[1, 10]])
    config = source_config(source, ["Clave", "Importe"])
    widget.apply_base_profile_config({"config": config, "nombre_perfil": "Ejemplo"})
    restored = widget.build_config()
    assert restored["hoja"]["nombre_detectado"] == "Datos"
    assert restored["columna_busqueda"]["header_detectado"] == "Clave"
    assert [c["alias"] for c in restored["columnas"]] == ["Clave", "Importe"]
    assert widget.excel_session is None
    assert not errors
    check_writable(source)


def test_invalid_sheet_in_panel_releases_file(panel, excel_file, monkeypatch):
    widget, errors = panel
    source = excel_file("maestro.xlsx", ["Clave"], [[1]])
    monkeypatch.setattr(gui.filedialog, "askopenfilename", lambda **kwargs: str(source))
    widget.select_file()
    widget.sheet_entry.delete(0, "end")
    widget.sheet_entry.insert(0, "NoExiste")
    widget.validate_sheet()
    assert len(errors) == 1
    check_writable(source)
