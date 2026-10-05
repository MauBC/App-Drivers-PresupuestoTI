import pytest
from conftest import source_config

from src.core.excel_headers import clean_header, prepare_headers
from src.core.excel_inspector import read_headers
from src.core.driver_executor import run_driver, read_excel_with_origin
from src.core.driver_executor import resolve_dataframe_column as driver_column
from src.core.completion_executor import resolve_dataframe_column as completion_column
from src.core.completion_executor import run_completion
from src.core.search_executor import run_search


@pytest.mark.parametrize("resolve", [driver_column, completion_column])
def test_header_lookup_preserves_peruvian_names_and_symbols(resolve):
    headers = ["  Área de Gestión (%)\u200b", "\ufeffAño / Niñez\u00a0", "N° de Documento"]
    assert resolve(headers, "área de gestión (%)") == headers[0]
    assert resolve(headers, "Año / Niñez") == headers[1]
    assert resolve(headers, "N° DE DOCUMENTO") == headers[2]
    with pytest.raises(ValueError, match="No se encontro"):
        resolve(headers, "Ano / Ninez")


def test_unicode_composition_and_internal_spaces():
    assert clean_header(" A\u0301rea / An\u0303o (%) ") == "Área / Año (%)"
    assert clean_header("Centro  de costo") == "Centro  de costo"
    assert prepare_headers([None, "Año", "Ano"]) == ["COL_1", "Año", "Ano"]


@pytest.mark.parametrize("headers", [["DNI", " dni\u00a0"], ["Área", "A\u0301rea"], [None, "COL_1"]])
def test_real_excel_rejects_ambiguity_before_pandas_renames_it(excel_file, headers):
    path = excel_file("ambiguo.xlsx", headers, [[1, 2]])
    before = path.read_bytes()
    for read in (read_headers, read_excel_with_origin):
        with pytest.raises(ValueError, match="columnas 1 y 2"):
            read(path, "Datos", 1)
    assert path.read_bytes() == before


@pytest.mark.parametrize("driver_type", ["cantidad", "porcentaje"])
def test_driver_dirty_headers_keeps_validation_distribution_and_extra_names(excel_file, driver_type):
    result = excel_file("resultado.xlsx", [" DNI ", "CECO\u00a0", " PRECIO ", "\ufeffvalidado", " Área / Año (%) "],
                        [["001", "ORIG", 100, 1, "Niñez"], ["002", "OTRO", 40, 0, "Perú"]])
    base = excel_file("maestro.xlsx", ["\u200bDNI", " CECO ", "PORCENTAJE\u2060"],
                      [["001", "A", .25], ["001", "B", .75]])
    originals = [p.read_bytes() for p in (result, base)]
    assert read_headers(result, "Datos", 1)["headers"][-1] == "Área / Año (%)"
    config = {"driver_type": driver_type,
              "resultado": {"file_path": str(result), "sheet_name": "Datos", "header_row": 1,
                            "dni_col": "dni", "ceco_col": "CECO", "price_col": "PRECIO"},
              "base": {"file_path": str(base), "sheet_name": "Datos", "header_row": 1,
                       "dni_col": "DNI", "ceco_col": "CECO", "percentage_col": "PORCENTAJE"}}
    df, _, _ = run_driver(config)
    assert df["CECO2"].tolist() == ["A", "B"]
    assert df["PORCENTAJE"].tolist() == [.25, .75]
    if driver_type == "cantidad":
        assert df["PRECIO_DISTRIBUIDO"].sum() == 100
    assert [p.read_bytes() for p in (result, base)] == originals


def test_match_and_completion_with_dirty_spanish_headers(excel_file):
    template = excel_file("consulta.xlsx", ["\ufeffNombre", " Área / Año (%) "], [["Peña", "Perú"]])
    base = excel_file("base.xlsx", ["Nombre\u200b", " Importe "], [["Peña", 100]])
    config = {"template": source_config(template, ["Nombre", "Área / Año (%)"]),
              "base": source_config(base, ["Nombre", "Importe"]),
              "match": {"tipo_busqueda": "exacta", "score_exacto": 100, "score_aproximado": 75},
              "salida": {"incluir_columnas_tecnicas": True}}
    df, _ = run_search(config)
    assert df["Área / Año (%)"].tolist() == ["Perú"]
    assert df["Importe"].tolist() == [100]
    df.loc[0, "VALIDADO"] = 0
    df.loc[0, "Importe"] = None
    headers = ["\ufeff " + str(h).lower() + "\u00a0" if str(h) == "VALIDADO"
               else " " + str(h) + "\u200b" for h in df.columns]
    previous = excel_file("previo.xlsx", headers, list(df.itertuples(index=False, name=None)))
    before = previous.read_bytes()
    config["previous_result"] = {"archivo": {"ruta": str(previous)},
                                 "hoja": {"nombre_detectado": "Datos"},
                                 "fila_header": 1, "columna_busqueda": "VALOR_BUSCADO"}
    config["column_mappings"] = [{"base_alias": "Importe", "base_header_detectado": "Importe",
                                  "destination_column_ref": "importe"}]
    completed, summary = run_completion(config)
    assert completed["Importe"].tolist() == [100]
    assert completed["VALIDADO"].tolist() == [1]
    assert completed["Área / Año (%)"].tolist() == ["Perú"]
    assert summary["filas_reprocesadas"] == 1
    assert previous.read_bytes() == before
