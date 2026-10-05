from copy import deepcopy
from src.gui.background_tasks import tasks_for
from pathlib import Path
import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox, filedialog

from src.gui.widgets.excel_config_panel import ExcelConfigPanel
from src.gui.widgets.completion_panel import CompletionPanel
from src.gui.widgets.driver_panel import DriverPanel
from src.core.search_profile_store import save_search_profile
from src.core.search_executor import run_search, export_result
from src.core.completion_executor import (
    read_result_headers,
    run_completion,
    export_completion_result,
)


class SearchBuilderView(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.search_type_var = tk.StringVar(value="similaridad")
        self.output_mode_var = tk.StringVar(value="excel_nuevo")
        self.include_technical_cols_var = tk.BooleanVar(value=True)

        self.completion_file_path = None
        self.completion_header_info = None

        self._build_layout()

    def _build_layout(self):
        title = ctk.CTkLabel(
            self,
            text="BUSCADOR DE CECOS / GENERADOR DE DRIVERS",
            font=ctk.CTkFont(size=30, weight="bold"),
        )
        title.grid(row=0, column=0, sticky="w", padx=20, pady=(18, 8))

        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))

        self.tab_template = self.tabs.add("Template")
        self.tab_base = self.tabs.add("Maestro")
        self.tab_match = self.tabs.add("Tipo de busqueda")
        self.tab_complete = self.tabs.add("Completar resultado")
        self.tab_save = self.tabs.add("Ejecutar y logs")
        self.tab_driver = self.tabs.add("Generar Drivers")

        for tab in [self.tab_template, self.tab_base, self.tab_match, self.tab_complete, self.tab_driver, self.tab_save]:
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)

        self.template_panel = ExcelConfigPanel(
            self.tab_template,
            title="Template / Excel de entrada",
            role="template",
            log_callback=self.log,
        )
        self.template_panel.grid(row=0, column=0, sticky="nsew")

        self.base_panel = ExcelConfigPanel(
            self.tab_base,
            title="Base / Excel donde se busca",
            role="base",
            log_callback=self.log,
        )
        self.base_panel.grid(row=0, column=0, sticky="nsew")

        self._build_match_tab()
        self._build_completion_tab()
        self._build_execute_tab()
        self._build_driver_tab()

    def _build_match_tab(self):
        frame = ctk.CTkScrollableFrame(self.tab_match)
        frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        frame.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(frame, text="Configuracion de match", font=ctk.CTkFont(size=24, weight="bold"))
        title.grid(row=0, column=0, sticky="w", padx=16, pady=(16, 8))

        search_box = ctk.CTkFrame(frame)
        search_box.grid(row=1, column=0, sticky="ew", padx=16, pady=8)
        search_box.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(search_box, text="Tipo de busqueda", font=ctk.CTkFont(size=18, weight="bold"))
        label.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        radio_1 = ctk.CTkRadioButton(
            search_box,
            text="Similaridad / Fuzzy",
            variable=self.search_type_var,
            value="similaridad",
        )
        radio_1.grid(row=1, column=0, sticky="w", padx=12, pady=6)

        radio_2 = ctk.CTkRadioButton(
            search_box,
            text="Exacta / ==",
            variable=self.search_type_var,
            value="exacta",
        )
        radio_2.grid(row=2, column=0, sticky="w", padx=12, pady=(6, 12))

        threshold_box = ctk.CTkFrame(frame)
        threshold_box.grid(row=2, column=0, sticky="ew", padx=16, pady=8)
        threshold_box.grid_columnconfigure((0, 1), weight=1)

        label_threshold = ctk.CTkLabel(
            threshold_box,
            text="Rangos para similaridad",
            font=ctk.CTkFont(size=18, weight="bold"),
        )
        label_threshold.grid(row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(12, 6))

        ctk.CTkLabel(threshold_box, text="EXACTO desde").grid(row=1, column=0, sticky="w", padx=12, pady=4)
        self.score_exacto_entry = ctk.CTkEntry(threshold_box)
        self.score_exacto_entry.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 12))
        self.score_exacto_entry.insert(0, "92")

        ctk.CTkLabel(threshold_box, text="APROXIMADO desde").grid(row=1, column=1, sticky="w", padx=12, pady=4)
        self.score_aproximado_entry = ctk.CTkEntry(threshold_box)
        self.score_aproximado_entry.grid(row=2, column=1, sticky="ew", padx=12, pady=(0, 12))
        self.score_aproximado_entry.insert(0, "75")

        info = ctk.CTkLabel(
            threshold_box,
            text="Regla: score >= EXACTO => EXACTO | score >= APROXIMADO => APROXIMADO | menor => NO ENCONTRADO",
            anchor="w",
        )
        info.grid(row=3, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))

        output_box = ctk.CTkFrame(frame)
        output_box.grid(row=3, column=0, sticky="ew", padx=16, pady=8)
        output_box.grid_columnconfigure(0, weight=1)

        label_output = ctk.CTkLabel(output_box, text="Salida", font=ctk.CTkFont(size=18, weight="bold"))
        label_output.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        out_1 = ctk.CTkRadioButton(
            output_box,
            text="Generar Excel nuevo",
            variable=self.output_mode_var,
            value="excel_nuevo",
        )
        out_1.grid(row=1, column=0, sticky="w", padx=12, pady=6)

        out_2 = ctk.CTkRadioButton(
            output_box,
            text="Agregar hoja al Excel template",
            variable=self.output_mode_var,
            value="nueva_hoja_template",
        )
        out_2.grid(row=2, column=0, sticky="w", padx=12, pady=6)

        ctk.CTkLabel(output_box, text="Nombre de hoja resultado").grid(row=3, column=0, sticky="w", padx=12, pady=(10, 2))
        self.output_sheet_entry = ctk.CTkEntry(output_box)
        self.output_sheet_entry.grid(row=4, column=0, sticky="ew", padx=12, pady=4)
        self.output_sheet_entry.insert(0, "RESULTADO_BUSQUEDA")

        tech_check = ctk.CTkCheckBox(
            output_box,
            text="Incluir columnas tecnicas: ESTADO_MATCH, SCORE_MATCH, VALIDADO, FUENTE, OBSERVACION",
            variable=self.include_technical_cols_var,
        )
        tech_check.grid(row=5, column=0, sticky="w", padx=12, pady=(10, 12))

    def _build_completion_tab(self):
        self.completion_panel = CompletionPanel(
            self.tab_complete,
            base_config_getter=self.base_panel.build_config,
            match_config_getter=self.build_match_config_only,
            log_callback=self.log,
        )
        self.completion_panel.grid(row=0, column=0, sticky="nsew")

    def _build_driver_tab(self):
        self.driver_panel = DriverPanel(
            self.tab_driver,
            log_callback=self.log,
        )
        self.driver_panel.grid(row=0, column=0, sticky="nsew")

    def build_match_config_only(self) -> dict:
        score_exacto = self.score_exacto_entry.get().strip()
        score_aproximado = self.score_aproximado_entry.get().strip()

        if not score_exacto.isdigit() or not score_aproximado.isdigit():
            raise ValueError("Los rangos de score deben ser numeros enteros")

        score_exacto = int(score_exacto)
        score_aproximado = int(score_aproximado)

        if score_aproximado > score_exacto:
            raise ValueError("APROXIMADO desde no puede ser mayor que EXACTO desde")

        if score_exacto < 0 or score_exacto > 100 or score_aproximado < 0 or score_aproximado > 100:
            raise ValueError("Los scores deben estar entre 0 y 100")

        return {
            "tipo_busqueda": self.search_type_var.get(),
            "score_exacto": score_exacto,
            "score_aproximado": score_aproximado,
        }

    def _build_execute_tab(self):
        self.tab_save.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(self.tab_save)
        top.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 8))
        top.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(top, text="Ejecutar busqueda", font=ctk.CTkFont(size=24, weight="bold"))
        title.grid(row=0, column=0, sticky="w", padx=16, pady=(16, 8))

        ctk.CTkLabel(
            top,
            text="Puedes ejecutar sin guardar. El nombre del perfil solo se usa si deseas reutilizar esta configuracion luego.",
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 8))

        ctk.CTkLabel(top, text="Nombre del perfil opcional").grid(row=2, column=0, sticky="w", padx=16, pady=(8, 2))
        self.profile_name_entry = ctk.CTkEntry(top, placeholder_text="Ej: Busqueda_Personal_Maestro")
        self.profile_name_entry.grid(row=3, column=0, sticky="ew", padx=16, pady=4)

        buttons = ctk.CTkFrame(top)
        buttons.grid(row=4, column=0, sticky="ew", padx=16, pady=(10, 16))
        buttons.grid_columnconfigure((0, 1, 2), weight=1)

        btn_validate = ctk.CTkButton(buttons, text="Validar configuracion", command=self.validate_runtime_config)
        btn_validate.grid(row=0, column=0, sticky="ew", padx=6, pady=6)

        self.btn_execute = ctk.CTkButton(buttons, text="Ejecutar busqueda", command=self.execute_search)
        self.btn_execute.grid(row=0, column=1, sticky="ew", padx=6, pady=6)

        btn_save = ctk.CTkButton(buttons, text="Guardar perfil", command=self.save_profile)
        btn_save.grid(row=0, column=2, sticky="ew", padx=6, pady=6)

        self.log_textbox = ctk.CTkTextbox(self.tab_save)
        self.log_textbox.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.set_log_text("Listo. Configura Template, Base, Match y luego ejecuta la busqueda.")

    def set_textbox(self, textbox, text: str):
        textbox.configure(state="normal")
        textbox.delete("1.0", "end")
        textbox.insert("1.0", text)
        textbox.configure(state="disabled")

    def set_log_text(self, text: str):
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.insert("1.0", text)
        self.log_textbox.configure(state="disabled")

    def log(self, text: str):
        self.log_textbox.configure(state="normal")
        current = self.log_textbox.get("1.0", "end").strip()
        if current:
            new_text = current + "\n" + text
        else:
            new_text = text
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.insert("1.0", new_text)
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def build_runtime_config(self) -> dict:
        score_exacto = self.score_exacto_entry.get().strip()
        score_aproximado = self.score_aproximado_entry.get().strip()

        if not score_exacto.isdigit() or not score_aproximado.isdigit():
            raise ValueError("Los rangos de score deben ser numeros enteros")

        score_exacto = int(score_exacto)
        score_aproximado = int(score_aproximado)

        if score_aproximado > score_exacto:
            raise ValueError("APROXIMADO desde no puede ser mayor que EXACTO desde")

        if score_exacto < 0 or score_exacto > 100 or score_aproximado < 0 or score_aproximado > 100:
            raise ValueError("Los scores deben estar entre 0 y 100")

        output_sheet = self.output_sheet_entry.get().strip()
        if not output_sheet:
            raise ValueError("Coloca un nombre para la hoja resultado")

        template_config = self.template_panel.build_config()
        base_config = self.base_panel.build_config()

        return {
            "tipo": "CUSTOM_SEARCH_RUNTIME",
            "version": 1,
            "checkpoint": 2,
            "template": template_config,
            "base": base_config,
            "match": {
                "tipo_busqueda": self.search_type_var.get(),
                "score_exacto": score_exacto,
                "score_aproximado": score_aproximado,
            },
            "salida": {
                "modo": self.output_mode_var.get(),
                "nombre_hoja_resultado": output_sheet,
                "incluir_columnas_tecnicas": bool(self.include_technical_cols_var.get()),
            },
        }

    def build_profile_data(self) -> dict:
        profile_name = self.profile_name_entry.get().strip()
        if not profile_name:
            raise ValueError("Coloca un nombre para guardar el perfil")

        profile = self.build_runtime_config()
        profile["tipo"] = "CUSTOM_SEARCH_PROFILE"
        profile["nombre_perfil"] = profile_name

        return profile

    def validate_runtime_config(self):
        try:
            config = self.build_runtime_config()

            lines = [
                "CONFIGURACION VALIDA",
                "",
                f"Tipo busqueda: {config['match']['tipo_busqueda']}",
                f"Rango EXACTO desde: {config['match']['score_exacto']}",
                f"Rango APROXIMADO desde: {config['match']['score_aproximado']}",
                "",
                "Template:",
                f"- Archivo: {config['template']['archivo']['nombre']}",
                f"- Hoja: {config['template']['hoja']['nombre_detectado']}",
                f"- Header: {config['template']['fila_header']}",
                f"- Col busqueda: {config['template']['columna_busqueda']['alias']} "
                f"({config['template']['columna_busqueda']['header_detectado']})",
                f"- Columnas configuradas: {len(config['template']['columnas'])}",
                "",
                "Base:",
                f"- Archivo: {config['base']['archivo']['nombre']}",
                f"- Hoja: {config['base']['hoja']['nombre_detectado']}",
                f"- Header: {config['base']['fila_header']}",
                f"- Col busqueda: {config['base']['columna_busqueda']['alias']} "
                f"({config['base']['columna_busqueda']['header_detectado']})",
                f"- Columnas configuradas: {len(config['base']['columnas'])}",
                "",
                "Salida:",
                f"- Modo: {config['salida']['modo']}",
                f"- Hoja resultado: {config['salida']['nombre_hoja_resultado']}",
            ]

            self.set_log_text("\n".join(lines))
            messagebox.showinfo("OK", "La configuracion es valida")
        except Exception as error:
            self.log(f"ERROR VALIDACION: {error}")
            messagebox.showerror("Error", str(error))

    def execute_search(self):
        def failed(error):
            self.log(f"ERROR EJECUCION: {error}")
            messagebox.showerror("Error", str(error))

        try:
            tasks = tasks_for(self)
            if tasks.busy:
                raise RuntimeError("Ya hay un proceso en curso. Espere a que termine antes de ejecutar otro.")
            config = deepcopy(self.build_runtime_config())
            self.log("Iniciando busqueda...")
        except Exception as error:
            failed(error)
            return

        def work(report):
            result_df, summary = run_search(config, progress=report)
            report("Guardando resultado...")
            output_path = export_result(result_df, config)
            return output_path, summary, result_df.attrs["excel_output_sheets"]

        def finished(result):
            output_path, summary, output_sheets = result
            self.log("Busqueda terminada.")
            self.log(f"Filas template: {summary['template_rows']}")
            self.log(f"Filas base: {summary['base_rows']}")
            self.log(f"Resultados: {summary['result_rows']}")
            self.log(f"EXACTO: {summary['exactos']}")
            self.log(f"APROXIMADO: {summary['aproximados']}")
            self.log(f"NO ENCONTRADO: {summary['no_encontrados']}")

            self.log(f"Excel generado: {output_path}")
            self.log(f"Hojas generadas: {', '.join(output_sheets)}")
            messagebox.showinfo("Busqueda terminada", f"Resultado generado en:\n{output_path}\nHojas: {', '.join(output_sheets)}")

        tasks.start(
            work, self.btn_execute, "Ejecutar busqueda", finished, failed,
            lambda text: self.log(text),
        )

    def save_profile(self):
        try:
            profile = self.build_profile_data()
            path = save_search_profile(profile["nombre_perfil"], profile)

            self.log(f"Perfil guardado correctamente: {path}")
            messagebox.showinfo("Perfil guardado", f"Perfil guardado en:\n{path}")
        except Exception as error:
            self.log(f"ERROR GUARDANDO PERFIL: {error}")
            messagebox.showerror("Error", str(error))

    def select_completion_file(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar resultado anterior",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )

        if not file_path:
            return

        self.completion_file_path = file_path
        self.completion_header_info = None
        self.completion_file_label.configure(text=Path(file_path).name)
        self.log(f"[completar] Resultado anterior seleccionado: {Path(file_path).name}")

    def load_completion_headers(self):
        try:
            if not self.completion_file_path:
                raise ValueError("Primero selecciona un resultado anterior")

            sheet_ref = self.completion_sheet_entry.get().strip()

            if not self.completion_header_entry.get().strip().isdigit():
                raise ValueError("La fila header debe ser un numero")

            header_row = int(self.completion_header_entry.get().strip())

            info = read_result_headers(
                file_path=self.completion_file_path,
                sheet_ref=sheet_ref,
                header_row=header_row,
            )

            self.completion_header_info = info

            lines = [
                f"Hoja detectada: {info['sheet_index'] + 1} - {info['sheet_name']}",
                f"Fila header: {info['header_row']}",
                f"Filas: {info['max_row']} | Columnas: {info['max_column']}",
                "",
                "Primeras columnas detectadas:",
            ]

            for idx, header in enumerate(info["first_10_headers"], start=1):
                lines.append(f"{idx}. {header}")

            self.set_textbox(self.completion_headers_textbox, "\n".join(lines))
            self.log("[completar] Headers del resultado cargados correctamente")
        except Exception as error:
            self.log(f"[completar] ERROR LEYENDO RESULTADO: {error}")
            messagebox.showerror("Error", str(error))

    def build_completion_config(self) -> dict:
        if not self.completion_file_path:
            raise ValueError("Selecciona un resultado anterior")

        if not self.completion_header_info:
            raise ValueError("Primero lee las columnas del resultado anterior")

        validado_col = self.completion_validado_entry.get().strip()
        search_col = self.completion_search_entry.get().strip()

        if not validado_col:
            raise ValueError("Coloca la columna VALIDADO")

        if not search_col:
            raise ValueError("Coloca la columna valor a buscar")

        score_exacto = self.score_exacto_entry.get().strip()
        score_aproximado = self.score_aproximado_entry.get().strip()

        if not score_exacto.isdigit() or not score_aproximado.isdigit():
            raise ValueError("Los rangos de score deben ser numeros enteros")

        score_exacto = int(score_exacto)
        score_aproximado = int(score_aproximado)

        if score_aproximado > score_exacto:
            raise ValueError("APROXIMADO desde no puede ser mayor que EXACTO desde")

        base_config = self.base_panel.build_config()

        return {
            "tipo": "CUSTOM_COMPLETION_RUNTIME",
            "version": 1,
            "checkpoint": 3,
            "previous_result": {
                "archivo": {
                    "ruta": str(Path(self.completion_file_path)),
                    "nombre": Path(self.completion_file_path).name,
                },
                "hoja": {
                    "entrada_usuario": self.completion_sheet_entry.get().strip(),
                    "indice_0": self.completion_header_info["sheet_index"],
                    "indice_1": self.completion_header_info["sheet_index"] + 1,
                    "nombre_detectado": self.completion_header_info["sheet_name"],
                },
                "fila_header": self.completion_header_info["header_row"],
                "columna_validado": validado_col,
                "columna_busqueda": search_col,
            },
            "base": base_config,
            "match": {
                "tipo_busqueda": self.search_type_var.get(),
                "score_exacto": score_exacto,
                "score_aproximado": score_aproximado,
            },
        }

    def validate_completion_config(self):
        try:
            config = self.build_completion_config()

            lines = [
                "CONFIGURACION DE COMPLETADO VALIDA",
                "",
                "Resultado anterior:",
                f"- Archivo: {config['previous_result']['archivo']['nombre']}",
                f"- Hoja: {config['previous_result']['hoja']['nombre_detectado']}",
                f"- Header: {config['previous_result']['fila_header']}",
                f"- Columna VALIDADO: {config['previous_result']['columna_validado']}",
                f"- Columna a buscar: {config['previous_result']['columna_busqueda']}",
                "",
                "Nueva base:",
                f"- Archivo: {config['base']['archivo']['nombre']}",
                f"- Hoja: {config['base']['hoja']['nombre_detectado']}",
                f"- Col busqueda: {config['base']['columna_busqueda']['alias']} "
                f"({config['base']['columna_busqueda']['header_detectado']})",
                "",
                f"Tipo busqueda: {config['match']['tipo_busqueda']}",
            ]

            self.set_log_text("\n".join(lines))
            messagebox.showinfo("OK", "La configuracion de completado es valida")
        except Exception as error:
            self.log(f"[completar] ERROR VALIDACION: {error}")
            messagebox.showerror("Error", str(error))

    def execute_completion(self):
        def failed(error):
            self.log(f"[completar] ERROR EJECUCION: {error}")
            messagebox.showerror("Error", str(error))

        try:
            tasks = tasks_for(self)
            if tasks.busy:
                raise RuntimeError("Ya hay un proceso en curso. Espere a que termine antes de ejecutar otro.")
            config = deepcopy(self.build_completion_config())
            self.log("[completar] Iniciando completado...")
        except Exception as error:
            failed(error)
            return

        def work(report):
            result_df, summary = run_completion(config, progress=report)
            report("Guardando resultado...")
            output_path = export_completion_result(result_df, config)
            return output_path, summary, result_df.attrs["excel_output_sheets"]

        def finished(result):
            output_path, summary, output_sheets = result
            self.log("[completar] Proceso terminado.")
            self.log(f"[completar] Filas resultado anterior: {summary['total_filas_resultado']}")
            self.log(f"[completar] Filas reprocesadas: {summary['filas_reprocesadas']}")
            self.log(f"[completar] Filas base: {summary['base_rows']}")
            self.log(f"[completar] EXACTO nuevos: {summary['exactos_nuevos']}")
            self.log(f"[completar] APROXIMADO nuevos: {summary['aproximados_nuevos']}")
            self.log(f"[completar] NO ENCONTRADO finales: {summary['no_encontrados_finales']}")

            self.log(f"[completar] Excel completado generado: {output_path}")
            self.log(f"[completar] Hojas generadas: {', '.join(output_sheets)}")
            messagebox.showinfo("Completado terminado", f"Resultado generado en:\n{output_path}\nHojas: {', '.join(output_sheets)}")

        tasks.start(
            work, self.btn_complete, "Completar resultado", finished, failed,
            lambda text: self.log(text),
        )

    def close_resources(self):
        self.template_panel.close_session()
        self.base_panel.close_session()

        if hasattr(self, "driver_panel"):
            self.driver_panel.close_resources()
