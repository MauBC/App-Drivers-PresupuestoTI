from copy import deepcopy
from src.gui.background_tasks import tasks_for
from pathlib import Path
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox

from src.core.completion_executor import (
    get_result_sheet_names,
    read_result_sheet_sample,
    read_result_headers,
    run_completion,
    export_completion_result,
    validate_destination_column,
    REQUIRED_TECHNICAL_COLUMNS,
)


class CompletionPanel(ctk.CTkScrollableFrame):
    def __init__(
        self,
        master,
        base_config_getter,
        match_config_getter,
        log_callback=None,
    ):
        super().__init__(master)

        self.base_config_getter = base_config_getter
        self.match_config_getter = match_config_getter
        self.log_callback = log_callback

        self.completion_file_path = None
        self.sheet_names = []
        self.header_info = None
        self.mapping_rows = {}
        self.mapping_counter = 0

        self.output_mode_var = tk.StringVar(value="excel_nuevo")

        self.grid_columnconfigure(0, weight=1)

        self._build_layout()

    def _log(self, text: str):
        if self.log_callback:
            self.log_callback(text)

    def _build_layout(self):
        title = ctk.CTkLabel(
            self,
            text="Completar resultado existente",
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        title.grid(row=0, column=0, sticky="w", padx=16, pady=(16, 8))

        info = ctk.CTkLabel(
            self,
            text="Reprocesa solo filas con VALIDADO = 0. No crea columnas de datos nuevas; solo actualiza columnas existentes.",
            anchor="w",
        )
        info.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))

        self._build_result_file_section(row=2)
        self._build_mapping_section(row=3)
        self._build_output_section(row=4)

    def _build_result_file_section(self, row: int):
        box = ctk.CTkFrame(self)
        box.grid(row=row, column=0, sticky="ew", padx=16, pady=8)
        box.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            box,
            text="1. Resultado anterior",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        self.file_label = ctk.CTkLabel(box, text="Ningun archivo seleccionado", anchor="w")
        self.file_label.grid(row=1, column=0, sticky="ew", padx=12, pady=4)

        btn_file = ctk.CTkButton(box, text="Seleccionar resultado anterior", command=self.select_file)
        btn_file.grid(row=2, column=0, sticky="ew", padx=12, pady=6)

        ctk.CTkLabel(box, text="Hoja resultado: posicion o nombre").grid(
            row=3, column=0, sticky="w", padx=12, pady=(10, 2)
        )
        self.sheet_entry = ctk.CTkEntry(box, placeholder_text="Ej: 1 o RESULTADO_BUSQUEDA")
        self.sheet_entry.grid(row=4, column=0, sticky="ew", padx=12, pady=4)
        self.sheet_entry.insert(0, "1")

        btn_validate_sheet = ctk.CTkButton(
            box,
            text="Validar hoja y ver muestra",
            command=self.validate_sheet,
        )
        btn_validate_sheet.grid(row=5, column=0, sticky="ew", padx=12, pady=8)

        self.sheet_result_label = ctk.CTkLabel(box, text="Hoja detectada: -", anchor="w")
        self.sheet_result_label.grid(row=6, column=0, sticky="ew", padx=12, pady=(2, 4))

        self.sheet_textbox = ctk.CTkTextbox(box, height=145)
        self.sheet_textbox.grid(row=7, column=0, sticky="ew", padx=12, pady=(4, 12))
        self._set_textbox(self.sheet_textbox, "Aqui apareceran las hojas disponibles y la muestra del resultado anterior.")

        ctk.CTkLabel(box, text="Fila header del resultado").grid(
            row=8, column=0, sticky="w", padx=12, pady=(8, 2)
        )
        self.header_entry = ctk.CTkEntry(box)
        self.header_entry.grid(row=9, column=0, sticky="ew", padx=12, pady=4)
        self.header_entry.insert(0, "1")

        btn_headers = ctk.CTkButton(box, text="Leer columnas del resultado", command=self.load_headers)
        btn_headers.grid(row=10, column=0, sticky="ew", padx=12, pady=8)

        self.headers_textbox = ctk.CTkTextbox(box, height=150)
        self.headers_textbox.grid(row=11, column=0, sticky="ew", padx=12, pady=(4, 12))
        self._set_textbox(self.headers_textbox, "Aqui apareceran los primeros 30 headers del resultado anterior.")

    def _build_mapping_section(self, row: int):
        box = ctk.CTkFrame(self)
        box.grid(row=row, column=0, sticky="ew", padx=16, pady=8)
        box.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            box,
            text="2. Mapeo hacia columnas existentes",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        info = ctk.CTkLabel(
            box,
            text="Genera un mapeo desde las columnas marcadas como Mostrar en la Base. Cada columna de Base debe tener una columna destino existente.",
            anchor="w",
        )
        info.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))

        controls = ctk.CTkFrame(box)
        controls.grid(row=2, column=0, sticky="ew", padx=12, pady=6)
        controls.grid_columnconfigure((0, 1), weight=1)

        btn_generate = ctk.CTkButton(
            controls,
            text="Generar mapeo desde Base",
            command=self.generate_mapping_from_base,
        )
        btn_generate.grid(row=0, column=0, sticky="ew", padx=6, pady=6)

        btn_resolve = ctk.CTkButton(
            controls,
            text="Resolver mapeo",
            command=self.resolve_mapping,
        )
        btn_resolve.grid(row=0, column=1, sticky="ew", padx=6, pady=6)

        self.mapping_container = ctk.CTkFrame(box)
        self.mapping_container.grid(row=3, column=0, sticky="ew", padx=12, pady=(8, 12))
        self.mapping_container.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(self.mapping_container)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=2)
        header.grid_columnconfigure(1, weight=2)
        header.grid_columnconfigure(2, weight=2)
        header.grid_columnconfigure(3, weight=3)

        labels = ["Base alias", "Base header", "Destino", "Destino detectado"]
        for col, text in enumerate(labels):
            ctk.CTkLabel(header, text=text, font=ctk.CTkFont(weight="bold")).grid(
                row=0, column=col, sticky="w", padx=6, pady=6
            )

    def _build_output_section(self, row: int):
        box = ctk.CTkFrame(self)
        box.grid(row=row, column=0, sticky="ew", padx=16, pady=8)
        box.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            box,
            text="3. Salida del completado",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        radio_1 = ctk.CTkRadioButton(
            box,
            text="Generar nuevo Excel completado",
            variable=self.output_mode_var,
            value="excel_nuevo",
        )
        radio_1.grid(row=1, column=0, sticky="w", padx=12, pady=6)

        radio_2 = ctk.CTkRadioButton(
            box,
            text="Agregar hoja nueva al mismo Excel resultado",
            variable=self.output_mode_var,
            value="hoja_nueva_mismo_excel",
        )
        radio_2.grid(row=2, column=0, sticky="w", padx=12, pady=6)

        ctk.CTkLabel(box, text="Nombre de hoja completada").grid(
            row=3, column=0, sticky="w", padx=12, pady=(10, 2)
        )
        self.output_sheet_entry = ctk.CTkEntry(box)
        self.output_sheet_entry.grid(row=4, column=0, sticky="ew", padx=12, pady=4)
        self.output_sheet_entry.insert(0, "RESULTADO_COMPLETADO")

        buttons = ctk.CTkFrame(box)
        buttons.grid(row=5, column=0, sticky="ew", padx=12, pady=(12, 14))
        buttons.grid_columnconfigure((0, 1), weight=1)

        btn_validate = ctk.CTkButton(buttons, text="Validar completado", command=self.validate_config)
        btn_validate.grid(row=0, column=0, sticky="ew", padx=6, pady=6)

        self.btn_complete = ctk.CTkButton(buttons, text="Completar resultado", command=self.execute_completion)
        self.btn_complete.grid(row=0, column=1, sticky="ew", padx=6, pady=6)

    def _set_textbox(self, textbox, text: str):
        textbox.configure(state="normal")
        textbox.delete("1.0", "end")
        textbox.insert("1.0", text)
        textbox.configure(state="disabled")

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar resultado anterior",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )

        if not file_path:
            return

        try:
            self.completion_file_path = file_path
            self.header_info = None
            self.sheet_names = get_result_sheet_names(file_path)
            self.file_label.configure(text=Path(file_path).name)

            lines = ["Hojas detectadas:"]
            for index, sheet_name in enumerate(self.sheet_names, start=1):
                lines.append(f"{index}. {sheet_name}")

            self._set_textbox(self.sheet_textbox, "\n".join(lines))
            self._set_textbox(self.headers_textbox, "Coloca la hoja y la fila header para leer columnas.")
            self._log(f"[completar] Resultado anterior seleccionado: {Path(file_path).name}")

        except Exception as error:
            self._log(f"[completar] ERROR CARGANDO RESULTADO: {error}")
            messagebox.showerror("Error", str(error))

    def validate_sheet(self):
        try:
            if not self.completion_file_path:
                raise ValueError("Primero selecciona un resultado anterior")

            sheet_ref = self.sheet_entry.get().strip()
            sample = read_result_sheet_sample(self.completion_file_path, sheet_ref)

            lines = [
                f"Hoja detectada: {sample['sheet_index'] + 1} - {sample['sheet_name']}",
                f"Filas: {sample['max_row']} | Columnas: {sample['max_column']}",
                "",
                "Muestra de las primeras 5 filas:",
            ]

            for row in sample["rows"]:
                lines.append(" | ".join(row))

            self.sheet_result_label.configure(
                text=f"Hoja detectada: {sample['sheet_index'] + 1} - {sample['sheet_name']}"
            )
            self._set_textbox(self.sheet_textbox, "\n".join(lines))
            self._log(f"[completar] Hoja validada: {sample['sheet_name']}")

        except Exception as error:
            self._log(f"[completar] ERROR VALIDANDO HOJA: {error}")
            messagebox.showerror("Error", str(error))

    def load_headers(self):
        try:
            if not self.completion_file_path:
                raise ValueError("Primero selecciona un resultado anterior")

            sheet_ref = self.sheet_entry.get().strip()

            if not self.header_entry.get().strip().isdigit():
                raise ValueError("La fila header debe ser un numero")

            header_row = int(self.header_entry.get().strip())

            info = read_result_headers(
                file_path=self.completion_file_path,
                sheet_ref=sheet_ref,
                header_row=header_row,
            )

            self.header_info = info

            lines = [
                f"Hoja detectada: {info['sheet_index'] + 1} - {info['sheet_name']}",
                f"Fila header: {info['header_row']}",
                f"Filas: {info['max_row']} | Columnas: {info['max_column']}",
                "",
                "Primeros 30 headers detectados:",
            ]

            for idx, header in enumerate(info.get("first_30_headers", info["headers"][:30]), start=1):
                lines.append(f"{idx}. {header}")

            if info["missing_required_technical"]:
                lines.append("")
                lines.append("FALTAN COLUMNAS TECNICAS:")
                for col in info["missing_required_technical"]:
                    lines.append(f"- {col}")
            else:
                lines.append("")
                lines.append("Columnas tecnicas obligatorias: OK")

            self._set_textbox(self.headers_textbox, "\n".join(lines))
            self._log("[completar] Headers del resultado cargados correctamente")
        except Exception as error:
            self._log(f"[completar] ERROR LEYENDO RESULTADO: {error}")
            messagebox.showerror("Error", str(error))

    def clear_mapping_rows(self):
        for row_id, data in list(self.mapping_rows.items()):
            data["frame"].destroy()
            del self.mapping_rows[row_id]

        self.mapping_counter = 0

    def generate_mapping_from_base(self):
        try:
            base_config = self.base_config_getter()

            selected_cols = [
                col for col in base_config["columnas"]
                if col.get("mostrar_salida", True)
            ]

            if not selected_cols:
                raise ValueError("La Base no tiene columnas marcadas como Mostrar")

            self.clear_mapping_rows()

            for col in selected_cols:
                self.add_mapping_row(col)

            self._log(f"[completar] Mapeo generado desde Base: {len(selected_cols)} columnas")
        except Exception as error:
            self._log(f"[completar] ERROR GENERANDO MAPEO: {error}")
            messagebox.showerror("Error", str(error))

    def add_mapping_row(self, base_col: dict):
        self.mapping_counter += 1
        row_id = f"map_{self.mapping_counter}"

        frame = ctk.CTkFrame(self.mapping_container)
        frame.grid(row=self.mapping_counter, column=0, sticky="ew", pady=4)
        frame.grid_columnconfigure(0, weight=2)
        frame.grid_columnconfigure(1, weight=2)
        frame.grid_columnconfigure(2, weight=2)
        frame.grid_columnconfigure(3, weight=3)

        alias_label = ctk.CTkLabel(frame, text=base_col["alias"], anchor="w")
        alias_label.grid(row=0, column=0, sticky="ew", padx=6, pady=6)

        header_label = ctk.CTkLabel(frame, text=base_col["header_detectado"], anchor="w")
        header_label.grid(row=0, column=1, sticky="ew", padx=6, pady=6)

        destination_entry = ctk.CTkEntry(frame, placeholder_text="Nombre, letra o numero")
        destination_entry.grid(row=0, column=2, sticky="ew", padx=6, pady=6)

        detected_label = ctk.CTkLabel(frame, text="-", anchor="w")
        detected_label.grid(row=0, column=3, sticky="ew", padx=6, pady=6)

        self.mapping_rows[row_id] = {
            "frame": frame,
            "base_alias": base_col["alias"],
            "base_header_detectado": base_col["header_detectado"],
            "destination_entry": destination_entry,
            "detected_label": detected_label,
        }

    def resolve_mapping(self):
        try:
            if not self.header_info:
                raise ValueError("Primero lee las columnas del resultado anterior")

            if self.header_info["missing_required_technical"]:
                raise ValueError(
                    "El Excel no es valido. Faltan columnas tecnicas: "
                    + ", ".join(self.header_info["missing_required_technical"])
                )

            headers = self.header_info["headers"]
            used_destinations = set()
            errors = []
            resolved_count = 0

            for row_id, data in self.mapping_rows.items():
                value = data["destination_entry"].get().strip()

                if not value:
                    data["detected_label"].configure(text="-")
                    errors.append(f"{data['base_alias']}: falta columna destino")
                    continue

                try:
                    destination = validate_destination_column(headers, value)
                    key = destination.strip().upper()

                    if key in used_destinations:
                        raise ValueError("columna destino duplicada")

                    used_destinations.add(key)
                    data["detected_label"].configure(text=destination)
                    resolved_count += 1
                except Exception as error:
                    data["detected_label"].configure(text="ERROR")
                    errors.append(f"{data['base_alias']}: {error}")

            if errors:
                self._log("[completar] Errores en mapeo:\n" + "\n".join(errors))
            else:
                self._log(f"[completar] Mapeo resuelto correctamente: {resolved_count} columnas")
        except Exception as error:
            self._log(f"[completar] ERROR RESOLVIENDO MAPEO: {error}")
            messagebox.showerror("Error", str(error))

    def build_config(self) -> dict:
        if not self.completion_file_path:
            raise ValueError("Selecciona un resultado anterior")

        if not self.header_info:
            raise ValueError("Primero lee las columnas del resultado anterior")

        if self.header_info["missing_required_technical"]:
            raise ValueError(
                "El Excel no es valido. Faltan columnas tecnicas: "
                + ", ".join(self.header_info["missing_required_technical"])
            )

        if not self.mapping_rows:
            raise ValueError("Primero genera el mapeo desde Base")

        output_sheet = self.output_sheet_entry.get().strip()
        if not output_sheet:
            raise ValueError("Coloca un nombre para la hoja completada")

        base_config = self.base_config_getter()
        match_config = self.match_config_getter()

        mappings = []
        headers = self.header_info["headers"]

        for row_id, data in self.mapping_rows.items():
            destination_ref = data["destination_entry"].get().strip()

            if not destination_ref:
                raise ValueError(f"Falta destino para la columna base: {data['base_alias']}")

            destination = validate_destination_column(headers, destination_ref)

            mappings.append({
                "base_alias": data["base_alias"],
                "base_header_detectado": data["base_header_detectado"],
                "destination_column_ref": destination_ref,
                "destination_column_resolved": destination,
            })

        return {
            "tipo": "CUSTOM_COMPLETION_RUNTIME",
            "version": 1,
            "checkpoint": 3,
            "previous_result": {
                "expected_headers": list(headers),
                "archivo": {
                    "ruta": str(Path(self.completion_file_path)),
                    "nombre": Path(self.completion_file_path).name,
                },
                "hoja": {
                    "entrada_usuario": self.sheet_entry.get().strip(),
                    "indice_0": self.header_info["sheet_index"],
                    "indice_1": self.header_info["sheet_index"] + 1,
                    "nombre_detectado": self.header_info["sheet_name"],
                },
                "fila_header": self.header_info["header_row"],
                "columna_busqueda": "VALOR_BUSCADO",
            },
            "base": base_config,
            "match": match_config,
            "column_mappings": mappings,
            "completion_output": {
                "modo": self.output_mode_var.get(),
                "nombre_hoja": output_sheet,
            },
        }

    def validate_config(self):
        try:
            config = self.build_config()

            lines = [
                "CONFIGURACION DE COMPLETADO VALIDA",
                "",
                "Resultado anterior:",
                f"- Archivo: {config['previous_result']['archivo']['nombre']}",
                f"- Hoja: {config['previous_result']['hoja']['nombre_detectado']}",
                f"- Header: {config['previous_result']['fila_header']}",
                "- Se reprocesa: VALIDADO = 0",
                "- Valor a buscar: VALOR_BUSCADO",
                "",
                "Nueva base:",
                f"- Archivo: {config['base']['archivo']['nombre']}",
                f"- Hoja: {config['base']['hoja']['nombre_detectado']}",
                f"- Col busqueda: {config['base']['columna_busqueda']['alias']} "
                f"({config['base']['columna_busqueda']['header_detectado']})",
                "",
                "Mapeo:",
            ]

            for item in config["column_mappings"]:
                lines.append(
                    f"- {item['base_alias']} -> {item['destination_column_resolved']}"
                )

            lines.extend([
                "",
                f"Tipo busqueda: {config['match']['tipo_busqueda']}",
                f"Salida: {config['completion_output']['modo']}",
            ])

            self._log("\n".join(lines))
            messagebox.showinfo("OK", "La configuracion de completado es valida")
        except Exception as error:
            self._log(f"[completar] ERROR VALIDACION: {error}")
            messagebox.showerror("Error", str(error))

    def execute_completion(self):
        def failed(error):
            self._log(f"[completar] ERROR EJECUCION: {error}")
            messagebox.showerror("Error", str(error))

        try:
            tasks = tasks_for(self)
            if tasks.busy:
                raise RuntimeError("Ya hay un proceso en curso. Espere a que termine antes de ejecutar otro.")
            config = deepcopy(self.build_config())
            self._log("[completar] Iniciando completado...")
        except Exception as error:
            failed(error)
            return

        def work(report):
            result_df, summary = run_completion(config, progress=report)
            report("Guardando resultado...")
            output_path = export_completion_result(result_df, config)
            return output_path, summary

        def finished(result):
            output_path, summary = result
            self._log("[completar] Proceso terminado.")
            self._log(f"[completar] Filas resultado anterior: {summary['total_filas_resultado']}")
            self._log(f"[completar] Filas reprocesadas: {summary['filas_reprocesadas']}")
            self._log(f"[completar] Filas base: {summary['base_rows']}")
            self._log(f"[completar] EXACTO nuevos: {summary['exactos_nuevos']}")
            self._log(f"[completar] APROXIMADO nuevos: {summary['aproximados_nuevos']}")
            self._log(f"[completar] NO ENCONTRADO finales: {summary['no_encontrados_finales']}")

            self._log(f"[completar] Resultado generado: {output_path}")
            messagebox.showinfo("Completado terminado", f"Resultado generado en:\n{output_path}")

        tasks.start(
            work, self.btn_complete, "Completar resultado", finished, failed,
            lambda text: self._log(text),
        )
