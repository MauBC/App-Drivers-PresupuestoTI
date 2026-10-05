from pathlib import Path
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox

from src.core.base_profile_store import save_base_profile, load_base_profile
from src.core.excel_inspector import (
    get_sheet_names,
    read_sheet_sample,
    read_headers,
    resolve_column,
)


class ExcelConfigPanel(ctk.CTkScrollableFrame):
    def __init__(self, master, title: str, role: str, log_callback=None):
        super().__init__(master)

        self.title = title
        self.role = role
        self.log_callback = log_callback
        self.enable_base_profiles = "base" in str(role).lower()

        self.file_path = None
        self.excel_session = None
        self.sheet_names = []
        self.header_info = None
        self.headers = []

        self.row_counter = 0
        self.column_rows = {}
        self.search_var = tk.StringVar(value="")

        self.grid_columnconfigure(0, weight=1)

        self._build_layout()

    def close_session(self):
        if self.excel_session is not None:
            self.excel_session.close()
            self.excel_session = None

    def _log(self, text: str):
        if self.log_callback:
            self.log_callback(text)

    def _build_layout(self):
        title_label = ctk.CTkLabel(
            self,
            text=self.title,
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        title_label.grid(row=0, column=0, sticky="w", padx=16, pady=(16, 8))

        current_row = 1

        if self.enable_base_profiles:
            self._build_profile_section(row=current_row)
            current_row += 1

        self._build_file_section(row=current_row)
        current_row += 1

        self._build_header_section(row=current_row)
        current_row += 1

        self._build_columns_section(row=current_row)

    def _build_profile_section(self, row: int):
        section = ctk.CTkFrame(self)
        section.grid(row=row, column=0, sticky="ew", padx=16, pady=8)
        section.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(
            section,
            text="0. Configuracion reutilizable de Base",
            font=ctk.CTkFont(size=16, weight="bold"),
        )
        label.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        info = ctk.CTkLabel(
            section,
            text="Guarda o carga archivo, hoja, header, columnas, alias y columna de busqueda de esta Base.",
            anchor="w",
        )
        info.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))

        ctk.CTkLabel(section, text="Nombre de configuracion").grid(
            row=2, column=0, sticky="w", padx=12, pady=(4, 2)
        )

        self.base_profile_name_entry = ctk.CTkEntry(
            section,
            placeholder_text="Ej: Maestro_Personal, Maestro_Altas, Base_Drivers"
        )
        self.base_profile_name_entry.grid(row=3, column=0, sticky="ew", padx=12, pady=4)

        buttons = ctk.CTkFrame(section)
        buttons.grid(row=4, column=0, sticky="ew", padx=12, pady=(8, 12))
        buttons.grid_columnconfigure((0, 1), weight=1)

        btn_save = ctk.CTkButton(
            buttons,
            text="Guardar config base",
            command=self.save_current_base_profile,
        )
        btn_save.grid(row=0, column=0, sticky="ew", padx=6, pady=6)

        btn_load = ctk.CTkButton(
            buttons,
            text="Cargar config base",
            command=self.load_base_profile_file,
        )
        btn_load.grid(row=0, column=1, sticky="ew", padx=6, pady=6)

    def _build_file_section(self, row: int):
        section = ctk.CTkFrame(self)
        section.grid(row=row, column=0, sticky="ew", padx=16, pady=8)
        section.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(section, text="1. Archivo y hoja", font=ctk.CTkFont(size=16, weight="bold"))
        label.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        self.file_label = ctk.CTkLabel(section, text="Ningun archivo seleccionado", anchor="w")
        self.file_label.grid(row=1, column=0, sticky="ew", padx=12, pady=4)

        btn_file = ctk.CTkButton(section, text="Seleccionar Excel", command=self.select_file)
        btn_file.grid(row=2, column=0, sticky="ew", padx=12, pady=6)

        sheet_label = ctk.CTkLabel(section, text="Hoja: posicion 1, 2, 3... o nombre")
        sheet_label.grid(row=3, column=0, sticky="w", padx=12, pady=(10, 2))

        self.sheet_entry = ctk.CTkEntry(section, placeholder_text="Ej: 1 o Personal")
        self.sheet_entry.grid(row=4, column=0, sticky="ew", padx=12, pady=4)

        btn_sheet = ctk.CTkButton(section, text="Validar hoja y ver muestra", command=self.validate_sheet)
        btn_sheet.grid(row=5, column=0, sticky="ew", padx=12, pady=6)

        self.sheet_result_label = ctk.CTkLabel(section, text="Hoja detectada: -", anchor="w")
        self.sheet_result_label.grid(row=6, column=0, sticky="ew", padx=12, pady=(4, 4))

        self.sheet_textbox = ctk.CTkTextbox(section, height=150)
        self.sheet_textbox.grid(row=7, column=0, sticky="ew", padx=12, pady=(4, 12))
        self._set_textbox(self.sheet_textbox, "Aqui apareceran las hojas y la muestra del archivo.")

    def _build_header_section(self, row: int):
        section = ctk.CTkFrame(self)
        section.grid(row=row, column=0, sticky="ew", padx=16, pady=8)
        section.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(section, text="2. Header", font=ctk.CTkFont(size=16, weight="bold"))
        label.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        header_label = ctk.CTkLabel(section, text="Fila donde esta el header")
        header_label.grid(row=1, column=0, sticky="w", padx=12, pady=(6, 2))

        self.header_entry = ctk.CTkEntry(section, placeholder_text="Ej: 6")
        self.header_entry.grid(row=2, column=0, sticky="ew", padx=12, pady=4)

        btn_header = ctk.CTkButton(section, text="Leer header", command=self.load_headers)
        btn_header.grid(row=3, column=0, sticky="ew", padx=12, pady=6)

        self.header_textbox = ctk.CTkTextbox(section, height=150)
        self.header_textbox.grid(row=4, column=0, sticky="ew", padx=12, pady=(4, 12))
        self._set_textbox(self.header_textbox, "Aqui apareceran los primeros 30 headers detectados.")

    def _build_columns_section(self, row: int):
        section = ctk.CTkFrame(self)
        section.grid(row=row, column=0, sticky="ew", padx=16, pady=8)
        section.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(section, text="3. Columnas", font=ctk.CTkFont(size=16, weight="bold"))
        label.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        info = ctk.CTkLabel(
            section,
            text="Agrega columnas por letra o numero. Debes marcar exactamente 1 como columna de busqueda.",
            anchor="w",
        )
        info.grid(row=1, column=0, sticky="ew", padx=12, pady=(2, 8))

        buttons = ctk.CTkFrame(section)
        buttons.grid(row=2, column=0, sticky="ew", padx=12, pady=6)
        buttons.grid_columnconfigure((0, 1, 2, 3), weight=1)

        btn_add = ctk.CTkButton(buttons, text="Anadir columna", command=self.add_column_row)
        btn_add.grid(row=0, column=0, sticky="ew", padx=4, pady=4)

        btn_all = ctk.CTkButton(buttons, text="Usar todas", command=self.use_all_headers)
        btn_all.grid(row=0, column=1, sticky="ew", padx=4, pady=4)

        btn_resolve = ctk.CTkButton(buttons, text="Resolver", command=self.resolve_all_columns)
        btn_resolve.grid(row=0, column=2, sticky="ew", padx=4, pady=4)

        btn_clear = ctk.CTkButton(buttons, text="Limpiar", command=self.clear_column_rows)
        btn_clear.grid(row=0, column=3, sticky="ew", padx=4, pady=4)

        self.rows_container = ctk.CTkFrame(section)
        self.rows_container.grid(row=3, column=0, sticky="ew", padx=12, pady=(8, 12))
        self.rows_container.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(self.rows_container)
        header.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 4))
        header.grid_columnconfigure(0, weight=2)
        header.grid_columnconfigure(1, weight=1)
        header.grid_columnconfigure(2, weight=3)
        header.grid_columnconfigure(3, weight=1)
        header.grid_columnconfigure(4, weight=1)
        header.grid_columnconfigure(5, weight=1)

        labels = ["Alias salida", "Col", "Header detectado", "Mostrar", "Buscar", "Quitar"]
        for col, text in enumerate(labels):
            ctk.CTkLabel(header, text=text, font=ctk.CTkFont(weight="bold")).grid(
                row=0, column=col, sticky="w", padx=6, pady=6
            )

        self.add_column_row()

    def _set_textbox(self, textbox, text: str):
        textbox.configure(state="normal")
        textbox.delete("1.0", "end")
        textbox.insert("1.0", text)
        textbox.configure(state="disabled")

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar Excel",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )

        if not file_path:
            return

        try:
            self._log(f"[{self.role}] Cargando archivo. Esto puede tardar unos segundos si el Excel es grande...")

            sheet_names = get_sheet_names(file_path)

            self.close_session()
            self.file_path = file_path
            self.sheet_names = sheet_names

            self.header_info = None
            self.headers = []
            self.search_var.set("")

            self.file_label.configure(text=Path(file_path).name)

            lines = ["Hojas detectadas:"]
            for index, sheet_name in enumerate(self.sheet_names, start=1):
                lines.append(f"{index}. {sheet_name}")

            self._set_textbox(self.sheet_textbox, "\n".join(lines))
            self._set_textbox(self.header_textbox, "Coloca la hoja y la fila header para leer headers.")

            self._log(f"[{self.role}] Hojas consultadas y archivo cerrado: {Path(file_path).name}")
        except Exception as error:
            messagebox.showerror("Error", str(error))
            self._log(f"[{self.role}] Error al cargar archivo: {error}")

    def validate_sheet(self):
        try:
            if not self.file_path:
                raise ValueError("Primero selecciona un archivo Excel")

            sheet_ref = self.sheet_entry.get().strip()
            sample = read_sheet_sample(self.file_path, sheet_ref)

            self.sheet_result_label.configure(
                text=f"Hoja detectada: {sample['sheet_index'] + 1} - {sample['sheet_name']}"
            )

            lines = [
                f"Hoja detectada: {sample['sheet_index'] + 1} - {sample['sheet_name']}",
                f"Filas: {sample['max_row']} | Columnas: {sample['max_column']}",
                "",
                "Muestra de las primeras filas:",
            ]

            for row in sample["rows"]:
                lines.append(" | ".join(row))

            self._set_textbox(self.sheet_textbox, "\n".join(lines))
            self._log(f"[{self.role}] Hoja consultada y archivo cerrado: {sample['sheet_name']}")
        except Exception as error:
            messagebox.showerror("Error", str(error))
            self._log(f"[{self.role}] Error al validar hoja: {error}")

    def load_headers(self):
        try:
            if not self.file_path:
                raise ValueError("Primero selecciona un archivo Excel")

            sheet_ref = self.sheet_entry.get().strip()

            if not self.header_entry.get().strip().isdigit():
                raise ValueError("La fila header debe ser un numero")

            header_row = int(self.header_entry.get().strip())

            self.header_info = read_headers(
                file_path=self.file_path,
                sheet_ref=sheet_ref,
                header_row=header_row,
            )

            self.headers = self.header_info["headers"]

            lines = [
                f"Hoja detectada: {self.header_info['sheet_index'] + 1} - {self.header_info['sheet_name']}",
                f"Fila header: {self.header_info['header_row']}",
                f"Total headers detectados: {len(self.headers)}",
                "",
                "Primeros 30 headers:",
            ]

            for index, header in enumerate(self.header_info.get("first_30_headers", self.headers[:30]), start=1):
                lines.append(f"{index}. {header}")

            self._set_textbox(self.header_textbox, "\n".join(lines))
            self._log(f"[{self.role}] Encabezados consultados y archivo cerrado: {len(self.headers)}")
        except Exception as error:
            messagebox.showerror("Error", str(error))
            self._log(f"[{self.role}] Error al leer headers: {error}")

    def add_column_row(self, alias: str = "", column_ref: str = "", mostrar: bool = True):
        self.row_counter += 1
        row_id = f"{self.role}_{self.row_counter}"

        frame = ctk.CTkFrame(self.rows_container)
        frame.grid(row=self.row_counter, column=0, sticky="ew", padx=0, pady=4)
        frame.grid_columnconfigure(0, weight=2)
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_columnconfigure(2, weight=3)
        frame.grid_columnconfigure(3, weight=1)
        frame.grid_columnconfigure(4, weight=1)
        frame.grid_columnconfigure(5, weight=1)

        alias_entry = ctk.CTkEntry(frame, placeholder_text="Ej: DNI")
        alias_entry.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        alias_entry.insert(0, alias)

        col_entry = ctk.CTkEntry(frame, placeholder_text="A o 1", width=80)
        col_entry.grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        col_entry.insert(0, column_ref)

        detected_label = ctk.CTkLabel(frame, text="-", anchor="w")
        detected_label.grid(row=0, column=2, sticky="ew", padx=6, pady=6)

        mostrar_var = tk.BooleanVar(value=mostrar)
        check = ctk.CTkCheckBox(frame, text="", variable=mostrar_var)
        check.grid(row=0, column=3, sticky="w", padx=6, pady=6)

        radio = ctk.CTkRadioButton(frame, text="", variable=self.search_var, value=row_id)
        radio.grid(row=0, column=4, sticky="w", padx=6, pady=6)

        btn_delete = ctk.CTkButton(
            frame,
            text="X",
            width=40,
            command=lambda rid=row_id: self.delete_column_row(rid),
        )
        btn_delete.grid(row=0, column=5, sticky="ew", padx=6, pady=6)

        self.column_rows[row_id] = {
            "frame": frame,
            "alias_entry": alias_entry,
            "col_entry": col_entry,
            "detected_label": detected_label,
            "mostrar_var": mostrar_var,
        }

        return row_id

    def delete_column_row(self, row_id: str):
        if row_id not in self.column_rows:
            return

        if self.search_var.get() == row_id:
            self.search_var.set("")

        self.column_rows[row_id]["frame"].destroy()
        del self.column_rows[row_id]

    def clear_column_rows(self):
        for row_id in list(self.column_rows.keys()):
            self.delete_column_row(row_id)

        self.row_counter = 0
        self.add_column_row()
        self._log(f"[{self.role}] Columnas limpiadas")

    def use_all_headers(self):
        try:
            if not self.headers:
                raise ValueError("Primero debes leer el header")

            for row_id in list(self.column_rows.keys()):
                self.delete_column_row(row_id)

            self.row_counter = 0

            for index, header in enumerate(self.headers, start=1):
                self.add_column_row(alias=header, column_ref=str(index), mostrar=True)

            self._log(f"[{self.role}] Se agregaron todas las columnas: {len(self.headers)}")
        except Exception as error:
            messagebox.showerror("Error", str(error))
            self._log(f"[{self.role}] Error al usar todas las columnas: {error}")

    def resolve_all_columns(self):
        try:
            if not self.headers:
                raise ValueError("Primero debes leer el header")

            resolved_count = 0
            errors = []

            for row_id, data in self.column_rows.items():
                value = data["col_entry"].get().strip()

                if not value:
                    data["detected_label"].configure(text="-")
                    continue

                try:
                    resolved = resolve_column(self.headers, value)
                    data["detected_label"].configure(
                        text=f"{resolved['letra_excel']} / {resolved['indice_1']} - {resolved['header_detectado']}"
                    )

                    if not data["alias_entry"].get().strip():
                        data["alias_entry"].insert(0, resolved["header_detectado"])

                    resolved_count += 1
                except Exception as error:
                    data["detected_label"].configure(text="ERROR")
                    errors.append(f"Fila {row_id}: {error}")

            if errors:
                self._log(f"[{self.role}] Errores al resolver:\n" + "\n".join(errors))
            else:
                self._log(f"[{self.role}] Columnas resueltas correctamente: {resolved_count}")
        except Exception as error:
            messagebox.showerror("Error", str(error))
            self._log(f"[{self.role}] Error al resolver columnas: {error}")

    def save_current_base_profile(self):
        try:
            if not self.enable_base_profiles:
                raise ValueError("Este panel no soporta perfiles de base")

            profile_name = self.base_profile_name_entry.get().strip()

            if not profile_name:
                raise ValueError("Coloca un nombre para la configuracion base")

            config = self.build_config()
            output_path = save_base_profile(profile_name, config)

            self._log(f"[{self.role}] Configuracion base guardada: {output_path}")
            messagebox.showinfo("Config base guardada", f"Configuracion guardada en:\n{output_path}")

        except Exception as error:
            self._log(f"[{self.role}] ERROR GUARDANDO CONFIG BASE: {error}")
            messagebox.showerror("Error", str(error))

    def load_base_profile_file(self):
        try:
            initial_dir = Path("data") / "perfiles_base"
            initial_dir.mkdir(parents=True, exist_ok=True)

            file_path = filedialog.askopenfilename(
                title="Seleccionar configuracion base",
                initialdir=str(initial_dir),
                filetypes=[
                    ("JSON files", "*.json"),
                    ("All files", "*.*"),
                ],
            )

            if not file_path:
                return

            profile_data = load_base_profile(file_path)
            self.apply_base_profile_config(profile_data)

            self._log(f"[{self.role}] Configuracion base cargada: {Path(file_path).name}")
            messagebox.showinfo("Config base cargada", "Configuracion base cargada correctamente")

        except Exception as error:
            self._log(f"[{self.role}] ERROR CARGANDO CONFIG BASE: {error}")
            messagebox.showerror("Error", str(error))

    def apply_base_profile_config(self, profile_data: dict):
        config = profile_data["config"]

        if hasattr(self, "base_profile_name_entry"):
            self.base_profile_name_entry.delete(0, "end")
            self.base_profile_name_entry.insert(0, profile_data.get("nombre_perfil", ""))

        file_path = Path(config["archivo"]["ruta"])

        if not file_path.exists():
            selected = filedialog.askopenfilename(
                title="No se encontro el archivo original. Selecciona el Excel de esta base",
                filetypes=[
                    ("Excel files", "*.xlsx *.xlsm"),
                    ("All files", "*.*"),
                ],
            )

            if not selected:
                raise ValueError("No se selecciono archivo Excel para cargar la base")

            file_path = Path(selected)

        sheet_names = get_sheet_names(str(file_path))

        self.close_session()
        self.file_path = str(file_path)
        self.sheet_names = sheet_names

        self.header_info = None
        self.headers = []
        self.search_var.set("")

        self.file_label.configure(text=file_path.name)

        sheet_ref = (
            config.get("hoja", {}).get("entrada_usuario")
            or config.get("hoja", {}).get("nombre_detectado")
            or str(config.get("hoja", {}).get("indice_1", "1"))
        )

        self.sheet_entry.delete(0, "end")
        self.sheet_entry.insert(0, str(sheet_ref))

        header_row = int(config["fila_header"])
        self.header_entry.delete(0, "end")
        self.header_entry.insert(0, str(header_row))

        sample = read_sheet_sample(self.file_path, str(sheet_ref))
        self.sheet_result_label.configure(
            text=f"Hoja detectada: {sample['sheet_index'] + 1} - {sample['sheet_name']}"
        )

        sheet_lines = [
            f"Hoja detectada: {sample['sheet_index'] + 1} - {sample['sheet_name']}",
            f"Filas: {sample['max_row']} | Columnas: {sample['max_column']}",
            "",
            "Muestra de las primeras filas:",
        ]

        for row in sample["rows"]:
            sheet_lines.append(" | ".join(row))

        self._set_textbox(self.sheet_textbox, "\n".join(sheet_lines))

        self.header_info = read_headers(
            file_path=self.file_path,
            sheet_ref=str(sheet_ref),
            header_row=header_row,
        )
        self.headers = self.header_info["headers"]

        header_lines = [
            f"Hoja detectada: {self.header_info['sheet_index'] + 1} - {self.header_info['sheet_name']}",
            f"Fila header: {self.header_info['header_row']}",
            f"Total headers detectados: {len(self.headers)}",
            "",
            "Primeros 30 headers:",
        ]

        for index, header in enumerate(self.header_info.get("first_30_headers", self.headers[:30]), start=1):
            header_lines.append(f"{index}. {header}")

        self._set_textbox(self.header_textbox, "\n".join(header_lines))

        for row_id in list(self.column_rows.keys()):
            self.delete_column_row(row_id)

        self.row_counter = 0
        self.search_var.set("")

        search_col = config.get("columna_busqueda", {})
        search_index_0 = search_col.get("indice_0")

        for col in config.get("columnas", []):
            column_ref = (
                col.get("posicion_usuario")
                or col.get("letra_excel")
                or str(col.get("indice_1", ""))
            )

            row_id = self.add_column_row(
                alias=col.get("alias", ""),
                column_ref=str(column_ref),
                mostrar=bool(col.get("mostrar_salida", True)),
            )

            if col.get("es_busqueda") or (
                search_index_0 is not None and col.get("indice_0") == search_index_0
            ):
                self.search_var.set(row_id)

        if not self.column_rows:
            self.add_column_row()

        self.resolve_all_columns()

    def build_config(self) -> dict:
        if not self.file_path:
            raise ValueError(f"[{self.role}] Falta seleccionar archivo")

        if not self.header_info or not self.headers:
            raise ValueError(f"[{self.role}] Falta leer headers")

        selected_search_id = self.search_var.get().strip()
        if not selected_search_id:
            raise ValueError(f"[{self.role}] Debes marcar exactamente 1 columna de busqueda")

        columnas = []
        columna_busqueda = None

        aliases = set()

        for row_id, data in self.column_rows.items():
            column_ref = data["col_entry"].get().strip()

            if not column_ref:
                continue

            resolved = resolve_column(self.headers, column_ref)

            alias = data["alias_entry"].get().strip()
            if not alias:
                alias = resolved["header_detectado"]

            if alias in aliases:
                raise ValueError(f"[{self.role}] Alias duplicado: {alias}")

            aliases.add(alias)

            es_busqueda = row_id == selected_search_id

            item = {
                "row_id": row_id,
                "alias": alias,
                "posicion_usuario": resolved["posicion_usuario"],
                "letra_excel": resolved["letra_excel"],
                "indice_0": resolved["indice_0"],
                "indice_1": resolved["indice_1"],
                "header_detectado": resolved["header_detectado"],
                "mostrar_salida": bool(data["mostrar_var"].get()),
                "es_busqueda": es_busqueda,
            }

            if es_busqueda:
                columna_busqueda = item

            columnas.append(item)

        if not columnas:
            raise ValueError(f"[{self.role}] Debes configurar al menos una columna")

        if columna_busqueda is None:
            raise ValueError(f"[{self.role}] La columna marcada como busqueda no tiene columna asignada")

        return {
            "archivo": {
                "ruta": str(Path(self.file_path)),
                "nombre": Path(self.file_path).name,
            },
            "hoja": {
                "entrada_usuario": self.sheet_entry.get().strip(),
                "indice_0": self.header_info["sheet_index"],
                "indice_1": self.header_info["sheet_index"] + 1,
                "nombre_detectado": self.header_info["sheet_name"],
            },
            "fila_header": self.header_info["header_row"],
            "total_headers_detectados": len(self.headers),
            "columnas": columnas,
            "columna_busqueda": columna_busqueda,
        }
