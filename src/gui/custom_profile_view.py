from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox

from src.core.excel_inspector import (
    get_sheet_names,
    parse_sheet_ref,
    read_headers,
    resolve_column,
)
from src.core.profile_store import save_profile


class CustomProfileView(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)

        self.file_path = None
        self.sheet_names = []
        self.header_info = None
        self.headers = []

        self.column_entries = {}
        self.column_result_labels = {}
        self.output_vars = {}

        self._build_layout()

    def _build_layout(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(1, weight=1)

        title = ctk.CTkLabel(
            self,
            text="Perfil Excel personalizado",
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        title.grid(row=0, column=0, columnspan=2, sticky="w", padx=20, pady=(20, 10))

        left_panel = ctk.CTkFrame(self)
        left_panel.grid(row=1, column=0, sticky="nsew", padx=(20, 10), pady=10)
        left_panel.grid_columnconfigure(0, weight=1)

        right_panel = ctk.CTkFrame(self)
        right_panel.grid(row=1, column=1, sticky="nsew", padx=(10, 20), pady=10)
        right_panel.grid_columnconfigure(0, weight=1)
        right_panel.grid_rowconfigure(1, weight=1)

        self._build_file_section(left_panel)
        self._build_header_section(left_panel)
        self._build_save_section(left_panel)

        self._build_mapping_section(right_panel)
        self._build_log_section(right_panel)

    def _build_file_section(self, parent):
        section = ctk.CTkFrame(parent)
        section.grid(row=0, column=0, sticky="ew", padx=12, pady=12)
        section.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(section, text="1. Archivo y hoja", font=ctk.CTkFont(size=16, weight="bold"))
        label.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        self.file_label = ctk.CTkLabel(section, text="Ningun archivo seleccionado", anchor="w")
        self.file_label.grid(row=1, column=0, sticky="ew", padx=12, pady=4)

        file_button = ctk.CTkButton(section, text="Seleccionar Excel", command=self.select_file)
        file_button.grid(row=2, column=0, sticky="ew", padx=12, pady=6)

        sheet_label = ctk.CTkLabel(section, text="Hoja: posicion 1, 2, 3... o nombre")
        sheet_label.grid(row=3, column=0, sticky="w", padx=12, pady=(10, 2))

        self.sheet_entry = ctk.CTkEntry(section, placeholder_text="Ej: 1 o Personal")
        self.sheet_entry.grid(row=4, column=0, sticky="ew", padx=12, pady=4)

        validate_button = ctk.CTkButton(section, text="Validar hoja", command=self.validate_sheet)
        validate_button.grid(row=5, column=0, sticky="ew", padx=12, pady=6)

        self.sheet_result_label = ctk.CTkLabel(section, text="Hoja detectada: -", anchor="w")
        self.sheet_result_label.grid(row=6, column=0, sticky="ew", padx=12, pady=(4, 12))

        self.sheets_textbox = ctk.CTkTextbox(section, height=120)
        self.sheets_textbox.grid(row=7, column=0, sticky="ew", padx=12, pady=(4, 12))
        self.sheets_textbox.insert("1.0", "Aqui apareceran las hojas del archivo.")
        self.sheets_textbox.configure(state="disabled")

    def _build_header_section(self, parent):
        section = ctk.CTkFrame(parent)
        section.grid(row=1, column=0, sticky="ew", padx=12, pady=12)
        section.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(section, text="2. Header", font=ctk.CTkFont(size=16, weight="bold"))
        label.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        header_label = ctk.CTkLabel(section, text="Fila donde esta el header")
        header_label.grid(row=1, column=0, sticky="w", padx=12, pady=(6, 2))

        self.header_entry = ctk.CTkEntry(section, placeholder_text="Ej: 6")
        self.header_entry.grid(row=2, column=0, sticky="ew", padx=12, pady=4)

        header_button = ctk.CTkButton(section, text="Leer header", command=self.load_headers)
        header_button.grid(row=3, column=0, sticky="ew", padx=12, pady=6)

        self.header_textbox = ctk.CTkTextbox(section, height=120)
        self.header_textbox.grid(row=4, column=0, sticky="ew", padx=12, pady=(4, 12))
        self.header_textbox.insert("1.0", "Aqui apareceran los primeros 5 headers.")
        self.header_textbox.configure(state="disabled")

    def _build_save_section(self, parent):
        section = ctk.CTkFrame(parent)
        section.grid(row=2, column=0, sticky="ew", padx=12, pady=12)
        section.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(section, text="4. Guardar perfil", font=ctk.CTkFont(size=16, weight="bold"))
        label.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        self.profile_name_entry = ctk.CTkEntry(section, placeholder_text="Ej: Maestro_Personal_Ransa")
        self.profile_name_entry.grid(row=1, column=0, sticky="ew", padx=12, pady=4)

        save_button = ctk.CTkButton(section, text="Guardar perfil personalizado", command=self.save_current_profile)
        save_button.grid(row=2, column=0, sticky="ew", padx=12, pady=(8, 12))

    def _build_mapping_section(self, parent):
        title = ctk.CTkLabel(
            parent,
            text="3. Mapeo de columnas",
            font=ctk.CTkFont(size=18, weight="bold"),
        )
        title.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        mapping_frame = ctk.CTkScrollableFrame(parent)
        mapping_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=8)
        mapping_frame.grid_columnconfigure(1, weight=1)
        mapping_frame.grid_columnconfigure(2, weight=2)

        headers = ["Campo", "Columna", "Header detectado", "Salida"]
        for col, text in enumerate(headers):
            label = ctk.CTkLabel(mapping_frame, text=text, font=ctk.CTkFont(weight="bold"))
            label.grid(row=0, column=col, sticky="w", padx=8, pady=6)

        fields = [
            ("nombre", "Nombre"),
            ("dni", "DNI"),
            ("ceco", "CECO"),
            ("porcentaje", "Porcentaje"),
            ("estado", "Estado"),
            ("correo", "Correo"),
            ("pais", "Pais"),
            ("observacion", "Observacion"),
        ]

        default_output = {
            "nombre": True,
            "dni": True,
            "ceco": True,
            "porcentaje": True,
            "estado": True,
            "correo": False,
            "pais": False,
            "observacion": False,
        }

        for row_index, (field_key, field_label) in enumerate(fields, start=1):
            label = ctk.CTkLabel(mapping_frame, text=field_label)
            label.grid(row=row_index, column=0, sticky="w", padx=8, pady=6)

            entry = ctk.CTkEntry(mapping_frame, placeholder_text="A o 1")
            entry.grid(row=row_index, column=1, sticky="ew", padx=8, pady=6)

            result_label = ctk.CTkLabel(mapping_frame, text="-", anchor="w")
            result_label.grid(row=row_index, column=2, sticky="ew", padx=8, pady=6)

            output_var = ctk.BooleanVar(value=default_output.get(field_key, False))
            checkbox = ctk.CTkCheckBox(mapping_frame, text="Mostrar", variable=output_var)
            checkbox.grid(row=row_index, column=3, sticky="w", padx=8, pady=6)

            self.column_entries[field_key] = entry
            self.column_result_labels[field_key] = result_label
            self.output_vars[field_key] = output_var

        resolve_button = ctk.CTkButton(parent, text="Resolver columnas", command=self.resolve_columns)
        resolve_button.grid(row=2, column=0, sticky="ew", padx=12, pady=(4, 12))

    def _build_log_section(self, parent):
        self.log_textbox = ctk.CTkTextbox(parent, height=130)
        self.log_textbox.grid(row=3, column=0, sticky="ew", padx=12, pady=(4, 12))
        self.log_textbox.insert("1.0", "Listo para configurar un perfil.")
        self.log_textbox.configure(state="disabled")

    def set_textbox_text(self, textbox, text):
        textbox.configure(state="normal")
        textbox.delete("1.0", "end")
        textbox.insert("1.0", text)
        textbox.configure(state="disabled")

    def log(self, text):
        self.set_textbox_text(self.log_textbox, text)

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
            self.file_path = file_path
            self.sheet_names = get_sheet_names(file_path)

            self.file_label.configure(text=Path(file_path).name)

            lines = []
            for index, sheet_name in enumerate(self.sheet_names, start=1):
                lines.append(f"{index}. {sheet_name}")

            self.set_textbox_text(self.sheets_textbox, "\n".join(lines))
            self.log("Archivo cargado correctamente. Ahora coloca la hoja y la fila header.")
        except Exception as error:
            messagebox.showerror("Error", str(error))
            self.log(f"Error al leer archivo: {error}")

    def validate_sheet(self):
        try:
            if not self.file_path:
                raise ValueError("Primero selecciona un archivo Excel")

            if not self.sheet_names:
                self.sheet_names = get_sheet_names(self.file_path)

            sheet_ref = self.sheet_entry.get()
            sheet_index, sheet_name = parse_sheet_ref(sheet_ref, self.sheet_names)

            self.sheet_result_label.configure(
                text=f"Hoja detectada: {sheet_index + 1} - {sheet_name}"
            )
            self.log("Hoja validada correctamente.")
        except Exception as error:
            messagebox.showerror("Error", str(error))
            self.log(f"Error al validar hoja: {error}")

    def load_headers(self):
        try:
            if not self.file_path:
                raise ValueError("Primero selecciona un archivo Excel")

            sheet_ref = self.sheet_entry.get()

            if not self.header_entry.get().strip().isdigit():
                raise ValueError("La fila header debe ser un numero")

            header_row = int(self.header_entry.get().strip())

            self.header_info = read_headers(
                file_path=self.file_path,
                sheet_ref=sheet_ref,
                header_row=header_row,
            )

            self.headers = self.header_info["headers"]

            first_headers = self.header_info["first_5_headers"]
            lines = [
                f"Hoja detectada: {self.header_info['sheet_index'] + 1} - {self.header_info['sheet_name']}",
                f"Fila header: {self.header_info['header_row']}",
                "",
                "Primeros 5 headers detectados:",
            ]

            for index, header in enumerate(first_headers, start=1):
                lines.append(f"{index}. {header}")

            self.set_textbox_text(self.header_textbox, "\n".join(lines))
            self.log("Headers cargados correctamente. Ahora mapea las columnas por letra o numero.")
        except Exception as error:
            messagebox.showerror("Error", str(error))
            self.log(f"Error al leer headers: {error}")

    def resolve_columns(self):
        try:
            if not self.headers:
                raise ValueError("Primero debes leer el header")

            resolved_count = 0
            errors = []

            for field_key, entry in self.column_entries.items():
                value = entry.get().strip()

                if not value:
                    self.column_result_labels[field_key].configure(text="-")
                    continue

                try:
                    resolved = resolve_column(self.headers, value)
                    text = f"{resolved['indice_1']} - {resolved['header_detectado']}"
                    self.column_result_labels[field_key].configure(text=text)
                    resolved_count += 1
                except Exception as error:
                    self.column_result_labels[field_key].configure(text="ERROR")
                    errors.append(f"{field_key}: {error}")

            if errors:
                self.log("Algunas columnas tienen error:\n" + "\n".join(errors))
            else:
                self.log(f"Columnas resueltas correctamente: {resolved_count}")
        except Exception as error:
            messagebox.showerror("Error", str(error))
            self.log(f"Error al resolver columnas: {error}")

    def save_current_profile(self):
        try:
            if not self.file_path:
                raise ValueError("Primero selecciona un archivo Excel")

            if not self.header_info:
                raise ValueError("Primero debes leer el header")

            profile_name = self.profile_name_entry.get().strip()
            if not profile_name:
                raise ValueError("Coloca un nombre para el perfil")

            columnas = {}

            for field_key, entry in self.column_entries.items():
                value = entry.get().strip()

                if not value:
                    continue

                resolved = resolve_column(self.headers, value)
                resolved["mostrar_salida"] = bool(self.output_vars[field_key].get())
                columnas[field_key] = resolved

            if not columnas:
                raise ValueError("Debes mapear al menos una columna")

            profile_data = {
                "tipo": "CUSTOM_EXCEL_PROFILE",
                "version": 1,
                "nombre_perfil": profile_name,
                "archivo_ejemplo": Path(self.file_path).name,
                "hoja": {
                    "entrada_usuario": self.sheet_entry.get().strip(),
                    "indice_0": self.header_info["sheet_index"],
                    "indice_1": self.header_info["sheet_index"] + 1,
                    "nombre_detectado": self.header_info["sheet_name"],
                },
                "fila_header": self.header_info["header_row"],
                "columnas": columnas,
            }

            output_path = save_profile(profile_name, profile_data)

            messagebox.showinfo("Perfil guardado", f"Perfil guardado en:\n{output_path}")
            self.log(f"Perfil guardado correctamente:\n{output_path}")
        except Exception as error:
            messagebox.showerror("Error", str(error))
            self.log(f"Error al guardar perfil: {error}")
