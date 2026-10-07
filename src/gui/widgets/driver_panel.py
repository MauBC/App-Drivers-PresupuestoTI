from copy import deepcopy
from src.gui.background_tasks import tasks_for
from pathlib import Path
import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox, filedialog

from src.gui.widgets.excel_config_panel import ExcelConfigPanel
from src.core.driver_rules_store import save_ceco_rules_profile, load_ceco_rules_profile
from src.core.driver_executor import (
    DRIVER_CANTIDAD,
    DRIVER_PORCENTAJE,
    run_driver,
    export_driver_result,
    build_driver_debug_report,
    read_excel_with_origin,
    clean_cell_text,
)


def find_config_column_by_alias(config: dict, alias: str) -> dict:
    wanted = str(alias).strip().lower()

    if not wanted:
        raise ValueError("El alias de columna no puede estar vacio")

    for col in config["columnas"]:
        if str(col["alias"]).strip().lower() == wanted:
            return col

    raise ValueError(f"No se encontro una columna configurada con alias: {alias}")


class DriverPanel(ctk.CTkFrame):
    def __init__(self, master, log_callback=None):
        super().__init__(master)

        self.log_callback = log_callback

        self.driver_type_var = tk.StringVar(value=DRIVER_CANTIDAD)
        self.rule_rows = {}
        self.rule_counter = 0

        self.filter_enabled_var = tk.BooleanVar(value=False)
        self.filter_values_vars = {}
        self.filter_groups = []
        self.filter_group_counter = 0

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._build_layout()

    def _log(self, text: str):
        if self.log_callback:
            self.log_callback(text)

    def _build_layout(self):
        title = ctk.CTkLabel(
            self,
            text="Generar Drivers de Presupuesto",
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        title.grid(row=0, column=0, sticky="w", padx=16, pady=(16, 8))

        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        self.tab_resultado = self.tabs.add("Template")
        self.tab_base = self.tabs.add("Base")
        self.tab_reglas = self.tabs.add("Reglas CECO")
        self.tab_ejecutar = self.tabs.add("Ejecutar driver")

        for tab in [self.tab_resultado, self.tab_base, self.tab_reglas, self.tab_ejecutar]:
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)

        self.result_panel = ExcelConfigPanel(
            self.tab_resultado,
            title="Resultado validado / archivo origen",
            role="driver_resultado",
            log_callback=self._log,
        )
        self.result_panel.grid(row=0, column=0, sticky="nsew")

        self.base_panel = ExcelConfigPanel(
            self.tab_base,
            title="Base grande por DNI",
            role="driver_base",
            log_callback=self._log,
        )
        self.base_panel.grid(row=0, column=0, sticky="nsew")

        self._build_rules_tab()
        self._build_execute_tab()

    def _build_rules_tab(self):
        frame = ctk.CTkScrollableFrame(self.tab_reglas)
        frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        frame.grid_columnconfigure(0, weight=1)

        info = ctk.CTkLabel(
            frame,
            text=(
                "Configura los CECO especiales. Si el CECO original empieza con un prefijo, "
                "se reemplaza por CECO2 y se asigna porcentaje 1. En ese caso no se busca por DNI."
            ),
            anchor="w",
            wraplength=1050,
        )
        info.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 10))

        profile_box = ctk.CTkFrame(frame)
        profile_box.grid(row=1, column=0, sticky="ew", padx=16, pady=8)
        profile_box.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            profile_box,
            text="Perfil de reglas CECO",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        self.ceco_rules_profile_entry = ctk.CTkEntry(
            profile_box,
            placeholder_text="Ej: Reglas_CECO_Ransa"
        )
        self.ceco_rules_profile_entry.grid(row=1, column=0, sticky="ew", padx=12, pady=4)

        profile_buttons = ctk.CTkFrame(profile_box)
        profile_buttons.grid(row=2, column=0, sticky="ew", padx=12, pady=(8, 12))
        profile_buttons.grid_columnconfigure((0, 1), weight=1)

        btn_save_profile = ctk.CTkButton(
            profile_buttons,
            text="Guardar reglas CECO",
            command=self.save_current_ceco_rules_profile,
        )
        btn_save_profile.grid(row=0, column=0, sticky="ew", padx=6, pady=6)

        btn_load_profile = ctk.CTkButton(
            profile_buttons,
            text="Cargar reglas CECO",
            command=self.load_ceco_rules_profile_file,
        )
        btn_load_profile.grid(row=0, column=1, sticky="ew", padx=6, pady=6)

        buttons = ctk.CTkFrame(frame)
        buttons.grid(row=2, column=0, sticky="ew", padx=16, pady=8)
        buttons.grid_columnconfigure((0, 1), weight=1)

        btn_add = ctk.CTkButton(buttons, text="Anadir regla", command=self.add_rule_row)
        btn_add.grid(row=0, column=0, sticky="ew", padx=6, pady=6)

        btn_clear = ctk.CTkButton(buttons, text="Limpiar reglas", command=self.clear_rule_rows)
        btn_clear.grid(row=0, column=1, sticky="ew", padx=6, pady=6)

        self.rules_container = ctk.CTkFrame(frame)
        self.rules_container.grid(row=3, column=0, sticky="ew", padx=16, pady=(8, 16))
        self.rules_container.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(self.rules_container)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        header.grid_columnconfigure(0, weight=2)
        header.grid_columnconfigure(1, weight=3)
        header.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(header, text="Empieza con", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, sticky="w", padx=6, pady=6
        )
        ctk.CTkLabel(header, text="Reemplazar por CECO2", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=1, sticky="w", padx=6, pady=6
        )
        ctk.CTkLabel(header, text="Quitar", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=2, sticky="w", padx=6, pady=6
        )

        self.add_rule_row(prefijo="259", ceco2="51AD000FA7")

    def add_rule_row(self, prefijo: str = "", ceco2: str = ""):
        self.rule_counter += 1
        row_id = f"rule_{self.rule_counter}"

        frame = ctk.CTkFrame(self.rules_container)
        frame.grid(row=self.rule_counter, column=0, sticky="ew", pady=4)
        frame.grid_columnconfigure(0, weight=2)
        frame.grid_columnconfigure(1, weight=3)
        frame.grid_columnconfigure(2, weight=1)

        prefijo_entry = ctk.CTkEntry(frame, placeholder_text="Ej: 259")
        prefijo_entry.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        prefijo_entry.insert(0, prefijo)

        ceco2_entry = ctk.CTkEntry(frame, placeholder_text="Ej: 51AD000FA7")
        ceco2_entry.grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        ceco2_entry.insert(0, ceco2)

        btn_delete = ctk.CTkButton(
            frame,
            text="X",
            width=45,
            command=lambda rid=row_id: self.delete_rule_row(rid),
        )
        btn_delete.grid(row=0, column=2, sticky="ew", padx=6, pady=6)

        self.rule_rows[row_id] = {
            "frame": frame,
            "prefijo_entry": prefijo_entry,
            "ceco2_entry": ceco2_entry,
        }

    def delete_rule_row(self, row_id: str):
        if row_id not in self.rule_rows:
            return

        self.rule_rows[row_id]["frame"].destroy()
        del self.rule_rows[row_id]

    def clear_rule_rows(self):
        for row_id in list(self.rule_rows.keys()):
            self.delete_rule_row(row_id)

        self.rule_counter = 0
        self._log("[drivers] Reglas de CECO limpiadas")

    def _build_execute_tab(self):
        frame = ctk.CTkScrollableFrame(self.tab_ejecutar)
        frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        frame.grid_columnconfigure(0, weight=1)

        tipo_box = ctk.CTkFrame(frame)
        tipo_box.grid(row=0, column=0, sticky="ew", padx=16, pady=8)
        tipo_box.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            tipo_box,
            text="1. Tipo de driver",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        ctk.CTkRadioButton(
            tipo_box,
            text="Driver por cantidad: requiere DNI + CECO + PRECIO",
            variable=self.driver_type_var,
            value=DRIVER_CANTIDAD,
        ).grid(row=1, column=0, sticky="w", padx=12, pady=6)

        ctk.CTkRadioButton(
            tipo_box,
            text="Driver por porcentaje: requiere DNI + CECO",
            variable=self.driver_type_var,
            value=DRIVER_PORCENTAJE,
        ).grid(row=2, column=0, sticky="w", padx=12, pady=(6, 12))

        map_box = ctk.CTkFrame(frame)
        map_box.grid(row=1, column=0, sticky="ew", padx=16, pady=8)
        map_box.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkLabel(
            map_box,
            text="2. Alias de columnas configuradas",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 6))

        ctk.CTkLabel(
            map_box,
            text=(
                "En las pestañas Resultado validado y Base por DNI, configura las columnas y usa estos alias. "
                "Luego este panel los ubicara automaticamente."
            ),
            anchor="w",
            wraplength=1050,
        ).grid(row=1, column=0, columnspan=3, sticky="ew", padx=12, pady=(0, 10))

        ctk.CTkLabel(map_box, text="Resultado: alias DNI").grid(row=2, column=0, sticky="w", padx=12, pady=2)
        self.result_dni_alias_entry = ctk.CTkEntry(map_box)
        self.result_dni_alias_entry.grid(row=3, column=0, sticky="ew", padx=12, pady=4)
        self.result_dni_alias_entry.insert(0, "DNI")

        ctk.CTkLabel(map_box, text="Resultado: alias CECO").grid(row=2, column=1, sticky="w", padx=12, pady=2)
        self.result_ceco_alias_entry = ctk.CTkEntry(map_box)
        self.result_ceco_alias_entry.grid(row=3, column=1, sticky="ew", padx=12, pady=4)
        self.result_ceco_alias_entry.insert(0, "CECO")

        ctk.CTkLabel(map_box, text="Resultado: alias PRECIO").grid(row=2, column=2, sticky="w", padx=12, pady=2)
        self.result_price_alias_entry = ctk.CTkEntry(map_box)
        self.result_price_alias_entry.grid(row=3, column=2, sticky="ew", padx=12, pady=4)
        self.result_price_alias_entry.insert(0, "PRECIO")

        ctk.CTkLabel(map_box, text="Base: alias DNI").grid(row=4, column=0, sticky="w", padx=12, pady=(12, 2))
        self.base_dni_alias_entry = ctk.CTkEntry(map_box)
        self.base_dni_alias_entry.grid(row=5, column=0, sticky="ew", padx=12, pady=4)
        self.base_dni_alias_entry.insert(0, "DNI")

        ctk.CTkLabel(map_box, text="Base: alias CECO").grid(row=4, column=1, sticky="w", padx=12, pady=(12, 2))
        self.base_ceco_alias_entry = ctk.CTkEntry(map_box)
        self.base_ceco_alias_entry.grid(row=5, column=1, sticky="ew", padx=12, pady=4)
        self.base_ceco_alias_entry.insert(0, "CECO")

        ctk.CTkLabel(map_box, text="Base: alias PORCENTAJE").grid(row=4, column=2, sticky="w", padx=12, pady=(12, 2))
        self.base_percentage_alias_entry = ctk.CTkEntry(map_box)
        self.base_percentage_alias_entry.grid(row=5, column=2, sticky="ew", padx=12, pady=4)
        self.base_percentage_alias_entry.insert(0, "PORCENTAJE")

        self._build_filter_box(frame, row=2)

        buttons_box = ctk.CTkFrame(frame)
        buttons_box.grid(row=3, column=0, sticky="ew", padx=16, pady=12)
        buttons_box.grid_columnconfigure((0, 1), weight=1)

        btn_validate = ctk.CTkButton(
            buttons_box,
            text="Validar driver",
            command=self.validate_config,
        )
        btn_validate.grid(row=0, column=0, sticky="ew", padx=6, pady=10)

        self.btn_execute = ctk.CTkButton(
            buttons_box,
            text="Generar driver",
            command=self.execute_driver,
        )
        self.btn_execute.grid(row=0, column=1, sticky="ew", padx=6, pady=10)

    def _build_filter_box(self, frame, row: int):
        filter_box = ctk.CTkFrame(frame)
        filter_box.grid(row=row, column=0, sticky="ew", padx=16, pady=8)
        filter_box.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            filter_box,
            text="3. Filtro opcional por arrendamiento / categoria",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        ctk.CTkCheckBox(
            filter_box,
            text="Usar filtro para generar drivers separados por grupo",
            variable=self.filter_enabled_var,
        ).grid(row=1, column=0, sticky="w", padx=12, pady=6)

        ctk.CTkLabel(
            filter_box,
            text=(
                "Ejemplo: usar la columna ARRENDAMIENTO y crear grupos como ARR_1, ARR_2 o ARR_3_4. "
                "El driver final agregara GRUPO_FILTRO y VALOR_FILTRO."
            ),
            anchor="w",
            wraplength=1050,
        ).grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 10))

        top = ctk.CTkFrame(filter_box)
        top.grid(row=3, column=0, sticky="ew", padx=12, pady=6)
        top.grid_columnconfigure(0, weight=1)
        top.grid_columnconfigure(1, weight=1)
        top.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(top, text="Resultado: alias columna filtro").grid(
            row=0, column=0, sticky="w", padx=6, pady=2
        )
        self.filter_alias_entry = ctk.CTkEntry(top)
        self.filter_alias_entry.grid(row=1, column=0, sticky="ew", padx=6, pady=4)
        self.filter_alias_entry.insert(0, "FILTRO")

        ctk.CTkLabel(top, text="Nombre del grupo a crear").grid(
            row=0, column=1, sticky="w", padx=6, pady=2
        )
        self.filter_group_name_entry = ctk.CTkEntry(top, placeholder_text="Ej: ARR_3_4")
        self.filter_group_name_entry.grid(row=1, column=1, sticky="ew", padx=6, pady=4)

        btn_load_values = ctk.CTkButton(
            top,
            text="Cargar valores unicos",
            command=self.load_filter_values_from_result,
        )
        btn_load_values.grid(row=1, column=2, sticky="ew", padx=6, pady=4)

        self.filter_values_container = ctk.CTkFrame(filter_box)
        self.filter_values_container.grid(row=4, column=0, sticky="ew", padx=12, pady=8)
        self.filter_values_container.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkLabel(
            self.filter_values_container,
            text="Carga valores unicos para seleccionarlos aqui.",
            anchor="w",
        ).grid(row=0, column=0, columnspan=3, sticky="ew", padx=8, pady=8)

        btn_add_group = ctk.CTkButton(
            filter_box,
            text="Agregar grupo con valores seleccionados",
            command=self.add_filter_group,
        )
        btn_add_group.grid(row=5, column=0, sticky="ew", padx=12, pady=6)

        self.filter_groups_container = ctk.CTkFrame(filter_box)
        self.filter_groups_container.grid(row=6, column=0, sticky="ew", padx=12, pady=(8, 12))
        self.filter_groups_container.grid_columnconfigure(0, weight=1)

        self.render_filter_groups()

    def load_filter_values_from_result(self):
        try:
            result_config = self.result_panel.build_config()

            filter_col = find_config_column_by_alias(
                result_config,
                self.filter_alias_entry.get(),
            )

            self.result_panel.close_session()

            df = read_excel_with_origin(
                file_path=result_config["archivo"]["ruta"],
                sheet_name=result_config["hoja"]["nombre_detectado"],
                header_row=int(result_config["fila_header"]),
            )

            header = filter_col["header_detectado"]

            if header not in df.columns:
                raise ValueError(f"No se encontro la columna filtro en el resultado: {header}")

            values = sorted(
                {
                    clean_cell_text(value)
                    for value in df[header].tolist()
                    if clean_cell_text(value)
                },
                key=lambda value: value.lower(),
            )

            self.filter_values_vars.clear()

            for child in self.filter_values_container.winfo_children():
                child.destroy()

            if not values:
                ctk.CTkLabel(
                    self.filter_values_container,
                    text="No se encontraron valores en la columna filtro.",
                    anchor="w",
                ).grid(row=0, column=0, sticky="ew", padx=8, pady=8)
                return

            ctk.CTkLabel(
                self.filter_values_container,
                text=f"Valores encontrados: {len(values)}",
                font=ctk.CTkFont(weight="bold"),
            ).grid(row=0, column=0, columnspan=3, sticky="w", padx=8, pady=(8, 4))

            for index, value in enumerate(values, start=1):
                var = tk.BooleanVar(value=False)
                self.filter_values_vars[value] = var

                checkbox = ctk.CTkCheckBox(
                    self.filter_values_container,
                    text=value,
                    variable=var,
                )
                checkbox.grid(
                    row=((index - 1) // 3) + 1,
                    column=(index - 1) % 3,
                    sticky="w",
                    padx=8,
                    pady=4,
                )

            self._log(f"[drivers] Valores filtro cargados: {len(values)}")

        except Exception as error:
            self._log(f"[drivers] ERROR CARGANDO VALORES FILTRO: {error}")
            messagebox.showerror("Error", str(error))

    def add_filter_group(self):
        try:
            selected_values = [
                value
                for value, var in self.filter_values_vars.items()
                if var.get()
            ]

            if not selected_values:
                raise ValueError("Selecciona al menos un valor para el grupo")

            group_name = self.filter_group_name_entry.get().strip()

            if not group_name:
                self.filter_group_counter += 1
                group_name = f"GRUPO_{self.filter_group_counter}"

            self.filter_groups.append({
                "name": group_name,
                "values": selected_values,
            })

            self.filter_group_name_entry.delete(0, "end")

            for var in self.filter_values_vars.values():
                var.set(False)

            self.render_filter_groups()
            self._log(f"[drivers] Grupo filtro agregado: {group_name}")

        except Exception as error:
            self._log(f"[drivers] ERROR AGREGANDO GRUPO FILTRO: {error}")
            messagebox.showerror("Error", str(error))

    def delete_filter_group(self, index: int):
        if index < 0 or index >= len(self.filter_groups):
            return

        group = self.filter_groups.pop(index)
        self.render_filter_groups()
        self._log(f"[drivers] Grupo filtro eliminado: {group.get('name')}")

    def render_filter_groups(self):
        if not hasattr(self, "filter_groups_container"):
            return

        for child in self.filter_groups_container.winfo_children():
            child.destroy()

        ctk.CTkLabel(
            self.filter_groups_container,
            text="Grupos configurados",
            font=ctk.CTkFont(weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=8, pady=(8, 4))

        if not self.filter_groups:
            ctk.CTkLabel(
                self.filter_groups_container,
                text="Aun no hay grupos. Puedes generar driver sin filtro o agregar grupos.",
                anchor="w",
            ).grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))
            return

        for index, group in enumerate(self.filter_groups):
            row = ctk.CTkFrame(self.filter_groups_container)
            row.grid(row=index + 1, column=0, sticky="ew", padx=6, pady=4)
            row.grid_columnconfigure(0, weight=1)

            values_text = ", ".join(group.get("values", []))

            ctk.CTkLabel(
                row,
                text=f"{group.get('name')} -> {values_text}",
                anchor="w",
                wraplength=900,
            ).grid(row=0, column=0, sticky="ew", padx=8, pady=6)

            ctk.CTkButton(
                row,
                text="X",
                width=45,
                command=lambda idx=index: self.delete_filter_group(idx),
            ).grid(row=0, column=1, sticky="e", padx=8, pady=6)

    def build_filter_config(self, result_config: dict) -> dict:
        if not self.filter_enabled_var.get():
            return {
                "enabled": False,
                "column": "",
                "groups": [],
            }

        filter_col = find_config_column_by_alias(
            result_config,
            self.filter_alias_entry.get(),
        )

        if not self.filter_groups:
            raise ValueError("Activaste filtro, pero no configuraste ningun grupo")

        return {
            "enabled": True,
            "column": filter_col["header_detectado"],
            "groups": self.filter_groups,
        }


    def save_current_ceco_rules_profile(self):
        try:
            profile_name = self.ceco_rules_profile_entry.get().strip()

            if not profile_name:
                raise ValueError("Coloca un nombre para el perfil de reglas CECO")

            rules = self.build_special_rules()

            if not rules:
                raise ValueError("No hay reglas CECO para guardar")

            output_path = save_ceco_rules_profile(profile_name, rules)

            self._log(f"[drivers] Reglas CECO guardadas: {output_path}")
            messagebox.showinfo("Reglas CECO guardadas", f"Reglas guardadas en:\n{output_path}")

        except Exception as error:
            self._log(f"[drivers] ERROR GUARDANDO REGLAS CECO: {error}")
            messagebox.showerror("Error", str(error))

    def load_ceco_rules_profile_file(self):
        try:
            initial_dir = Path("data") / "perfiles_drivers" / "reglas_ceco"
            initial_dir.mkdir(parents=True, exist_ok=True)

            file_path = filedialog.askopenfilename(
                title="Seleccionar perfil de reglas CECO",
                initialdir=str(initial_dir),
                filetypes=[
                    ("JSON files", "*.json"),
                    ("All files", "*.*"),
                ],
            )

            if not file_path:
                return

            profile_data = load_ceco_rules_profile(file_path)
            self.apply_ceco_rules_profile(profile_data)

            self._log(f"[drivers] Reglas CECO cargadas: {Path(file_path).name}")
            messagebox.showinfo("Reglas CECO cargadas", "Reglas CECO cargadas correctamente")

        except Exception as error:
            self._log(f"[drivers] ERROR CARGANDO REGLAS CECO: {error}")
            messagebox.showerror("Error", str(error))

    def apply_ceco_rules_profile(self, profile_data: dict):
        if hasattr(self, "ceco_rules_profile_entry"):
            self.ceco_rules_profile_entry.delete(0, "end")
            self.ceco_rules_profile_entry.insert(0, profile_data.get("nombre_perfil", ""))

        self.clear_rule_rows()

        for rule in profile_data.get("rules", []):
            self.add_rule_row(
                prefijo=rule.get("prefijo", ""),
                ceco2=rule.get("ceco2", ""),
            )

        if not self.rule_rows:
            self.add_rule_row()

    def build_special_rules(self) -> list[dict]:
        rules = []

        for _, data in self.rule_rows.items():
            prefijo = data["prefijo_entry"].get().strip()
            ceco2 = data["ceco2_entry"].get().strip()

            if not prefijo and not ceco2:
                continue

            if not prefijo or not ceco2:
                raise ValueError("Todas las reglas CECO deben tener prefijo y CECO2")

            rules.append({
                "prefijo": prefijo,
                "ceco2": ceco2,
            })

        return rules

    def build_driver_config(self) -> dict:
        driver_type = self.driver_type_var.get()

        result_config = self.result_panel.build_config()
        base_config = self.base_panel.build_config()

        result_dni_col = find_config_column_by_alias(
            result_config,
            self.result_dni_alias_entry.get(),
        )
        result_ceco_col = find_config_column_by_alias(
            result_config,
            self.result_ceco_alias_entry.get(),
        )

        base_dni_col = find_config_column_by_alias(
            base_config,
            self.base_dni_alias_entry.get(),
        )
        base_ceco_col = find_config_column_by_alias(
            base_config,
            self.base_ceco_alias_entry.get(),
        )
        base_percentage_col = find_config_column_by_alias(
            base_config,
            self.base_percentage_alias_entry.get(),
        )

        filter_config = self.build_filter_config(result_config)

        result_data = {
            "expected_headers": list(self.result_panel.headers),
            "file_path": result_config["archivo"]["ruta"],
            "sheet_name": result_config["hoja"]["nombre_detectado"],
            "header_row": result_config["fila_header"],
            "dni_col": result_dni_col["header_detectado"],
            "ceco_col": result_ceco_col["header_detectado"],
        }

        if driver_type == DRIVER_CANTIDAD:
            result_price_col = find_config_column_by_alias(
                result_config,
                self.result_price_alias_entry.get(),
            )
            result_data["price_col"] = result_price_col["header_detectado"]

        if filter_config.get("enabled"):
            result_data["filter_col"] = filter_config["column"]

        used_result_headers = {
            result_data.get("dni_col"),
            result_data.get("ceco_col"),
            result_data.get("price_col"),
            result_data.get("filter_col"),
            "VALIDADO",
        }
        used_result_headers = {str(value).strip() for value in used_result_headers if str(value).strip()}

        extra_cols = []
        for col in result_config["columnas"]:
            header = str(col.get("header_detectado", "")).strip()
            alias = str(col.get("alias", "")).strip() or header

            if not header:
                continue

            if header in used_result_headers:
                continue

            extra_cols.append({
                "header": header,
                "alias": alias,
            })

        result_data["extra_cols"] = extra_cols

        return {
            "driver_type": driver_type,
            "resultado": result_data,
            "base": {
                "expected_headers": list(self.base_panel.headers),
                "file_path": base_config["archivo"]["ruta"],
                "sheet_name": base_config["hoja"]["nombre_detectado"],
                "header_row": base_config["fila_header"],
                "dni_col": base_dni_col["header_detectado"],
                "ceco_col": base_ceco_col["header_detectado"],
                "percentage_col": base_percentage_col["header_detectado"],
            },
            "special_ceco_rules": self.build_special_rules(),
            "filter": {
                "enabled": filter_config.get("enabled", False),
                "groups": filter_config.get("groups", []),
            },
            "output": {
                "dir": str(Path(result_data["file_path"]).parent),
                "name": Path(result_data["file_path"]).stem + "_driver",
            },
        }

    def validate_config(self):
        try:
            config = self.build_driver_config()

            lines = [
                "CONFIGURACION DE DRIVER VALIDA",
                "",
                f"Tipo driver: {config['driver_type']}",
                "",
                "Resultado validado:",
                f"- Archivo: {Path(config['resultado']['file_path']).name}",
                f"- Hoja: {config['resultado']['sheet_name']}",
                f"- Header: {config['resultado']['header_row']}",
                f"- DNI: {config['resultado']['dni_col']}",
                f"- CECO: {config['resultado']['ceco_col']}",
            ]

            if config["driver_type"] == DRIVER_CANTIDAD:
                lines.append(f"- PRECIO: {config['resultado']['price_col']}")

            extra_cols = config["resultado"].get("extra_cols", [])
            if extra_cols:
                lines.append(f"- Columnas extra al DRIVER: {', '.join(col.get('alias', col.get('header', '')) for col in extra_cols)}")
            else:
                lines.append("- Columnas extra al DRIVER: ninguna")

            filter_cfg = config.get("filter", {})
            if filter_cfg.get("enabled"):
                lines.extend([
                    f"- FILTRO: {config['resultado']['filter_col']}",
                    "",
                    "Grupos filtro:",
                ])

                for group in filter_cfg.get("groups", []):
                    lines.append(f"- {group.get('name')}: {', '.join(group.get('values', []))}")
            else:
                lines.append("- FILTRO: no usado")

            lines.extend([
                "",
                "Base por DNI:",
                f"- Archivo: {Path(config['base']['file_path']).name}",
                f"- Hoja: {config['base']['sheet_name']}",
                f"- Header: {config['base']['header_row']}",
                f"- DNI: {config['base']['dni_col']}",
                f"- CECO: {config['base']['ceco_col']}",
                f"- PORCENTAJE: {config['base']['percentage_col']}",
                "",
                f"Reglas CECO especiales: {len(config['special_ceco_rules'])}",
                "Salida: mismo Excel origen",
            ])

            self._log("\n".join(lines))
            messagebox.showinfo("OK", "La configuracion del driver es valida")

        except Exception as error:
            self._log(f"[drivers] ERROR VALIDACION: {error}")
            messagebox.showerror("Error", str(error))

    def execute_driver(self):
        def failed(error):
            self._log(f"[drivers] ERROR EJECUCION: {error}")
            messagebox.showerror("Error", str(error))

        try:
            tasks = tasks_for(self)
            if tasks.busy:
                raise RuntimeError("Ya hay un proceso en curso. Espere a que termine antes de ejecutar otro.")
            config = deepcopy(self.build_driver_config())
            self._log("[drivers] Iniciando generacion de driver...")
        except Exception as error:
            failed(error)
            return

        def work(report):
            df_driver, df_observados, summary = run_driver(config, progress=report)
            report("Guardando drivers...")
            output_path = export_driver_result(
                df_driver=df_driver, df_observados=df_observados, summary=summary,
                output_dir=config["output"]["dir"], output_name=config["output"]["name"],
                source_file_path=config["resultado"]["file_path"],
            )
            return output_path, summary, df_driver.attrs["excel_output_sheets"]

        def finished(result):
            output_path, summary, output_sheets = result
            self._log("[drivers] Driver generado correctamente")
            self._log(f"[drivers] Filas driver: {summary['filas_driver']}")
            self._log(f"[drivers] Filtro usado: {summary.get('filtro_usado', 'NO')}")
            if summary.get("filtro_usado") == "SI":
                self._log(f"[drivers] Grupos filtro: {summary.get('grupos_filtro', 0)}")
                self._log(f"[drivers] Filas excluidas por filtro: {summary.get('filas_excluidas_por_filtro', 0)}")
            self._log(f"[drivers] Observados: {summary['filas_observadas']}")
            self._log(f"[drivers] CECO especiales: {summary['cecos_especiales']}")
            self._log(f"[drivers] DNI vacios: {summary['dni_vacios']}")
            self._log(f"[drivers] DNI no encontrados: {summary['dni_no_encontrados']}")
            self._log(f"[drivers] Filas expandidas por DNI: {summary['filas_expandidas_por_dni']}")
            self._log(f"[drivers] Excel origen actualizado: {output_path}")
            self._log(f"[drivers] Hojas generadas: {', '.join(output_sheets)}")

            messagebox.showinfo("Driver generado", f"Guardado en este equipo:\n{output_path}\nHojas: {', '.join(output_sheets)}")

        tasks.start(
            work, self.btn_execute, "Generar driver", finished, failed,
            lambda text: self._log(text),
        )

    def close_resources(self):
        self.result_panel.close_session()
        self.base_panel.close_session()
