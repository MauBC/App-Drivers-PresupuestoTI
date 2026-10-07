import customtkinter as ctk

from src.gui.search_builder_view import SearchBuilderView
from src.core.excel_safety import configure_logging
from src.core import excel_recovery
from src.gui.recovery_panel import RecoveryPanel


class PresupuestoApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Automatizacion Presupuesto")
        self.geometry("1320x860")
        self.minsize(1150, 740)

        # Fondo verde oscuro tipo Ransa / olivo
        self.configure(fg_color="#161C12")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.view = SearchBuilderView(self)
        self.view.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        self.pending_button = ctk.CTkButton(self, text="Resultados pendientes", command=self.open_pending)
        self.pending_button.grid(row=1, column=0, sticky="e", padx=12, pady=6)
        self.refresh_pending()

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def refresh_pending(self):
        self.pending_button.configure(text=f"Resultados pendientes ({len(excel_recovery.pending())})")

    def open_pending(self):
        panel = getattr(self, "recovery_panel", None)
        if panel is None or not panel.winfo_exists():
            self.recovery_panel = RecoveryPanel(self)
        else:
            panel.lift()


    def on_close(self):
        tasks = getattr(self, "_background_tasks", None)
        if tasks is not None and tasks.busy:
            self.view.log("Hay un proceso en curso. Espere a que termine y vuelva a cerrar la ventana.")
            return
        try:
            self.view.close_resources()
        finally:
            self.destroy()


if __name__ == "__main__":
    diagnostic_path = configure_logging()
    ctk.set_appearance_mode("dark")

    # Tema interno estable de CustomTkinter.
    # Evita errores por JSON incompleto.
    ctk.set_default_color_theme("green")

    # Agranda un poco toda la interfaz.
    ctk.set_widget_scaling(1.10)
    ctk.set_window_scaling(1.0)

    app = PresupuestoApp()
    app.view.log(f"Diagnostico de archivos: {diagnostic_path}")
    app.mainloop()
