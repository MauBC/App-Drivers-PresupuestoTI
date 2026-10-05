import customtkinter as ctk

from src.gui.search_builder_view import SearchBuilderView


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

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_close(self):
        try:
            self.view.close_resources()
        finally:
            self.destroy()


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")

    # Tema interno estable de CustomTkinter.
    # Evita errores por JSON incompleto.
    ctk.set_default_color_theme("green")

    # Agranda un poco toda la interfaz.
    ctk.set_widget_scaling(1.10)
    ctk.set_window_scaling(1.0)

    app = PresupuestoApp()
    app.mainloop()
