from pathlib import Path
from tkinter import filedialog, messagebox
import customtkinter as ctk

from src.core import excel_recovery
from src.gui.background_tasks import tasks_for


class RecoveryPanel(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Resultados pendientes")
        self.geometry("900x500")
        self.protocol("WM_DELETE_WINDOW", self.close)
        ctk.CTkLabel(self, text="Recupere el resultado calculado sin ejecutar nuevamente el proceso.").pack(pady=10)
        ctk.CTkLabel(self, text="El guardado es local; la sincronización con la nube depende de OneDrive.").pack()
        self.content = ctk.CTkScrollableFrame(self)
        self.content.pack(fill="both", expand=True, padx=12, pady=12)
        self.refresh()

    def close(self):
        if not tasks_for(self.master).busy:
            self.destroy()

    def refresh(self):
        for child in self.content.winfo_children():
            child.destroy()
        records = excel_recovery.pending()
        if not records:
            ctk.CTkLabel(self.content, text="No hay resultados pendientes.").pack(pady=20)
        for record in records:
            row = ctk.CTkFrame(self.content)
            row.pack(fill="x", pady=6)
            ctk.CTkLabel(row, text=f"{Path(record['destination']).name}\nDestino: {record['destination']}",
                         wraplength=800, justify="left").pack(anchor="w", padx=10, pady=6)
            ctk.CTkButton(row, text="Reintentar guardado",
                         command=lambda r=record: self.run(r)).pack(side="left", padx=10, pady=10)
            ctk.CTkButton(row, text="Guardar en otro archivo",
                         command=lambda r=record: self.save_elsewhere(r)).pack(side="left", padx=10, pady=10)
        self.status = ctk.CTkButton(self.content, text="Actualizar lista", command=self.refresh)
        self.status.pack(pady=10)

    def save_elsewhere(self, record):
        suffix = Path(record["payload"]).suffix
        selected = filedialog.asksaveasfilename(
            parent=self, title="Guardar copia completa del libro generado (nombre nuevo)",
            initialfile=Path(record["destination"]).stem + "_recuperado" + suffix,
            defaultextension=suffix, filetypes=[("Libro Excel", "*" + suffix)],
        )
        if selected:
            self.run(record, selected)

    def run(self, record, alternate=None):
        manager = tasks_for(self.master)
        if manager.busy:
            messagebox.showinfo("Proceso en curso", "Espere a que termine el proceso actual.", parent=self)
            return
        def work(report):
            report("Reintentando guardado...")
            return excel_recovery.restore(record["id"], alternate)
        def finished(path):
            self.refresh()
            self.master.refresh_pending()
            messagebox.showinfo("Guardado en este equipo", f"Archivo: {path}\nLa sincronización con la nube depende de OneDrive.", parent=self)
        def failed(error):
            messagebox.showerror("Resultado conservado", str(error), parent=self)
        manager.start(work, self.status, "Actualizar lista", finished, failed, lambda text: None)
