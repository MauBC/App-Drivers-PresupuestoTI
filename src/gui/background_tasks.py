"""Un trabajo por ventana. El worker comunica eventos sin llamar a Tkinter."""
from queue import Queue, Empty
from threading import Thread, get_ident
from time import perf_counter

from src.core.excel_safety import logger


class BackgroundTasks:
    def __init__(self, root):
        self.root = root
        self.busy = False
        self.thread = None
        self._ui_thread = get_ident()
        self._events = Queue()

    def start(self, worker, button, idle_text, on_success, on_error, on_progress):
        if get_ident() != self._ui_thread:
            raise RuntimeError("Los trabajos deben iniciarse desde la interfaz")
        if self.busy:
            raise RuntimeError("Ya hay un proceso en curso. Espere a que termine antes de ejecutar otro.")
        self.busy = True
        self._button, self._idle_text = button, idle_text
        self._success, self._error, self._progress = on_success, on_error, on_progress
        button.configure(state="disabled", text="Preparando...")

        def run():
            started = perf_counter()
            phase, phase_started = "Preparacion", started
            def report(text):
                nonlocal phase, phase_started
                now = perf_counter()
                logger.info("Etapa terminada etapa=%s segundos=%.3f", phase, now - phase_started)
                phase, phase_started = text, now
                self._events.put(("progress", text))
            try:
                result = worker(report)
            except Exception as error:
                logger.exception("Fallo en trabajo de fondo")
                self._events.put(("error", error))
            else:
                self._events.put(("success", result))
            finally:
                logger.info("Etapa terminada etapa=%s segundos=%.3f", phase, perf_counter() - phase_started)
                logger.info("Trabajo de fondo terminado segundos=%.3f", perf_counter() - started)

        self.thread = Thread(target=run, name="presupuesto-worker", daemon=False)
        try:
            self.thread.start()
        except Exception:
            self.busy = False
            button.configure(state="normal", text=idle_text)
            raise
        self.root.after(40, self._poll)

    def _poll(self):
        try:
            while True:
                event, value = self._events.get_nowait()
                if event == "progress":
                    self._button.configure(text=value)
                    self._progress(value)
                else:
                    self.busy = False
                    self._button.configure(state="normal", text=self._idle_text)
                    if event == "success":
                        self._success(value)
                    else:
                        self._error(value)
                    return
        except Empty:
            pass
        finally:
            if self.busy:
                self.root.after(40, self._poll)


def tasks_for(widget):
    root = widget.winfo_toplevel()
    if not hasattr(root, "_background_tasks"):
        root._background_tasks = BackgroundTasks(root)
    return root._background_tasks
