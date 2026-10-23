"""Conversor Markdown Portable. SPDX-License-Identifier: AGPL-3.0-or-later."""
from __future__ import annotations

import argparse
import json
import os
import queue
import sys
import threading
import time
from dataclasses import asdict
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from converter import Cancelled, Options, convert


class App:
    def __init__(self, root):
        self.root = root
        root.title("Conversor a Markdown · Portable")
        root.geometry("960x740")
        root.minsize(800, 650)
        root.configure(bg="#f3f6f8")
        self.paths = []
        self.events = queue.Queue()
        self.cancel = threading.Event()
        self.running = False
        self.done_count = 0
        self.error_count = 0
        self.last_outputs = []
        self.output = tk.StringVar()
        self.status = tk.StringVar(value="Listo para convertir")
        self.markers = tk.BooleanVar(value=True)
        self.tables = tk.BooleanVar(value=True)
        self.headings = tk.BooleanVar(value=True)
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background="#f3f6f8")
        style.configure("TLabel", background="#f3f6f8", foreground="#203341", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 25, "bold"))
        style.configure("Sub.TLabel", foreground="#536979")
        style.configure("TButton", font=("Segoe UI", 10), padding=(12, 8))
        style.configure("Accent.TButton", background="#087e83", foreground="white", font=("Segoe UI", 11, "bold"))
        style.map("Accent.TButton", background=[("active", "#09696d"), ("disabled", "#a1b5ba")])
        style.configure("TCheckbutton", background="#f3f6f8", font=("Segoe UI", 10))
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=30, background="white", fieldbackground="white")
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), padding=7)
        style.configure("Horizontal.TProgressbar", background="#087e83", troughcolor="#dce6eb")
        body = ttk.Frame(root, padding=24)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Documentos → Markdown", style="Title.TLabel").pack(anchor="w")
        ttk.Label(body, text="PDF y Word .docx  ·  Conversión local  ·  Sin límite artificial de páginas", style="Sub.TLabel").pack(anchor="w", pady=(4, 18))
        toolbar = ttk.Frame(body)
        toolbar.pack(fill="x", pady=(0, 8))
        self.add = ttk.Button(toolbar, text="＋ Agregar documentos", command=self.add_files)
        self.add.pack(side="left")
        self.remove = ttk.Button(toolbar, text="Quitar seleccionados", command=self.remove_files)
        self.remove.pack(side="left", padx=8)
        self.count = ttk.Label(toolbar, text="0 archivos", style="Sub.TLabel")
        self.count.pack(side="right")
        frame = ttk.Frame(body)
        frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(frame, columns=("file", "size", "state"), show="headings", height=6)
        for col, label, width in [("file", "Documento", 450), ("size", "Tamaño", 95), ("state", "Estado", 200)]:
            self.tree.heading(col, text=label)
            self.tree.column(col, width=width, minwidth=80, stretch=col != "size")
        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        ttk.Label(body, text="Carpeta de destino", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(16, 5))
        destination = ttk.Frame(body)
        destination.pack(fill="x")
        self.entry = ttk.Entry(destination, textvariable=self.output, font=("Segoe UI", 10))
        self.entry.pack(side="left", fill="x", expand=True, ipady=7)
        self.choose = ttk.Button(destination, text="Elegir carpeta…", command=self.choose_folder)
        self.choose.pack(side="left", padx=(8, 0))
        opts = ttk.Frame(body)
        opts.pack(fill="x", pady=(12, 8))
        self.checks = []
        for text, var in [("Marcar páginas PDF", self.markers), ("Detectar tablas PDF", self.tables), ("Inferir títulos PDF", self.headings)]:
            check = ttk.Checkbutton(opts, text=text, variable=var)
            check.pack(side="left", padx=(0, 20))
            self.checks.append(check)
        ttk.Label(body, text="PDF escaneados: requieren OCR previo. Se exporta el texto; las imágenes no se incluyen.", style="Sub.TLabel").pack(anchor="w")
        ttk.Label(body, text="En diseños complejos, revisa el orden de lectura y las tablas del resultado.", style="Sub.TLabel").pack(anchor="w", pady=(2, 12))
        self.progress = ttk.Progressbar(body, mode="determinate", maximum=100)
        self.progress.pack(fill="x")
        ttk.Label(body, textvariable=self.status, style="Sub.TLabel").pack(anchor="w", pady=(5, 8))
        actions = ttk.Frame(body)
        actions.pack(fill="x")
        self.start = ttk.Button(actions, text="Convertir a Markdown", style="Accent.TButton", command=self.start_conversion)
        self.start.pack(side="left")
        self.stop = ttk.Button(actions, text="Cancelar", command=self.stop_conversion, state="disabled")
        self.stop.pack(side="left", padx=8)
        ttk.Button(actions, text="Abrir destino", command=self.open_folder).pack(side="right")
        self.log = tk.Text(body, height=5, wrap="word", font=("Segoe UI", 9), bg="white", fg="#425868", relief="flat", padx=10, pady=7, state="disabled")
        self.log.pack(fill="both", pady=(12, 0))
        self.log_message("Selecciona documentos y una carpeta de destino. Los archivos existentes se conservan.")
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.after(80, self.poll)

    def log_message(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def add_files(self):
        for name in filedialog.askopenfilenames(title="Seleccionar PDF o Word", filetypes=[("PDF y Word", "*.pdf *.docx"), ("PDF", "*.pdf"), ("Word", "*.docx")]):
            path = Path(name)
            if path not in self.paths:
                self.paths.append(path)
                self.tree.insert("", "end", iid=str(path), values=(path.name, f"{path.stat().st_size / 1048576:.1f} MB", "Pendiente"))
        self.count.configure(text=f"{len(self.paths)} archivos")
        if self.paths and not self.output.get():
            self.output.set(str(self.paths[0].parent / "Markdown"))

    def remove_files(self):
        for iid in self.tree.selection():
            self.paths.remove(Path(iid))
            self.tree.delete(iid)
        self.count.configure(text=f"{len(self.paths)} archivos")

    def choose_folder(self):
        folder = filedialog.askdirectory(title="Carpeta para los archivos Markdown")
        if folder:
            self.output.set(folder)

    def open_folder(self):
        folder = Path(self.output.get()) if self.output.get() else None
        if folder and folder.is_dir():
            os.startfile(str(folder.resolve()))
        else:
            messagebox.showinfo("Destino", "Elige una carpeta de destino. Se creará al convertir.")

    def set_running(self, running):
        self.running = running
        for widget in [self.add, self.remove, self.choose, self.entry, self.start] + self.checks:
            widget.configure(state="disabled" if running else "normal")
        self.stop.configure(state="normal" if running else "disabled")

    def start_conversion(self):
        if not self.paths:
            messagebox.showinfo("Documentos", "Agrega al menos un PDF o Word .docx.")
            return
        if not self.output.get().strip():
            messagebox.showinfo("Destino", "Selecciona una carpeta de destino.")
            return
        folder = Path(self.output.get().strip()).resolve()
        if folder.exists() and not folder.is_dir():
            messagebox.showerror("Destino", "El destino debe ser una carpeta.")
            return
        self.cancel.clear()
        self.done_count = self.error_count = 0
        self.last_outputs = []
        self.progress["value"] = 0
        self.set_running(True)
        options = Options(self.markers.get(), self.tables.get(), self.headings.get())
        for path in self.paths:
            self.tree.set(str(path), "state", "Pendiente")
        self.status.set("Iniciando conversión…")
        threading.Thread(target=self.worker, args=(list(self.paths), folder, options), daemon=True).start()

    def worker(self, paths, folder, options):
        started = time.monotonic()
        try:
            for index, path in enumerate(paths):
                if self.cancel.is_set():
                    break
                self.events.put(("state", str(path), "Convirtiendo…"))
                last = [0.0]
                def progress(done, total, label):
                    now = time.monotonic()
                    if now - last[0] > .1 or done == total:
                        last[0] = now
                        self.events.put(("progress", 100 * (index + done / max(total, 1)) / len(paths), f"{index + 1}/{len(paths)} · {path.name} · {label}"))
                try:
                    result = convert(path, folder, options, progress, self.cancel)
                    self.events.put(("result", str(path), result))
                except Cancelled:
                    self.events.put(("state", str(path), "Cancelado"))
                    break
                except Exception as exc:
                    self.events.put(("error", str(path), str(exc)))
        finally:
            self.events.put(("finished", time.monotonic() - started))

    def poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "state":
                    self.tree.set(event[1], "state", event[2])
                elif kind == "progress":
                    self.progress["value"] = event[1]
                    self.status.set(event[2])
                elif kind == "result":
                    result = event[2]
                    self.done_count += 1
                    self.last_outputs.append(result.output)
                    self.tree.set(event[1], "state", "Listo · revisar avisos" if result.warnings else "Listo")
                    self.log_message(f"Guardado: {result.output}")
                    for warning in result.warnings:
                        self.log_message("  Aviso: " + warning)
                    if result.empty_pages:
                        self.log_message("  Páginas sin texto: " + ", ".join(map(str, result.empty_pages)))
                elif kind == "error":
                    self.error_count += 1
                    self.tree.set(event[1], "state", "Error")
                    self.log_message(f"Error en {Path(event[1]).name}: {event[2]}")
                elif kind == "finished":
                    self.set_running(False)
                    label = "Cancelado" if self.cancel.is_set() else "Finalizado"
                    if not self.cancel.is_set():
                        self.progress["value"] = 100
                    self.status.set(f"{label} · {self.done_count} convertido(s) · {self.error_count} error(es) · {event[1]:.1f} s")
                    for p in self.paths:
                        if self.tree.set(str(p), "state") == "Pendiente":
                            self.tree.set(str(p), "state", "Sin procesar")
        except queue.Empty:
            pass
        self.root.after(80, self.poll)

    def stop_conversion(self):
        self.cancel.set()
        self.stop.configure(state="disabled")
        self.status.set("Cancelando al terminar el bloque actual…")

    def close(self):
        if self.running:
            messagebox.showinfo("Conversión en curso", "Pulsa Cancelar y espera a que finalice antes de cerrar.")
        else:
            self.root.destroy()


def main():
    parser = argparse.ArgumentParser(description="Conversor local PDF / DOCX a Markdown")
    parser.add_argument("--convert", nargs="+", metavar="ARCHIVO")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--report", type=Path, help="Informe JSON de la conversión")
    parser.add_argument("--no-tables", action="store_true")
    parser.add_argument("--no-headings", action="store_true")
    parser.add_argument("--no-page-markers", action="store_true")
    parser.add_argument("--check-ui", type=Path, help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.check_ui:
        root = tk.Tk()
        App(root)
        root.update()
        args.check_ui.write_text("UI OK", encoding="utf-8")
        root.destroy()
        return 0
    if args.convert:
        if not args.output:
            parser.error("--output es obligatorio con --convert")
        results = []
        failed = False
        for name in args.convert:
            try:
                result = convert(Path(name), args.output, Options(not args.no_page_markers, not args.no_tables, not args.no_headings))
                results.append({"status": "ok", **asdict(result)})
            except Exception as exc:
                failed = True
                results.append({"status": "error", "source": name, "error": str(exc)})
        report = json.dumps(results, ensure_ascii=False, indent=2)
        if args.report:
            args.report.write_text(report, encoding="utf-8")
        if sys.stdout:
            print(report)
        return int(failed)
    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
