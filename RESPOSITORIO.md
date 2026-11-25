import os
import tempfile
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


# Cada diseño tiene un color principal y un fondo para las ideas clave.
DISENOS = {
    "Lavanda": ("634395", "EFE8FA"),
    "Azul académico": ("22577A", "E5F1F8"),
    "Verde natural": ("32634A", "E7F3EB"),
}


def crear_word(ruta, titulo, contenido, diseno):
    principal, fondo = DISENOS[diseno]
    color = RGBColor.from_string(principal)
    documento = Document()
    normal = documento.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    encabezado = documento.sections[0].header.paragraphs[0]
    encabezado.add_run("CUADERNO DE CLASE").bold = True

    portada = documento.add_heading(titulo.strip() or "Mis apuntes", 0)
    portada.runs[0].font.color.rgb = color
    documento.add_paragraph(date.today().strftime("%d/%m/%Y"))

    for linea in contenido.splitlines():
        if linea.startswith("## "):
            tema = documento.add_heading(linea[3:].strip() or "Tema", level=1)
            tema.runs[0].font.color.rgb = color
        elif linea.startswith("! "):
            celda = documento.add_table(rows=1, cols=1).cell(0, 0)
            sombreado = OxmlElement("w:shd")
            sombreado.set(qn("w:fill"), fondo)
            celda._tc.get_or_add_tcPr().append(sombreado)
            parrafo = celda.paragraphs[0]
            etiqueta = parrafo.add_run("IDEA CLAVE   ")
            etiqueta.bold = True
            etiqueta.font.color.rgb = color
            parrafo.add_run(linea[2:])
        elif linea.startswith("- "):
            documento.add_paragraph(linea[2:], style="List Bullet")
        else:
            documento.add_paragraph(linea)

    # El archivo anterior se reemplaza solo al terminar la escritura.
    with tempfile.NamedTemporaryFile(dir=ruta.parent, delete=False) as archivo:
        temporal = Path(archivo.name)
    try:
        documento.save(temporal)
        os.replace(temporal, ruta)
    finally:
        temporal.unlink(missing_ok=True)


class Cuaderno:
    def __init__(self, ventana):
        self.ventana = ventana
        self.ruta = None
        self.pendiente = None
        self.cambios = False
        ventana.title("Apuntes con Estilo")
        ventana.geometry("700x520")

        panel = ttk.Frame(ventana, padding=16)
        panel.pack(fill="both", expand=True)

        ttk.Label(panel, text="Título de la clase").pack(anchor="w")
        self.titulo = tk.StringVar(value="Mis apuntes de clase")
        ttk.Entry(panel, textvariable=self.titulo).pack(fill="x", pady=(3, 8))

        self.diseno = tk.StringVar(value="Lavanda")
        ttk.Label(panel, text="Diseño del Word").pack(anchor="w")
        ttk.Combobox(
            panel, textvariable=self.diseno, values=list(DISENOS),
            state="readonly"
        ).pack(anchor="w", pady=(3, 8))

        barra = ttk.Frame(panel)
        barra.pack(fill="x")
        for nombre, marca in [("Tema", "## "), ("Idea clave", "! "), ("Tarea", "- ")]:
            ttk.Button(
                barra, text=nombre,
                command=lambda prefijo=marca: self.insertar(prefijo)
            ).pack(side="left", padx=(0, 5))

        self.editor = tk.Text(panel, wrap="word", undo=True, font=("Calibri", 13))
        self.editor.pack(fill="both", expand=True, pady=8)
        ttk.Button(panel, text="Guardar Word…", command=self.guardar).pack(anchor="w")

        self.estado = tk.StringVar(value="Elige dónde guardar para activar el autoguardado.")
        ttk.Label(panel, textvariable=self.estado, wraplength=640).pack(anchor="w", pady=5)

        self.editor.bind("<<Modified>>", self.al_editar)
        self.titulo.trace_add("write", self.programar)
        self.diseno.trace_add("write", self.programar)
        ventana.protocol("WM_DELETE_WINDOW", self.cerrar)

    def insertar(self, marca):
        # Aplica el formato a la línea donde está el cursor.
        inicio = self.editor.index("insert linestart")
        linea = self.editor.get(inicio, "insert lineend")
        for anterior in ("## ", "! ", "- "):
            if linea.startswith(anterior):
                self.editor.delete(inicio, f"{inicio}+{len(anterior)}c")
                break
        self.editor.insert(inicio, marca)
        self.editor.mark_set("insert", f"{inicio} lineend")
        self.editor.focus_set()

    def al_editar(self, evento):
        if self.editor.edit_modified():
            self.editor.edit_modified(False)
            self.programar()

    def programar(self, *_):
        self.cambios = True
        if self.pendiente is not None:
            self.ventana.after_cancel(self.pendiente)
            self.pendiente = None
        if self.ruta:
            self.estado.set("Cambios pendientes…")
            self.pendiente = self.ventana.after(1200, self.guardar)

    def guardar(self):
        if self.pendiente is not None:
            self.ventana.after_cancel(self.pendiente)
            self.pendiente = None
        if not self.ruta:
            destino = filedialog.asksaveasfilename(
                defaultextension=".docx", initialfile="Apuntes.docx",
                filetypes=[("Documento de Word", "*.docx")]
            )
            if not destino:
                return False
            self.ruta = Path(destino)

        try:
            crear_word(
                self.ruta, self.titulo.get(),
                self.editor.get("1.0", "end-1c"), self.diseno.get()
            )
        except PermissionError:
            self.estado.set("Cierra el documento en Word y pulsa Guardar Word para reintentar.")
            return False
        except (OSError, ValueError) as error:
            self.estado.set(f"No se pudo guardar: {error}")
            return False

        self.cambios = False
        self.estado.set(f"Guardado en {self.ruta}")
        return True

    def cerrar(self):
        if self.cambios:
            respuesta = messagebox.askyesnocancel(
                "Cerrar cuaderno", "¿Quieres guardar los últimos cambios antes de salir?"
            )
            if respuesta is None or (respuesta and not self.guardar()):
                return
        self.ventana.destroy()


if __name__ == "__main__":
    raiz = tk.Tk()
    Cuaderno(raiz)
    raiz.mainloop()
