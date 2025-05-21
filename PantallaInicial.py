import tkinter as tk
from tkinter import filedialog

class RutaSelector:
    def __init__(self, root):
        self.root = root
        self.root.title("Creación de estimados Rolling")
        self.root.state('zoomed')
        self.root.configure(bg="#f9f9f9")

        self.rutas = {}
        self.valores = {}

        # ---------- Título principal ----------
        titulo = tk.Label(
            self.root,
            text="Creación de estimados Rolling",
            fg="#FF69B4",
            bg="#f9f9f9",
            font=("Segoe UI", 32, "bold")
        )
        titulo.pack(pady=40)

        # ---------- Contenedor principal ----------
        frame_contenido = tk.Frame(self.root, bg="#f9f9f9")
        frame_contenido.pack(pady=10, padx=20, fill="both", expand=True)

        # ---------- Archivos ----------
        frame_archivos = tk.LabelFrame(
            frame_contenido,
            text="Archivos",
            font=("Segoe UI", 14, "bold"),
            bg="#ffffff",
            padx=20,
            pady=20
        )
        frame_archivos.pack(fill="x", pady=10)

        campos_archivos = [
            ("RutaArchivoGlobal", "Archivo Global"),
            ("RutaArchivoCDL", "Archivo CDL"),
            ("RutaArchivoCrecimientos", "Archivo Crecimientos"),
            ("RutaArchivoVentaHistorica", "Venta Histórica"),
            ("RutaMacrosNovoApp", "Macros NovoApp"),
            ("RutaMacrosRolling", "Macros Rolling Forecast")
        ]

        for clave, etiqueta in campos_archivos:
            self.crear_campo_archivo(frame_archivos, clave, etiqueta)

        # ---------- Variables ----------
        frame_variables = tk.LabelFrame(
            frame_contenido,
            text="Variables",
            font=("Segoe UI", 14, "bold"),
            bg="#ffffff",
            padx=20,
            pady=20
        )
        frame_variables.pack(fill="x", pady=10)

        campos_texto = [
            ("InicioRollingCORP", "Inicio Rolling CORP"),
            ("InicioRollingPR", "Inicio Rolling PR"),
            ("AñoFinRolling", "Año Fin Rolling"),
            ("TipoEstimado", "Tipo Estimado")
        ]

        for clave, etiqueta in campos_texto:
            self.crear_campo_texto(frame_variables, clave, etiqueta)

        # ---------- Botones ----------
        # Usamos pack() para apilarlos uno debajo del otro
        self.crear_boton("Imprimir valores", self.mostrar_valores)
        self.crear_boton("Guardar configuración", self.guardar_configuracion)

    def crear_boton(self, texto, comando):
        boton = tk.Button(
            self.root,
            text=texto,
            command=comando,
            bg="#FF69B4",
            fg="white",
            font=("Segoe UI", 14, "bold"),
            padx=20,
            pady=10
        )
        boton.pack(pady=10, fill="x")

    def crear_campo_archivo(self, parent, nombre, texto):
        frame = tk.Frame(parent, bg="#ffffff")
        frame.pack(fill="x", pady=10)

        tk.Label(
            frame, text=texto + ":", bg="#ffffff",
            font=("Segoe UI", 12)
        ).pack(side="left", padx=10)

        btn = tk.Button(
            frame, text="Buscar", command=lambda: self.seleccionar_archivo(nombre),
            font=("Segoe UI", 12)
        )
        btn.pack(side="left", padx=5)

        lbl = tk.Label(
            frame, text="No seleccionado", bg="#ffffff",
            anchor="w", width=100, font=("Segoe UI", 12)
        )
        lbl.pack(side="left", padx=10)
        self.rutas[nombre] = {"ruta": "", "label": lbl}

    def crear_campo_texto(self, parent, nombre, texto):
        frame = tk.Frame(parent, bg="#ffffff")
        frame.pack(fill="x", pady=10)

        tk.Label(
            frame, text=texto + ":", bg="#ffffff",
            font=("Segoe UI", 12)
        ).pack(side="left", padx=10)

        entrada = tk.Entry(frame, width=50, font=("Segoe UI", 12))
        entrada.pack(side="left", padx=10)
        self.valores[nombre] = entrada

    def seleccionar_archivo(self, nombre):
        ruta = filedialog.askopenfilename()
        if ruta:
            self.rutas[nombre]["ruta"] = ruta
            self.rutas[nombre]["label"].config(text=ruta)

    def mostrar_valores(self):
        print("---- Rutas de archivos ----")
        for clave, valor in self.rutas.items():
            print(f"{clave} = {valor['ruta']}")

        print("\n---- Valores de entrada ----")
        for clave, entrada in self.valores.items():
            print(f"{clave} = {entrada.get()}")

    def guardar_configuracion(self):
        # Puedes agregar el código para guardar la configuración en un archivo .json o .xlsx
        print("Configuración guardada")

# Ejecutar interfaz
if __name__ == "__main__":
    root = tk.Tk()
    app = RutaSelector(root)
    root.mainloop()
