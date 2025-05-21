import tkinter as tk
from tkinter import filedialog, messagebox
from Eliminador import Eliminador

def seleccionar_archivo():
    archivo = filedialog.askopenfilename(title="Seleccionar Archivo CDL",
                                         filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo:
        entry_cdl.delete(0, tk.END)
        entry_cdl.insert(0, archivo)

def ejecutar_funcion_Inicio():
    ruta_cdl = entry_cdl.get().strip()
    inicio_corp = entry_inicio_corp.get().strip()
    inicio_pr = entry_inicio_pr.get().strip()
    fin_carga = entry_fin_carga.get().strip()

    if not ruta_cdl or not inicio_corp or not inicio_pr or not fin_carga:
        messagebox.showwarning("Advertencia", "Todos los campos deben estar llenos antes de iniciar la eliminación.")
        return
    if not inicio_corp.isdigit() or not inicio_pr.isdigit() or not fin_carga.isdigit():
        messagebox.showerror("Error", "Los valores de inicio Rolling Corporativo, Inicio Rolling PR y Fin Carga Rolling deben ser números.")
        return
    
    C(ruta_cdl, inicio_corp, inicio_pr, fin_carga)

def C(ruta_cdl, inicio_corp, inicio_pr, fin_carga):
    """Ejemplo de función que procesa los datos"""
    messagebox.showinfo("Ejecución", f"Ejecutando función C con:\n\n"
                                     f"Ruta CDL: {ruta_cdl}\n"
                                     f"Inicio Rolling Corporativo: {inicio_corp}\n"
                                     f"Inicio Rolling PR: {inicio_pr}\n"
                                     f"Fin Carga Rolling: {fin_carga}")
    E= Eliminador(ruta_cdl,inicio_corp,inicio_pr,fin_carga)

# Valores iniciales
inicioRollingCorporativo = ""
inicioRollingPR = ""
FinCargaRolling = ""

# Crear la ventana
root = tk.Tk()
root.title("Eliminación de Estimados Rolling")
root.geometry("700x400")  # ⬆ Aumentar resolución para mejor visibilidad
root.config(bg="white")

# Panel rosa (Encabezado)
panel_titulo = tk.Frame(root, bg="#A6269D", height=60)
panel_titulo.pack(fill="x")

titulo_label = tk.Label(panel_titulo, text="Eliminación de Estimados Rolling",
                        font=("Arial", 16, "bold"), bg="#A6269D", fg="black")
titulo_label.pack(pady=15)

# Contenedor principal
frame_contenido = tk.Frame(root, bg="white")
frame_contenido.pack(pady=20, padx=20, fill="both", expand=True)

# Configurar columnas para que se expandan con la ventana
frame_contenido.columnconfigure(1, weight=1)

# Etiquetas y Entradas
tk.Label(frame_contenido, text="Ruta CDL:", bg="white", font=("Arial", 12)).grid(row=0, column=0, sticky="w", padx=10, pady=10)
entry_cdl = tk.Entry(frame_contenido, width=60, font=("Arial", 12))
entry_cdl.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

btn_archivo = tk.Button(frame_contenido, text="Buscar Archivo", command=seleccionar_archivo, 
                        bg="#A6269D", fg="black", font=("Arial", 12, "bold"), width=15)
btn_archivo.grid(row=0, column=2, padx=10, pady=10)

tk.Label(frame_contenido, text="Inicio Rolling Corporativo:", bg="white", font=("Arial", 12)).grid(row=1, column=0, sticky="w", padx=10, pady=10)
entry_inicio_corp = tk.Entry(frame_contenido, width=20, font=("Arial", 12))
entry_inicio_corp.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
entry_inicio_corp.insert(0, str(inicioRollingCorporativo))

tk.Label(frame_contenido, text="Inicio Rolling PR:", bg="white", font=("Arial", 12)).grid(row=2, column=0, sticky="w", padx=10, pady=10)
entry_inicio_pr = tk.Entry(frame_contenido, width=20, font=("Arial", 12))
entry_inicio_pr.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
entry_inicio_pr.insert(0, str(inicioRollingPR))

tk.Label(frame_contenido, text="Fin Carga Rolling:", bg="white", font=("Arial", 12)).grid(row=3, column=0, sticky="w", padx=10, pady=10)
entry_fin_carga = tk.Entry(frame_contenido, width=20, font=("Arial", 12))
entry_fin_carga.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
entry_fin_carga.insert(0, str(FinCargaRolling))

# Botón para ejecutar la función C
btn_ejecutar = tk.Button(root, text="Ejecutar Función", command=ejecutar_funcion_Inicio, 
                         bg="#A6269D", fg="black", font=("Arial", 12, "bold"), width=20, height=2)
btn_ejecutar.pack(pady=20)

# Ejecutar la ventana
root.mainloop()
