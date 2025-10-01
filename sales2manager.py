import tkinter as tk
from tkinter import filedialog, ttk, messagebox, font
import pandas as pd
import unicodedata
import os
import threading
from datetime import datetime

# --- DETECCIÓN DE MODO WINDOWS/EXCEL ---
try:
    import win32com.client as win32
    WINDOWS_MODE = True
except ImportError:
    WINDOWS_MODE = False

# --- Variables Globales ---
df_global = pd.DataFrame()
ventana_carga = None

# --- Funciones de Lógica ---
def detectar_separador(file_path):
    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        primera_linea = f.readline()
        posibles_separadores = [",", ";", "|", "\t"]
        conteos = {sep: primera_linea.count(sep) for sep in posibles_separadores}
        if max(conteos.values()) == 0: return ","
        return max(conteos, key=conteos.get)

def mostrar_ventana_carga():
    global ventana_carga
    ventana_carga = tk.Toplevel(root)
    ventana_carga.title("Procesando"); ventana_carga.geometry("300x120"); ventana_carga.resizable(False, False)
    ventana_carga.transient(root); ventana_carga.grab_set()
    x = root.winfo_x() + (root.winfo_width()//2) - 150
    y = root.winfo_y() + (root.winfo_height()//2) - 60
    ventana_carga.geometry(f"+{x}+{y}")
    tk.Label(ventana_carga, text="Procesando grilla, por favor espere...", font=("Arial", 11)).pack(pady=30)
    root.update()

def cerrar_ventana_carga():
    global ventana_carga
    if ventana_carga:
        ventana_carga.destroy()
        ventana_carga = None
        
def cargar_archivo():
    file_path = filedialog.askopenfilename(
        title="Seleccionar archivo", filetypes=[("Archivos de datos", "*.csv *.xlsx *.xls")]
    )
    if not file_path:
        return
    
    mostrar_ventana_carga()

    def procesar_datos():
        global df_global
        try:
            # 1. Carga el archivo original
            if file_path.endswith(".csv"):
                sep = detectar_separador(file_path)
                df_cargado = pd.read_csv(file_path, sep=sep, on_bad_lines="warn")
            else:
                df_cargado = pd.read_excel(file_path)

            # 2. Define las columnas que deben existir en el archivo original
            columnas_fuente_necesarias = ["RUT_COMPRADOR", "DV_COMPRADOR"]
            
            # 3. Verifica que esas columnas existan
            columnas_faltantes = [col for col in columnas_fuente_necesarias if col not in df_cargado.columns]
            
            if columnas_faltantes:
                messagebox.showerror(
                    "Error de Columnas",
                    f"El archivo cargado no contiene las columnas necesarias: {', '.join(columnas_faltantes)}.\n\nPor favor, verifique el archivo."
                )
                limpiar_grilla()
                return

            # 4. Crea el DataFrame final desde cero
            df_final = pd.DataFrame()

            # 5. Construye la columna del RUT completo (RUT + DV)
            rut_completo = df_cargado["RUT_COMPRADOR"].astype(str) + '-' + df_cargado["DV_COMPRADOR"].astype(str)

            # 6. Construye el DataFrame final con nombres de columna únicos
            df_final["Vacio1"] = ""
            df_final["Vacio2"] = ""
            df_final["RutClie"] = rut_completo
            df_final["RutFact"] = rut_completo
            df_final["Fecha de Documento"] = "DINAMICA"
            df_final["NUMERO DE DOCUMENTO"] = "DINAMICA"
            df_final["FECHA DE VENCIMIENTO"] = "DINAMICA"
            df_final["Moneda"] = "$"
            df_final["Desc-Gral"] = "0"
            df_final["Tipo-Desc-Gral"] = "1"
            df_final["Codigo Postal"] = "DINAMICA"
            df_final["Cantidad"] = "DINAMICA"
            df_final["Precio Unitario"] = "DINAMICA"
            df_final["Descuento item"] = ""
            df_final["Bodega"] = "01"
            df_final["Cuenta Venta"] = "DINAMICA"
            df_final["Centro de Costos"] = "101"
            df_final["Observacion"] = ""
            df_final["Descripcion producto"] = "DINAMICA"
            df_final["Vacio3"] = ""
            df_final["Vacio4"] = ""
            df_final["Vacio5"] = ""
            df_final["Numero de OC"] = ""
            df_final["Codigo vendedor"] = "577"
            df_final["Codigo Sucursal"] = "01"
            df_final["Codigo Forma Pago"] = "DINAMICA"
            df_final["Glosa de pago"] = "Contado"
            df_final["Dias de vencimiento"] = ""
            df_final["Obs FAV"] = ""
            df_final["Fecha Entrega"] = "DINAMICA"
            df_final["Tipo de Venta"] = "E-COMMERCE"
            df_final["Obs Guia"] = ""
            df_final["Oc Referencia"] = ""
            df_final["Fecha OC Referencia"] = ""
            df_final["HES Referencia"] = ""
            df_final["Fecha HES Referencia"] = ""
            df_final["Fecha Guia Desp Ref"] = ""
            df_final["N° Contrato"] = ""
            df_final["Fecha Contrato"] = ""
            df_final["N° Pedido"] = ""
            df_final["Fecha Pedido"] = ""
            df_final["Aprobado"] = ""
            df_final["Contrato de arriendo"] = ""
            df_final["Atributo1"] = "DINAMICA"
            df_final["Atributo2"] = "DINAMICA"
            df_final["Atributo3"] = "DINAMICA"
            df_final["Atributo4"] = "NO APLICA"
            df_final["Atributo5"] = "2"
            df_final["Atributo6"] = "PDQ"
            df_final["Atributo7"] = "DINAMICA"
            
            # 7. Asigna el resultado a la variable global y muéstralo en la grilla
            df_global = df_final.copy()
            mostrar_grilla(df_global)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar o procesar el archivo:\n{e}")
        finally:
            cerrar_ventana_carga()

    # Llama a la función de procesamiento después de un breve instante
    root.after(50, procesar_datos)

def escribir_directo_a_excel95(df_datos, ruta_archivo):
    excel_app = None
    try:
        excel_app = win32.Dispatch("Excel.Application")
        excel_app.Visible = False
        wb = excel_app.Workbooks.Add()
        ws = wb.Worksheets(1)

        # Escribir encabezados
        for c_num, col_name in enumerate(df_datos.columns):
            # Si el nombre interno de la columna empieza con "Vacio" (Vacio1, Vacio2, etc.),
            # escribe la palabra explícita "Vacio" en el encabezado de Excel.
            nombre_final_columna = "Vacio" if str(col_name).startswith("Vacio") else col_name
            ws.Cells(1, c_num + 1).Value = nombre_final_columna

        # Escribir datos
        for r_num, row_data in enumerate(df_datos.itertuples(index=False), start=2):
            for c_num, cell_value in enumerate(row_data, start=1):
                ws.Cells(r_num, c_num).Value = str(cell_value).upper()

        # Guardar en formato Excel 95 (código 39)
        wb.SaveAs(os.path.abspath(ruta_archivo), FileFormat=39)
        wb.Close(False)
    finally:
        if excel_app:
            excel_app.Quit()

def generar_archivo_manager():
    if df_global.empty:
        messagebox.showwarning("Atención", "No hay datos en la tabla para exportar.")
        return

    def ejecutar_exportacion(ventana_emergente):
        fecha_hoy = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_sugerido = f"IMP_NV_Web_{fecha_hoy}.xls"
        
        file_path_validos = filedialog.asksaveasfilename(
            initialfile=nombre_sugerido, defaultextension=".xls",
            filetypes=[("Excel 5.0/95 Workbook (*.xls)", "*.xls")],
            title="Guardar archivo para Manager"
        )
        
        if not file_path_validos:
            ventana_emergente.destroy()
            return

        ventana_emergente.destroy()

        ventana_proceso = tk.Toplevel(root)
        ventana_proceso.title("Procesando")
        ventana_proceso.geometry("350x100")
        ventana_proceso.transient(root)
        ventana_proceso.grab_set()
        tk.Label(ventana_proceso, text="Creando archivo con Excel...\nEsto puede tardar un momento.", pady=20).pack()
        root.update_idletasks()

        def finalizar_con_exito(mensaje):
            ventana_proceso.destroy()
            messagebox.showinfo("Éxito", mensaje)

        def finalizar_con_error(e):
            ventana_proceso.destroy()
            messagebox.showerror("Error de Exportación", f"No se pudo crear el archivo con Excel.\n\nDetalle: {e}")

        def thread_exportar():
            try:
                escribir_directo_a_excel95(df_global, file_path_validos)
                mensaje_exito = f"Archivo guardado correctamente en:\n{file_path_validos}"
                root.after(0, finalizar_con_exito, mensaje_exito)
            except Exception as e:
                root.after(0, finalizar_con_error, e)

        hilo = threading.Thread(target=thread_exportar)
        hilo.start()

    ventana = tk.Toplevel(root)
    ventana.title("Confirmar Exportación"); ventana.geometry("400x150"); ventana.resizable(False, False)
    ventana.transient(root); ventana.grab_set()
    texto_info = f"Se exportarán {len(df_global)} filas.\n\n¿Desea continuar?"
    tk.Label(ventana, text=texto_info, font=("Arial", 11), justify=tk.LEFT).pack(pady=15, padx=20)
    frame_botones = tk.Frame(ventana); frame_botones.pack(pady=10)
    tk.Button(frame_botones, text="Cancelar", command=ventana.destroy, width=12).pack(side="left", padx=10)
    btn_generar_confirmado = tk.Button(frame_botones, text="Generar", command=lambda: ejecutar_exportacion(ventana), width=12)
    btn_generar_confirmado.pack(side="right", padx=10)
    root.wait_window(ventana)

# --- Funciones de la GUI ---
def bloquear_redimension(event):
    if tree.identify_region(event.x, event.y) == "separator": return "break"

def limpiar_grilla():
    global df_global
    tree.delete(*tree.get_children())
    tree["columns"] = ()
    df_global = pd.DataFrame()

def ajustar_ancho_columnas(tree):
    font_obj = font.Font()
    for col in tree["columns"]:
        if col == '#': continue
        max_width = font_obj.measure(tree.heading(col)["text"])
        for item in tree.get_children():
            try:
                col_index = tree["columns"].index(col)
                cell_text = str(tree.item(item, "values")[col_index])
                max_width = max(max_width, font_obj.measure(cell_text))
            except (IndexError, TypeError, ValueError): continue
        tree.column(col, width=max_width + 25, anchor="w", stretch=False)

def mostrar_grilla(df):
    limpiar_grilla()
    tree["columns"] = ["#"] + list(df.columns)
    tree["show"] = "headings"
    tree.heading("#", text="#"); tree.column("#", width=50, anchor="center", stretch=False)
    for col in df.columns:
        if col == '#': continue
        tree.heading(col, text=col, anchor='w'); tree.column(col, width=100, anchor="w")
    
    for idx, row in enumerate(df.itertuples(index=False)):
        valores_fila = list(row)
        fila_mayus = [str(v).upper() if pd.notna(v) else "" for v in valores_fila]
        tag = "oddrow" if (idx + 1) % 2 else "evenrow"
        tree.insert("", tk.END, values=[idx + 1] + fila_mayus, tags=(tag,))
    
    tree.tag_configure("oddrow", background="white")
    tree.tag_configure("evenrow", background="#f0f0f0")
    ajustar_ancho_columnas(tree)

# --- Configuración de la Ventana Principal y Widgets ---
root = tk.Tk()
root.title("Sales2Manager")
root.geometry("1200x800")
frame_botones_superiores = tk.Frame(root)
frame_botones_superiores.pack(pady=15)
btn_cargar = tk.Button(frame_botones_superiores, text="Cargar Archivo CSV/Excel", command=cargar_archivo, width=30, height=2)
btn_cargar.pack(side="left", padx=10)
btn_limpiar = tk.Button(frame_botones_superiores, text="Limpiar Grilla", command=limpiar_grilla, width=20, height=2)
btn_limpiar.pack(side="left", padx=10)
frame_grilla = tk.Frame(root)
frame_grilla.pack(padx=10, pady=10, expand=True, fill="both")
scroll_y = ttk.Scrollbar(frame_grilla, orient="vertical"); scroll_y.pack(side="right", fill="y")
scroll_x = ttk.Scrollbar(frame_grilla, orient="horizontal"); scroll_x.pack(side="bottom", fill="x")
tree = ttk.Treeview(frame_grilla, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
tree.pack(expand=True, fill="both")
tree.bind('<B1-Motion>', bloquear_redimension)
scroll_y.config(command=tree.yview); scroll_x.config(command=tree.xview)
btn_generar = tk.Button(root, text="Generar Archivo para Manager", command=generar_archivo_manager, width=30, height=2)
btn_generar.pack(pady=20)

if not WINDOWS_MODE:
    btn_generar.config(state="disabled")
    
footer = tk.Label(root, text="LLANTAS DEL PACIFICO - Departamento de Informática", font=("Arial", 10, "italic"), fg="gray")
footer.pack(side="bottom", pady=10)

root.mainloop()