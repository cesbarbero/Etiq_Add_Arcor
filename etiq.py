import win32print
import win32api
import tkinter as tk
from tkinter import ttk

def centrar_texto(texto, ancho):
    """Centra el texto en un espacio de ancho dado."""
    if len(texto) >= ancho:
        return texto[:ancho]  # Recorta si el texto es más largo que el ancho
    espacios = (ancho - len(texto)) // 2
    return " " * espacios + texto

def centrar_verticalmente(texto, alto, ancho):
    """Crea un texto centrado verticalmente en un espacio de altura dada."""
    lineas_vacias = (alto - 1) // 2  # Líneas vacías arriba y abajo del texto
    linea_centrada = centrar_texto(texto, ancho)
    return ["" * ancho] * lineas_vacias + [linea_centrada] + ["" * ancho] * lineas_vacias

def generar_etiqueta(producto, ancho_mm, alto_mm):
    """Genera una etiqueta con bordes finos y el texto centrado horizontal y verticalmente."""
    dpi = 203  # Resolución estándar (203 DPI)
    ancho_caracteres = int(ancho_mm / 2.5)  # Aproximación: 1 carácter ~ 2.5mm de ancho
    alto_lineas = int(alto_mm / 6)  # Aproximación: 1 línea ~ 6mm de alto

    etiqueta = ["-" * ancho_caracteres]  # Línea superior del borde
    etiqueta += centrar_verticalmente(producto, alto_lineas, ancho_caracteres)
    etiqueta.append("-" * ancho_caracteres)  # Línea inferior del borde

    return "\n".join(etiqueta) + "\n"

def imprimir_etiquetas(producto, cantidad, ancho_mm, alto_mm):
    """Imprime etiquetas centradas para el producto seleccionado."""
    etiqueta = generar_etiqueta(producto, ancho_mm, alto_mm)

    # Obtener la impresora predeterminada
    impresora = win32print.GetDefaultPrinter()
    print(f"Imprimiendo en: {impresora}")

    try:
        # Enviar las etiquetas a la impresora
        for _ in range(cantidad):
            # Abre la impresora y escribe los datos
            handle = win32print.OpenPrinter(impresora)
            job = win32print.StartDocPrinter(handle, 1, ("Etiqueta", None, "RAW"))
            win32print.StartPagePrinter(handle)
            win32print.WritePrinter(handle, etiqueta.encode("utf-8"))
            win32print.EndPagePrinter(handle)
            win32print.EndDocPrinter(handle)
            win32print.ClosePrinter(handle)
        print(f"{cantidad} etiquetas impresas correctamente.")
    except Exception as e:
        print(f"Error al imprimir: {e}")

def seleccionar_producto(productos):
    """Muestra una ventana con una lista desplegable para seleccionar un producto."""
    def on_select():
        nonlocal producto_seleccionado
        producto_seleccionado = combo.get()
        root.destroy()

    root = tk.Tk()
    root.title("Seleccionar Producto")
    root.geometry("800x600")  # Aumentar el tamaño de la ventana

    tk.Label(root, text="Seleccione un producto:", font=("Arial", 16)).pack(pady=40)  # Aumentar el tamaño del texto

    producto_seleccionado = None
    combo = ttk.Combobox(root, values=productos, state="readonly", font=("Arial", 16), width=40)  # Aumentar tamaño del combo
    combo.pack(pady=20)
    combo.current(0)  # Seleccionar el primer producto por defecto

    tk.Button(root, text="Aceptar", command=on_select, font=("Arial", 16)).pack(pady=40)  # Aumentar el tamaño del botón

    root.mainloop()

    if producto_seleccionado:
        return producto_seleccionado
    else:
        raise ValueError("No se seleccionó ningún producto.")

if __name__ == "__main__":
    # Lista de productos
    productos = ["FRUIT DISC", "SPEARMINT", "BUTTERSCOTCH", "CINNAMON", "PEPPERMINT", "STRAWBERRY"]
    ancho_etiqueta_mm = 80  # Ancho de la etiqueta en milímetros
    alto_etiqueta_mm = 50   # Alto de la etiqueta en milímetros

    try:
        # Selección del producto con lista desplegable
        producto_seleccionado = seleccionar_producto(productos)

        # Solicitar cantidad de copias
        cantidad = int(input("Ingrese la cantidad de copias a imprimir: "))
        if cantidad <= 0:
            raise ValueError("La cantidad debe ser mayor a 0.")

        # Imprimir etiquetas
        imprimir_etiquetas(producto_seleccionado, cantidad, ancho_etiqueta_mm, alto_etiqueta_mm)

    except ValueError as e:
        print(f"Error: {e}")
