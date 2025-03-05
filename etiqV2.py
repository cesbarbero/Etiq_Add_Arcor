#Final OK sin bordes

import win32print
import win32ui
from win32con import *

def centrarHorizontalmente(texto, anchoEtiquetaPx, font):
    # Crear un DC temporal para medir el tamaño del texto
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(win32print.GetDefaultPrinter())
    dc.SelectObject(font)
    
    # Medir el ancho del texto en píxeles
    textoAncho, _ = dc.GetTextExtent(texto)
    
    # Calcular la posición X para centrar horizontalmente
    x = (anchoEtiquetaPx - textoAncho) // 2  # Centrado horizontal
    
    return x

def imprimirEtiqueta(texto):
    # Obtener la impresora predeterminada
    impresora = win32print.GetDefaultPrinter()
    
    # Abrir la impresora
    hprinter = win32print.OpenPrinter(impresora)
    pdc = win32ui.CreateDC()
    pdc.CreatePrinterDC(impresora)
    
    # Crear fuente Arial, en itálica, negrita y tamaño 80 (ajustado)
    font = win32ui.CreateFont({
        "name": "Arial",  # Fuente estándar
        "height": 80,     # Tamaño de fuente en puntos
        "weight": FW_BOLD,  # Negrita
        "italic": True,   # Estilo itálica
    })
    
    # Dimensiones de la etiqueta en milímetros
    anchoEtiquetaMM = 80
    altoEtiquetaMM = 50
    
    # Convertir milímetros a píxeles (203 DPI)
    anchoEtiquetaPx = int(anchoEtiquetaMM * 203 / 25.4)
    altoEtiquetaPx = int(altoEtiquetaMM * 203 / 25.4)
    
    # Calcular las coordenadas para centrar el texto horizontalmente y verticalmente
    # Ojo, configurar el tamaño 80x50 en las Preferencias de la etiq. en Windows
    x = centrarHorizontalmente(texto, anchoEtiquetaPx, font)
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(win32print.GetDefaultPrinter())
    dc.SelectObject(font)
    _, textoAlto = dc.GetTextExtent(texto)
    y = (altoEtiquetaPx - textoAlto) // 2  # Centrado vertical

    pdc.SelectObject(font)
    
    pdc.StartDoc("Etiqueta")
    pdc.StartPage()
    
    # Escribir el texto en las coordenadas calculadas
    pdc.TextOut(x, y, texto)
    
    # Terminar la página y documento
    pdc.EndPage()
    pdc.EndDoc()
    pdc.DeleteDC()

def main():
    print("Seleccione un producto para la etiqueta:")
    productos = ["FRUIT DISC", "SPEARMINT", "BUTTERSCOTCH", "CINNAMON", "PEPPERMINT", "STRAWBERRY"]
    for i, producto in enumerate(productos, 1):
        print(f"{i}. {producto}")

    seleccion = int(input("Ingrese el número del producto: "))
    productoSeleccionado = productos[seleccion - 1]

    cantidad = int(input("Ingrese la cantidad de etiquetas a imprimir: "))

    if cantidad <= 0:
        print("Error: La cantidad debe ser mayor a 0.")
        return

    print("Imprimiendo etiquetas...")
    for i in range(cantidad):
        imprimirEtiqueta(productoSeleccionado)
        print(f"Etiqueta {i + 1} impresa.")

    print("Impresión completada.")

if __name__ == "__main__":
    main()
