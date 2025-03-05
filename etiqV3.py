#VER LINEA INFERIOR NO SALE, Y LINEA DERECHA SALE MUY AL BORDE

import win32print
import win32ui
from win32con import *

def centrarHorizontalmente(texto, anchoEtiquetaPx, font):
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(win32print.GetDefaultPrinter())
    dc.SelectObject(font)
    
    textoAncho, _ = dc.GetTextExtent(texto)
    x = (anchoEtiquetaPx - textoAncho) // 2  # Centrado horizontal
    return x

def imprimirEtiqueta(texto):
    # Obtener la impresora predeterminada
    # Ojo, configurar el tamaño 80x50 en las Preferencias de la etiq. en Windows
    impresora = win32print.GetDefaultPrinter()
    
    # Abrir la impresora
    pdc = win32ui.CreateDC()
    pdc.CreatePrinterDC(impresora)
    
    # Crear fuente Arial, en itálica, negrita y tamaño 80
    font = win32ui.CreateFont({
        "name": "Arial",
        "height": 75,
        "weight": FW_BOLD,
        "italic": True,
    })
    
    # Dimensiones de la etiqueta en milímetros
    anchoEtiquetaMM = 80
    altoEtiquetaMM = 50
    margenMM = 1  # Margen de 1 mm en todos los lados
    grosorLineaPx = 5  # Grosor de la línea en píxeles

    # Convertir milímetros a píxeles (203 DPI)
    anchoEtiquetaPx = int(anchoEtiquetaMM * 203 / 25.4)
    altoEtiquetaPx = int(altoEtiquetaMM * 203 / 25.4)
    margenPx = int(margenMM * 203 / 25.4)
    
    # Calcular la posición del texto
    x = centrarHorizontalmente(texto, anchoEtiquetaPx, font)
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(win32print.GetDefaultPrinter())
    dc.SelectObject(font)
    _, textoAlto = dc.GetTextExtent(texto)
    y = (altoEtiquetaPx - textoAlto) // 2  # Centrado vertical
    
    pdc.SelectObject(font)
    
    pdc.StartDoc("Etiqueta")
    pdc.StartPage()
    
    # Crear la pluma (pen) para el borde
    pen = win32ui.CreatePen(PS_SOLID, grosorLineaPx, 0x000000)  # Línea negra sólida
    pdc.SelectObject(pen)
    
    # Dibujar el rectángulo manualmente con márgenes
    pdc.MoveTo((margenPx, margenPx))  # Esquina superior izquierda
    pdc.LineTo((anchoEtiquetaPx - margenPx, margenPx))  # Línea superior
    pdc.LineTo((anchoEtiquetaPx - margenPx, altoEtiquetaPx - margenPx))  # Línea derecha
    pdc.LineTo((margenPx, altoEtiquetaPx - margenPx))  # Línea inferior
    pdc.LineTo((margenPx, margenPx))  # Línea izquierda, cerrando el rectángulo
    
    # Dibujar el texto centrado
    pdc.TextOut(x, y, texto)
    
    # Finalizar la página y el documento
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
