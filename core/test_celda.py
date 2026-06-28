import xlwings as xw

ruta  = r"C:\Proyectos\cierre_202606.xlsx"  # ← tu ruta real
hoja  = "CIERRE"
celda = "E23"

with xw.App(visible=False, add_book=False) as app:
    wb = app.books.open(ruta)
    ws = wb.sheets[hoja]
    valor = ws.range(celda).value
    print(f"Valor xlwings: {valor}")
    wb.close()