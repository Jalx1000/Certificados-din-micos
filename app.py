from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfbase import pdfmetrics
import pandas as pd
from reportlab.pdfbase.ttfonts import TTFont
import qrcode
import os
from io import BytesIO
from reportlab.lib.units import inch
import tempfile
from tqdm import tqdm

def camel_case(s):
    return ' '.join(word.capitalize() for word in s.split())

def text_url(s):
    return '-'.join(word.lower() for word in s.split())

def agregar_texto_vertical(c, x, y, texto, size, color, vertical=False):
    c.setFont("arialmt", size)  # Cambia la fuente y el tamaño según sea necesario
    c.setFillColor(color)  # Cambia el color si es necesario

    if vertical:
        # Guardar el estado actual del lienzo
        c.saveState()
        # Aplicar rotación y trasladar el origen
        c.translate(x, y)
        c.rotate(-90)
        # Dibujar el texto en la posición rotada
        c.drawString(0, 0, texto)
        # Restaurar el estado original del lienzo
        c.restoreState()
    else:
        # Dibujar el texto en la posición horizontal
        c.drawString(x, y, texto)

def agregar_texto(x,y,texto,size,color,):
    c.setFont("ArialMTExtraBold", size)  # Cambia la fuente y el tamaño según sea necesario
    c.setFillColor(color)  # Cambia el color si es necesario
    c.drawString(x, y, texto)
    c.drawString(x, y, texto) 

pdfmetrics.registerFont(TTFont('arialmt', 'arialmt.ttf'))
pdfmetrics.registerFont(TTFont('ArialMTExtraBold', 'ARIALMTEXTRABOLD.ttf'))

# Cargar la lista de nombres desde un archivo de Excel
df = pd.read_excel('nombre2.xlsx')  # Asegúrate de que el archivo .xlsx tenga una columna 'Nombre'
print(df.columns)

# Definir la posición del nombre en el PDF (ajustar según tu diseño)
nombre_posicion = (98, 630)  # (x, y) desde la esquina inferior izquierda
qr_posicion = (114, 81)  # (x, y) desde la esquina inferior izquierda, ajustar según sea necesario

# Cargar la plantilla del certificado
template_path = 'plantilla2.pdf'

for index, row in tqdm(df.iterrows()):
    # Obtener el nombre desde la fila actual
    nombre = row['Nombre']
    nombre_camelcase = camel_case(nombre)
    nombre_url = text_url(nombre)
    codigo = row['Codigo']
    codigo_url = text_url(codigo)
    # Generar la URL dinámica para el código QR
    url = f'https://wibel.net/{nombre_url}-{codigo_url}'

    # Crear el código QR
    qr = qrcode.QRCode(version=1, box_size=6, border=0.5)
    qr.add_data(url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill='black', back_color='white')

    # Guardar el código QR en un archivo temporal
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as temp_qr_file:
        qr_img.save(temp_qr_file, format='PNG')
        temp_qr_path = temp_qr_file.name

    # Crear un nuevo PDF con el nombre añadido
    temp_pdf = f'temp/{nombre}_temp.pdf'
    c = canvas.Canvas(temp_pdf, pagesize=(1500,2000))

    # Dibujar el nombre en el PDF en la posición deseada
    c.setFont("ArialMTExtraBold", 38)  # Cambia la fuente y el tamaño según sea necesario
    c.setFillColor("#303030")  # Cambia el color si es necesario
    c.drawString(nombre_posicion[0], nombre_posicion[1], nombre_camelcase)
    # agregar_texto(604,288,nombre_camelcase,15,"#333333")
    # Insertar el código QR en el PDF en la posición deseada
    c.drawImage(temp_qr_path, qr_posicion[0], qr_posicion[1], width=1.7*inch, height=1.7*inch)

    # Ejemplo de uso dentro del bucle
    agregar_texto_vertical(c, 90, 123, codigo, 8, "#333333", vertical=True)


    # Finalizar el PDF temporal
    c.save()

    # Eliminar el archivo temporal del código QR
    os.remove(temp_qr_path)

    # Leer el PDF de la plantilla y el PDF temporal con el nombre
    template_pdf = PdfReader(template_path)
    name_pdf = PdfReader(temp_pdf)

    # Crear un nuevo PDF con la combinación
    output = PdfWriter()
    for page in range(len(template_pdf.pages)):
        template_page = template_pdf.pages[page]
        if page == 0:
            template_page.merge_page(name_pdf.pages[0])  # Fusionar la primera página con el nombre
        output.add_page(template_page)

    # Guardar el certificado con el nombre del destinatario
    with open(f'certificados/{nombre_url}-{codigo_url}.pdf', 'wb') as outputStream:
        output.write(outputStream)

    # Eliminar el archivo PDF temporal
    os.remove(temp_pdf)

print("Certificados generados con éxito.")
