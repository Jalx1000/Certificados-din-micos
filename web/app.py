from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import pandas as pd
import os
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import qrcode
import tempfile
from PyPDF2 import PdfReader, PdfWriter

app = FastAPI()

# Montar la carpeta estática para servir archivos estáticos
app.mount("/static", StaticFiles(directory="static"), name="static")

pdfmetrics.registerFont(TTFont('arialmt', 'arialmt.ttf'))
pdfmetrics.registerFont(TTFont('ArialMTExtraBold', 'ARIALMTEXTRABOLD.ttf'))


def camel_case(s: str) -> str:
    return ' '.join(word.capitalize() for word in s.split())


def text_url(s: str) -> str:
    return '-'.join(word.lower() for word in s.split())


def agregar_texto_vertical(c, x, y, texto, size, color, vertical=False):
    c.setFont("arialmt", size)
    c.setFillColor(color)
    if vertical:
        c.saveState()
        c.translate(x, y)
        c.rotate(-90)
        c.drawString(0, 0, texto)
        c.restoreState()
    else:
        c.drawString(x, y, texto)


@app.get("/", response_class=HTMLResponse)
def read_root():
    return FileResponse("static/upload.html")


@app.post("/upload/")
async def upload_file(file: UploadFile = File(...)):
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="Invalid file type")

    # Save the uploaded file
    file_path = os.path.join('data', 'nombre3.xlsx')
    os.makedirs('data', exist_ok=True)
    with open(file_path, 'wb') as f:
        f.write(await file.read())

    # Generar certificados
    df = pd.read_excel(file_path)
    template_path = 'Plantilla-bpl-julio.pdf'
    os.makedirs('certificados', exist_ok=True)

    for index, row in df.iterrows():
        nombre = row['Nombre']
        nombre_camelcase = camel_case(nombre)
        nombre_url = text_url(nombre)
        codigo = row['Codigo']
        codigo_url = text_url(codigo)
        url = f'https://wibel.net/{nombre_url}-{codigo_url}'

        qr = qrcode.QRCode(version=1, box_size=6, border=0.5)
        qr.add_data(url)
        qr.make(fit=True)
        qr_img = qr.make_image(fill='black', back_color='white')

        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as temp_qr_file:
            qr_img.save(temp_qr_file, format='PNG')
            temp_qr_path = temp_qr_file.name

        temp_pdf = f'temp/{nombre}_temp.pdf'
        os.makedirs('temp', exist_ok=True)
        c = canvas.Canvas(temp_pdf, pagesize=(1500, 2000))
        c.setFont("ArialMTExtraBold", 38)
        c.setFillColor("#303030")
        c.drawString(98, 630, nombre_camelcase)
        c.drawImage(temp_qr_path, 114, 81, width=1.7*inch, height=1.7*inch)
        agregar_texto_vertical(c, 104.2, 123, codigo, 8,
                               "#333333", vertical=True)
        c.save()

        os.remove(temp_qr_path)

        template_pdf = PdfReader(template_path)
        name_pdf = PdfReader(temp_pdf)
        output = PdfWriter()

        for page in range(len(template_pdf.pages)):
            template_page = template_pdf.pages[page]
            if page == 0:
                template_page.merge_page(name_pdf.pages[0])
            output.add_page(template_page)

        output_path = os.path.join(
            'certificados', f'{nombre_url}-{codigo_url}.pdf')
        with open(output_path, 'wb') as outputStream:
            output.write(outputStream)

        os.remove(temp_pdf)

    return {"message": "Certificates generated successfully"}


@app.get("/certificates/{filename}")
def get_certificate(filename: str):
    file_path = os.path.join('certificados', filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(file_path)
