from flask import Flask, request, render_template, redirect, url_for
import pdfrw
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import os

app = Flask(__name__)

PDF_TEMPLATE_PATH = 'C:/Users/sgnca/Documents/DEBITOCTSGL/DEBITOCTSGL.pdf'

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/submit_form', methods=['POST'])
def submit_form():
    # Obtener datos del formulario
    titular_cuenta = request.form['titular_cuenta']
    cbu = request.form['cbu']
    alias = request.form['alias']
    nombre_cliente = request.form['nombre_cliente']
    dni = request.form['dni']
    numero_abonado = request.form['numero_abonado']

    # Crear el nombre del archivo PDF de salida basado en nombre_cliente
    output_pdf_path = f"{nombre_cliente}.pdf"

    # Crear diccionario de datos
    data = {
        'titular_cuenta': titular_cuenta,
        'cbu': cbu,
        'alias': alias,
        'nombre_cliente': nombre_cliente,
        'dni': dni,
        'Numero de Abonado y Servicios': numero_abonado,
        'dia': datetime.now().strftime('%d'),
        'mes': datetime.now().strftime('%m'),
        'año': datetime.now().strftime('%Y')
    }

    # Rellenar el PDF con los datos del formulario
    rellenar_pdf(PDF_TEMPLATE_PATH, output_pdf_path, data)

    # Enviar el PDF rellenado por correo electrónico
    enviar_correo(output_pdf_path)

    # Mostrar ventana emergente y redirigir a la página principal
    return render_template('message.html')

def rellenar_pdf(template_path, output_path, data_dict):
    template_pdf = pdfrw.PdfReader(template_path)
    annotations = template_pdf.pages[0]['/Annots']

    for annotation in annotations:
        if annotation['/T']:
            key = annotation['/T'][1:-1]
            if key in data_dict:
                annotation.update(
                    pdfrw.PdfDict(V=f'{data_dict[key]}')
                )

    pdfrw.PdfWriter().write(output_path, template_pdf)

def enviar_correo(pdf_path):
    from_email = "solicitudes@cooptelsangenaro.com.ar"
    from_password = "@Ctsgl2146"
    to_email = "cooptelsangenaro@gmail.com"

    # Crear el mensaje
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = "Solicitud de Débito Automático"

    body = "Adjunto se encuentra la solicitud de débito automático."
    msg.attach(MIMEText(body, 'plain'))

    # Adjuntar PDF
    with open(pdf_path, "rb") as pdf_file:
        attach = MIMEApplication(pdf_file.read(), _subtype="pdf")
        attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_path))
        msg.attach(attach)

    # Enviar el correo
    try:
        with smtplib.SMTP_SSL('c1930171.ferozo.com', 465) as server:
            server.set_debuglevel(1)
            server.login(from_email, from_password)
            server.sendmail(from_email, to_email, msg.as_string())
    except Exception as e:
        print(f"Error al enviar correo: {e}")

if __name__ == "__main__":
    app.run(debug=True)
