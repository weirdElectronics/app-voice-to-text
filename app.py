from flask import Flask, request, render_template
import speech_recognition as sr
from docx import Document
import os
import base64
import openpyxl
import re

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/guardar_audio', methods=['POST'])
def guardar_audio():
    audio_b64 = request.form['audio']
    modo = request.form.get('modo', 'texto')
    audio_data = audio_b64.split(',')[1]
    audio_bytes = base64.b64decode(audio_data)

    webm_path = "grabacion.webm"
    wav_path = "grabacion.wav"
    with open(webm_path, "wb") as f:
        f.write(audio_bytes)

    os.system(f"ffmpeg -y -i {webm_path} -ar 16000 -ac 1 -f wav {wav_path}")

    r = sr.Recognizer()
    with sr.AudioFile(wav_path) as source:
        audio_rec = r.record(source)
    try:
        texto = r.recognize_google(audio_rec, language="es-AR")
    except Exception as e:
        return f"Error al transcribir: {e}"

    if modo == "texto":
        doc_path = os.path.expanduser("~/Desktop/transcripciones.docx")
        if os.path.exists(doc_path):
            doc = Document(doc_path)
        else:
            doc = Document()

        # Agregar párrafo con fuente tipo máquina de escribir
        p = doc.add_paragraph(texto)
        run = p.runs[0]
        run.font.name = "Courier New"   # fuente estilo máquina de escribir

        doc.save(doc_path)
        return f"Texto guardado en documento: {texto}"

    elif modo == "suma":
        excel_path = os.path.expanduser("~/Desktop/gastos.xlsx")
        if os.path.exists(excel_path):
            wb = openpyxl.load_workbook(excel_path)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Gastos"
            ws.append(["Descripción", "Monto"])

        # Buscar número en el texto
        match = re.search(r"(\d+(?:[.,]\d+)*)", texto.lower())
        if match:
            monto_str = match.group(1).replace(",", ".")
            try:
                monto = float(monto_str)
            except:
                monto = 0.0

            # Si aparece la palabra "mil" en el texto, multiplicar por 1000
            if "mil" in texto.lower():
                monto *= 1000
        else:
            monto = 0.0

        # Descripción sin el número
        descripcion = texto.replace(match.group(1), "").strip() if match else texto

        # Agregar gasto
        ws.append([descripcion, monto])

        # Calcular total en Python (sin duplicar filas TOTAL)
        total = sum(
            cell.value for cell in ws["B"][1:] if isinstance(cell.value, (int, float))
        )

        # Escribir total siempre en la primera fila
        ws["A1"] = "TOTAL"
        ws["B1"] = total

        wb.save(excel_path)
        return f"Gasto registrado: {descripcion} (monto: {monto})"


@app.route('/reset_documento', methods=['POST'])
def reset_documento():
    doc_path = os.path.expanduser("~/Desktop/transcripciones.docx")
    if os.path.exists(doc_path):
        os.remove(doc_path)

    # Crear documento vacío, sin encabezado
    doc = Document()
    doc.save(doc_path)

    excel_path = os.path.expanduser("~/Desktop/gastos.xlsx")
    if os.path.exists(excel_path):
        os.remove(excel_path)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gastos"
    ws.append(["Descripción", "Monto"])
    wb.save(excel_path)

    return "Documentos reiniciados. Word y Excel están vacíos y listos para nuevas transcripciones."


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
