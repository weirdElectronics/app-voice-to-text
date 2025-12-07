from flask import Flask, request, render_template, send_file
import speech_recognition as sr
from docx import Document
import os
import base64
import openpyxl
import re
import io

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
        # Crear documento Word en memoria
        doc = Document()
        p = doc.add_paragraph(texto)
        run = p.runs[0]
        run.font.name = "Courier New"

        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)

        return send_file(
            file_stream,
            as_attachment=True,
            download_name="transcripcion.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    elif modo == "suma":
        # Crear Excel en memoria
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
            if "mil" in texto.lower():
                monto *= 1000
        else:
            monto = 0.0

        descripcion = texto.replace(match.group(1), "").strip() if match else texto
        ws.append([descripcion, monto])

        total = sum(
            cell.value for cell in ws["B"][1:] if isinstance(cell.value, (int, float))
        )
        ws["A1"] = "TOTAL"
        ws["B1"] = total

        file_stream = io.BytesIO()
        wb.save(file_stream)
        file_stream.seek(0)

        return send_file(
            file_stream,
            as_attachment=True,
            download_name="gastos.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

@app.route('/reset_documento', methods=['POST'])
def reset_documento():
    # Ya no se necesita borrar archivos en servidor,
    # simplemente devolvemos documentos vacíos si se quiere reiniciar.
    doc = Document()
    file_stream_doc = io.BytesIO()
    doc.save(file_stream_doc)
    file_stream_doc.seek(0)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gastos"
    ws.append(["Descripción", "Monto"])
    file_stream_xlsx = io.BytesIO()
    wb.save(file_stream_xlsx)
    file_stream_xlsx.seek(0)

    return "Documentos reiniciados (se generarán vacíos al descargar)."

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)

