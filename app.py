
from flask import Flask, request, render_template, send_file, session, redirect, url_for
import speech_recognition as sr
from docx import Document
import os
import base64
import openpyxl
import re
import uuid
from io import BytesIO

# Google Drive (OAuth)
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = "clave_supersecreta"

# Configuración Google Drive
SCOPES = ['https://www.googleapis.com/auth/drive.file']
CREDENTIALS_FILE = 'credentials.json'
FOLDER_ID = os.getenv("DRIVE_FOLDER_ID")

# -------------------------------
# Parseo de montos en español
# -------------------------------
import re

def parse_amount_es(texto: str):
    """
    Devuelve (monto_float, descripcion_sin_montos).
    Soporta:
      - "500.850" -> 500850.0
      - "1.200,50" -> 1200.50
      - "2.000.000" -> 2000000.0
      - "2000.000" (ASR raro) -> 2000000.0
      - "350 mil" -> 350000.0
      - "350 mil 500" -> 350500.0
    """
    t = texto.lower()

    # Caso "X mil Y" o "X mil"
    m_compuesto = re.search(r'(\d+)\s*mil(?:\s*(\d+))?', t)
    if m_compuesto:
        x = int(m_compuesto.group(1))
        y = int(m_compuesto.group(2)) if m_compuesto.group(2) else 0
        monto = x * 1000 + y
        desc = re.sub(r'(\d+\s*mil(?:\s*\d+)?)', '', t)
        desc = re.sub(r'\b(pesos?|ars|argentinos?)\b', '', desc)
        desc = re.sub(r'\s+', ' ', desc).strip()
        return float(monto), desc

    # Número con separadores
    m_num = re.search(r'\d+(?:[.,]\d+)*', t)
    if m_num:
        raw = m_num.group(0)
        s = raw.replace(' ', '')

        if ',' in s and '.' in s:
            # puntos como miles, coma como decimales: 1.234,56 -> 1234.56
            s = s.replace('.', '').replace(',', '.')
        elif '.' in s:
            parts = s.split('.')
            # múltiples puntos o último bloque de 3 dígitos => puntos como miles
            if s.count('.') > 1 or (len(parts) > 1 and len(parts[-1]) == 3):
                s = ''.join(parts)  # 500.850 -> 500850, 2000.000 -> 2000000
        elif ',' in s:
            s = s.replace(',', '.')

        try:
            monto = float(s)
        except:
            monto = 0.0

        # “X mil” sin compuesto explícito (ej: "400 mil")
        if 'mil' in t and s.isdigit():
            monto *= 1000

        # Quitar el número y palabras de moneda de la descripción
        desc = re.sub(re.escape(raw), '', t)
        desc = re.sub(r'\bmil\b', '', desc)
        desc = re.sub(r'\b(pesos?|ars|argentinos?)\b', '', desc)
        desc = re.sub(r'\s+', ' ', desc).strip()

        return monto, desc

    # Sin número: devolver texto limpio
    desc = re.sub(r'\b(pesos?|ars|argentinos?)\b', '', t)
    desc = re.sub(r'\s+', ' ', desc).strip()
    return 0.0, desc

# -------------------------------
# OAuth y Drive helpers
# -------------------------------
def get_credentials():
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        if creds and creds.valid:
            return creds
    flow = Flow.from_client_secrets_file(
        CREDENTIALS_FILE,
        scopes=SCOPES,
        redirect_uri=url_for("oauth2callback", _external=True)
    )
    auth_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent"
    )
    session["oauth_state"] = state
    return redirect(auth_url)

@app.route("/oauth2callback")
def oauth2callback():
    state = session.get("oauth_state")
    flow = Flow.from_client_secrets_file(
        CREDENTIALS_FILE,
        scopes=SCOPES,
        redirect_uri=url_for("oauth2callback", _external=True)
    )
    flow.fetch_token(authorization_response=request.url)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return redirect(url_for("index"))

def build_drive_service(creds):
    return build("drive", "v3", credentials=creds)

def get_user_id():
    user_id = session.get("user_id")
    if not user_id:
        user_id = str(uuid.uuid4())
        session["user_id"] = user_id
    return user_id

# -------------------------------
# Guardar Word acumulativo
# -------------------------------
def save_word_to_drive(user_id, new_doc):
    creds = get_credentials()
    if not isinstance(creds, Credentials):
        return creds
    drive_service = build_drive_service(creds)

    filename = f"{user_id}_transcripciones.docx"
    query = f"name='{filename}'"
    if FOLDER_ID:
        query += f" and '{FOLDER_ID}' in parents"

    results = drive_service.files().list(q=query, fields="files(id)").execute()
    items = results.get("files", [])

    if items:
        file_id = items[0]["id"]
        request = drive_service.files().get_media(fileId=file_id)
        existing_buffer = BytesIO()
        downloader = MediaIoBaseDownload(existing_buffer, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        existing_buffer.seek(0)

        existing_doc = Document(existing_buffer)
        for p in new_doc.paragraphs:
            existing_doc.add_paragraph(p.text)

        buffer = BytesIO()
        existing_doc.save(buffer)
        buffer.seek(0)
        media = MediaIoBaseUpload(
            buffer,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            resumable=True
        )
        drive_service.files().update(fileId=file_id, media_body=media).execute()
        return file_id
    else:
        buffer = BytesIO()
        new_doc.save(buffer)
        buffer.seek(0)
        file_metadata = {"name": filename}
        if FOLDER_ID:
            file_metadata["parents"] = [FOLDER_ID]
        created = drive_service.files().create(
            body=file_metadata,
            media_body=MediaIoBaseUpload(
                buffer,
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                resumable=True
            ),
            fields="id"
        ).execute()
        return created.get("id")

# -------------------------------
# Guardar Excel acumulativo (único TOTAL al final)
# -------------------------------
def save_excel_to_drive(user_id, new_wb):
    creds = get_credentials()
    if not isinstance(creds, Credentials):
        return creds
    drive_service = build_drive_service(creds)

    filename = f"{user_id}_gastos.xlsx"
    query = f"name='{filename}'"
    if FOLDER_ID:
        query += f" and '{FOLDER_ID}' in parents"

    results = drive_service.files().list(q=query, fields="files(id)").execute()
    items = results.get("files", [])

    if items:
        # Descargar y abrir existente
        file_id = items[0]["id"]
        request = drive_service.files().get_media(fileId=file_id)
        existing_buffer = BytesIO()
        downloader = MediaIoBaseDownload(existing_buffer, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        existing_buffer.seek(0)

        existing_wb = openpyxl.load_workbook(existing_buffer)
        ws = existing_wb.active

        # 1) Eliminar filas TOTAL previas
        for row_idx in range(ws.max_row, 1, -1):
            val = ws.cell(row=row_idx, column=1).value
            if isinstance(val, str) and val.strip().upper() == "TOTAL":
                ws.delete_rows(row_idx, 1)

        # 2) Asegurar encabezados
        header_a1 = ws.cell(row=1, column=1).value
        header_b1 = ws.cell(row=1, column=2).value
        has_desc = isinstance(header_a1, str) and "descrip" in header_a1.lower()
        has_monto = isinstance(header_b1, str) and "monto" in header_b1.lower()
        if not (has_desc and has_monto):
            ws.delete_rows(1, 1)
            ws.insert_rows(1)
            ws.cell(row=1, column=1, value="Descripción")
            ws.cell(row=1, column=2, value="Monto")

        # 3) Agregar filas nuevas (sin encabezados)
        for i, row in enumerate(new_wb.active.iter_rows(values_only=True)):
            if i == 0 and row and len(row) >= 2:
                header_like = (
                    isinstance(row[0], str) and "descrip" in row[0].lower()
                    or isinstance(row[1], str) and "monto" in row[1].lower()
                )
                if header_like:
                    continue
            ws.append(row)

        # 4) Calcular el total ignorando encabezados
        total = 0.0
        for row_idx in range(2, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=2).value
            if isinstance(val, (int, float)):
                total += float(val)

        # 5) Agregar ÚNICA fila TOTAL al final
        ws.append(["TOTAL", total])

        # Subir actualización
        buffer = BytesIO()
        existing_wb.save(buffer)
        buffer.seek(0)
        media = MediaIoBaseUpload(
            buffer,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=True
        )
        drive_service.files().update(fileId=file_id, media_body=media).execute()
        return file_id

    else:
        # Crear nuevo Excel con encabezados y TOTAL
        base_wb = openpyxl.Workbook()
        base_ws = base_wb.active
        base_ws.title = "Gastos"
        base_ws.append(["Descripción", "Monto"])

        for row in new_wb.active.iter_rows(values_only=True):
            base_ws.append(row)

        total = 0.0
        for row_idx in range(2, base_ws.max_row + 1):
            val = base_ws.cell(row=row_idx, column=2).value
            if isinstance(val, (int, float)):
                total += float(val)

        base_ws.append(["TOTAL", total])

        buffer = BytesIO()
        base_wb.save(buffer)
        buffer.seek(0)
        file_metadata = {"name": filename}
        if FOLDER_ID:
            file_metadata["parents"] = [FOLDER_ID]
        created = drive_service.files().create(
            body=file_metadata,
            media_body=MediaIoBaseUpload(
                buffer,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                resumable=True
            ),
            fields="id"
        ).execute()
        return created.get("id")



# -------------------------------
# Rutas base
# -------------------------------
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

    user_id = get_user_id()

    if modo == "texto":
        doc = Document()
        doc.add_paragraph(texto)
        save_word_to_drive(user_id, doc)
        return f"Texto guardado en documento: {texto}"
        
    elif modo == "suma":
        # Crear workbook temporal con la fila del nuevo gasto
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Gastos"

        monto, descripcion = parse_amount_es(texto)
        if not descripcion:
            descripcion = "Gasto"

        ws.append([descripcion, monto])

        save_excel_to_drive(user_id, wb)
        return f"Gasto registrado: {descripcion} (monto: {monto})"


        ws.append([descripcion, monto])

        save_excel_to_drive(user_id, wb)
        return f"Gasto registrado: {descripcion} (monto: {monto})"


# -------------------------------
# Ver online
# -------------------------------
@app.route('/ver_word')
def ver_word():
    user_id = get_user_id()
    creds = get_credentials()
    if not isinstance(creds, Credentials):
        return creds
    drive_service = build_drive_service(creds)

    filename = f"{user_id}_transcripciones.docx"
    query = f"name='{filename}'"
    if FOLDER_ID:
        query += f" and '{FOLDER_ID}' in parents"

    results = drive_service.files().list(q=query, fields="files(id)").execute()
    items = results.get("files", [])
    if not items:
        return render_template("ver_word.html", contenido=[])

    file_id = items[0]["id"]
    request = drive_service.files().get_media(fileId=file_id)
    buffer = BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    buffer.seek(0)

    doc = Document(buffer)
    contenido = [p.text for p in doc.paragraphs]
    return render_template("ver_word.html", contenido=contenido)

@app.route('/ver_excel')
def ver_excel():
    user_id = get_user_id()
    creds = get_credentials()
    if not isinstance(creds, Credentials):
        return creds
    drive_service = build_drive_service(creds)

    filename = f"{user_id}_gastos.xlsx"
    query = f"name='{filename}'"
    if FOLDER_ID:
        query += f" and '{FOLDER_ID}' in parents"

    results = drive_service.files().list(q=query, fields="files(id)").execute()
    items = results.get("files", [])
    if not items:
        return render_template("ver_excel.html", filas=[])

    file_id = items[0]["id"]
    request = drive_service.files().get_media(fileId=file_id)
    buffer = BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    buffer.seek(0)

    wb = openpyxl.load_workbook(buffer)
    ws = wb.active
    filas = [row for row in ws.iter_rows(values_only=True)]
    return render_template("ver_excel.html", filas=filas)

# -------------------------------
# Descargar
# -------------------------------
@app.route('/descargar_doc')
def descargar_doc():
    user_id = get_user_id()
    creds = get_credentials()
    if not isinstance(creds, Credentials):
        return creds
    drive_service = build_drive_service(creds)

    filename = f"{user_id}_transcripciones.docx"
    query = f"name='{filename}'"
    if FOLDER_ID:
        query += f" and '{FOLDER_ID}' in parents"

    results = drive_service.files().list(q=query, fields="files(id)").execute()
    items = results.get("files", [])
    if not items:
        return "No hay documento aún."

    file_id = items[0]["id"]
    request = drive_service.files().get_media(fileId=file_id)
    buffer = BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    buffer.seek(0)

    return send_file(buffer, as_attachment=True, download_name="transcripciones.docx")

@app.route('/descargar_excel')
def descargar_excel():
    user_id = get_user_id()
    creds = get_credentials()
    if not isinstance(creds, Credentials):
        return creds
    drive_service = build_drive_service(creds)

    filename = f"{user_id}_gastos.xlsx"
    query = f"name='{filename}'"
    if FOLDER_ID:
        query += f" and '{FOLDER_ID}' in parents"

    results = drive_service.files().list(q=query, fields="files(id)").execute()
    items = results.get("files", [])
    if not items:
        return "No hay Excel aún."

    file_id = items[0]["id"]
    request = drive_service.files().get_media(fileId=file_id)
    buffer = BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    buffer.seek(0)

    return send_file(buffer, as_attachment=True, download_name="gastos.xlsx")

# -------------------------------
# Reset documentos
# -------------------------------
@app.route('/reset_documento', methods=['POST'])
def reset_documento():
    user_id = get_user_id()
    creds = get_credentials()
    if not isinstance(creds, Credentials):
        return creds
    drive_service = build_drive_service(creds)

    filenames = [f"{user_id}_transcripciones.docx", f"{user_id}_gastos.xlsx"]
    borrados = []

    for filename in filenames:
        query = f"name='{filename}'"
        if FOLDER_ID:
            query += f" and '{FOLDER_ID}' in parents"
        results = drive_service.files().list(q=query, fields="files(id)").execute()
        items = results.get("files", [])
        for item in items:
            drive_service.files().delete(fileId=item["id"]).execute()
            borrados.append(filename)

    if borrados:
        return f"Documentos reiniciados: {', '.join(set(borrados))}"
    else:
        return "No había documentos para reiniciar."

# -------------------------------
# Arranque
# -------------------------------
if __name__ == "__main__":
    # Para pruebas locales sin HTTPS:
    # export OAUTHLIB_INSECURE_TRANSPORT=1
    app.run(debug=True, host="0.0.0.0", port=5000)
