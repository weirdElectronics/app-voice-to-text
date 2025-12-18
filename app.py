
from flask import Flask, request, render_template, send_file, session, redirect, url_for
import speech_recognition as sr
from docx import Document
import os
import base64
import openpyxl
import re
import uuid
from io import BytesIO

# Librerías de Google Drive (OAuth)
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

# --- Autenticación OAuth ---
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

# --- user_id por sesión ---
def get_user_id():
    user_id = session.get("user_id")
    if not user_id:
        user_id = str(uuid.uuid4())
        session["user_id"] = user_id
    return user_id

# --- Guardar Word acumulativo ---
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

# --- Guardar Excel acumulativo ---
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
        # Si ya existe el Excel en Drive, lo descargamos
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

        # Agregar filas nuevas desde el workbook temporal
        for row in new_wb.active.iter_rows(values_only=True):
            ws.append(row)

        # Recalcular total correctamente
        filas = list(ws.iter_rows(values_only=True))
        montos = []
        for fila in filas[1:]:  # ignoramos encabezados
            if fila[0] and str(fila[0]).upper() == "TOTAL":
                continue  # ignoramos fila TOTAL existente
            if isinstance(fila[1], (int, float)):
                montos.append(fila[1])

        total = sum(montos)

        # Si la última fila ya es TOTAL, actualizamos su valor
        ultima_fila = filas[-1]
        if ultima_fila[0] and str(ultima_fila[0]).upper() == "TOTAL":
            ws.cell(row=len(filas), column=2, value=total)
        else:
            # Agregamos una nueva fila TOTAL al final
            ws.append(["TOTAL", total])

        # Guardar y subir actualización
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
        # Si no existe, creamos un nuevo Excel desde cero
        buffer = BytesIO()
        new_wb.save(buffer)
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


# --- Rutas base ---
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
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Gastos"
        ws.append(["Descripción", "Monto"])

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

        total = sum(cell.value for cell in ws["B"][1:] if isinstance(cell.value, (int, float)))
        ws["A1"] = "TOTAL"
        ws["B1"] = total

        save_excel_to_drive(user_id, wb)
        return f"Gasto registrado: {descripcion} (monto: {monto})"

# --- Ver documentos online ---
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

# --- Descargar documentos ---
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

# --- Resetear documentos del usuario ---
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

# --- Arranque ---
if __name__ == "__main__":
    # Para pruebas locales con HTTP:
    # export OAUTHLIB_INSECURE_TRANSPORT=1
    app.run(debug=True, host="0.0.0.0", port=5000)

