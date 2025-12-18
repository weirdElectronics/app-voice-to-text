
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
from googleapiclient.http import MediaIoBaseUpload

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = "clave_supersecreta"  # Necesario para manejar sesiones

# Configuración Google Drive
SCOPES = ['https://www.googleapis.com/auth/drive.file']
CREDENTIALS_FILE = 'credentials.json'  # archivo OAuth descargado
FOLDER_ID = os.getenv("DRIVE_FOLDER_ID")  # carpeta AppData en tu Drive

# --- Funciones de autenticación OAuth ---
def get_credentials():
    # Si ya existe token.json y es válido
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        if creds and creds.valid:
            return creds

    # Si no existe token.json, iniciar flujo OAuth
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

    # Guardamos token para futuras ejecuciones
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return redirect(url_for("test_drive"))

def build_drive_service(creds):
    return build("drive", "v3", credentials=creds)

# --- Función para obtener user_id ---
def get_user_id():
    user_id = session.get("user_id")
    if not user_id:
        user_id = str(uuid.uuid4())
        session["user_id"] = user_id
    return user_id

# --- Funciones para subir a Drive ---
def upload_to_drive(user_id, file_bytes, filename, mime_type):
    creds = get_credentials()
    if not isinstance(creds, Credentials):
        return creds  # redirige a Google OAuth si falta autorización

    drive_service = build_drive_service(creds)
    file_metadata = {
        'name': f"{user_id}_{filename}"
    }
    if FOLDER_ID:
        file_metadata['parents'] = [FOLDER_ID]

    media = MediaIoBaseUpload(BytesIO(file_bytes), mimetype=mime_type, resumable=True)
    file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()
    return file.get('id')

def save_word_to_drive(user_id, doc):
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return upload_to_drive(user_id, buffer.getvalue(), "transcripciones.docx",
                           "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

def save_excel_to_drive(user_id, wb):
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return upload_to_drive(user_id, buffer.getvalue(), "gastos.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- Rutas ---
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
        p = doc.add_paragraph(texto)
        run = p.runs[0]
        run.font.name = "Courier New"

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

        total = sum(
            cell.value for cell in ws["B"][1:] if isinstance(cell.value, (int, float))
        )
        ws["A1"] = "TOTAL"
        ws["B1"] = total

        save_excel_to_drive(user_id, wb)
        return f"Gasto registrado: {descripcion} (monto: {monto})"

@app.route('/test_drive')
def test_drive():
    user_id = get_user_id()
    content = b"Hola Micaela, esto es una prueba."
    try:
        file_id = upload_to_drive(user_id, content, "prueba.txt", "text/plain")
        return f"Archivo subido a Drive con ID: {file_id}"
    except Exception as e:
        return f"Error al subir a Drive: {e}"

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)

