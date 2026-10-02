import os
import re
import json
import subprocess
from datetime import datetime
from pptx import Presentation
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ==========================================
# 1. PARSER DE FECHAS
# ==========================================
class DateParser:
    """Clase orientada al procesamiento y extracción de fechas en formatos en español."""
    
    MESES = {
        1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
        5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
        9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
    }

    @classmethod
    def parse_fecha_fin(cls, fecha_str: str) -> tuple[str, str]:
        """
        Procesa fechas tipo 'dd/mm/yyyy' o 'dd de mmmm de yyyy'
        Devuelve una tupla: (mes, año)
        """
        fecha_str = str(fecha_str).strip().lower()
        
        # Formato dd/mm/yyyy o dd-mm-yyyy
        if "/" in fecha_str or "-" in fecha_str:
            delimitador = "/" if "/" in fecha_str else "-"
            partes = fecha_str.split(delimitador)
            if len(partes) == 3:
                num_mes = int(partes[1])
                anio = partes[2]
                mes = cls.MESES.get(num_mes, "Enero")
                return mes, anio

        # Formato 'dd de mmmm de yyyy'
        match = re.search(r'(\d+)\s+de\s+([a-zA-Záéíóúñ]+)\s+de\s+(\d{4})', fecha_str)
        if match:
            mes = match.group(2).capitalize()
            anio = match.group(3)
            return mes, anio

        # Fallback defensivo si el formato no coincide
        fecha_actual = datetime.now()
        return cls.MESES[fecha_actual.month], str(fecha_actual.year)

# ==========================================
# 2. TRANSFORMADOR DE PLANTILLAS PPTX
# ==========================================
class PPTXTransformer:
    """Maneja la modificación de archivos PowerPoint preservando formatos."""

    def __init__(self, template_path: str):
        self.template_path = template_path

    def generate_presentation(self, replacements: dict, output_path: str) -> str:
        prs = Presentation(self.template_path)

        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    self._replace_in_text_frame(shape.text_frame, replacements)
                
                if shape.has_table:
                    for cell in shape.table.iter_cells():
                        if cell.text_frame:
                            self._replace_in_text_frame(cell.text_frame, replacements)

        prs.save(output_path)
        return output_path

    def _replace_in_text_frame(self, text_frame, replacements: dict):
        """Reemplaza placeholders buscando a nivel de párrafo y run para no perder fuentes/estilos."""
        for paragraph in text_frame.paragraphs:
            for key, value in replacements.items():
                if key in paragraph.text:
                    # Intenta reemplazo a nivel de Run para preservar formato exacto
                    replaced = False
                    for run in paragraph.runs:
                        if key in run.text:
                            run.text = run.text.replace(key, str(value))
                            replaced = True
                    
                    # Fallback por si el placeholder quedó dividido entre varios runs
                    if not replaced:
                        paragraph.text = paragraph.text.replace(key, str(value))

# ==========================================
# 3. CONVERTIDOR DE PPTX A PDF
# ==========================================
class PDFConverter:
    """Maneja la conversión de archivos de Office a PDF mediante LibreOffice Headless."""

    @staticmethod
    def convert_to_pdf(input_pptx_path: str, output_dir: str) -> str:
        command = [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            input_pptx_path
        ]
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        
        if result.returncode != 0:
            raise RuntimeError(f"Error al convertir a PDF mediante LibreOffice: {result.stderr}")
        
        base_name = os.path.splitext(os.path.basename(input_pptx_path))[0]
        return os.path.join(output_dir, f"{base_name}.pdf")

# ==========================================
# 4. GESTOR DE GOOGLE DRIVE API
# ==========================================
class GoogleDriveManager:
    """Administra la creación de directorios y subida de archivos en Google Drive."""

    SCOPES = ['https://www.googleapis.com/auth/drive']

    def __init__(self, service_account_json_str: str):
        info = json.loads(service_account_json_str)
        creds = Credentials.from_service_account_info(info, scopes=self.SCOPES)
        self.service = build('drive', 'v3', credentials=creds)

    def create_folder(self, folder_name: str, parent_id: str = None) -> str:
        metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        if parent_id:
            metadata['parents'] = [parent_id]

        folder = self.service.files().create(body=metadata, fields='id').execute()
        return folder.get('id')

    def upload_file(self, file_path: str, target_folder_id: str, mime_type: str) -> str:
        file_name = os.path.basename(file_path)
        metadata = {
            'name': file_name,
            'parents': [target_folder_id]
        }
        media = MediaFileUpload(file_path, mimetype=mime_type, resumable=True)
        uploaded = self.service.files().create(body=metadata, media_body=media, fields='id').execute()
        return uploaded.get('id')

# ==========================================
# 5. ORQUESTADOR DEL PROCESO
# ==========================================
class DiplomaOrchestrator:
    """Orquesta todo el flujo end-to-end recibiendo el payload cargado."""

    PARENT_DRIVE_ID = "1J7LU585mNaco3hD5lmKwuhXf9Eu5tmBC"
    TEMPLATE_PATH = "UCAL_Plantilla.pptx"

    def __init__(self, payload: dict, gdrive_json: str):
        self.payload = payload
        self.drive_manager = GoogleDriveManager(gdrive_json)
        self.transformer = PPTXTransformer(self.TEMPLATE_PATH)

    def run(self):
        # 1. Extraer datos generales
        programa_raw = self.payload.get("programa", "")
        fecha_fin_raw = self.payload.get("fecha_fin", "")
        total_horas = self.payload.get("total_horas", "")
        estudiantes = self.payload.get("estudiantes", [])

        # Transformación condicional del Nombre del Programa
        if str(programa_raw).strip() == "Coaching Profesional":
            programa_display = "COACHING PROFESIONAL: ACOMPAÑANDO LA TRANSFORMACIÓN"
        else:
            programa_display = programa_raw

        # Procesar Mes y Año
        mes, anio = DateParser.parse_fecha_fin(fecha_fin_raw)

        # 2. Configurar la estructura de carpetas en Google Drive
        # Formato raíz: YYMMDD_[Programa] (ejemplo: 261002_Coaching Profesional)
        prefix_date = datetime.now().strftime("%y%m%d")
        root_folder_name = f"{prefix_date}_{programa_raw}".strip()

        print(f"📁 Creando estructura de carpetas en Drive: {root_folder_name}")
        root_folder_id = self.drive_manager.create_folder(root_folder_name, self.PARENT_DRIVE_ID)
        editables_folder_id = self.drive_manager.create_folder("Editables", root_folder_id)
        finales_folder_id = self.drive_manager.create_folder("Finales", root_folder_id)

        # Directorios temporales en local
        os.makedirs("output/pptx", exist_ok=True)
        os.makedirs("output/pdf", exist_ok=True)

        # 3. Procesar cada estudiante
        for est in estudiantes:
            nombre = est.get("nombres_apellidos", "").strip()
            codigo = est.get("codigo_diploma", "").strip()

            print(f"🎓 Procesando Diploma: {nombre} ({codigo})")

            replacements = {
                "{{nombre}}": nombre,
                "{{horas}}": total_horas,
                "{{mes}}": mes,
                "{{anio}}": anio,
                "{{codigo_diploma}}": codigo,
                "{{programa}}": programa_display
            }

            # Nombres de archivo limpitos de caracteres extraños
            safe_filename = re.sub(r'[^\w\s-]', '', f"{codigo}_{nombre}").replace(" ", "_")
            local_pptx = f"output/pptx/{safe_filename}.pptx"

            # Reemplazar datos en la plantilla PPTX
            self.transformer.generate_presentation(replacements, local_pptx)

            # Convertir PPTX a PDF usando LibreOffice Headless
            local_pdf = PDFConverter.convert_to_pdf(local_pptx, "output/pdf")

            # Subir a Google Drive
            print(f"  ⬆️ Subiendo Editable (.pptx)...")
            self.drive_manager.upload_file(
                local_pptx, 
                editables_folder_id, 
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

            print(f"  ⬆️ Subiendo Final (.pdf)...")
            self.drive_manager.upload_file(
                local_pdf, 
                finales_folder_id, 
                "application/pdf"
            )

        print("🚀 ¡Proceso completado con éxito!")


if __name__ == "__main__":
    # GitHub Actions almacena el evento del dispatch en la ruta apuntada por GITHUB_EVENT_PATH
    event_path = os.environ.get("GITHUB_EVENT_PATH")
    service_account_key = os.environ.get("GDRIVE_SERVICE_ACCOUNT_KEY")

    if not event_path or not os.path.exists(event_path):
        raise ValueError("No se encontró el archivo del evento de GitHub.")
    
    if not service_account_key:
        raise ValueError("Falta el secret 'GDRIVE_SERVICE_ACCOUNT_KEY' en GitHub.")

    with open(event_path, "r", encoding="utf-8") as f:
        event_data = json.load(f)

    # Extraemos el payload que enviamos desde Google Apps Script
    client_payload = event_data.get("client_payload", {})

    orchestrator = DiplomaOrchestrator(client_payload, service_account_key)
    orchestrator.run()
