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
    """Procesamiento defensivo para formateo de fechas en español."""
    
    MESES = {
        1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
        5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
        9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
    }

    @classmethod
    def parse_fecha_fin(cls, fecha_str: str) -> tuple[str, str]:
        """
        Procesa fechas tipo 'dd/mm/yyyy' o 'dd de mmmm de yyyy'
        Retorna: (mes, año)
        """
        fecha_clean = str(fecha_str).strip().lower()
        
        # Formato numérico: dd/mm/yyyy o dd-mm-yyyy
        if "/" in fecha_clean or "-" in fecha_clean:
            delimitador = "/" if "/" in fecha_clean else "-"
            partes = fecha_clean.split(delimitador)
            if len(partes) == 3:
                try:
                    num_mes = int(partes[1])
                    anio = partes[2]
                    mes = cls.MESES.get(num_mes, "Enero")
                    return mes, anio
                except ValueError:
                    pass

        # Formato texto: 'dd de mmmm de yyyy'
        match = re.search(r'(\d+)\s+de\s+([a-zA-Záéíóúñ]+)\s+de\s+(\d{4})', fecha_clean)
        if match:
            mes = match.group(2).capitalize()
            anio = match.group(3)
            return mes, anio

        # Fallback de seguridad al año/mes actual si la cadena es inválida
        now = datetime.now()
        return cls.MESES[now.month], str(now.year)

# ==========================================
# 2. TRANSFORMADOR DE PLANTILLAS PPTX
# ==========================================
class PPTXTransformer:
    """Manipulación de presentaciones conservando estilos y fuentes a nivel de Run."""

    def __init__(self, template_path: str):
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"No se encontró la plantilla en: {template_path}")
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
        for paragraph in text_frame.paragraphs:
            for key, value in replacements.items():
                if key in paragraph.text:
                    self._replace_in_paragraph(paragraph, key, str(value))

    @staticmethod
    def _replace_in_paragraph(paragraph, key: str, value: str):
        """
        Reemplaza preservando el formato de fuente.
        Si la etiqueta está contenida dentro de un solo 'run', modifica solo ese run.
        Si la etiqueta fue dividida por PowerPoint entre varios 'runs', realiza el reemplazo en el párrafo.
        """
        for run in paragraph.runs:
            if key in run.text:
                run.text = run.text.replace(key, value)
                return
        
        # Fallback si el placeholder quedó fragmentado por el motor interno de PPTX
        paragraph.text = paragraph.text.replace(key, value)

# ==========================================
# 3. CONVERTIDOR HEADLESS (LIBREOFFICE)
# ==========================================
class PDFConverter:
    """Ejecución aislada de LibreOffice para exportación PDF."""

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
            raise RuntimeError(f"Fallo en la conversión a PDF vía LibreOffice: {result.stderr}")
        
        base_name = os.path.splitext(os.path.basename(input_pptx_path))[0]
        return os.path.join(output_dir, f"{base_name}.pdf")

# ==========================================
# 4. GESTOR API GOOGLE DRIVE
# ==========================================
class GoogleDriveManager:
    """Cliente SDK v3 para Google Drive."""

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
# 5. ORQUESTADOR DE PROCESO
# ==========================================
class DiplomaOrchestrator:
    """Orquestador principal del pipeline."""

    PARENT_DRIVE_ID = "1J7LU585mNaco3hD5lmKwuhXf9Eu5tmBC"
    TEMPLATE_PATH = "UCAL_Plantilla.pptx"

    def __init__(self, payload: dict, gdrive_json: str):
        self.payload = payload
        self.drive_manager = GoogleDriveManager(gdrive_json)
        self.transformer = PPTXTransformer(self.TEMPLATE_PATH)

    def run(self):
        # 1. Extracción y mapeo de variables generales
        programa_raw = self.payload.get("programa", "")
        fecha_fin_raw = self.payload.get("fecha_fin", "")
        total_horas = self.payload.get("total_horas", "")
        estudiantes = self.payload.get("estudiantes", [])

        # Lógica de mapeo condicional para el programa
        if str(programa_raw).strip() == "Coaching Profesional":
            programa_display = "COACHING PROFESIONAL: ACOMPAÑANDO LA TRANSFORMACIÓN"
        else:
            programa_display = programa_raw

        mes, anio = DateParser.parse_fecha_fin(fecha_fin_raw)

        # 2. Creación de la jerarquía en Google Drive: yymmdd_[Programa]
        prefix_date = datetime.now().strftime("%y%m%d")
        root_folder_name = f"{prefix_date}_{programa_raw}".strip()

        print(f"📁 Creando directorio raíz en Drive: {root_folder_name}")
        root_folder_id = self.drive_manager.create_folder(root_folder_name, self.PARENT_DRIVE_ID)
        editables_folder_id = self.drive_manager.create_folder("Editables", root_folder_id)
        finales_folder_id = self.drive_manager.create_folder("Finales", root_folder_id)

        # Directorios temporales de compilación
        os.makedirs("output/pptx", exist_ok=True)
        os.makedirs("output/pdf", exist_ok=True)

        # 3. Iteración sobre estudiantes
        for est in estudiantes:
            nombre = est.get("nombres_apellidos", "").strip()
            codigo = est.get("codigo_diploma", "").strip()

            print(f"🎓 Procesando: {nombre} | Código: {codigo}")

            replacements = {
                "{{nombre}}": nombre,
                "{{horas}}": total_horas,
                "{{mes}}": mes,
                "{{anio}}": anio,
                "{{codigo_diploma}}": codigo,
                "{{programa}}": programa_display
            }

            # Sanitización del nombre de archivo para evitar caracteres inválidos
            safe_filename = re.sub(r'[^\w\s-]', '', f"{codigo}_{nombre}").strip().replace(" ", "_")
            local_pptx = f"output/pptx/{safe_filename}.pptx"

            # Generar el editable .pptx
            self.transformer.generate_presentation(replacements, local_pptx)

            # Generar el final .pdf
            local_pdf = PDFConverter.convert_to_pdf(local_pptx, "output/pdf")

            # Cargar a las carpetas correspondientes en Google Drive
            print("  ⬆️ Subiendo editable (.pptx)...")
            self.drive_manager.upload_file(
                local_pptx, 
                editables_folder_id, 
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

            print("  ⬆️ Subiendo final (.pdf)...")
            self.drive_manager.upload_file(
                local_pdf, 
                finales_folder_id, 
                "application/pdf"
            )

        print("🚀 ¡Proceso finalizado exitosamente!")


if __name__ == "__main__":
    event_path = os.environ.get("GITHUB_EVENT_PATH")
    service_account_key = os.environ.get("GDRIVE_SERVICE_ACCOUNT_KEY")

    if not event_path or not os.path.exists(event_path):
        raise ValueError("No se detectó el archivo de evento de GitHub Actions.")
    
    if not service_account_key:
        raise ValueError("El Secret 'GDRIVE_SERVICE_ACCOUNT_KEY' no está configurado.")

    with open(event_path, "r", encoding="utf-8") as f:
        event_data = json.load(f)

    client_payload = event_data.get("client_payload", {})

    orchestrator = DiplomaOrchestrator(client_payload, service_account_key)
    orchestrator.run()
