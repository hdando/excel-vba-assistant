from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List, Dict, Any, Optional
import asyncio
import aiofiles
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import json
import uuid
import os
import google.generativeai as genai
import tempfile
import shutil
from datetime import datetime
from dotenv import load_dotenv
import re

# Charger les variables d'environnement
load_dotenv()

# Configuration
UPLOAD_MAX_SIZE = 50 * 1024 * 1024  # 50MB
ALLOWED_EXTENSIONS = {'.xlsx', '.xlsm'}
SESSION_TIMEOUT = 7200  # 2 hours
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

# Initialize Gemini
if not GEMINI_API_KEY:
    raise Exception("❌ ERREUR CRITIQUE : Clé API Gemini manquante ! Ajoutez GEMINI_API_KEY dans le fichier .env")

genai.configure(api_key=GEMINI_API_KEY)

# Utiliser le bon modèle Gemini - essayer dans l'ordre
model = None
models_to_try = ['gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-pro']

for model_name in models_to_try:
    try:
        model = genai.GenerativeModel(model_name)
        # Tester que le modèle fonctionne
        test_response = model.generate_content("Test")
        print(f"✅ Modèle {model_name} initialisé avec succès")
        break
    except Exception as e:
        print(f"⚠️ Le modèle {model_name} n'est pas disponible : {str(e)[:100]}")
        continue

if not model:
    raise Exception("❌ Impossible d'initialiser un modèle Gemini. Vérifiez votre clé API.")

# Stockage en mémoire (remplace Redis)
sessions_store = {}

# Note: requirements.txt devrait contenir:
# fastapi==0.104.1
# uvicorn[standard]==0.24.0
# python-multipart==0.0.6
# aiofiles==23.2.1
# pandas==2.1.3
# openpyxl==3.1.2
# google-generativeai==0.3.0
# python-dotenv==1.0.0

app = FastAPI()

# CORS configuration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Models
class ChatMessage(BaseModel):
    session_id: str
    message: str
    context: Optional[Dict[str, Any]] = None

class AnalyzeRequest(BaseModel):
    session_id: str
    module_name: Optional[str] = None
    sheet_name: Optional[str] = None

class CellUpdate(BaseModel):
    session_id: str
    sheet_name: str
    row: int
    col: int
    value: str

# Helper functions
async def get_session(session_id: str):
    return sessions_store.get(session_id)

async def save_session(session_id: str, data: dict):
    sessions_store[session_id] = data

def extract_vba_code(workbook_path: str) -> Dict[str, str]:
    """Extract VBA code from Excel file"""
    vba_modules = {}
    try:
        # Simulation pour le moment
        vba_modules = {
            "Module1": "Sub Example()\n    ' Sample VBA code\n    MsgBox \"Hello\"\nEnd Sub",
            "ThisWorkbook": "Private Sub Workbook_Open()\n    ' Workbook open event\nEnd Sub"
        }
    except Exception as e:
        print(f"Error extracting VBA: {e}")
    return vba_modules

def analyze_excel_structure(workbook_path: str) -> Dict[str, Any]:
    """Analyze Excel file structure"""
    wb = load_workbook(workbook_path, data_only=True)
    
    structure = {
        "sheets": [],
        "total_sheets": len(wb.sheetnames),
        "has_vba": workbook_path.endswith('.xlsm'),
        "file_size": os.path.getsize(workbook_path),
        "created": datetime.now().isoformat()
    }
    
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        sheet_info = {
            "name": sheet_name,
            "max_row": sheet.max_row,
            "max_column": sheet.max_column,
            "has_data": sheet.max_row > 1 or sheet.max_column > 1,
            "formulas": [],
            "charts": [],
            "tables": [],
            "data": []
        }
        
        # Get headers (first row)
        headers = []
        for cell in sheet[1]:
            headers.append(str(cell.value) if cell.value else "")
        
        # Get data (first 100 rows)
        sheet_data = []
        for row in sheet.iter_rows(min_row=1, max_row=min(100, sheet.max_row), values_only=True):
            sheet_data.append([str(cell) if cell is not None else "" for cell in row])
        
        sheet_info["headers"] = headers
        sheet_info["data"] = sheet_data
        
        # Check for formulas
        for row in sheet.iter_rows(max_row=min(100, sheet.max_row)):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    sheet_info["formulas"].append({
                        "cell": cell.coordinate,
                        "formula": cell.value
                    })
        
        structure["sheets"].append(sheet_info)
    
    wb.close()
    return structure

async def generate_initial_analysis(filename: str, structure: Dict[str, Any], vba_modules: Dict[str, str]) -> str:
    """Generate initial analysis using Gemini"""
    prompt = f"""
    Tu es un expert Excel qui analyse des fichiers pour des utilisateurs métiers.
    Analyse ce fichier Excel et fournis un résumé clair et utile en français.
    
    Nom du fichier : {filename}
    Nombre de feuilles : {structure['total_sheets']}
    Contient du VBA : {'Oui' if structure['has_vba'] else 'Non'}
    
    Détails des feuilles :
    {json.dumps(structure['sheets'], indent=2)}
    
    Modules VBA trouvés :
    {list(vba_modules.keys()) if vba_modules else 'Aucun'}
    
    Fournis :
    1. Un aperçu rapide du fichier et son utilité probable
    2. Les points clés sur la structure des données
    3. Les problèmes potentiels ou améliorations possibles
    4. Des suggestions d'actions pour l'utilisateur
    
    Sois concis, clair et utilise un langage accessible. Utilise des émojis pour rendre le texte plus agréable.
    """
    
    response = await model.generate_content_async(prompt)
    return response.text

# API Endpoints
@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload and analyze Excel file"""
    # Validate file
    if not any(file.filename.endswith(ext) for ext in ALLOWED_EXTENSIONS):
        raise HTTPException(400, "Invalid file type. Only .xlsx and .xlsm files are allowed.")
    
    # Check file size
    contents = await file.read()
    if len(contents) > UPLOAD_MAX_SIZE:
        raise HTTPException(400, f"File too large. Maximum size is {UPLOAD_MAX_SIZE // 1024 // 1024}MB")
    
    # Generate session ID
    session_id = str(uuid.uuid4())
    
    # Save file temporarily
    temp_dir = tempfile.mkdtemp()
    file_path = os.path.join(temp_dir, file.filename)
    
    async with aiofiles.open(file_path, 'wb') as f:
        await f.write(contents)
    
    # Analyze file structure
    structure = analyze_excel_structure(file_path)
    vba_modules = extract_vba_code(file_path) if file.filename.endswith('.xlsm') else {}
    
    # Generate initial analysis
    initial_analysis = await generate_initial_analysis(file.filename, structure, vba_modules)
    
    # Store session data
    session_data = {
        "filename": file.filename,
        "file_path": file_path,
        "structure": structure,
        "vba_modules": vba_modules,
        "created": datetime.now().isoformat(),
        "chat_history": []
    }
    
    await save_session(session_id, session_data)
    
    return {
        "session_id": session_id,
        "filename": file.filename,
        "structure": structure,  # Ceci contient maintenant les données
        "vba_modules": list(vba_modules.keys()),
        "initial_analysis": initial_analysis
    }

@app.post("/api/chat")
async def chat_with_agent(message: ChatMessage):
    """Chat endpoint for conversational interactions"""
    session_data = await get_session(message.session_id)
    
    if not session_data:
        raise HTTPException(404, "Session not found or expired")
    
    # Build context for Gemini with function calling
    # Analyser le message pour détecter une demande de modification
    message_lower = message.message.lower()
    # Liste étendue de mots-clés pour détecter les modifications
    modification_keywords = [
        'écris', 'ecris', 'écrire', 'ecrire',
        'mets', 'mettre', 'met',
        'change', 'changer', 'modifie', 'modifier',
        'insère', 'insere', 'insérer', 'inserer',
        'ajoute', 'ajouter',
        'remplace', 'remplacer',
        'saisis', 'saisir',
        'entre', 'entrer',
        'place', 'placer',
        'définis', 'definir', 'défini'
    ]
    
    # Vérifier aussi les patterns comme "dans [cellule]" ou "en [cellule]"
    has_cell_reference = bool(re.search(r'\b[A-Z]+\d+\b', message.message))
    is_modification = any(word in message_lower for word in modification_keywords) and has_cell_reference
    
    if is_modification:
        # Prompt spécifique pour les modifications
        system_prompt = f"""
Tu es un assistant Excel. L'utilisateur veut modifier une cellule.
Réponds UNIQUEMENT avec ce format JSON, rien d'autre :

{{
    "action": "update_cell",
    "sheet": "Feuil1",
    "cell": "[LA_CELLULE]",
    "value": "[LA_VALEUR]",
    "message": "✅ J'ai modifié la cellule [LA_CELLULE] avec la valeur '[LA_VALEUR]'. La modification est sauvegardée."
}}

Exemples :
- "écris 18 dans AC1" → {{"action": "update_cell", "sheet": "Feuil1", "cell": "AC1", "value": "18", "message": "✅ J'ai écrit 18 dans la cellule AC1. La modification est sauvegardée."}}
- "mets 500 en B3" → {{"action": "update_cell", "sheet": "Feuil1", "cell": "B3", "value": "500", "message": "✅ J'ai mis 500 dans la cellule B3. La modification est sauvegardée."}}
- "change A1 pour Titre" → {{"action": "update_cell", "sheet": "Feuil1", "cell": "A1", "value": "Titre", "message": "✅ J'ai changé A1 pour 'Titre'. La modification est sauvegardée."}}

Message de l'utilisateur : {message.message}

IMPORTANT : Réponds UNIQUEMENT avec le JSON, pas de texte avant ou après !
"""
    else:
        # Prompt normal pour les autres demandes
        system_prompt = f"""
Tu es un expert Excel/VBA qui aide des utilisateurs métiers en français.
Contexte : Fichier Excel "{session_data['filename']}" avec {len(session_data['structure']['sheets'])} feuilles.

Structure du fichier :
- Feuilles : {[sheet['name'] for sheet in session_data['structure']['sheets']]}
- Taille : {session_data['structure']['sheets'][0]['max_row']} lignes × {session_data['structure']['sheets'][0]['max_column']} colonnes

Tu peux analyser, expliquer, et donner des conseils sur le fichier Excel.
Utilise des émojis pour rendre la conversation agréable.

Message de l'utilisateur : {message.message}
"""
    
    # Generate response with Gemini
    response = await model.generate_content_async(system_prompt)
    response_text = response.text.strip()
    
    print(f"[DEBUG] Réponse brute de Gemini : {response_text[:200]}...")
    
    # Vérifier si la réponse contient une action
    try:
        # Chercher un JSON valide dans la réponse
        json_match = re.search(r'\{[^{}]*"action"[^{}]*\}', response_text, re.DOTALL)
        if json_match:
            print(f"[DEBUG] JSON trouvé : {json_match.group()}")
            action_data = json.loads(json_match.group())
            
            if action_data.get('action') == 'update_cell':
                print(f"[DEBUG] Action update_cell détectée")
                # Effectuer la modification
                sheet_name = action_data.get('sheet', 'Feuil1')
                cell_ref = action_data.get('cell', '').upper()
                value = action_data.get('value', '')
                
                print(f"[DEBUG] Modification : {sheet_name} / {cell_ref} = {value}")
                
                # Convertir la référence de cellule (ex: AC1) en indices
                col_letters = ''.join(filter(str.isalpha, cell_ref))
                row_num = int(''.join(filter(str.isdigit, cell_ref))) - 1
                
                # Convertir les lettres en index de colonne (A=0, B=1, ..., Z=25, AA=26, AB=27, ...)
                col_num = 0
                for char in col_letters:
                    col_num = col_num * 26 + (ord(char.upper()) - ord('A') + 1)
                col_num -= 1  # Ajuster pour index 0
                
                print(f"[DEBUG] Indices : ligne {row_num}, colonne {col_num}")
                
                # Mettre à jour la cellule
                if 0 <= row_num < len(session_data['structure']['sheets'][0]['data']) and \
                   0 <= col_num < len(session_data['structure']['sheets'][0]['data'][0]):
                    
                    # Mettre à jour en mémoire
                    session_data['structure']['sheets'][0]['data'][row_num][col_num] = str(value)
                    
                    # Mettre à jour le fichier Excel
                    wb = load_workbook(session_data['file_path'])
                    ws = wb[sheet_name]
                    ws[cell_ref] = value
                    wb.save(session_data['file_path'])
                    wb.close()
                    
                    print(f"[DEBUG] Modification effectuée avec succès")
                    
                    # Sauvegarder la session
                    await save_session(message.session_id, session_data)
                    
                    # Utiliser le message personnalisé de l'IA
                    response_text = action_data.get('message', f"✅ J'ai modifié la cellule {cell_ref} avec la valeur '{value}'")
                else:
                    response_text = f"❌ Erreur : La cellule {cell_ref} est hors limites du tableau (max: {len(session_data['structure']['sheets'][0]['data'][0])} colonnes, {len(session_data['structure']['sheets'][0]['data'])} lignes)."
    except Exception as e:
        print(f"[DEBUG] Pas d'action détectée ou erreur : {e}")
        # Si ce n'est pas une action, utiliser la réponse normale
        pass
    
    # Update chat history
    session_data['chat_history'].append({
        "user": message.message,
        "assistant": response_text,
        "timestamp": datetime.now().isoformat()
    })
    
    # Save updated session
    await save_session(message.session_id, session_data)
    
    # Stream response
    async def generate():
        for chunk in response_text.split(' '):
            yield f"data: {json.dumps({'chunk': chunk + ' '})}\n\n"
            await asyncio.sleep(0.01)
        yield f"data: {json.dumps({'done': True})}\n\n"
    
    return StreamingResponse(generate(), media_type="text/event-stream")

@app.post("/api/update-cell")
async def update_cell(update: CellUpdate):
    """Update a single cell value"""
    session_data = await get_session(update.session_id)
    
    if not session_data:
        raise HTTPException(404, "Session not found or expired")
    
    try:
        # Update in memory
        for sheet in session_data['structure']['sheets']:
            if sheet['name'] == update.sheet_name:
                if update.row < len(sheet['data']) and update.col < len(sheet['data'][update.row]):
                    sheet['data'][update.row][update.col] = update.value
                    break
        
        # Update the actual Excel file
        wb = load_workbook(session_data['file_path'])
        ws = wb[update.sheet_name]
        # Excel uses 1-based indexing
        ws.cell(row=update.row + 1, column=update.col + 1, value=update.value)
        wb.save(session_data['file_path'])
        wb.close()
        
        # Save updated session
        await save_session(update.session_id, session_data)
        
        return {"status": "success", "message": "Cell updated"}
    except Exception as e:
        raise HTTPException(500, f"Error updating cell: {str(e)}")

@app.post("/api/export")
async def export_file(session_id: str):
    """Export modified Excel file"""
    session_data = await get_session(session_id)
    
    if not session_data:
        raise HTTPException(404, "Session not found or expired")
    
    file_path = session_data['file_path']
    
    if not os.path.exists(file_path):
        raise HTTPException(404, "File not found")
    
    async def iterfile():
        async with aiofiles.open(file_path, 'rb') as f:
            while chunk := await f.read(8192):
                yield chunk
    
    return StreamingResponse(
        iterfile(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename=modified_{session_data['filename']}"
        }
    )

@app.get("/api/session/{session_id}")
async def get_session_data(session_id: str):
    """Get session data"""
    session_data = await get_session(session_id)
    
    if not session_data:
        raise HTTPException(404, "Session not found or expired")
    
    return session_data

@app.get("/")
async def root():
    return {
        "message": "Excel VBA Agent API is running!",
        "gemini_status": "✅ Gemini AI activé" if model else "❌ Gemini AI non disponible"
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)