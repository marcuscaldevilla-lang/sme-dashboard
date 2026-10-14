import os, base64, urllib.parse, csv, io, logging, secrets
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta, timezone
from typing import List, Dict, Any, Optional, Tuple
import streamlit as st
import requests
import psycopg2
from psycopg2.extras import RealDictCursor
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

st.set_page_config(page_title="Qween | Acompanhamento da Execução", page_icon="📊", layout="wide", initial_sidebar_state="collapsed")

ADO_ORG = "SME-Spassu"
ADO_BASE_URL = f"https://dev.azure.com/{ADO_ORG}"
ALL_WI_TYPES = ["Task", "Product Backlog Item", "Melhoria", "HotFix", "Design", "Bug", "BugFix", "Débito Técnico"]
ALL_STATES = ["To Do", "In Progress", "Resolved", "Done"]
DEFAULT_WI_TYPES = [t for t in ALL_WI_TYPES if t not in ["Task", "Débito Técnico"]]

SPASSU_LOGO_URL = "https://raw.githubusercontent.com/marcuscaldevilla-lang/sme-dashboard/main/spassu4.png"
SME_LOGO_URL = "https://raw.githubusercontent.com/marcuscaldevilla-lang/sme-dashboard/main/sme.png"

def inject_custom_css():
    css = """
    <style>
        #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
        .stApp { background: radial-gradient(circle at 15% 45%, #ff5d2b 0%, #f24e1e 50%, #df3a09 100%) !important; font-family: 'Montserrat', sans-serif; }
        .dotted-rings { position: fixed; top: 50%; left: -150px; transform: translateY(-50%); width: 900px; height: 900px; pointer-events: none; z-index: 0; }
        .ring { position: absolute; top: 50%; left: 50%; border-radius: 50%; border: 3px dotted rgba(255, 255, 255, 0.85); transform: translate(-50%, -50%); }
        .ring-1 { width: 300px; height: 300px; } .ring-2 { width: 440px; height: 440px; border-width: 4px; }
        .ring-3 { width: 580px; height: 580px; border-width: 4px; } .ring-4 { width: 720px; height: 720px; border-width: 5px; opacity: 0.9; }
        .ring-5 { width: 880px; height: 880px; border-width: 5px; opacity: 0.7; }
        .stTextInput > div > div > input, .stNumberInput > div > div > input, .stSelectbox > div > div > div > input { background-color: rgba(255, 255, 255, 0.95) !important; color: #222222 !important; border-radius: 10px !important; height: 48px !important; font-weight: 500 !important; }
        .stButton > button { width: 100%; height: 50px; background-color: #ffffff !important; color: #f24e1e !important; font-weight: 700 !important; border-radius: 10px !important; border: none !important; }
        .hero-title { color: #ffffff; font-size: 3.2rem; font-weight: 800; line-height: 1.15; margin-bottom: 20px; text-shadow: 0 4px 14px rgba(0, 0, 0, 0.15); }
        .hero-subtitle { color: #e6e6e6; font-size: 1.25rem; font-weight: 500; }
        .badge-qween { display: inline-block; background: rgba(255, 255, 255, 0.22); backdrop-filter: blur(8px); padding: 6px 16px; border-radius: 30px; font-size: 0.85rem; font-weight: 700; letter-spacing: 1.5px; text-transform: uppercase; margin-bottom: 20px; color: #ffffff; border: 1px solid rgba(255, 255, 255, 0.35); }
        .auth-card { background: rgba(255, 255, 255, 0.12); backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px); border: 1px solid rgba(255, 255, 255, 0.25); border-radius: 20px; padding: 40px; box-shadow: 0 20px 45px rgba(0, 0, 0, 0.18); color: #ffffff; }
        .token-box { background: rgba(0,0,0,0.4); padding: 15px; border-radius: 8px; font-family: monospace; font-size: 1rem; word-break: break-all; margin: 15px 0; border: 2px solid rgba(255,255,255,0.5); text-align: center; color: #fff; }
    </style>
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    """
    st.markdown(css, unsafe_allow_html=True)

def authenticate_user(email: str, password: str) -> Optional[Dict]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT id, nome, email, tipo, apikey, status FROM usuarios WHERE email = %s AND senha_hash = %s LIMIT 1", (email.lower().strip(), password))
            user = cur.fetchone()
        conn.close()
        return user
    except Exception as e:
        logger.error(f"Erro na autenticação: {e}")
        return None

def criar_usuario_com_apikey(nome: str, email: str, senha: str, apikey: str) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM usuarios WHERE email = %s", (email.lower().strip(),))
            if cur.fetchone():
                conn.close()
                return False, "Este e-mail já está cadastrado."
            cur.execute("INSERT INTO usuarios (nome, email, senha_hash, apikey, tipo, status) VALUES (%s, %s, %s, %s, 'User', 'Ativo')", (nome.strip(), email.lower().strip(), senha, apikey))
            conn.commit()
        conn.close()
        return True, "Conta criada com sucesso! Faça login com seu e-mail e senha."
    except psycopg2.IntegrityError as e:
        if conn: conn.rollback()
        if conn: conn.close()
        return False, f"Erro de integridade: {str(e)}"
    except Exception as e:
        if conn: conn.close()
        return False, f"Erro ao criar conta: {str(e)}"

def solicitar_reset_senha(email: str) -> Tuple[bool, str, Optional[str]]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado.", None
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT id, nome FROM usuarios WHERE email = %s", (email.lower().strip(),))
            usuario = cur.fetchone()
            if not usuario:
                conn.close()
                return False, "E-mail não cadastrado no sistema.", None
            token = secrets.token_urlsafe(32)
            expires = datetime.now(timezone.utc) + timedelta(hours=1)
            cur.execute("UPDATE usuarios SET reset_token = %s, reset_token_expires = %s WHERE id = %s", (token, expires, usuario['id']))
            conn.commit()
        conn.close()
        return True, f"Token de recuperação gerado para {usuario['nome']}.", token
    except Exception as e:
        if conn: conn.close()
        return False, f"Erro ao processar solicitação: {str(e)}", None

def redefinir_senha(token: str, nova_senha: str) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT id, reset_token, reset_token_expires FROM usuarios WHERE reset_token = %s", (token.strip(),))
            usuario = cur.fetchone()
            if not usuario:
                conn.close()
                return False, "Token inválido ou expirado."
            if usuario['reset_token_expires'] < datetime.now(timezone.utc):
                conn.close()
                return False, "Token expirado. Solicite um novo."
            cur.execute("UPDATE usuarios SET senha_hash = %s, reset_token = NULL, reset_token_expires = NULL WHERE id = %s", (nova_senha, usuario['id']))
            conn.commit()
        conn.close()
        return True, "Senha redefinida com sucesso! Faça login com a nova senha."
    except Exception as e:
        if conn: conn.close()
        return False, f"Erro ao redefinir senha: {str(e)}"

def get_active_projects_from_db() -> List[str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT name FROM devops_projects WHERE is_active = TRUE ORDER BY name")
            projects = [row['name'] for row in cur.fetchall()]
        conn.close()
        return projects
    except: return []

def get_colaboradores_from_db() -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url:
        logger.error("NEON_DATABASE_URL não configurado")
        return pd.DataFrame(columns=["colaborador_azure", "colaborador_ponto", "cargo", "nivel", "data_inicio", "promocao", "reposicao", "data_saida"])
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT colaborador_azure, colaborador_ponto, cargo, nivel, 
                       data_inicio, promocao, reposicao, data_saida
                FROM colaboradores 
                ORDER BY colaborador_ponto
            """)
            rows = cur.fetchall()
            logger.info(f"get_colaboradores_from_db: {len(rows)} colaboradores encontrados")
            df = pd.DataFrame(rows)
        conn.close()
        return df
    except Exception as e:
        logger.error(f"Erro ao buscar colaboradores: {e}")
        return pd.DataFrame(columns=["colaborador_azure", "colaborador_ponto", "cargo", "nivel", "data_inicio", "promocao", "reposicao", "data_saida"])

def get_horasponto_from_db(ano: int, mes: int) -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT colaborador_ponto, horas_ponto, afastamento FROM horasponto WHERE ano_base = %s AND mes_base = %s", (ano, mes))
            df = pd.DataFrame(cur.fetchall())
        conn.close()
        return df
    except: return pd.DataFrame()

def get_dias_uteis_from_db(ano: int, mes: int) -> float:
    db_url = os.getenv("NEON_DATABASE_URL")
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT horas_uteis FROM dias_uteis_mensais WHERE ano = %s AND mes = %s LIMIT 1", (ano, mes))
            row = cur.fetchone()
        conn.close()
        return float(row['horas_uteis']) if row else 0.0
    except: return 0.0

def get_valores_from_db() -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT nivel, valor FROM valores")
            df = pd.DataFrame(cur.fetchall())
        conn.close()
        return df
    except: return pd.DataFrame(columns=["nivel", "valor"])

def get_capacity_by_nivel(nivel: str) -> float:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url or not nivel:
        return 0.0
    nivel_limpo = "".join(nivel.split())
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT capacity FROM capacity 
                WHERE REPLACE(REPLACE(REPLACE(TRIM(CAST(nivel AS VARCHAR)), E'\\n', ''), E'\\r', ''), E'\\t', '') = %s
                LIMIT 1
            """, (nivel_limpo,))
            row = cur.fetchone()
        conn.close()
        if row:
            return float(row['capacity'])
        else:
            logger.warning(f"Nível '{nivel}' não encontrado na tabela capacity")
            return 0.0
    except Exception as e:
        logger.error(f"Erro ao buscar capacity para nível '{nivel}': {e}")
        return 0.0

def limpar_valor_monetario(valor) -> float:
    if valor is None or pd.isna(valor):
        return 0.0
    valor_str = str(valor).strip()
    valor_str = valor_str.replace('$', '').replace('R$', '').replace(' ', '')
    if ',' in valor_str and '.' in valor_str:
        ultima_virgula = valor_str.rfind(',')
        ultimo_ponto = valor_str.rfind('.')
        if ultima_virgula > ultimo_ponto:
            valor_str = valor_str.replace('.', '').replace(',', '.')
        else:
            valor_str = valor_str.replace(',', '')
    elif ',' in valor_str:
        valor_str = valor_str.replace(',', '.')
    try:
        return float(valor_str)
    except (ValueError, TypeError):
        logger.warning(f"Não foi possível converter valor: '{valor}'")
        return 0.0

def listar_usuarios(tipo_usuario: str, usuario_id: int) -> List[Dict]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return []
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            if tipo_usuario == 'Admin':
                cur.execute("SELECT id, nome, email, tipo, status, apikey, criado_em FROM usuarios ORDER BY nome")
            else:
                cur.execute("SELECT id, nome, email, tipo, status, apikey, criado_em FROM usuarios WHERE id = %s", (usuario_id,))
            usuarios = cur.fetchall()
        conn.close()
        return usuarios
    except: return []

def atualizar_usuario(usuario_id: int, dados: Dict, tipo_usuario_logado: str, usuario_logado_id: int) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            if tipo_usuario_logado != 'Admin' and usuario_id != usuario_logado_id:
                conn.close()
                return False, "Você não tem permissão para editar este usuário."
            campos_permitidos = ['nome', 'email', 'senha_hash', 'apikey', 'tipo', 'status'] if tipo_usuario_logado == 'Admin' else ['nome', 'email', 'senha_hash', 'apikey']
            campos_atualizar = {k: v for k, v in dados.items() if k in campos_permitidos and v is not None}
            if not campos_atualizar:
                conn.close()
                return False, "Nenhum campo válido para atualizar."
            set_clause = ", ".join([f"{campo} = %s" for campo in campos_atualizar.keys()])
            valores = list(campos_atualizar.values()) + [usuario_id]
            cur.execute(f"UPDATE usuarios SET {set_clause} WHERE id = %s", valores)
            if cur.rowcount == 0:
                conn.close()
                return False, "Usuário não encontrado."
            conn.commit()
        conn.close()
        return True, "Dados atualizados com sucesso!"
    except Exception as e:
        if conn: conn.rollback()
        if conn: conn.close()
        return False, f"Erro ao atualizar: {str(e)}"

def deletar_usuario(usuario_id: int, tipo_usuario_logado: str) -> Tuple[bool, str]:
    if tipo_usuario_logado != 'Admin':
        return False, "Apenas administradores podem deletar usuários."
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM usuarios WHERE id = %s", (usuario_id,))
            if cur.rowcount == 0:
                conn.close()
                return False, "Usuário não encontrado."
            conn.commit()
        conn.close()
        return True, "Usuário deletado com sucesso!"
    except Exception as e:
        if conn: conn.rollback()
        if conn: conn.close()
        return False, f"Erro ao deletar: {str(e)}"

def get_ado_headers(pat: str) -> Dict[str, str]:
    auth_string = f":{pat}"
    return {"Authorization": f"Basic {base64.b64encode(auth_string.encode('ascii')).decode('ascii')}", "Content-Type": "application/json"}

def validate_ado_api_key(pat: str) -> Tuple[bool, str]:
    if not pat: return False, "API Key vazia."
    url = f"{ADO_BASE_URL}/_apis/projects?api-version=7.1&$top=1"
    try:
        response = requests.get(url, headers=get_ado_headers(pat), timeout=15)
        if response.status_code == 401: return False, "API Azure inválida: autenticação não autorizada (401)."
        elif response.status_code == 403: return False, "API Azure inválida: acesso proibido (403)."
        elif response.status_code == 404: return False, "API Azure inválida: organização não encontrada (404)."
        elif response.status_code != 200: return False, f"API Azure inválida: erro {response.status_code}."
        return True, "API Key válida."
    except requests.exceptions.ConnectionError:
        return False, "API Azure inválida: não foi possível conectar."
    except Exception as e:
        return False, f"API Azure inválida: {str(e)}"

def parse_date(date_str: str) -> Optional[datetime]:
    if not date_str: return None
    try: return datetime.fromisoformat(date_str.replace('Z', ''))
    except: return None

def fetch_all_sprints_for_project(pat: str, project: str, year: int, month: int) -> List[Dict[str, Any]]:
    url = f"{ADO_BASE_URL}/{urllib.parse.quote(project, safe='')}/_apis/work/teamsettings/iterations?api-version=7.1"
    try:
        response = requests.get(url, headers=get_ado_headers(pat), timeout=15)
        if response.status_code in [401, 403]: return []
        response.raise_for_status()
        data = response.json()
        month_start = datetime(year, month, 1, 0, 0, 0)
        month_end = datetime(year + 1, 1, 1, 0, 0, 0) - timedelta(seconds=1) if month == 12 else datetime(year, month + 1, 1, 0, 0, 0) - timedelta(seconds=1)
        valid_sprints = []
        for item in data.get("value", []):
            attrs = item.get("attributes", {})
            start_date = parse_date(attrs.get("startDate"))
            if start_date and month_start <= start_date <= month_end:
                valid_sprints.append({"project": project, "name": item["name"], "path": item["path"], "start": start_date.strftime("%d/%m/%Y"), "finish": (attrs.get("finishDate") or "").replace("T", " ").split(".")[0]})
        return valid_sprints
    except: return []

def fetch_all_workitems_ids_by_project(pat: str, project: str, sprint_paths: List[str]) -> Tuple[List[int], Optional[str]]:
    if not sprint_paths: return [], "Nenhum path"
    formatted_paths = ", ".join([f"'{p.replace(chr(92), chr(92)+chr(92))}'" for p in sprint_paths])
    wiql_query = f"SELECT [System.Id] FROM WorkItems WHERE [System.IterationPath] IN ({formatted_paths}) ORDER BY [System.Id]"
    try:
        response = requests.post(f"{ADO_BASE_URL}/{urllib.parse.quote(project, safe='')}/_apis/wit/wiql?api-version=7.1", headers=get_ado_headers(pat), json={"query": wiql_query}, timeout=60)
        if response.status_code in [401, 403]: return [], "API Key inválida"
        response.raise_for_status()
        return [w["id"] for w in response.json().get("workItems", [])], None
    except: return [], "Erro"

def fetch_workitems_ids_by_project(pat: str, project: str, sprint_paths: List[str]) -> Tuple[List[int], Optional[str]]:
    if not sprint_paths: return [], "Nenhum path"
    formatted_paths = ", ".join([f"'{p.replace(chr(92), chr(92)+chr(92))}'" for p in sprint_paths])
    all_types_sql = ", ".join([f"'{t}'" for t in ALL_WI_TYPES])
    wiql_query = f"SELECT [System.Id] FROM WorkItems WHERE [System.WorkItemType] IN ({all_types_sql}) AND [System.IterationPath] IN ({formatted_paths}) ORDER BY [System.Id]"
    try:
        response = requests.post(f"{ADO_BASE_URL}/{urllib.parse.quote(project, safe='')}/_apis/wit/wiql?api-version=7.1", headers=get_ado_headers(pat), json={"query": wiql_query}, timeout=60)
        if response.status_code in [401, 403]: return [], "API Key inválida"
        response.raise_for_status()
        return [w["id"] for w in response.json().get("workItems", [])], None
    except: return [], "Erro"

def fetch_hierarchy_links_by_project(pat: str, project: str, sprint_paths: List[str]) -> Tuple[List[Dict], Optional[str]]:
    if not sprint_paths: return [], "Nenhum path"
    formatted_paths = ", ".join([f"'{p.replace(chr(92), chr(92)+chr(92))}'" for p in sprint_paths])
    wiql_query = f"SELECT [System.Id] FROM workitemLinks WHERE ([Source].[System.IterationPath] IN ({formatted_paths})) AND ([System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Forward' OR [System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Reverse') AND ([Target].[System.IterationPath] IN ({formatted_paths})) MODE (MustContain)"
    try:
        response = requests.post(f"{ADO_BASE_URL}/{urllib.parse.quote(project, safe='')}/_apis/wit/wiql?api-version=7.1", headers=get_ado_headers(pat), json={"query": wiql_query}, timeout=60)
        if response.status_code in [401, 403]: return [], "API Key inválida"
        response.raise_for_status()
        return response.json().get("workItemRelations", []), None
    except: return [], "Erro"

def fetch_workitems_details(pat: str, project: str, ids: List[int]) -> Tuple[Dict[int, Dict], Optional[str]]:
    if not ids: return {}, "Nenhum ID"
    url = f"{ADO_BASE_URL}/{urllib.parse.quote(project, safe='')}/_apis/wit/workitemsbatch?api-version=7.1"
    fields = ["System.Id", "System.Title", "System.WorkItemType", "System.State", "System.IterationPath", "System.CreatedDate", "Microsoft.VSTS.Common.ClosedDate", "Microsoft.VSTS.Scheduling.OriginalEstimate", "System.TeamProject", "Microsoft.VSTS.Scheduling.CompletedWork", "System.AssignedTo"]
    all_items = {}
    try:
        for i in range(0, len(ids), 200):
            response = requests.post(url, headers=get_ado_headers(pat), json={"ids": ids[i:i+200], "fields": fields}, timeout=20)
            if response.status_code in [401, 403]: return all_items, "API Key inválida"
            response.raise_for_status()
            for item in response.json().get("value", []): all_items[item["id"]] = item
        return all_items, None
    except: return all_items, "Erro"

def format_date(date_str: Optional[str]) -> str:
    if not date_str: return "-"
    try: return datetime.fromisoformat(date_str.replace('Z', '+00:00')).strftime("%d/%m/%Y %H:%M:%S")
    except: return "-"

def format_hours(hours: Optional[float]) -> str: return f"{float(hours):.2f}" if hours is not None else "0.00"

def format_moeda_brl(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def esc(text: str) -> str:
    if text is None: return ""
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&#39;")

def state_to_css_class(state: str) -> str: return state.replace(" ", "-").replace("'", "")

def get_assigned_to_name(assigned_to: Any) -> str:
    if not assigned_to: return "Não atribuído"
    if isinstance(assigned_to, dict): return assigned_to.get("displayName", "Não atribuído")
    return str(assigned_to)

def encontrar_colaborador_ponto(assigned_name, azure_to_ponto):
    if assigned_name == "Não atribuído":
        return assigned_name
    azure_key = assigned_name.strip().lower()
    if "<" in azure_key:
        azure_key = azure_key.split("<")[0].strip()
    if azure_key in azure_to_ponto:
        return azure_to_ponto[azure_key]
    for azure_name, ponto_name in azure_to_ponto.items():
        if azure_key in azure_name or azure_name in azure_key:
            return ponto_name
    palavras_azure = azure_key.split()
    for azure_name, ponto_name in azure_to_ponto.items():
        palavras_db = azure_name.split()
        palavras_comuns = set(palavras_azure) & set(palavras_db)
        if len(palavras_comuns) >= 2:
            return ponto_name
    return assigned_name

def generate_csv_parents_only(parents_data, children_map, all_workitems):
    output = io.StringIO()
    writer = csv.writer(output, delimiter=';', quoting=csv.QUOTE_ALL)
    writer.writerow(["Projeto", "Sprint", "ID", "Título", "Tipo", "Situação", "Horas Previstas", "Link", "Dt Criação", "Dt Término", "Qtd Filhos"])
    for p in parents_data:
        f = p.get("fields", {})
        pid, proj = p["id"], p.get("_project", "")
        sprint, title = str(f.get("System.IterationPath", "N/A")), f.get("System.Title", "")
        ptype, pstate = f.get("System.WorkItemType", ""), f.get("System.State", "")
        hours = format_hours(f.get("Microsoft.VSTS.Scheduling.OriginalEstimate"))
        link = f"{ADO_BASE_URL}/{urllib.parse.quote(proj, safe='')}/_workitems/edit/{pid}"
        writer.writerow([proj, sprint, pid, title, ptype, pstate, hours, link, format_date(f.get("System.CreatedDate")), format_date(f.get("Microsoft.VSTS.Common.ClosedDate")), len(children_map.get(pid, []))])
    return "\ufeff" + output.getvalue()

def generate_csv_complete(parents_data, children_map, all_workitems):
    output = io.StringIO()
    writer = csv.writer(output, delimiter=';', quoting=csv.QUOTE_ALL)
    writer.writerow(["Projeto", "Sprint", "ID", "Título", "Tipo", "Situação", "Horas Previstas", "Link", "Dt Criação", "Dt Término", "Nível", "ID Pai"])
    for p in parents_data:
        f = p.get("fields", {})
        pid, proj = p["id"], p.get("_project", "")
        sprint, title = str(f.get("System.IterationPath", "N/A")), f.get("System.Title", "")
        ptype, pstate = f.get("System.WorkItemType", ""), f.get("System.State", "")
        hours = format_hours(f.get("Microsoft.VSTS.Scheduling.OriginalEstimate"))
        link = f"{ADO_BASE_URL}/{urllib.parse.quote(proj, safe='')}/_workitems/edit/{pid}"
        writer.writerow([proj, sprint, pid, title, ptype, pstate, hours, link, format_date(f.get("System.CreatedDate")), format_date(f.get("Microsoft.VSTS.Common.ClosedDate")), "PAI", ""])
        for cid in children_map.get(pid, []):
            if cid not in all_workitems: continue
            cf = all_workitems[cid].get("fields", {})
            writer.writerow([proj, sprint, cid, cf.get("System.Title", ""), cf.get("System.WorkItemType", ""), cf.get("System.State", ""), format_hours(cf.get("Microsoft.VSTS.Scheduling.OriginalEstimate")), f"{ADO_BASE_URL}/{urllib.parse.quote(proj, safe='')}/_workitems/edit/{cid}", format_date(cf.get("System.CreatedDate")), format_date(cf.get("Microsoft.VSTS.Common.ClosedDate")), "FILHO", pid])
    return "\ufeff" + output.getvalue()

def generate_csv_backlogs_nao_done(parents_data):
    output = io.StringIO()
    writer = csv.writer(output, delimiter=';', quoting=csv.QUOTE_ALL)
    writer.writerow(["Projeto", "Sprint", "ID", "Descrição da História", "Assign To", "Status"])
    count = 0
    for p in parents_data:
        f = p.get("fields", {})
        if f.get("System.WorkItemType") == "Product Backlog Item" and f.get("System.State") != "Done":
            proj = p.get("_project", "")
            sprint = str(f.get("System.IterationPath", "N/A")).split("\\")[-1] if f.get("System.IterationPath") else "N/A"
            writer.writerow([proj, sprint, p["id"], f.get("System.Title", ""), get_assigned_to_name(f.get("System.AssignedTo")), f.get("System.State", "")])
            count += 1
    return "\ufeff" + output.getvalue(), count

def generate_csv_tasks_done_sem_horas(parents_data):
    output = io.StringIO()
    writer = csv.writer(output, delimiter=';', quoting=csv.QUOTE_ALL)
    writer.writerow(["Projeto", "Sprint", "ID", "Descrição da Tarefa", "Assign To", "Status", "Completed Work"])
    count = 0
    for p in parents_data:
        f = p.get("fields", {})
        if f.get("System.WorkItemType") == "Task" and f.get("System.State") == "Done":
            completed_work = f.get("Microsoft.VSTS.Scheduling.CompletedWork")
            completed_work_float = float(completed_work) if completed_work is not None else 0.0
            if completed_work_float == 0:
                proj = p.get("_project", "")
                sprint = str(f.get("System.IterationPath", "N/A")).split("\\")[-1] if f.get("System.IterationPath") else "N/A"
                writer.writerow([proj, sprint, p["id"], f.get("System.Title", ""), get_assigned_to_name(f.get("System.AssignedTo")), f.get("System.State", ""), f"{completed_work_float:.2f}"])
                count += 1
    return "\ufeff" + output.getvalue(), count

def generate_excel_faturamento(df_faturamento: pd.DataFrame, mes: str, ano: int, dias_uteis: float, total_faturado: float) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Faturamento"
    
    header_font = Font(name='Montserrat', size=14, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='FF6600', end_color='FF6600', fill_type='solid')
    info_font = Font(name='Montserrat', size=11, bold=True, color='333333')
    data_font = Font(name='Montserrat', size=10)
    money_font = Font(name='Montserrat', size=10, bold=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    ws.merge_cells('A1:M1')
    cell_title = ws['A1']
    cell_title.value = f"FATURAMENTO - {mes}/{ano}"
    cell_title.font = header_font
    cell_title.fill = header_fill
    cell_title.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 35
    
    ws.merge_cells('A2:M2')
    cell_info = ws['A2']
    cell_info.value = f"Dias Úteis: {dias_uteis:.0f} | Total Faturado: {format_moeda_brl(total_faturado)}"
    cell_info.font = info_font
    cell_info.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 25
    
    headers = [
        "Profissional", "Cargo", "Nível", "Data de Início", "Promoção de Cargo",
        "Novo Colaborador", "Data de Saída", "Horas Trabalhadas", "Horas Afast Acordado",
        "Horas Afast Legal", "Horas Arredondadas", "Valor Profissional", "Valor Recalculado"
    ]
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_num)
        cell.value = header
        cell.font = Font(name='Montserrat', size=10, bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='2C5282', end_color='2C5282', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    
    ws.row_dimensions[4].height = 35
    
    column_widths = [35, 20, 15, 15, 20, 25, 15, 18, 22, 18, 20, 20, 20]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    
    for row_num, (_, row) in enumerate(df_faturamento.iterrows(), 5):
        ws.cell(row=row_num, column=1, value=row["Profissional"]).font = data_font
        ws.cell(row=row_num, column=2, value=row["Cargo"]).font = data_font
        ws.cell(row=row_num, column=3, value=row["Nível"]).font = data_font
        ws.cell(row=row_num, column=4, value=row["Data de Início"]).font = data_font
        ws.cell(row=row_num, column=5, value=row["Promoção de Cargo"]).font = data_font
        ws.cell(row=row_num, column=6, value=row["Novo Colaborador"]).font = data_font
        ws.cell(row=row_num, column=7, value=row["Data de Saída"]).font = data_font
        ws.cell(row=row_num, column=8, value=row["Horas Trabalhadas"]).font = data_font
        ws.cell(row=row_num, column=9, value=row["Horas Afast Acordado"]).font = data_font
        ws.cell(row=row_num, column=10, value=row["Horas Afast Legal"]).font = data_font
        ws.cell(row=row_num, column=11, value=row["Horas Arredondadas"]).font = data_font
        ws.cell(row=row_num, column=12, value=row["Valor Profissional"]).font = money_font
        ws.cell(row=row_num, column=13, value=row["Valor Recalculado"]).font = money_font
        
        ws.cell(row=row_num, column=12).number_format = 'R$ #,##0.00'
        ws.cell(row=row_num, column=13).number_format = 'R$ #,##0.00'
        
        for col in range(1, 14):
            ws.cell(row=row_num, column=col).border = thin_border
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

def render_table_html(parents_data, children_map, all_workitems, total_hours):
    css = """
    <style>
    .ado-tbl { width: 100%; border-collapse: collapse; font-family: 'Segoe UI', sans-serif; font-size: 13px; background: white; }
    .ado-tbl th { background-color: #FF6600; color: white; padding: 10px 8px; text-align: left; border: 1px solid #ddd; font-size: 12px; }
    .ado-tbl td { padding: 0; border: 1px solid #f0f0f0; }
    .parent-tbl { width: 100%; border-collapse: collapse; background: white; cursor: pointer; }
    .parent-tbl:hover { background: #fff5eb; }
    .parent-tbl td { padding: 10px 8px; border: none; vertical-align: middle; }
    .child-tbl { width: 100%; font-size: 12px; border-collapse: collapse; background: #fafafa; margin-top: 5px; }
    .child-tbl th { background: #e9ecef; color: #333; padding: 6px 8px; text-align: left; font-size: 11px; }
    .child-tbl td { padding: 6px 8px; border-bottom: 1px solid #e0e0e0; border-left: 3px solid #FF6600; }
    .ftr { background: #E0E0E0; font-weight: bold; }
    .ftr td { padding: 12px 8px; border-top: 2px solid #bbb; }
    .lnk { display: inline-block; padding: 4px 10px; background: #0078d4; color: white; text-decoration: none; border-radius: 4px; font-size: 11px; }
    .lnk:hover { background: #005a9e; }
    .badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }
    .b-Done { background: #d4edda; color: #155724; }
    .b-Resolved { background: #d1ecf1; color: #0c5460; }
    .b-In-Progress { background: #fff3cd; color: #856404; }
    .b-To-Do { background: #e2e3e5; color: #383d41; }
    details > summary { list-style: none; cursor: pointer; padding: 0; }
    details > summary::-webkit-details-marker { display: none; }
    details > summary::before { content: '▶'; display: inline-block; color: #FF6600; font-size: 10px; margin-right: 4px; }
    details[open] > summary::before { content: '▼'; }
    .dt { font-size: 11px; color: #666; }
    .sprint-path { font-size: 11px; color: #444; word-break: break-word; }
    </style>
    """
    rows = []
    for p in parents_data:
        f = p.get("fields", {})
        pid = p["id"]
        title = esc(f.get("System.Title", ""))
        ptype = esc(f.get("System.WorkItemType", ""))
        pstate = esc(f.get("System.State", ""))
        hours = format_hours(f.get("Microsoft.VSTS.Scheduling.OriginalEstimate"))
        sprint_path = esc(str(f.get("System.IterationPath", "N/A")))
        created = format_date(f.get("System.CreatedDate"))
        closed = format_date(f.get("Microsoft.VSTS.Common.ClosedDate"))
        proj = p.get("_project", "")
        link = f"{ADO_BASE_URL}/{urllib.parse.quote(proj, safe='')}/_workitems/edit/{pid}"
        child_ids = children_map.get(pid, [])
        children = [all_workitems[cid] for cid in child_ids if cid in all_workitems]
        state_class = state_to_css_class(pstate)
        parent_row = f"""<tr><td colspan="10" style="padding:0;border:none;">
<details><summary>
<table class="parent-tbl"><tr>
<td style="width:3%"></td>
<td style="width:22%" class="sprint-path"><strong>{sprint_path}</strong></td>
<td style="width:6%;color:#FF6600;font-weight:700;">#{pid}</td>
<td style="width:25%" title="{title}">{title[:50]}{'...' if len(title)>50 else ''}</td>
<td style="width:10%">{ptype}</td>
<td style="width:10%"><span class="badge b-{state_class}">{pstate}</span></td>
<td style="width:7%;font-weight:600;color:#2c5282;">{hours}</td>
<td style="width:6%"><a href="{link}" target="_blank" class="lnk" onclick="event.stopPropagation()">Abrir</a></td>
<td style="width:6%" class="dt">{created}</td>
<td style="width:6%" class="dt">{closed}</td>
</tr></table>
</summary>"""
        if children:
            parent_row += '<div style="padding:0 0 0 30px;"><table class="child-tbl"><thead><tr><th style="width:3%"></th><th style="width:22%">Sprint</th><th style="width:6%">ID</th><th style="width:25%">Título</th><th style="width:10%">Tipo</th><th style="width:10%">Situação</th><th style="width:7%">Horas</th><th style="width:6%">Link</th><th style="width:6%">Dt Criação</th><th style="width:6%">Dt Término</th></tr></thead><tbody>'
            for c in children:
                cf = c.get("fields", {})
                cstate = esc(cf.get("System.State", ""))
                cstate_class = state_to_css_class(cstate)
                parent_row += f"""<tr>
<td></td><td class="sprint-path">{sprint_path}</td>
<td style="color:#666;">#{c['id']}</td>
<td>{esc(cf.get('System.Title',''))}</td>
<td>{esc(cf.get('System.WorkItemType',''))}</td>
<td><span class="badge b-{cstate_class}">{cstate}</span></td>
<td>{format_hours(cf.get('Microsoft.VSTS.Scheduling.OriginalEstimate'))}</td>
<td>-</td>
<td class="dt">{format_date(cf.get('System.CreatedDate'))}</td>
<td class="dt">{format_date(cf.get('Microsoft.VSTS.Common.ClosedDate'))}</td>
</tr>"""
            parent_row += "</tbody></table></div>"
        else:
            parent_row += '<div style="padding:10px 0 10px 30px;color:#888;font-style:italic;font-size:12px;">Nenhuma atividade filha vinculada.</div>'
        parent_row += "</details></td></tr>"
        rows.append(parent_row)
    full_html = css + """
<table class="ado-tbl">
<thead><tr>
<th style="width:3%"></th><th style="width:22%">Sprint</th><th style="width:6%">ID</th>
<th style="width:25%">Título</th><th style="width:10%">Tipo</th><th style="width:10%">Situação</th>
<th style="width:7%">Horas</th><th style="width:6%">Link</th><th style="width:6%">Dt Criação</th><th style="width:6%">Dt Término</th>
</tr></thead>
<tbody>
""" + "\n".join(rows) + f"""
<tr class="ftr">
<td colspan="6" style="text-align:right;">TOTAL ACUMULADO DE HORAS PREVISTAS:</td>
<td style="color:#FF6600;font-size:16px;">{total_hours:.2f} h</td>
<td colspan="3"></td>
</tr>
</tbody>
</table>
"""
    st.markdown(full_html, unsafe_allow_html=True)

def render_colaborador_tab(all_workitems_colab, project_of_item, df_colaboradores, year, month_name):
    periodo_atual = f"{year}-{month_name}"
    if st.session_state.get('filtro_colab_periodo') != periodo_atual:
        st.session_state.filtro_colab_periodo = periodo_atual
        if 'filtro_colab_tab' in st.session_state:
            del st.session_state.filtro_colab_tab
    
    azure_to_ponto = {}
    for _, row in df_colaboradores.iterrows():
        azure_name = str(row["colaborador_azure"]).strip().lower()
        ponto_name = str(row["colaborador_ponto"]).strip()
        if azure_name:
            azure_to_ponto[azure_name] = ponto_name
    
    tarefas_list = []
    for item_id, item in all_workitems_colab.items():
        f = item.get("fields", {})
        if f.get("System.State") != "Done":
            continue
        assigned_to = f.get("System.AssignedTo")
        assigned_name = get_assigned_to_name(assigned_to)
        colaborador_ponto = encontrar_colaborador_ponto(assigned_name, azure_to_ponto)
        proj = project_of_item.get(item_id, "")
        completed_work = float(f.get("Microsoft.VSTS.Scheduling.CompletedWork") or 0)
        link = f"{ADO_BASE_URL}/{urllib.parse.quote(proj, safe='')}/_workitems/edit/{item_id}"
        tarefas_list.append({
            "Colaborador": colaborador_ponto,
            "Projeto": proj,
            "Sprint": str(f.get("System.IterationPath", "N/A")).split("\\")[-1],
            "ID": item_id,
            "Tipo": f.get("System.WorkItemType", ""),
            "Título": f.get("System.Title", ""),
            "Horas Apontadas": completed_work,
            "Link": link
        })
    
    if not tarefas_list:
        st.warning("⚠️ Nenhuma tarefa com State=Done encontrada.")
        return
    
    df_tarefas = pd.DataFrame(tarefas_list)
    filtro = st.sidebar.multiselect("👥 Filtrar por Colaborador", options=sorted(df_tarefas["Colaborador"].unique()), default=sorted(df_tarefas["Colaborador"].unique()), key="filtro_colab_tab")
    if filtro:
        df_tarefas = df_tarefas[df_tarefas["Colaborador"].isin(filtro)]
    
    df_sumario = df_tarefas.groupby("Colaborador").agg({"Horas Apontadas": "sum", "ID": "count"}).reset_index()
    df_sumario.columns = ["Colaborador", "Total Horas", "Qtd Tarefas"]
    df_sumario = df_sumario.sort_values("Colaborador", ascending=True)
    
    st.subheader("📊 Sumário por Colaborador (State: Done)")
    st.info(f"Total: **{len(df_tarefas)} tarefa(s)** | **{df_tarefas['Horas Apontadas'].sum():.2f}h**")
    st.dataframe(df_sumario, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.subheader(" Horas Apontadas por Colaborador")
    fig = px.bar(df_sumario, x="Colaborador", y="Total Horas", text="Total Horas", color="Colaborador", color_discrete_sequence=px.colors.qualitative.Set3)
    fig.update_layout(margin=dict(t=20, b=80, l=20, r=20), height=400, showlegend=False, xaxis_tickangle=-45)
    fig.update_traces(texttemplate='%{text:.1f}h', textposition='outside')
    st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("---")
    st.subheader("📋 Detalhamento por Colaborador")
    for colaborador in sorted(df_tarefas["Colaborador"].unique()):
        df_col = df_tarefas[df_tarefas["Colaborador"] == colaborador].sort_values(["Projeto", "Sprint"])
        with st.expander(f"👤 **{colaborador}** — {len(df_col)} tarefa(s) | {df_col['Horas Apontadas'].sum():.2f}h", expanded=False):
            st.dataframe(df_col[["Projeto", "Sprint", "ID", "Tipo", "Título", "Horas Apontadas", "Link"]], use_container_width=True, hide_index=True,
                column_config={
                    "Projeto": st.column_config.TextColumn("Projeto"),
                    "Sprint": st.column_config.TextColumn("Sprint"),
                    "ID": st.column_config.NumberColumn("ID"),
                    "Tipo": st.column_config.TextColumn("Tipo"),
                    "Título": st.column_config.TextColumn("Título", width="large"),
                    "Horas Apontadas": st.column_config.NumberColumn("Horas", format="%.2f h"),
                    "Link": st.column_config.LinkColumn("Link", display_text="Abrir")
                })

def render_horas_colaborador_tab(df_colaboradores, df_horasponto, horas_uteis_mensais, all_workitems_colab, project_of_item, year, month_name):
    periodo_atual = f"{year}-{month_name}"
    if st.session_state.get('filtro_horas_periodo') != periodo_atual:
        st.session_state.filtro_horas_periodo = periodo_atual
        if 'filtro_horas' in st.session_state:
            del st.session_state.filtro_horas
    
    colaborador_nivel = {}
    for _, row in df_colaboradores.iterrows():
        ponto_name = str(row["colaborador_ponto"]).strip()
        nivel = str(row.get("nivel", "")).strip()
        if ponto_name:
            colaborador_nivel[ponto_name] = nivel
    
    nivel_capacity = {}
    niveis_unicos = set(colaborador_nivel.values())
    for nivel in niveis_unicos:
        if nivel:
            nivel_capacity[nivel] = get_capacity_by_nivel(nivel)
    
    horas_azure = {}
    azure_to_ponto = {}
    for _, row in df_colaboradores.iterrows():
        azure_name = str(row["colaborador_azure"]).strip().lower()
        ponto_name = str(row["colaborador_ponto"]).strip()
        if azure_name:
            azure_to_ponto[azure_name] = ponto_name
    
    for item_id, item in all_workitems_colab.items():
        f = item.get("fields", {})
        if f.get("System.State") != "Done":
            continue
        assigned_name = get_assigned_to_name(f.get("System.AssignedTo"))
        colaborador_ponto = encontrar_colaborador_ponto(assigned_name, azure_to_ponto)
        horas_azure[colaborador_ponto] = horas_azure.get(colaborador_ponto, 0) + float(f.get("Microsoft.VSTS.Scheduling.CompletedWork") or 0)
    
    lista_consolidada = []
    for _, hp_row in df_horasponto.iterrows():
        ponto_name = str(hp_row["colaborador_ponto"]).strip()
        horas_ponto = float(hp_row["horas_ponto"]) if pd.notna(hp_row.get("horas_ponto")) else 0.0
        afastamento = float(hp_row["afastamento"]) if pd.notna(hp_row.get("afastamento")) else 0.0
        nivel = colaborador_nivel.get(ponto_name, "")
        capacity = nivel_capacity.get(nivel, 0.0)
        capacity_calculado = (horas_ponto / 8.0) * capacity if (horas_ponto > 0 and capacity > 0) else 0.0
        lista_consolidada.append({"Colaborador": ponto_name, "Ponto": horas_ponto, "Azure": horas_azure.get(ponto_name, 0.0), "Afastamento": afastamento, "Capacity": capacity_calculado})
    
    df_consolidado = pd.DataFrame(lista_consolidada).sort_values("Colaborador", ascending=True)
    filtro = st.sidebar.multiselect("👥 Filtrar por Colaborador", options=sorted(df_consolidado["Colaborador"].unique()), default=sorted(df_consolidado["Colaborador"].unique()), key="filtro_horas")
    if filtro:
        df_consolidado = df_consolidado[df_consolidado["Colaborador"].isin(filtro)]
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("🕐 Total Ponto", f"{df_consolidado['Ponto'].sum():.2f} h")
    col2.metric("💻 Total Azure", f"{df_consolidado['Azure'].sum():.2f} h")
    col3.metric("🏖️ Total Afastamento", f"{df_consolidado['Afastamento'].sum():.2f} h")
    col4.metric("📊 Total Capacity", f"{df_consolidado['Capacity'].sum():.2f} h")
    
    st.markdown("---")
    st.subheader("📊 Horas por Colaborador")
    st.info(f"Período: **{month_name}/{year}** | Horas úteis: **{horas_uteis_mensais:.2f}h**")
    st.dataframe(df_consolidado, use_container_width=True, hide_index=True,
        column_config={"Colaborador": st.column_config.TextColumn("Colaborador", width="large"),
                       "Ponto": st.column_config.NumberColumn("Ponto (h)", format="%.2f h"),
                       "Azure": st.column_config.NumberColumn("Azure (h)", format="%.2f h"),
                       "Afastamento": st.column_config.NumberColumn("Afastamento (h)", format="%.2f h"),
                       "Capacity": st.column_config.NumberColumn("Capacity (h)", format="%.2f h")})
    
    st.markdown("---")
    st.subheader("📈 Comparativo: Ponto vs Azure vs Capacity")
    df_grafico = df_consolidado.melt(id_vars=["Colaborador"], value_vars=["Ponto", "Azure", "Capacity"], var_name="Fonte", value_name="Horas")
    fig = px.bar(df_grafico, x="Colaborador", y="Horas", color="Fonte", barmode="group", color_discrete_map={"Ponto": "#28a745", "Azure": "#0078d4", "Capacity": "#FF6600"})
    fig.update_layout(margin=dict(t=20, b=80, l=20, r=20), height=450, xaxis_tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)

# =============================================================================
# FUNÇÕES DE PERSISTÊNCIA DE FATURAMENTO
# =============================================================================

def verificar_faturamento_existente(ano: int, mes: int) -> Optional[Dict]:
    """Verifica se já existe faturamento para o ano/mês no banco"""
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url:
        return None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT id, ano, mes, qt_dias_uteis, total_faturado, criado_em 
                FROM faturamento_mensal 
                WHERE ano = %s AND mes = %s
            """, (ano, mes))
            result = cur.fetchone()
        conn.close()
        return result
    except Exception as e:
        logger.error(f"Erro ao verificar faturamento existente: {e}")
        return None

def buscar_detalhes_faturamento(id_faturamento: int) -> pd.DataFrame:
    """Busca os detalhes do faturamento do banco"""
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url:
        return pd.DataFrame()
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT profissional, cargo, nivel, data_inicio, promocao_cargo,
                       novo_colaborador, data_saida, horas_trabalhadas, horas_afast_acordado,
                       horas_afast_legal, horas_arredondadas, valor_profissional, valor_recalculado
                FROM faturamento_detalhe
                WHERE id_faturamento = %s
                ORDER BY profissional
            """, (id_faturamento,))
            rows = cur.fetchall()
        conn.close()
        return pd.DataFrame(rows)
    except Exception as e:
        logger.error(f"Erro ao buscar detalhes do faturamento: {e}")
        return pd.DataFrame()

def salvar_faturamento_no_banco(ano: int, mes: int, dias_uteis: float, total_faturado: float, df_detalhes: pd.DataFrame) -> Tuple[bool, str]:
    """Salva o faturamento calculado no banco de dados"""
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url:
        return False, "Banco de dados não configurado."
    
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            # ✅ Converter valores NumPy para tipos Python nativos
            ano_int = int(ano)
            mes_int = int(mes)
            dias_uteis_float = float(dias_uteis)
            total_faturado_float = float(total_faturado)
            
            cur.execute("""
                INSERT INTO faturamento_mensal (ano, mes, qt_dias_uteis, total_faturado)
                VALUES (%s, %s, %s, %s)
                ON CONFLICT (ano, mes) DO UPDATE SET
                    qt_dias_uteis = EXCLUDED.qt_dias_uteis,
                    total_faturado = EXCLUDED.total_faturado,
                    atualizado_em = CURRENT_TIMESTAMP
                RETURNING id
            """, (ano_int, mes_int, dias_uteis_float, total_faturado_float))
            result = cur.fetchone()
            id_faturamento = result['id']
            
            cur.execute("DELETE FROM faturamento_detalhe WHERE id_faturamento = %s", (id_faturamento,))
            
            for _, row in df_detalhes.iterrows():
                data_ini = None
                data_saida = None
                if row["Data de Início"] and row["Data de Início"] != "":
                    try:
                        data_ini = datetime.strptime(str(row["Data de Início"]), "%d/%m/%Y").date()
                    except:
                        pass
                if row["Data de Saída"] and row["Data de Saída"] != "":
                    try:
                        data_saida = datetime.strptime(str(row["Data de Saída"]), "%d/%m/%Y").date()
                    except:
                        pass
                
                # ✅ Converter TODOS os valores NumPy para tipos Python nativos
                profissional = str(row["Profissional"])
                cargo = str(row["Cargo"])
                nivel = str(row["Nível"])
                promocao = str(row["Promoção de Cargo"]) if row["Promoção de Cargo"] else ""
                novo_colaborador = str(row["Novo Colaborador"]) if row["Novo Colaborador"] else ""
                
                horas_trabalhadas = float(row["Horas Trabalhadas"])
                horas_afast_acordado = float(row["Horas Afast Acordado"])
                horas_afast_legal = float(row["Horas Afast Legal"])
                horas_arredondadas = float(row["Horas Arredondadas"])
                valor_profissional = float(row["Valor Profissional"])
                valor_recalculado = float(row["Valor Recalculado"])
                
                cur.execute("""
                    INSERT INTO faturamento_detalhe 
                    (id_faturamento, profissional, cargo, nivel, data_inicio, promocao_cargo,
                     novo_colaborador, data_saida, horas_trabalhadas, horas_afast_acordado,
                     horas_afast_legal, horas_arredondadas, valor_profissional, valor_recalculado)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    int(id_faturamento),
                    profissional,
                    cargo,
                    nivel,
                    data_ini,
                    promocao,
                    novo_colaborador,
                    data_saida,
                    horas_trabalhadas,
                    horas_afast_acordado,
                    horas_afast_legal,
                    horas_arredondadas,
                    valor_profissional,
                    valor_recalculado
                ))
            
            conn.commit()
        conn.close()
        return True, f"Faturamento salvo com sucesso! ID: {id_faturamento}"
    except Exception as e:
        if conn: conn.rollback()
        if conn: conn.close()
        logger.error(f"Erro ao salvar faturamento: {e}")
        return False, f"Erro ao salvar: {str(e)}"

def deletar_faturamento_existente(ano: int, mes: int) -> Tuple[bool, str]:
    """Deleta faturamento existente do banco"""
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url:
        return False, "Banco de dados não configurado."
    
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("""
                DELETE FROM faturamento_mensal 
                WHERE ano = %s AND mes = %s
            """, (ano, mes))
            if cur.rowcount == 0:
                conn.close()
                return False, "Nenhum faturamento encontrado para deletar."
            conn.commit()
        conn.close()
        return True, "Faturamento deletado com sucesso."
    except Exception as e:
        if conn: conn.rollback()
        if conn: conn.close()
        logger.error(f"Erro ao deletar faturamento: {e}")
        return False, f"Erro ao deletar: {str(e)}"

def calcular_faturamento(df_colaboradores, df_horasponto, df_valores, dias_uteis):
    """Calcula o faturamento baseado nos dados fornecidos"""
    nivel_valor = {}
    for _, row in df_valores.iterrows():
        nivel = str(row["nivel"]).strip()
        valor = limpar_valor_monetario(row.get("valor"))
        if nivel:
            nivel_valor[nivel] = valor
    
    colaborador_horas = {}
    for _, row in df_horasponto.iterrows():
        ponto_name = str(row["colaborador_ponto"]).strip()
        horas_ponto = float(row["horas_ponto"]) if pd.notna(row.get("horas_ponto")) else 0.0
        afastamento = float(row["afastamento"]) if pd.notna(row.get("afastamento")) else 0.0
        if ponto_name:
            colaborador_horas[ponto_name] = {"horas_ponto": horas_ponto, "afastamento": afastamento}
    
    lista_faturamento = []
    for _, row in df_colaboradores.iterrows():
        colaborador_ponto = str(row["colaborador_ponto"]).strip()
        cargo = str(row.get("cargo", "")).strip()
        nivel = str(row.get("nivel", "")).strip()
        data_inicio = row.get("data_inicio")
        promocao = str(row.get("promocao", "")).strip() if pd.notna(row.get("promocao")) else ""
        reposicao = str(row.get("reposicao", "")).strip() if pd.notna(row.get("reposicao")) else ""
        data_saida = row.get("data_saida")
        
        horas_info = colaborador_horas.get(colaborador_ponto, {"horas_ponto": 0.0, "afastamento": 0.0})
        horas_trabalhadas = horas_info["horas_ponto"]
        horas_afast_legal = horas_info["afastamento"]
        horas_afast_acordado = 0.0
        
        horas_arredondadas = horas_afast_acordado + horas_trabalhadas
        valor_profissional = nivel_valor.get(nivel, 0.0)
        valor_diario = valor_profissional / dias_uteis if dias_uteis > 0 else 0.0
        valor_recalculado = round(horas_arredondadas * valor_diario, 2)
        
        lista_faturamento.append({
            "Profissional": colaborador_ponto,
            "Cargo": cargo,
            "Nível": nivel,
            "Data de Início": data_inicio.strftime("%d/%m/%Y") if pd.notna(data_inicio) else "",
            "Promoção de Cargo": promocao,
            "Novo Colaborador": reposicao,
            "Data de Saída": data_saida.strftime("%d/%m/%Y") if pd.notna(data_saida) else "",
            "Horas Trabalhadas": round(horas_trabalhadas, 2),
            "Horas Afast Acordado": round(horas_afast_acordado, 2),
            "Horas Afast Legal": round(horas_afast_legal, 2),
            "Horas Arredondadas": round(horas_arredondadas, 2),
            "Valor Profissional": round(valor_profissional, 2),
            "Valor Recalculado": valor_recalculado
        })
    
    return pd.DataFrame(lista_faturamento)

def render_faturamento_tab(df_colaboradores, df_horasponto, df_valores, dias_uteis, year, month_name, month_num):
    """Renderiza a aba de Faturamento com lógica de persistência"""
    st.subheader("💰 Faturamento por Colaborador")
    st.info(f"Período: **{month_name}/{year}** | Dias úteis: **{dias_uteis:.0f}**")
    
    # Verifica se já existe faturamento no banco para este período
    faturamento_existente = verificar_faturamento_existente(year, month_num)
    
    # Inicializa session_state para controle
    if 'faturamento_calculado' not in st.session_state:
        st.session_state.faturamento_calculado = False
    if 'df_faturamento_exibido' not in st.session_state:
        st.session_state.df_faturamento_exibido = None
    if 'total_faturado_exibido' not in st.session_state:
        st.session_state.total_faturado_exibido = 0.0
    if 'mostrar_alerta_recuperar' not in st.session_state:
        st.session_state.mostrar_alerta_recuperar = False
    
    # Se já existe no banco e ainda não foi calculado nesta sessão, carrega do banco
    if faturamento_existente and not st.session_state.faturamento_calculado:
        st.info(f"✅ Faturamento existente encontrado para **{month_name}/{year}** (ID: {faturamento_existente['id']})")
        df_detalhes = buscar_detalhes_faturamento(faturamento_existente['id'])
        if not df_detalhes.empty:
            # Converte datas para string no formato brasileiro
            for col in ['data_inicio', 'data_saida']:
                if col in df_detalhes.columns:
                    df_detalhes[col] = df_detalhes[col].apply(
                        lambda x: x.strftime("%d/%m/%Y") if pd.notna(x) else ""
                    )
            # Renomeia colunas para o formato esperado
            df_detalhes = df_detalhes.rename(columns={
                'data_inicio': 'Data de Início',
                'promocao_cargo': 'Promoção de Cargo',
                'novo_colaborador': 'Novo Colaborador',
                'data_saida': 'Data de Saída',
                'horas_trabalhadas': 'Horas Trabalhadas',
                'horas_afast_acordado': 'Horas Afast Acordado',
                'horas_afast_legal': 'Horas Afast Legal',
                'horas_arredondadas': 'Horas Arredondadas',
                'valor_profissional': 'Valor Profissional',
                'valor_recalculado': 'Valor Recalculado'
            })
            st.session_state.df_faturamento_exibido = df_detalhes
            st.session_state.total_faturado_exibido = faturamento_existente['total_faturado']
            st.session_state.faturamento_calculado = True
            st.rerun()
    
    # Botão Calcular Faturamento (acima do grid)
    st.markdown("---")
    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn1:
        if st.button(" Calcular Faturamento", use_container_width=True, type="primary"):
            if faturamento_existente:
                st.session_state.mostrar_alerta_recuperar = True
                st.rerun()
            else:
                with st.spinner("Calculando faturamento..."):
                    df_faturamento = calcular_faturamento(df_colaboradores, df_horasponto, df_valores, dias_uteis)
                    total_faturado = df_faturamento["Valor Recalculado"].sum()
                    
                    sucesso, msg = salvar_faturamento_no_banco(year, month_num, dias_uteis, total_faturado, df_faturamento)
                    
                    if sucesso:
                        st.session_state.df_faturamento_exibido = df_faturamento
                        st.session_state.total_faturado_exibido = total_faturado
                        st.session_state.faturamento_calculado = True
                        st.success(f"✅ {msg}")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg}")
    
    # Alerta de recálculo
    if st.session_state.get('mostrar_alerta_recuperar'):
        st.warning(f"⚠️ **Atenção:** Já existem dados de faturamento para **{month_name}/{year}** no banco de dados.")
        st.info(f"ID do faturamento existente: **{faturamento_existente['id']}** | Total: **{format_moeda_brl(faturamento_existente['total_faturado'])}**")
        
        col_sim, col_nao = st.columns(2)
        with col_sim:
            if st.button("✅ Sim, Recalcular", use_container_width=True, type="primary"):
                with st.spinner("Recalculando faturamento..."):
                    sucesso_del, msg_del = deletar_faturamento_existente(year, month_num)
                    if sucesso_del:
                        df_faturamento = calcular_faturamento(df_colaboradores, df_horasponto, df_valores, dias_uteis)
                        total_faturado = df_faturamento["Valor Recalculado"].sum()
                        
                        sucesso_save, msg_save = salvar_faturamento_no_banco(year, month_num, dias_uteis, total_faturado, df_faturamento)
                        
                        if sucesso_save:
                            st.session_state.df_faturamento_exibido = df_faturamento
                            st.session_state.total_faturado_exibido = total_faturado
                            st.session_state.faturamento_calculado = True
                            st.session_state.mostrar_alerta_recuperar = False
                            st.success(f"✅ Faturamento recalculado e salvo! {msg_save}")
                            st.rerun()
                        else:
                            st.error(f"❌ {msg_save}")
                    else:
                        st.error(f"❌ {msg_del}")
        
        with col_nao:
            if st.button("❌ Não, Cancelar", use_container_width=True):
                st.session_state.mostrar_alerta_recuperar = False
                st.rerun()
    
    # Exibe o grid se já foi calculado
    if st.session_state.faturamento_calculado and st.session_state.df_faturamento_exibido is not None:
        df_faturamento = st.session_state.df_faturamento_exibido
        total_faturado = st.session_state.total_faturado_exibido
        
        st.markdown("---")
        col_total1, col_total2 = st.columns([3, 1])
        with col_total1:
            st.markdown("### 💵 Total Faturado")
        with col_total2:
            st.metric(label="", value=format_moeda_brl(total_faturado))
        st.markdown("---")
        
        excel_data = generate_excel_faturamento(df_faturamento, month_name, year, dias_uteis, total_faturado)
        st.download_button(
            "⬇️ Exportar Faturamento (Excel)",
            data=excel_data,
            file_name=f"faturamento_{year}_{month_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.markdown("---")
        
        st.dataframe(
            df_faturamento,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Profissional": st.column_config.TextColumn("Profissional", width="large"),
                "Cargo": st.column_config.TextColumn("Cargo"),
                "Nível": st.column_config.TextColumn("Nível"),
                "Data de Início": st.column_config.TextColumn("Data de Início"),
                "Promoção de Cargo": st.column_config.TextColumn("Promoção de Cargo"),
                "Novo Colaborador": st.column_config.TextColumn("Novo Colaborador"),
                "Data de Saída": st.column_config.TextColumn("Data de Saída"),
                "Horas Trabalhadas": st.column_config.NumberColumn("Horas Trabalhadas", format="%.2f h"),
                "Horas Afast Acordado": st.column_config.NumberColumn("Horas Afast Acordado", format="%.2f h"),
                "Horas Afast Legal": st.column_config.NumberColumn("Horas Afast Legal", format="%.2f h"),
                "Horas Arredondadas": st.column_config.NumberColumn("Horas Arredondadas", format="%.2f h"),
                "Valor Profissional": st.column_config.NumberColumn("Valor Profissional", format="R$ %.2f"),
                "Valor Recalculado": st.column_config.NumberColumn("Valor Recalculado", format="R$ %.2f")
            }
        )
    else:
        st.info(" Clique em **🧮 Calcular Faturamento** para gerar o faturamento do período.")

def render_crud_perfil():
    usuario_info = st.session_state.user_info
    tipo_usuario = usuario_info.get('cargo', 'User')
    usuario_id = usuario_info.get('id')
    st.markdown("---")
    st.subheader("️ Gerenciamento de Perfil")
    usuarios = listar_usuarios(tipo_usuario, usuario_id)
    if not usuarios:
        st.warning("Nenhum usuário encontrado.")
        return
    st.markdown("### Usuários Cadastrados")
    df_usuarios = pd.DataFrame(usuarios)
    df_exibicao = df_usuarios.copy()
    df_exibicao['apikey'] = df_exibicao['apikey'].apply(lambda x: f"{x[:6]}...{x[-4:]}" if x and len(x) > 10 else x)
    df_exibicao['criado_em'] = pd.to_datetime(df_exibicao['criado_em']).dt.strftime('%d/%m/%Y %H:%M')
    st.dataframe(
        df_exibicao[['id', 'nome', 'email', 'tipo', 'status', 'apikey', 'criado_em']],
        use_container_width=True,
        hide_index=True,
        column_config={
            "id": st.column_config.NumberColumn("ID"),
            "nome": st.column_config.TextColumn("Nome"),
            "email": st.column_config.TextColumn("E-mail"),
            "tipo": st.column_config.TextColumn("Tipo"),
            "status": st.column_config.TextColumn("Status"),
            "apikey": st.column_config.TextColumn("API Key"),
            "criado_em": st.column_config.TextColumn("Criado em")
        }
    )
    st.markdown("---")
    if tipo_usuario == 'Admin':
        usuario_selecionado = st.selectbox("Selecione o usuário para editar:", options=[f"{u['id']} - {u['nome']} ({u['email']})" for u in usuarios], index=0)
        usuario_id_selecionado = int(usuario_selecionado.split(" - ")[0])
        usuario_dados = next((u for u in usuarios if u['id'] == usuario_id_selecionado), None)
    else:
        usuario_dados = usuarios[0]
        usuario_id_selecionado = usuario_id
        st.info(f"Editando seu próprio perfil: **{usuario_dados['nome']}**")
    if not usuario_dados:
        st.error("Usuário não encontrado.")
        return
    st.markdown("### Editar Dados")
    with st.form("form_editar_usuario"):
        nome = st.text_input("Nome Completo", value=usuario_dados.get('nome', ''))
        email = st.text_input("E-mail", value=usuario_dados.get('email', ''))
        apikey = st.text_input("API Key do Azure DevOps", value=usuario_dados.get('apikey', ''), type="password")
        senha_hash = st.text_input("Nova Senha (deixe em branco para manter)", type="password", placeholder="Mínimo 6 caracteres")
        if tipo_usuario == 'Admin':
            col1, col2 = st.columns(2)
            with col1:
                tipo = st.selectbox("Tipo de Usuário", options=['User', 'Admin', 'Gestor', 'Auditor'], index=['User', 'Admin', 'Gestor', 'Auditor'].index(usuario_dados.get('tipo', 'User')))
            with col2:
                status = st.selectbox("Status", options=['Ativo', 'Inativo', 'Bloqueado', 'Pendente'], index=['Ativo', 'Inativo', 'Bloqueado', 'Pendente'].index(usuario_dados.get('status', 'Ativo')))
        else:
            tipo = usuario_dados.get('tipo', 'User')
            status = usuario_dados.get('status', 'Ativo')
        submitted = st.form_submit_button("💾 Salvar Alterações")
        if submitted:
            erros = []
            if not nome: erros.append("Nome é obrigatório")
            if not email: erros.append("E-mail é obrigatório")
            if senha_hash and len(senha_hash) < 6: erros.append("Senha deve ter no mínimo 6 caracteres")
            if erros:
                for erro in erros: st.error(f"⚠️ {erro}")
            else:
                dados = {'nome': nome.strip(), 'email': email.lower().strip(), 'apikey': apikey.strip()}
                if senha_hash: dados['senha_hash'] = senha_hash
                if tipo_usuario == 'Admin':
                    dados['tipo'] = tipo
                    dados['status'] = status
                sucesso, msg = atualizar_usuario(usuario_id_selecionado, dados, tipo_usuario, usuario_id)
                if sucesso:
                    st.success(f"✅ {msg}")
                    if usuario_id_selecionado == usuario_id:
                        st.session_state.user_info['nome_completo'] = nome.strip()
                        st.session_state.user_info['email'] = email.lower().strip()
                        if apikey.strip(): st.session_state.user_info['apikey'] = apikey.strip()
                    st.rerun()
                else:
                    st.error(f"❌ {msg}")
    if tipo_usuario == 'Admin' and usuario_id_selecionado != usuario_id:
        st.markdown("---")
        st.markdown("### Zona de Perigo")
        st.warning(f"Você está prestes a deletar o usuário: **{usuario_dados['nome']}**")
        if st.button("🗑️ Deletar Usuário", type="secondary"):
            sucesso, msg = deletar_usuario(usuario_id_selecionado, tipo_usuario)
            if sucesso:
                st.success(f"✅ {msg}")
                st.rerun()
            else:
                st.error(f"❌ {msg}")

def render_login_screen():
    inject_custom_css()
    st.markdown('<div class="dotted-rings"><div class="ring ring-1"></div><div class="ring ring-2"></div><div class="ring ring-3"></div><div class="ring ring-4"></div><div class="ring ring-5"></div></div>', unsafe_allow_html=True)
    col1, col2 = st.columns([1.1, 0.9], gap="large")
    with col1:
        st.markdown('<div style="margin-top: 40px;"><span class="badge-qween">Qween Platform</span><h1 class="hero-title">Relatório Mensal de acompanhamento da Execução</h1><p class="hero-subtitle">Secretaria Municipal de Educação - SP</p></div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="auth-card">', unsafe_allow_html=True)
        view = st.session_state.get("login_view", "login")
        if view == "login":
            st.markdown("<h2 style='color:white; font-weight:700;'>Entrar</h2>", unsafe_allow_html=True)
            st.markdown("<p style='color:#e6e6e6; margin-bottom:30px;'>Acesse com suas credenciais institucionais</p>", unsafe_allow_html=True)
            with st.form("login_form"):
                email = st.text_input("E-mail ou Usuário", placeholder="nome@spassu.com.br")
                password = st.text_input("Senha", type="password", placeholder="••••••••")
                submitted = st.form_submit_button("Acessar Painel")
                if submitted:
                    if email and password:
                        user = authenticate_user(email, password)
                        if user:
                            if user.get('status') != 'Ativo':
                                st.error("⚠️ Usuário inativo. Contate o administrador.")
                            elif not user.get('apikey'):
                                st.error("⚠️ Seu usuário não possui API Key do Azure DevOps cadastrada.")
                            else:
                                st.session_state.logged_in = True
                                st.session_state.user_info = {'id': user['id'], 'nome_completo': user['nome'], 'email': user['email'], 'cargo': user['tipo'], 'apikey': user['apikey']}
                                st.session_state.login_view = "login"
                                st.rerun()
                        else: st.error("E-mail ou senha inválidos.")
                    else: st.warning("Preencha todos os campos.")
            st.markdown("")
            col_btn1, col_btn2 = st.columns([3, 1])
            with col_btn2:
                if st.button("Esqueci minha senha", use_container_width=True):
                    st.session_state.login_view = "forgot"; st.rerun()
            st.markdown('<div style="display:flex; align-items:center; text-align:center; margin:28px 0; color:rgba(255,255,255,0.7); font-size:0.8rem; font-weight:600;"><div style="flex:1; border-bottom:1px solid rgba(255,255,255,0.25); margin-right:12px;"></div>OU<div style="flex:1; border-bottom:1px solid rgba(255,255,255,0.25); margin-left:12px;"></div></div>', unsafe_allow_html=True)
            if st.button("Primeiro Acesso? Ativar Conta", use_container_width=True):
                st.session_state.login_view = "register"; st.rerun()
        elif view == "forgot":
            st.markdown("<h2 style='color:white; font-weight:700;'>Recuperação de Senha</h2>", unsafe_allow_html=True)
            st.markdown("<p style='color:#e6e6e6; margin-bottom:30px;'>Insira seu e-mail para receber o token de recuperação</p>", unsafe_allow_html=True)
            with st.form("forgot_form"):
                email = st.text_input("E-mail Cadastrado", placeholder="seu.email@spassu.com.br")
                submitted = st.form_submit_button("Gerar Token de Recuperação")
                if submitted:
                    if not email:
                        st.error("⚠️ Informe o e-mail.")
                    else:
                        sucesso, msg, token = solicitar_reset_senha(email)
                        if sucesso:
                            st.success(f"✅ {msg}")
                            st.markdown("---")
                            st.markdown("** Seu Token de Recuperação:**")
                            st.markdown(f'<div class="token-box">{token}</div>', unsafe_allow_html=True)
                            st.warning("⚠️ **Guarde este token!** Ele será necessário no próximo passo.")
                            st.info(f"⏰ Token válido por **1 hora**")
                            st.session_state.reset_token = token
                            st.session_state.reset_email = email
                            st.session_state.go_to_reset = True
                        else:
                            st.error(f"❌ {msg}")
            if st.session_state.get('go_to_reset'):
                if st.button("Redefinir Senha Agora →", use_container_width=True):
                    st.session_state.login_view = "reset_password"
                    st.session_state.go_to_reset = False
                    st.rerun()
            if st.button("← Voltar ao Login", use_container_width=True):
                st.session_state.login_view = "login"
                st.session_state.go_to_reset = False
                st.rerun()
        elif view == "reset_password":
            st.markdown("<h2 style='color:white; font-weight:700;'>Redefinir Senha</h2>", unsafe_allow_html=True)
            st.markdown("<p style='color:#e6e6e6; margin-bottom:30px;'>Insira o token recebido e sua nova senha</p>", unsafe_allow_html=True)
            with st.form("reset_form"):
                token = st.text_input("Token de Recuperação", placeholder="Cole o token aqui", value=st.session_state.get('reset_token', ''))
                nova_senha = st.text_input("Nova Senha", type="password", placeholder="Mínimo 6 caracteres")
                confirmar_senha = st.text_input("Confirme a Nova Senha", type="password", placeholder="Repita a senha")
                submitted = st.form_submit_button("Redefinir Senha")
                if submitted:
                    if not token:
                        st.error("⚠️ Token é obrigatório")
                    elif not nova_senha or len(nova_senha) < 6:
                        st.error("️ Senha deve ter no mínimo 6 caracteres")
                    elif nova_senha != confirmar_senha:
                        st.error("⚠️ Senhas não coincidem")
                    else:
                        sucesso, msg = redefinir_senha(token, nova_senha)
                        if sucesso:
                            st.success(f"✅ {msg}")
                            if 'reset_token' in st.session_state: del st.session_state.reset_token
                            if 'reset_email' in st.session_state: del st.session_state.reset_email
                            if 'go_to_reset' in st.session_state: del st.session_state.go_to_reset
                            st.session_state.login_view = "login"
                            st.rerun()
                        else:
                            st.error(f"❌ {msg}")
            if st.button("← Voltar ao Login", use_container_width=True):
                st.session_state.login_view = "login"
                st.session_state.go_to_reset = False
                st.rerun()
        elif view == "register":
            st.markdown("<h2 style='color:white; font-weight:700;'>Primeiro Acesso</h2>", unsafe_allow_html=True)
            st.markdown("<p style='color:#e6e6e6; margin-bottom:30px;'>Crie sua conta para acessar o sistema</p>", unsafe_allow_html=True)
            with st.form("register_form", clear_on_submit=True):
                nome = st.text_input("Nome Completo", placeholder="Seu nome completo")
                email = st.text_input("E-mail Corporativo", placeholder="colaborador@spassu.com.br")
                apikey = st.text_input("API Key do Azure DevOps", placeholder="Cole sua Personal Access Token (PAT)")
                st.caption("💡 Gere em: Azure DevOps → User Settings → Personal Access Tokens")
                senha = st.text_input("Crie uma Senha", type="password", placeholder="Mínimo 6 caracteres")
                confirmar_senha = st.text_input("Confirme a Senha", type="password", placeholder="Repita a senha")
                submitted = st.form_submit_button("Validar e Criar Conta")
                if submitted:
                    erros = []
                    if not nome: erros.append("Nome é obrigatório")
                    if not email: erros.append("E-mail é obrigatório")
                    if not apikey or len(apikey.strip()) < 10: erros.append("API Key do Azure DevOps é obrigatória (mínimo 10 caracteres)")
                    if not senha or len(senha) < 6: erros.append("Senha deve ter no mínimo 6 caracteres")
                    if senha != confirmar_senha: erros.append("Senhas não coincidem")
                    if erros:
                        for erro in erros: st.error(f"⚠️ {erro}")
                    else:
                        sucesso, msg = criar_usuario_com_apikey(nome, email, senha, apikey.strip())
                        if sucesso:
                            st.success(f"✅ {msg}")
                            st.info("🎉 Sua conta foi criada! Faça login com seu e-mail e senha.")
                            st.session_state.login_view = "login"
                            st.rerun()
                        else:
                            st.error(f"❌ {msg}")
            if st.button("Já possuo senha (Entrar)", use_container_width=True):
                st.session_state.login_view = "login"
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown(f'<div style="position:fixed;bottom:30px;right:40px;display:flex;align-items:center;gap:20px;z-index:100;"><img src="{SPASSU_LOGO_URL}" style="height:40px;"><div style="width:1px;height:40px;background:rgba(255,255,255,0.5);"></div><img src="{SME_LOGO_URL}" style="height:40px;"></div>', unsafe_allow_html=True)

def main():
    if 'logged_in' not in st.session_state: st.session_state.logged_in = False
    if 'login_view' not in st.session_state: st.session_state.login_view = "login"
    if 'show_perfil' not in st.session_state: st.session_state.show_perfil = False

    if not st.session_state.logged_in:
        render_login_screen()
        return

    st.set_page_config(page_title="Dashboard de Entregas - SME Spassu", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
    st.sidebar.markdown("---")
    st.sidebar.success(f"👤 {st.session_state.user_info['nome_completo']}")
    st.sidebar.caption(f"Tipo: {st.session_state.user_info.get('cargo', 'User')}")
    if st.sidebar.button("️ Perfil", use_container_width=True):
        st.session_state.show_perfil = True
        st.rerun()
    if st.sidebar.button("🚪 Sair do Sistema", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.user_info = None
        st.session_state.show_perfil = False
        st.rerun()
    if st.session_state.get('show_perfil'):
        st.title("⚙️ Gerenciamento de Perfil")
        render_crud_perfil()
        if st.button("← Voltar ao Dashboard"):
            st.session_state.show_perfil = False
            st.rerun()
        return

    st.title(" Dashboard de Acompanhamento de Entregas")
    st.caption("Azure DevOps • SME-Spassu • Sprint Tracker")
    st.markdown("---")
    st.sidebar.header("⚙️ Configuração & Filtros")
    pat = st.session_state.user_info.get('apikey', '')
    if not pat:
        st.error("⚠️ API Key do Azure DevOps não cadastrada para este usuário.")
        st.stop()
    with st.spinner("🔐 Validando API Key do Azure DevOps..."):
        api_valida, msg_erro = validate_ado_api_key(pat)
    if not api_valida:
        st.error(f"🚫 {msg_erro}")
        st.info("💡 **Como resolver:** Acesse **⚙️ Perfil** na barra lateral e atualize sua API Key do Azure DevOps com um Personal Access Token (PAT) válido.")
        st.stop()

    current_year = datetime.now().year
    year = st.sidebar.selectbox("📅 Ano", options=[current_year + 1, current_year, current_year - 1], index=1 if current_year == 2026 else 0)
    month_names = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    month_name = st.sidebar.selectbox("📅 Mês", options=month_names, index=datetime.now().month - 1)
    month_num = month_names.index(month_name) + 1

    all_projects = get_active_projects_from_db()
    selected_projects = st.sidebar.multiselect("🏢 Projetos", options=all_projects, default=all_projects)
    if not selected_projects: st.info("Selecione ao menos um projeto."); st.stop()

    with st.spinner("🔄 Buscando sprints..."):
        sprints_by_project = {proj: fetch_all_sprints_for_project(pat, proj, year, month_num) for proj in selected_projects}
        sprints_by_project = {k: v for k, v in sprints_by_project.items() if v}
    if not sprints_by_project: st.warning(f"⚠️ Nenhuma sprint encontrada em {month_name}/{year}."); st.stop()

    sprint_options = [f"{s['path']}" for proj, sprints in sprints_by_project.items() for s in sprints]
    selected_sprints_display = st.sidebar.multiselect(f"🏃 Sprints", options=sprint_options, default=sprint_options)
    sprint_to_info = {}
    for proj, sprints in sprints_by_project.items():
        for s in sprints: sprint_to_info[f"{proj}|||{s['path']}"] = (proj, s["path"])
    selected_sprint_paths, paths_by_project = [], {}
    for display_name in selected_sprints_display:
        for key, (proj, path) in sprint_to_info.items():
            if path == display_name:
                selected_sprint_paths.append((proj, path)); paths_by_project.setdefault(proj, []).append(path); break

    selected_wi_types = st.sidebar.multiselect("📋 Tipo de Work Item", options=ALL_WI_TYPES, default=DEFAULT_WI_TYPES)
    selected_states = st.sidebar.multiselect("🎯 Situação", options=ALL_STATES, default=ALL_STATES)

    with st.spinner(" Buscando work items..."):
        all_ids_by_project = {proj: fetch_workitems_ids_by_project(pat, proj, paths)[0] for proj, paths in paths_by_project.items()}
        all_workitems, project_of_item = {}, {}
        for proj, ids in all_ids_by_project.items():
            if ids:
                items, _ = fetch_workitems_details(pat, proj, ids)
                for item_id, item in items.items(): all_workitems[item_id] = item; project_of_item[item_id] = proj

    with st.spinner("🔍 Buscando work items (Colaboradores)..."):
        all_ids_colab = {proj: fetch_all_workitems_ids_by_project(pat, proj, paths)[0] for proj, paths in paths_by_project.items()}
        all_workitems_colab, project_of_item_colab = {}, {}
        for proj, ids in all_ids_colab.items():
            if ids:
                items, _ = fetch_workitems_details(pat, proj, ids)
                for item_id, item in items.items(): all_workitems_colab[item_id] = item; project_of_item_colab[item_id] = proj

    with st.spinner("🔗 Buscando hierarquia..."):
        children_map = {}
        for proj, paths in paths_by_project.items():
            relations, _ = fetch_hierarchy_links_by_project(pat, proj, paths)
            for rel in relations:
                if rel and rel.get("source") and rel.get("target"):
                    src, tgt = rel["source"].get("id"), rel["target"].get("id")
                    if rel.get("rel") == "System.LinkTypes.Hierarchy-Forward": children_map.setdefault(src, []).append(tgt)
                    elif rel.get("rel") == "System.LinkTypes.Hierarchy-Reverse": children_map.setdefault(tgt, []).append(src)

    parents_data = []
    valid_paths = set(p for _, p in selected_sprint_paths)
    for item_id, item in all_workitems.items():
        f = item.get("fields", {})
        if f.get("System.WorkItemType") in selected_wi_types and f.get("System.State") in selected_states and f.get("System.IterationPath", "") in valid_paths:
            item["_project"] = project_of_item.get(item_id, ""); parents_data.append(item)

    df_colaboradores = get_colaboradores_from_db()
    df_horasponto = get_horasponto_from_db(year, month_num)
    horas_uteis_mensais = get_dias_uteis_from_db(year, month_num)
    df_valores = get_valores_from_db()

    tipo_usuario = st.session_state.user_info.get('cargo', 'User')
    if tipo_usuario in ['Admin', 'Gestor']:
        tab_lista, tab_colaborador, tab_horas, tab_faturamento = st.tabs(["📋 Lista Geral", "👥 Tarefas por Colaborador", "🕐 Horas por Colaborador", " Faturamento"])
    else:
        tab_lista, tab_colaborador, tab_horas = st.tabs(["📋 Lista Geral", "👥 Tarefas por Colaborador", "🕐 Horas por Colaborador"])
    
    with tab_lista:
        total_items = len(parents_data)
        total_hours = 0.0
        for p in parents_data:
            f = p.get("fields", {})
            horas = f.get("Microsoft.VSTS.Scheduling.OriginalEstimate")
            if horas is not None:
                try: total_hours += float(horas)
                except: pass
        done_resolved = sum(1 for p in parents_data if p.get("fields", {}).get("System.State") in ["Done", "Resolved"])
        in_progress_todo = sum(1 for p in parents_data if p.get("fields", {}).get("System.State") in ["In Progress", "To Do"])
        historias_nao_done = sum(1 for p in parents_data if p.get("fields", {}).get("System.WorkItemType") == "Product Backlog Item" and p.get("fields", {}).get("System.State") != "Done")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📋 Total de Itens", total_items)
        col2.metric("⏱️ Total de Horas", f"{total_hours:.2f} h")
        col3.metric("📊 Situação", f"✅ {done_resolved} | 🔄 {in_progress_todo}")
        col4.metric("📖 Histórias em Aberto", historias_nao_done, help="Product Backlog Items com status diferente de Done")
        st.markdown("---")
        if parents_data:
            type_counts = {}
            for p in parents_data:
                t = p.get("fields", {}).get("System.WorkItemType", "Desconhecido")
                type_counts[t] = type_counts.get(t, 0) + 1
            df_types = pd.DataFrame([{"Tipo": k, "Quantidade": v} for k, v in sorted(type_counts.items(), key=lambda x: x[1], reverse=True)])
            state_counts = {}
            for p in parents_data:
                s = p.get("fields", {}).get("System.State", "Desconhecido")
                state_counts[s] = state_counts.get(s, 0) + 1
            df_states = pd.DataFrame([{"Situação": k, "Quantidade": v} for k, v in sorted(state_counts.items(), key=lambda x: x[1], reverse=True)])
            state_colors = {"Done": "#28a745", "Resolved": "#17a2b8", "In Progress": "#ffc107", "To Do": "#6c757d"}
            type_colors = px.colors.qualitative.Set2
            col_chart1, col_chart2 = st.columns(2)
            with col_chart1:
                st.subheader("📊 Total por Tipo")
                fig_type = px.pie(df_types, values="Quantidade", names="Tipo", hole=0.5, color_discrete_sequence=type_colors, hover_data=["Quantidade", "Tipo"])
                fig_type.update_layout(margin=dict(t=20, b=20, l=20, r=20), height=350, showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5))
                fig_type.update_traces(textposition='inside', textinfo='percent+label', textfont_size=11)
                st.plotly_chart(fig_type, use_container_width=True)
            with col_chart2:
                st.subheader(" Total por Situação")
                fig_state = px.pie(df_states, values="Quantidade", names="Situação", hole=0.5, color="Situação", color_discrete_map=state_colors, hover_data=["Quantidade", "Situação"])
                fig_state.update_layout(margin=dict(t=20, b=20, l=20, r=20), height=350, showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5))
                fig_state.update_traces(textposition='inside', textinfo='percent+label', textfont_size=11)
                st.plotly_chart(fig_state, use_container_width=True)
            st.markdown("---")
            st.subheader(" Exportar Dados")
            st.markdown("**📋 Work Items Gerais**")
            col_csv1, col_csv2 = st.columns(2)
            with col_csv1:
                csv_parents = generate_csv_parents_only(parents_data, children_map, all_workitems)
                st.download_button("⬇️ Baixar CSV (Apenas Pais)", data=csv_parents, file_name=f"sme_sprints_{year}_{month_name}_pais.csv", mime="text/csv", use_container_width=True)
            with col_csv2:
                csv_complete = generate_csv_complete(parents_data, children_map, all_workitems)
                st.download_button("⬇️ Baixar CSV (Completo)", data=csv_complete, file_name=f"sme_sprints_{year}_{month_name}_completo.csv", mime="text/csv", use_container_width=True)
            st.markdown("---")
            st.markdown("**📖 Histórias (Product Backlog Item) em Aberto**")
            st.caption("Product Backlog Items com status diferente de Done no período selecionado")
            csv_backlogs, count_backlogs = generate_csv_backlogs_nao_done(parents_data)
            col_csv3, col_csv4 = st.columns(2)
            with col_csv3:
                st.download_button(f"️ Baixar CSV - Histórias em Aberto ({count_backlogs} itens)", data=csv_backlogs, file_name=f"backlogs_abertos_{year}_{month_name}.csv", mime="text/csv", use_container_width=True, help="Exporta Product Backlog Items com status diferente de Done")
            with col_csv4:
                st.info(f"📊 **{count_backlogs}** histórias em aberto encontradas no período")
            st.markdown("---")
            st.markdown("**✅ Tasks Concluídas sem Horas Apontadas**")
            st.caption("Tasks com Status = Done e Completed Work = 0 ou vazio")
            csv_tasks, count_tasks = generate_csv_tasks_done_sem_horas(parents_data)
            col_csv5, col_csv6 = st.columns(2)
            with col_csv5:
                st.download_button(f"⬇️ Baixar CSV - Tasks Done sem Horas ({count_tasks} itens)", data=csv_tasks, file_name=f"tasks_done_sem_horas_{year}_{month_name}.csv", mime="text/csv", use_container_width=True, help="Exporta Tasks com State=Done e CompletedWork=0 ou vazio")
            with col_csv6:
                st.info(f"📊 **{count_tasks}** tasks concluídas sem horas apontadas")
            st.markdown("---")
            render_table_html(parents_data, children_map, all_workitems, total_hours)
        else:
            st.warning("️ Nenhum item após filtros.")
    
    with tab_colaborador:
        st.subheader(" Tarefas por Colaborador")
        st.info(f"Período: **{month_name}/{year}** | Apenas **State = Done**")
        if not df_colaboradores.empty: render_colaborador_tab(all_workitems_colab, project_of_item_colab, df_colaboradores, year, month_name)
        else: st.warning("⚠️ Tabela de colaboradores vazia.")
    
    with tab_horas:
        st.subheader("🕐 Horas por Colaborador")
        if df_colaboradores.empty: st.warning("⚠️ Tabela de colaboradores vazia.")
        elif df_horasponto.empty: st.warning(f"⚠️ Sem dados em `horasponto` para {month_name}/{year}.")
        elif horas_uteis_mensais == 0: st.warning(f"⚠️ Sem dados em `dias_uteis_mensais` para {month_name}/{year}.")
        else: render_horas_colaborador_tab(df_colaboradores, df_horasponto, horas_uteis_mensais, all_workitems_colab, project_of_item_colab, year, month_name)
    
    if tipo_usuario in ['Admin', 'Gestor']:
        with tab_faturamento:
            if df_colaboradores.empty:
                st.warning("⚠️ Tabela de colaboradores vazia.")
            elif df_horasponto.empty:
                st.warning(f"⚠️ Sem dados em `horasponto` para {month_name}/{year}.")
            elif df_valores.empty:
                st.warning("⚠️ Tabela de valores vazia.")
            else:
                # ✅ CORREÇÃO: passando month_num como parâmetro adicional
                render_faturamento_tab(df_colaboradores, df_horasponto, df_valores, horas_uteis_mensais, year, month_name, month_num)

if __name__ == "__main__":
    main()
