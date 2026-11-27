import os, base64, urllib.parse, csv, io, logging, secrets
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone
from typing import List, Dict, Any, Optional, Tuple
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import psycopg2
from psycopg2.extras import RealDictCursor, execute_values
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

ANO_ATUAL = datetime.now().year
MIN_DATA = datetime(ANO_ATUAL - 100, 1, 1).date()
MAX_DATA = datetime(ANO_ATUAL + 100, 12, 31).date()

MOTIVOS_DEMISSAO = ["", "pedido de demissão", "desligamento", "fim de alocação", "outro"]

@st.cache_resource
def get_http_session() -> requests.Session:
    session = requests.Session()
    retry_strategy = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retry_strategy, pool_connections=25, pool_maxsize=25)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session

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
    .destaque-card { background: white; border-radius: 12px; padding: 20px; border: 1px solid #e0e0e0; box-shadow: 0 2px 8px rgba(0,0,0,0.05); min-height: 280px; }
    .destaque-titulo { font-size: 11px; color: #FF6600; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; }
    .destaque-subtitulo { font-size: 18px; font-weight: 700; color: #1a1a2e; margin-bottom: 15px; }
    .destaque-item { padding: 10px 0; border-bottom: 1px solid #f0f0f0; }
    .destaque-item:last-child { border-bottom: none; }
    .destaque-nome { font-weight: 600; color: #1a1a2e; font-size: 14px; }
    .destaque-cargo { font-size: 12px; color: #666; }
    .destaque-data { font-size: 12px; color: #FF6600; font-weight: 600; }
    .badge-hoje { background: #FF6600; color: white; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }
    .badge-amanha { background: #28a745; color: white; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }
    .badge-semana { background: #e9ecef; color: #495057; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }
    </style>
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    """
    st.markdown(css, unsafe_allow_html=True)

def multiselect_with_toggle(label: str, options: list, default: list, key: str):
    if not options: options = []
    
    # Inicializa apenas na primeira vez (quando a key ainda não existe)
    if key not in st.session_state:
        st.session_state[key] = [d for d in default if d in options] if default else []
    else:
        # Mantém apenas valores que ainda estão nas opções disponíveis
        st.session_state[key] = [v for v in st.session_state[key] if v in options]
        # ⚠️ REMOVIDO: Não re-preenche automaticamente quando a lista fica vazia
        # O usuário deve poder ter seleção vazia se quiser
    
    st.sidebar.markdown(f"**{label}**")
    col_btn1, col_btn2 = st.sidebar.columns(2)
    with col_btn1:
        if st.sidebar.button("✅ Todos", key=f"{key}_btn_all", use_container_width=True):
            st.session_state[key] = list(options)
            st.rerun()
    with col_btn2:
        if st.sidebar.button("❌ Nenhum", key=f"{key}_btn_none", use_container_width=True):
            st.session_state[key] = []
            st.rerun()
    
    selected = st.sidebar.multiselect(
        label, 
        options=options, 
        default=st.session_state[key], 
        key=key, 
        label_visibility="collapsed"
    )
    
    # Sincroniza alterações manuais do usuário no multiselect
    if selected != st.session_state[key]:
        st.session_state[key] = selected
    
    return selected

# =============================================================================
# AUTENTICAÇÃO
# =============================================================================
def authenticate_user(email: str, password: str) -> Optional[Dict]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT id, nome, email, tipo, apikey, status FROM usuarios WHERE email = %s AND senha_hash = %s LIMIT 1", (email.lower().strip(), password))
            user = cur.fetchone()
        conn.close(); return user
    except Exception as e:
        logger.error(f"Erro na autenticação: {e}"); return None

def criar_usuario_com_apikey(nome: str, email: str, senha: str, apikey: str) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM usuarios WHERE email = %s", (email.lower().strip(),))
            if cur.fetchone(): conn.close(); return False, "Este e-mail já está cadastrado."
            cur.execute("INSERT INTO usuarios (nome, email, senha_hash, apikey, tipo, status) VALUES (%s, %s, %s, %s, 'User', 'Ativo')", (nome.strip(), email.lower().strip(), senha, apikey))
            conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Conta criada com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
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
            if not usuario: conn.close(); return False, "E-mail não cadastrado no sistema.", None
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
            if not usuario: conn.close(); return False, "Token inválido ou expirado."
            if usuario['reset_token_expires'] < datetime.now(timezone.utc): conn.close(); return False, "Token expirado."
            cur.execute("UPDATE usuarios SET senha_hash = %s, reset_token = NULL, reset_token_expires = NULL WHERE id = %s", (nova_senha, usuario['id']))
            conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Senha redefinida com sucesso!"
    except Exception as e:
        if conn: conn.close()
        return False, f"Erro ao redefinir senha: {str(e)}"

# =============================================================================
# AZURE DEVOPS
# =============================================================================
def get_ado_headers(pat: str) -> Dict[str, str]:
    auth_string = f":{pat}"
    return {"Authorization": f"Basic {base64.b64encode(auth_string.encode('ascii')).decode('ascii')}", "Content-Type": "application/json"}

def sync_projects_from_azure(pat: str) -> List[str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url or not pat: return []
    try:
        url = f"{ADO_BASE_URL}/_apis/projects?api-version=7.1&$top=100"
        session = get_http_session()
        response = session.get(url, headers=get_ado_headers(pat), timeout=20)
        if response.status_code != 200: return []
        azure_projects = response.json().get("value", [])
        project_names = [p["name"] for p in azure_projects]
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            for name in project_names:
                cur.execute("SELECT id FROM devops_projects WHERE name = %s", (name,))
                if not cur.fetchone(): cur.execute("INSERT INTO devops_projects (name, is_active) VALUES (%s, TRUE)", (name,))
        conn.commit(); conn.close(); st.cache_data.clear()
        return project_names
    except Exception as e:
        logger.error(f"Erro ao sincronizar projetos: {e}"); return []

@st.cache_data(ttl=300, show_spinner=False)
def get_active_projects_from_db(pat: str = None) -> List[str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT name FROM devops_projects WHERE is_active = TRUE ORDER BY name")
            projects = [row['name'] for row in cur.fetchall()]
        conn.close()
        if not projects and pat: projects = sync_projects_from_azure(pat)
        return projects
    except Exception as e:
        logger.error(f"Erro ao buscar projetos: {e}")
        return sync_projects_from_azure(pat) if pat else []

@st.cache_data(ttl=600, show_spinner=False)
def validate_ado_api_key(pat: str) -> Tuple[bool, str]:
    if not pat: return False, "API Key vazia."
    url = f"{ADO_BASE_URL}/_apis/projects?api-version=7.1&$top=1"
    try:
        session = get_http_session()
        response = session.get(url, headers=get_ado_headers(pat), timeout=15)
        if response.status_code == 401: return False, "API Azure inválida: autenticação não autorizada (401)."
        elif response.status_code == 403: return False, "API Azure inválida: acesso proibido (403)."
        elif response.status_code == 404: return False, "API Azure inválida: organização não encontrada (404)."
        elif response.status_code != 200: return False, f"API Azure inválida: erro {response.status_code}."
        return True, "API Key válida."
    except Exception as e:
        return False, f"API Azure inválida: {str(e)}"

def parse_date(date_str: str) -> Optional[datetime]:
    if not date_str: return None
    try: return datetime.fromisoformat(date_str.replace('Z', ''))
    except: return None

@st.cache_data(ttl=300, show_spinner=False)
def fetch_all_teams_for_project(pat: str, project: str) -> List[Dict[str, str]]:
    url = f"{ADO_BASE_URL}/_apis/projects/{urllib.parse.quote(project, safe='')}/teams?api-version=7.1&$top=100"
    try:
        session = get_http_session()
        response = session.get(url, headers=get_ado_headers(pat), timeout=15)
        if response.status_code in [401, 403]: return []
        response.raise_for_status()
        teams = response.json().get("value", [])
        return [{"id": t["id"], "name": t["name"]} for t in teams]
    except Exception as e:
        logger.error(f"Erro ao buscar times do projeto {project}: {e}"); return []

@st.cache_data(ttl=300, show_spinner=False)
def fetch_all_sprints_for_project(pat: str, project: str, year: int, month: int) -> List[Dict[str, Any]]:
    teams = fetch_all_teams_for_project(pat, project)
    if not teams: teams = [{"id": "", "name": "Default"}]
    month_start = datetime(year, month, 1, 0, 0, 0)
    if month == 12: month_end = datetime(year + 1, 1, 1, 0, 0, 0) - timedelta(seconds=1)
    else: month_end = datetime(year, month + 1, 1, 0, 0, 0) - timedelta(seconds=1)
    session = get_http_session(); valid_sprints = []; seen_paths = set()
    for team in teams:
        if team["id"]: team_part = f"/{urllib.parse.quote(team['id'], safe='')}"
        else: team_part = ""
        url = f"{ADO_BASE_URL}/{urllib.parse.quote(project, safe='')}{team_part}/_apis/work/teamsettings/iterations?api-version=7.1"
        try:
            response = session.get(url, headers=get_ado_headers(pat), timeout=15)
            if response.status_code in [401, 403]: continue
            response.raise_for_status()
            data = response.json()
            for item in data.get("value", []):
                path = item.get("path")
                if path in seen_paths: continue
                attrs = item.get("attributes", {})
                start_date = parse_date(attrs.get("startDate"))
                finish_date = attrs.get("finishDate")
                if start_date and month_start <= start_date <= month_end:
                    seen_paths.add(path)
                    valid_sprints.append({"project": project, "team": team["name"], "name": item["name"], "path": path, "start": start_date.strftime("%d/%m/%Y"), "finish": (finish_date or "").replace("T", " ").split(".")[0]})
        except Exception as e:
            logger.warning(f"Erro ao buscar sprints do time '{team['name']}' no projeto '{project}': {e}")
            continue
    return valid_sprints

@st.cache_data(ttl=300, show_spinner=False)
def fetch_all_workitems_ids_by_project(pat: str, project: str, sprint_paths: Tuple[str, ...]) -> Tuple[List[int], Optional[str]]:
    if not sprint_paths: return [], "Nenhum path"
    formatted_paths = ", ".join([f"'{p.replace(chr(92), chr(92)+chr(92))}'" for p in sprint_paths])
    wiql_query = f"SELECT [System.Id] FROM WorkItems WHERE [System.IterationPath] IN ({formatted_paths}) ORDER BY [System.Id]"
    try:
        session = get_http_session()
        response = session.post(f"{ADO_BASE_URL}/{urllib.parse.quote(project, safe='')}/_apis/wit/wiql?api-version=7.1", headers=get_ado_headers(pat), json={"query": wiql_query}, timeout=30)
        if response.status_code in [401, 403]: return [], "API Key inválida"
        response.raise_for_status()
        return [w["id"] for w in response.json().get("workItems", [])], None
    except: return [], "Erro"

@st.cache_data(ttl=300, show_spinner=False)
def fetch_hierarchy_links_by_project(pat: str, project: str, sprint_paths: Tuple[str, ...]) -> Tuple[List[Dict], Optional[str]]:
    if not sprint_paths: return [], "Nenhum path"
    formatted_paths = ", ".join([f"'{p.replace(chr(92), chr(92)+chr(92))}'" for p in sprint_paths])
    wiql_query = f"SELECT [System.Id] FROM workitemLinks WHERE ([Source].[System.IterationPath] IN ({formatted_paths})) AND ([System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Forward' OR [System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Reverse') AND ([Target].[System.IterationPath] IN ({formatted_paths})) MODE (MustContain)"
    try:
        session = get_http_session()
        response = session.post(f"{ADO_BASE_URL}/{urllib.parse.quote(project, safe='')}/_apis/wit/wiql?api-version=7.1", headers=get_ado_headers(pat), json={"query": wiql_query}, timeout=30)
        if response.status_code in [401, 403]: return [], "API Key inválida"
        response.raise_for_status()
        return response.json().get("workItemRelations", []), None
    except: return [], "Erro"

@st.cache_data(ttl=300, show_spinner=False)
def fetch_workitems_details_batch(pat: str, project: str, ids: Tuple[int, ...]) -> Tuple[Dict[int, Dict], Optional[str]]:
    if not ids: return {}, "Nenhum ID"
    url = f"{ADO_BASE_URL}/{urllib.parse.quote(project, safe='')}/_apis/wit/workitemsbatch?api-version=7.1"
    fields = ["System.Id", "System.Title", "System.WorkItemType", "System.State", "System.IterationPath",
              "System.CreatedDate", "Microsoft.VSTS.Common.ClosedDate",
              "Microsoft.VSTS.Scheduling.OriginalEstimate", "System.TeamProject",
              "Microsoft.VSTS.Scheduling.CompletedWork", "System.AssignedTo",
              "Microsoft.VSTS.Common.ActivatedDate"]
    all_items = {}
    try:
        session = get_http_session()
        for i in range(0, len(ids), 200):
            response = session.post(url, headers=get_ado_headers(pat), json={"ids": list(ids[i:i+200]), "fields": fields}, timeout=20)
            if response.status_code in [401, 403]: return all_items, "API Key inválida"
            response.raise_for_status()
            for item in response.json().get("value", []): all_items[item["id"]] = item
        return all_items, None
    except: return all_items, "Erro"

# =============================================================================
# HORAS / DIAS ÚTEIS / CAPACITY / VALORES
# =============================================================================
@st.cache_data(ttl=120, show_spinner=False)
def get_horasponto_from_db(ano: int, mes: int) -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'horasponto'")
            colunas_existentes = [row['column_name'] for row in cur.fetchall()]
            colunas_select = ['id_colaborador', 'colaborador_ponto', 'horas_ponto']
            if 'afastamento' in colunas_existentes: colunas_select.append('afastamento')
            if 'Folga' in colunas_existentes: colunas_select.append('"Folga"')
            cur.execute(f"SELECT {', '.join(colunas_select)} FROM horasponto WHERE ano_base = %s AND mes_base = %s", (ano, mes))
            df = pd.DataFrame(cur.fetchall())
        conn.close()
        df.columns = [col.lower() for col in df.columns]
        if 'afastamento' not in df.columns: df['afastamento'] = 0.0
        if 'folga' not in df.columns: df['folga'] = 0.0
        return df
    except Exception as e:
        logger.error(f"Erro ao buscar horasponto: {e}"); return pd.DataFrame()

@st.cache_data(ttl=120, show_spinner=False)
def get_horasponto_by_periodo(ano: int, mes: int) -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return pd.DataFrame()
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'horasponto'")
            colunas_existentes = [row['column_name'] for row in cur.fetchall()]
            colunas_select = ['id', 'ano_base', 'mes_base', 'colaborador_ponto', 'horas_ponto']
            if 'id_colaborador' in colunas_existentes: colunas_select.append('id_colaborador')
            if 'afastamento' in colunas_existentes: colunas_select.append('afastamento')
            if 'Folga' in colunas_existentes: colunas_select.append('"Folga"')
            cur.execute(f"SELECT {', '.join(colunas_select)} FROM horasponto WHERE ano_base = %s AND mes_base = %s ORDER BY colaborador_ponto", (ano, mes))
            rows = cur.fetchall()
        conn.close()
        if not rows: return pd.DataFrame()
        df = pd.DataFrame(rows)
        df.columns = [col.lower() for col in df.columns]
        if 'afastamento' not in df.columns: df['afastamento'] = 0.0
        if 'folga' not in df.columns: df['folga'] = 0.0
        return df
    except Exception as e:
        logger.error(f"Erro ao buscar horasponto do período: {e}"); return pd.DataFrame()

def deletar_horasponto_by_periodo(ano: int, mes: int) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM horasponto WHERE ano_base = %s AND mes_base = %s", (ano, mes))
            deletados = cur.rowcount; conn.commit()
        conn.close(); st.cache_data.clear()
        return True, f"{deletados} registros deletados com sucesso."
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao deletar: {str(e)}"

def inserir_horasponto_lote(registros: List[Dict]) -> Tuple[int, int, List[str]]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url or not registros: return 0, 0, ["Base vazia ou sem URL"]
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=15)
        with conn.cursor() as cur:
            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'horasponto' ORDER BY ordinal_position")
            colunas_existentes = [col[0] for col in cur.fetchall()]
            tem_afastamento = 'afastamento' in colunas_existentes
            tem_folga = 'Folga' in colunas_existentes
            cur.execute("SELECT COALESCE(MAX(id), 0) FROM horasponto")
            start_id = cur.fetchone()[0] + 1
            colunas_insert = ["id", "ano_base", "mes_base", "id_colaborador", "colaborador_ponto", "horas_ponto"]
            if tem_afastamento: colunas_insert.append("afastamento")
            if tem_folga: colunas_insert.append('"Folga"')
            colunas_str = ", ".join(colunas_insert)
            query = f"INSERT INTO horasponto ({colunas_str}) VALUES %s"
            dados_insert = []
            for i, r in enumerate(registros):
                valores = [start_id + i, r['ano'], r['mes'], r['id_colaborador'], r['colaborador_ponto'], r['horas_ponto']]
                if tem_afastamento: valores.append(r.get('afastamento', 0.0))
                if tem_folga: valores.append(r.get('folga', 0.0))
                dados_insert.append(tuple(valores))
            execute_values(cur, query, dados_insert); conn.commit()
        conn.close(); st.cache_data.clear()
        return len(dados_insert), 0, []
    except Exception as e:
        logger.error(f"❌ ERRO FATAL ao inserir: {e}")
        if conn: conn.rollback(); conn.close()
        return 0, len(registros), [str(e)]

@st.cache_data(ttl=600, show_spinner=False)
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

@st.cache_data(ttl=600, show_spinner=False)
def get_valores_from_db() -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT nivel, valor FROM valores"); df = pd.DataFrame(cur.fetchall())
        conn.close(); return df
    except: return pd.DataFrame(columns=["nivel", "valor"])

@st.cache_data(ttl=600, show_spinner=False)
def get_all_capacities() -> Dict[str, float]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return {}
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT nivel, capacity FROM capacity"); rows = cur.fetchall()
        conn.close()
        return {str(row['nivel']).strip(): float(row['capacity']) for row in rows}
    except Exception as e:
        logger.error(f"Erro ao buscar capacities: {e}"); return {}

def limpar_valor_monetario(valor) -> float:
    if valor is None or pd.isna(valor): return 0.0
    valor_str = str(valor).strip().replace('$', '').replace('R$', '').replace(' ', '')
    if ',' in valor_str and '.' in valor_str:
        if valor_str.rfind(',') > valor_str.rfind('.'): valor_str = valor_str.replace('.', '').replace(',', '.')
        else: valor_str = valor_str.replace(',', '')
    elif ',' in valor_str: valor_str = valor_str.replace(',', '.')
    try: return float(valor_str)
    except: return 0.0

# =============================================================================
# USUÁRIOS
# =============================================================================
@st.cache_data(ttl=60, show_spinner=False)
def listar_usuarios(tipo_usuario: str, usuario_id: int) -> List[Dict]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return []
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            if tipo_usuario == 'Admin': cur.execute("SELECT id, nome, email, tipo, status, apikey, criado_em FROM usuarios ORDER BY nome")
            else: cur.execute("SELECT id, nome, email, tipo, status, apikey, criado_em FROM usuarios WHERE id = %s", (usuario_id,))
            usuarios = cur.fetchall()
        conn.close(); return usuarios
    except: return []

@st.cache_data(ttl=300, show_spinner=False)
def listar_gestores_filtro() -> List[Dict]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return []
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT id, nome, email FROM usuarios WHERE tipo = 'Gestor' AND status = 'Ativo' ORDER BY nome")
            gestores = cur.fetchall()
        conn.close(); return gestores
    except Exception as e:
        logger.error(f"Erro ao buscar gestores para filtro: {e}"); return []

def atualizar_usuario(usuario_id: int, dados: Dict, tipo_usuario_logado: str, usuario_logado_id: int) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            if tipo_usuario_logado != 'Admin' and usuario_id != usuario_logado_id: conn.close(); return False, "Você não tem permissão para editar este usuário."
            campos_permitidos = ['nome', 'email', 'senha_hash', 'apikey', 'tipo', 'status'] if tipo_usuario_logado == 'Admin' else ['nome', 'email', 'senha_hash', 'apikey']
            campos_atualizar = {k: v for k, v in dados.items() if k in campos_permitidos and v is not None}
            if not campos_atualizar: conn.close(); return False, "Nenhum campo válido para atualizar."
            set_clause = ", ".join([f"{campo} = %s" for campo in campos_atualizar.keys()])
            valores = list(campos_atualizar.values()) + [usuario_id]
            cur.execute(f"UPDATE usuarios SET {set_clause} WHERE id = %s", valores); conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Dados atualizados com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao atualizar: {str(e)}"

def deletar_usuario(usuario_id: int, tipo_usuario_logado: str) -> Tuple[bool, str]:
    if tipo_usuario_logado != 'Admin': return False, "Apenas administradores podem deletar usuários."
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM usuarios WHERE id = %s", (usuario_id,)); conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Usuário deletado com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao deletar: {str(e)}"

# =============================================================================
# CARGOS
# =============================================================================
@st.cache_data(ttl=300, show_spinner=False)
def listar_cargos() -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return pd.DataFrame()
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT id, descricao FROM cargos ORDER BY descricao")
            rows = cur.fetchall()
        conn.close(); return pd.DataFrame(rows)
    except Exception as e:
        logger.error(f"Erro ao listar cargos: {e}"); return pd.DataFrame()

@st.cache_data(ttl=300, show_spinner=False)
def listar_cargos_options() -> List[Dict]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return []
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT id, descricao FROM cargos ORDER BY descricao")
            rows = cur.fetchall()
        conn.close(); return rows
    except Exception as e:
        logger.error(f"Erro ao buscar cargos options: {e}"); return []

# =============================================================================
# COLABORADORES - COM MOTIVO_DEMISSAO E OBSERVACAO
# =============================================================================
@st.cache_data(ttl=120, show_spinner=False)
def listar_colaboradores() -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return pd.DataFrame()
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT c.id, c.colaborador_azure, c.colaborador_ponto, 
                       c.cargo AS cargo_id, car.descricao AS cargo,
                       c.nivel, c.data_inicio, c.email, c.promocao, c.reposicao, 
                       c.data_saida, c.status, c.gestor, c.email_pessoal, c.cpf, 
                       c.telefone, c.local, c.linkedin, c.github, c.os, 
                       c.dt_nascimento, c.dia_aniv, c.mes_aniv, c.tecnologia,
                       c.motivo_demissao, c.observacao,
                       u.nome AS nome_gestor
                FROM colaboradores c
                LEFT JOIN cargos car ON c.cargo = car.id
                LEFT JOIN usuarios u ON c.gestor = u.id
                ORDER BY c.colaborador_ponto
            """)
            rows = cur.fetchall()
        conn.close(); return pd.DataFrame(rows)
    except Exception as e:
        logger.error(f"Erro ao listar colaboradores: {e}"); return pd.DataFrame()

@st.cache_data(ttl=300, show_spinner=False)
def listar_gestores() -> List[Dict]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return []
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT id, nome, email FROM usuarios WHERE tipo = 'Gestor' AND status = 'Ativo' ORDER BY nome")
            gestores = cur.fetchall()
        conn.close(); return gestores
    except Exception as e:
        logger.error(f"Erro ao buscar gestores: {e}"); return []

def criar_colaborador(dados: Dict) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM colaboradores WHERE colaborador_ponto = %s", (dados.get('colaborador_ponto', '').strip(),))
            if cur.fetchone(): conn.close(); return False, "Já existe um colaborador cadastrado com este nome."
            cur.execute("""
                INSERT INTO colaboradores
                (colaborador_azure, colaborador_ponto, cargo, nivel, data_inicio, email, promocao, reposicao, data_saida, status,
                 gestor, email_pessoal, cpf, telefone, local, linkedin, github, os, dt_nascimento, dia_aniv, mes_aniv, tecnologia,
                 motivo_demissao, observacao)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, 'Ativo',
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                dados.get('colaborador_azure', '').strip(),
                dados.get('colaborador_ponto', '').strip(),
                int(dados['cargo']) if dados.get('cargo') else None,
                dados.get('nivel', '').strip(),
                dados.get('data_inicio') if dados.get('data_inicio') else None,
                dados.get('email', '').strip(),
                dados.get('promocao') if dados.get('promocao') else None,
                dados.get('reposicao', '').strip(),
                dados.get('data_saida') if dados.get('data_saida') else None,
                int(dados['gestor']) if dados.get('gestor') else None,
                dados.get('email_pessoal', '').strip(),
                dados.get('cpf', '').strip(),
                dados.get('telefone', '').strip(),
                dados.get('local', '').strip(),
                dados.get('linkedin', '').strip(),
                dados.get('github', '').strip(),
                dados.get('os', '').strip(),
                dados.get('dt_nascimento') if dados.get('dt_nascimento') else None,
                int(dados['dia_aniv']) if dados.get('dia_aniv') else None,
                int(dados['mes_aniv']) if dados.get('mes_aniv') else None,
                dados.get('tecnologia', '').strip(),
                dados.get('motivo_demissao') if dados.get('motivo_demissao') else None,
                dados.get('observacao') if dados.get('observacao') else None
            ))
            conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Colaborador cadastrado com sucesso!"
    except psycopg2.errors.CheckViolation as e:
        if conn: conn.rollback(); conn.close()
        return False, "Motivo de demissão inválido. Valores aceitos: 'pedido de demissão', 'desligamento', 'fim de alocação', 'outro'"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao cadastrar: {str(e)}"

def atualizar_colaborador(colaborador_id: int, dados: Dict) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("""
                UPDATE colaboradores SET
                colaborador_azure = %s, colaborador_ponto = %s, cargo = %s, nivel = %s,
                data_inicio = %s, email = %s, promocao = %s, reposicao = %s, data_saida = %s, status = %s,
                gestor = %s, email_pessoal = %s, cpf = %s, telefone = %s, local = %s,
                linkedin = %s, github = %s, os = %s, dt_nascimento = %s, dia_aniv = %s, mes_aniv = %s, tecnologia = %s,
                motivo_demissao = %s, observacao = %s
                WHERE id = %s
            """, (
                dados.get('colaborador_azure', '').strip(),
                dados.get('colaborador_ponto', '').strip(),
                int(dados['cargo']) if dados.get('cargo') else None,
                dados.get('nivel', '').strip(),
                dados.get('data_inicio') if dados.get('data_inicio') else None,
                dados.get('email', '').strip(),
                dados.get('promocao') if dados.get('promocao') else None,
                dados.get('reposicao', '').strip(),
                dados.get('data_saida') if dados.get('data_saida') else None,
                dados.get('status', 'Ativo'),
                int(dados['gestor']) if dados.get('gestor') else None,
                dados.get('email_pessoal', '').strip(),
                dados.get('cpf', '').strip(),
                dados.get('telefone', '').strip(),
                dados.get('local', '').strip(),
                dados.get('linkedin', '').strip(),
                dados.get('github', '').strip(),
                dados.get('os', '').strip(),
                dados.get('dt_nascimento') if dados.get('dt_nascimento') else None,
                int(dados['dia_aniv']) if dados.get('dia_aniv') else None,
                int(dados['mes_aniv']) if dados.get('mes_aniv') else None,
                dados.get('tecnologia', '').strip(),
                dados.get('motivo_demissao') if dados.get('motivo_demissao') else None,
                dados.get('observacao') if dados.get('observacao') else None,
                colaborador_id
            ))
            conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Colaborador atualizado com sucesso!"
    except psycopg2.errors.CheckViolation as e:
        if conn: conn.rollback(); conn.close()
        return False, "Motivo de demissão inválido. Valores aceitos: 'pedido de demissão', 'desligamento', 'fim de alocação', 'outro'"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao atualizar: {str(e)}"

def deletar_colaborador(colaborador_id: int) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM colaboradores WHERE id = %s", (colaborador_id,)); conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Colaborador deletado com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao deletar: {str(e)}"

# =============================================================================
# HISTÓRICO DE ALTERAÇÕES
# =============================================================================
@st.cache_data(ttl=60, show_spinner=False)
def buscar_historico_colaborador(id_colaborador: int) -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return pd.DataFrame()
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT cla.id, cla.id_colaborador, cla.campo_alterado,
                       cla.valor_anterior, cla.valor_novo,
                       cla.alterado_em AS data_alteracao,
                       c.colaborador_ponto AS nome_colaborador
                FROM colaboradores_log_alteracoes cla
                LEFT JOIN colaboradores c ON cla.id_colaborador = c.id
                WHERE cla.id_colaborador = %s
                ORDER BY cla.alterado_em DESC
            """, (id_colaborador,))
            rows = cur.fetchall()
        conn.close()
        if rows:
            df = pd.DataFrame(rows)
            if 'data_alteracao' in df.columns:
                df['data_alteracao'] = pd.to_datetime(df['data_alteracao'], errors='coerce').dt.strftime('%d/%m/%Y %H:%M')
            return df
        return pd.DataFrame()
    except Exception as e:
        logger.error(f"Erro ao buscar histórico: {e}")
        if conn: conn.close()
        return pd.DataFrame()

@st.dialog("📜 Histórico de Alterações", width="large")
def mostrar_historico_modal(id_colaborador: int, nome_colaborador: str):
    st.markdown(f"### 👤 {nome_colaborador}")
    st.caption(f"ID do colaborador: **{id_colaborador}**")
    st.markdown("---")
    df_historico = buscar_historico_colaborador(id_colaborador)
    if df_historico.empty:
        st.info("📭 Nenhuma alteração registrada para este colaborador.")
    else:
        st.success(f"📊 **{len(df_historico)}** alteração(ões) encontrada(s)")
        st.markdown("---")
        colunas_desejadas = ['id', 'campo_alterado', 'valor_anterior', 'valor_novo', 'data_alteracao']
        colunas_existentes = [col for col in colunas_desejadas if col in df_historico.columns]
        st.dataframe(
            df_historico[colunas_existentes],
            width="stretch", hide_index=True,
            column_config={
                "id": st.column_config.NumberColumn("ID", width="small"),
                "campo_alterado": st.column_config.TextColumn("Campo Alterado", width="medium"),
                "valor_anterior": st.column_config.TextColumn("Valor Anterior", width="large"),
                "valor_novo": st.column_config.TextColumn("Valor Novo", width="large"),
                "data_alteracao": st.column_config.TextColumn("Data/Hora", width="medium")
            }
        )
    st.markdown("---")
    if st.button("✖️ Fechar", width="stretch", key="btn_fechar_historico"):
        st.rerun()

# =============================================================================
# DESTAQUES DA SEMANA (CARROSSEL)
# =============================================================================
def render_destaques_semana(df_colaboradores):
    if df_colaboradores is None or df_colaboradores.empty:
        return
    df_ativos = df_colaboradores.copy()
    if 'status' in df_ativos.columns:
        df_ativos = df_ativos[df_ativos['status'].astype(str).str.strip().str.lower() == 'ativo']
    if df_ativos.empty:
        return

    if 'carrossel_semana_offset' not in st.session_state:
        st.session_state.carrossel_semana_offset = 0

    offset = st.session_state.carrossel_semana_offset
    hoje = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    dia_semana = hoje.weekday()
    dias_para_domingo = 0 if dia_semana == 6 else dia_semana + 1
    inicio_semana_base = hoje - timedelta(days=dias_para_domingo)
    inicio_semana = inicio_semana_base + timedelta(weeks=offset)
    fim_semana = inicio_semana + timedelta(days=6)
    periodo_texto = f"{inicio_semana.strftime('%d/%m/%Y')} a {fim_semana.strftime('%d/%m/%Y')}"

    aniversariantes_semana = []
    for _, row in df_ativos.iterrows():
        dia_aniv, mes_aniv = row.get('dia_aniv'), row.get('mes_aniv')
        if pd.isna(dia_aniv) or pd.isna(mes_aniv):
            continue
        try:
            dia_aniv, mes_aniv = int(dia_aniv), int(mes_aniv)
            try:
                data_aniv_este_ano = datetime(inicio_semana.year, mes_aniv, dia_aniv)
            except ValueError:
                continue
            try:
                data_aniv_prox_ano = datetime(inicio_semana.year + 1, mes_aniv, dia_aniv)
            except ValueError:
                data_aniv_prox_ano = None
            data_valida = None
            if inicio_semana <= data_aniv_este_ano <= fim_semana:
                data_valida = data_aniv_este_ano
            elif data_aniv_prox_ano and inicio_semana <= data_aniv_prox_ano <= fim_semana:
                data_valida = data_aniv_prox_ano
            if data_valida:
                aniversariantes_semana.append({
                    'nome': row.get('colaborador_ponto', ''),
                    'cargo': row.get('cargo', ''),
                    'data_aniversario': data_valida
                })
        except Exception:
            continue
    aniversariantes_semana.sort(key=lambda x: x['data_aniversario'])

    aniversariantes_casa = []
    for _, row in df_ativos.iterrows():
        raw_data_inicio = row.get('data_inicio')
        if pd.isna(raw_data_inicio):
            continue
        try:
            data_inicio = pd.to_datetime(raw_data_inicio, dayfirst=True).to_pydatetime()
            data_inicio = data_inicio.replace(hour=0, minute=0, second=0, microsecond=0)
        except Exception:
            continue
        dia_inicio, mes_inicio, ano_inicio = data_inicio.day, data_inicio.month, data_inicio.year
        try:
            data_casa_este_ano = datetime(inicio_semana.year, mes_inicio, dia_inicio)
        except ValueError:
            continue
        try:
            data_casa_prox_ano = datetime(inicio_semana.year + 1, mes_inicio, dia_inicio)
        except ValueError:
            data_casa_prox_ano = None
        data_valida = None
        ano_referencia = inicio_semana.year
        if inicio_semana <= data_casa_este_ano <= fim_semana:
            data_valida = data_casa_este_ano
        elif data_casa_prox_ano and inicio_semana <= data_casa_prox_ano <= fim_semana:
            data_valida = data_casa_prox_ano
            ano_referencia = inicio_semana.year + 1
        if data_valida:
            anos_casa = ano_referencia - ano_inicio
            if anos_casa >= 1:
                aniversariantes_casa.append({
                    'nome': row.get('colaborador_ponto', ''),
                    'cargo': row.get('cargo', ''),
                    'anos_casa': anos_casa,
                    'data_aniversario': data_valida
                })
    aniversariantes_casa.sort(key=lambda x: x['data_aniversario'])

    st.markdown("### 🌟 Destaques da Semana")
    col_nav1, col_nav2, col_nav3 = st.columns([1, 4, 1])
    with col_nav1:
        pode_voltar = offset > 0
        if st.button("◀️ Anterior", disabled=not pode_voltar, use_container_width=True, key="btn_carrossel_anterior"):
            st.session_state.carrossel_semana_offset = offset - 1
            st.rerun()
    with col_nav2:
        st.markdown(f"**📅 {periodo_texto}**")
        if offset == 0:
            st.success("✅ Semana Atual")
        else:
            st.info(f"📆 Semana +{offset}")
    with col_nav3:
        pode_avancar = offset < 4
        if st.button("Próxima ▶️", disabled=not pode_avancar, use_container_width=True, key="btn_carrossel_proxima"):
            st.session_state.carrossel_semana_offset = offset + 1
            st.rerun()
    st.markdown("---")

    col_aniv, col_casa = st.columns(2)
    with col_aniv:
        html_aniv = '<div class="destaque-card"><div style="display:flex;align-items:center;gap:10px;margin-bottom:15px;"><span style="font-size:24px;">🎂</span><div><div class="destaque-titulo">Aniversários</div><div class="destaque-subtitulo">Aniversariantes da semana</div></div></div>'
        if not aniversariantes_semana:
            html_aniv += '<p style="color:#666;font-size:14px;">Nenhum aniversariante nesta semana.</p>'
        else:
            for aniv in aniversariantes_semana:
                dias_restantes = (aniv['data_aniversario'] - hoje).days
                if dias_restantes == 0:
                    badge = '<span class="badge-hoje">HOJE!</span>'
                elif dias_restantes < 0:
                    badge = f'<span class="badge-semana">Há {abs(dias_restantes)} dia(s)</span>'
                elif dias_restantes == 1:
                    badge = '<span class="badge-amanha">Amanhã</span>'
                else:
                    badge = f'<span class="badge-semana">Em {dias_restantes} dias</span>'
                data_formatada = aniv['data_aniversario'].strftime("%d/%m")
                html_aniv += f'<div class="destaque-item"><div style="display:flex;justify-content:space-between;align-items:center;"><div><div class="destaque-nome">🎉 {aniv["nome"]}</div><div class="destaque-cargo">{aniv["cargo"]}</div></div><div style="text-align:right;">{badge}<div class="destaque-data">{data_formatada}</div></div></div></div>'
        html_aniv += '</div>'
        st.markdown(html_aniv, unsafe_allow_html=True)

    with col_casa:
        html_casa = '<div class="destaque-card"><div style="display:flex;align-items:center;gap:10px;margin-bottom:15px;"><span style="font-size:24px;">🏢</span><div><div class="destaque-titulo">Tempo de Casa</div><div class="destaque-subtitulo">Aniversariantes de casa da semana</div></div></div>'
        if not aniversariantes_casa:
            html_casa += '<p style="color:#666;font-size:14px;">Nenhum aniversário de casa nesta semana.</p>'
        else:
            for casa in aniversariantes_casa:
                dias_restantes = (casa['data_aniversario'] - hoje).days
                if dias_restantes == 0:
                    badge = '<span class="badge-hoje">HOJE!</span>'
                elif dias_restantes < 0:
                    badge = f'<span class="badge-semana">Há {abs(dias_restantes)} dia(s)</span>'
                elif dias_restantes == 1:
                    badge = '<span class="badge-amanha">Amanhã</span>'
                else:
                    badge = f'<span class="badge-semana">Em {dias_restantes} dias</span>'
                data_formatada = casa['data_aniversario'].strftime("%d/%m")
                html_casa += f'<div class="destaque-item"><div style="display:flex;justify-content:space-between;align-items:center;"><div><div class="destaque-nome">🏆 {casa["nome"]}</div><div class="destaque-cargo">{casa["cargo"]} • <strong>{casa["anos_casa"]} ano(s)</strong></div></div><div style="text-align:right;">{badge}<div class="destaque-data">{data_formatada}</div></div></div></div>'
        html_casa += '</div>'
        st.markdown(html_casa, unsafe_allow_html=True)
    st.markdown("---")

# =============================================================================
# COLAB_PROJETOS
# =============================================================================
def listar_colab_projetos() -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return pd.DataFrame()
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT cp.id, cp.id_colab, cp.id_projeto, cp.created_at,
                       c.colaborador_ponto AS nome_colaborador, car.descricao AS cargo, p.name AS nome_projeto
                FROM colab_projetos cp
                LEFT JOIN colaboradores c ON cp.id_colab = c.id
                LEFT JOIN cargos car ON c.cargo = car.id
                LEFT JOIN devops_projects p ON cp.id_projeto = p.id
                ORDER BY c.colaborador_ponto, p.name
            """)
            rows = cur.fetchall()
        conn.close(); return pd.DataFrame(rows)
    except Exception as e:
        logger.error(f"Erro ao listar colab_projetos: {e}"); return pd.DataFrame()

@st.cache_data(ttl=300, show_spinner=False)
def listar_colaboradores_options() -> List[Dict]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return []
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT c.id, c.colaborador_ponto, car.descricao AS cargo 
                FROM colaboradores c
                LEFT JOIN cargos car ON c.cargo = car.id
                WHERE c.status = 'Ativo' ORDER BY c.colaborador_ponto
            """)
            rows = cur.fetchall()
        conn.close(); return rows
    except Exception as e:
        logger.error(f"Erro ao buscar colaboradores options: {e}"); return []

@st.cache_data(ttl=300, show_spinner=False)
def listar_projetos_options() -> List[Dict]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return []
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT id, name FROM devops_projects WHERE is_active = TRUE ORDER BY name")
            rows = cur.fetchall()
        conn.close(); return rows
    except Exception as e:
        logger.error(f"Erro ao buscar projetos options: {e}"); return []

def criar_colab_projeto(dados: Dict) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM colab_projetos WHERE id_colab = %s AND id_projeto = %s", (int(dados['id_colab']), int(dados['id_projeto'])))
            if cur.fetchone(): conn.close(); return False, "Este colaborador já está associado a este projeto."
            cur.execute("INSERT INTO colab_projetos (id_colab, id_projeto) VALUES (%s, %s)", (int(dados['id_colab']), int(dados['id_projeto'])))
            conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Associação criada com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao criar: {str(e)}"

def atualizar_colab_projeto(colab_projeto_id: int, dados: Dict) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM colab_projetos WHERE id_colab = %s AND id_projeto = %s AND id != %s", (int(dados['id_colab']), int(dados['id_projeto']), colab_projeto_id))
            if cur.fetchone(): conn.close(); return False, "Já existe associação com este colaborador e projeto."
            cur.execute("UPDATE colab_projetos SET id_colab = %s, id_projeto = %s WHERE id = %s", (int(dados['id_colab']), int(dados['id_projeto']), colab_projeto_id))
            conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Associação atualizada com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao atualizar: {str(e)}"

def deletar_colab_projeto(colab_projeto_id: int) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM colab_projetos WHERE id = %s", (colab_projeto_id,)); conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Associação removida com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao deletar: {str(e)}"

# =============================================================================
# COLAB_SQUAD
# =============================================================================
def listar_colab_squad() -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return pd.DataFrame()
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT cs.id, cs.id_colab, cs.id_squad, cs.created_at,
                       c.colaborador_ponto AS nome_colaborador,
                       car.descricao AS cargo,
                       s.descricao AS nome_squad,
                       co.descricao AS nome_coordenadoria
                FROM colab_squad cs
                LEFT JOIN colaboradores c ON cs.id_colab = c.id
                LEFT JOIN cargos car ON c.cargo = car.id
                LEFT JOIN squads s ON cs.id_squad = s.id
                LEFT JOIN coordenadoria_squad csq ON cs.id_squad = csq.id_squad
                LEFT JOIN coordenadorias co ON csq.id_coordenadoria = co.id
                ORDER BY c.colaborador_ponto, co.descricao, s.descricao
            """)
            rows = cur.fetchall()
        conn.close(); return pd.DataFrame(rows)
    except Exception as e:
        logger.error(f"Erro ao listar colab_squad: {e}"); return pd.DataFrame()

def criar_colab_squad(id_colab: int, id_squad: int) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM colab_squad WHERE id_colab = %s AND id_squad = %s", (int(id_colab), int(id_squad)))
            if cur.fetchone(): conn.close(); return False, "Este colaborador já está associado a este squad."
            cur.execute("INSERT INTO colab_squad (id_colab, id_squad) VALUES (%s, %s)", (int(id_colab), int(id_squad)))
            conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Associação criada com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao criar: {str(e)}"

def deletar_colab_squad(cs_id: int) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM colab_squad WHERE id = %s", (cs_id,)); conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Associação removida com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao deletar: {str(e)}"

@st.cache_data(ttl=120, show_spinner=False)
def buscar_distribuicao_squad() -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return pd.DataFrame()
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT s.descricao AS squad, COUNT(DISTINCT cs.id_colab) AS total
                FROM colab_squad cs
                LEFT JOIN squads s ON cs.id_squad = s.id
                LEFT JOIN colaboradores c ON cs.id_colab = c.id
                WHERE c.status = 'Ativo'
                GROUP BY s.descricao
                ORDER BY total DESC
            """)
            rows = cur.fetchall()
        conn.close(); return pd.DataFrame(rows)
    except Exception as e:
        logger.error(f"Erro ao buscar distribuição por squad: {e}"); return pd.DataFrame()

# =============================================================================
# REUNIÕES 1:1
# =============================================================================
@st.cache_data(ttl=120, show_spinner=False)
def listar_reunioes_11() -> pd.DataFrame:
    """
    Busca todas as reuniões 1:1 com:
    - Colaborador (colaborador_ponto da tabela colaboradores)
    - Gestor (nome da tabela usuarios)
    - Squads (descricao da tabela squads agregadas por vírgula via colab_squad)
    - data_reuniao, Info_Senior, Realizada da tabela colab_reuniao11
    """
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url:
        return pd.DataFrame()
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT 
                    r.id,
                    r.id_colab,
                    c.colaborador_ponto AS colaborador,
                    u.nome AS gestor,
                    COALESCE(
                        STRING_AGG(DISTINCT s.descricao, ', ' ORDER BY s.descricao), 
                        '-'
                    ) AS squads,
                    r.data_reuniao,
                    COALESCE(r."Info_Senior", false) AS "Info_Senior",
                    COALESCE(r."Realizada", false) AS "Realizada",
                    r."Observacao" AS observacao,
                    r.created_at
                FROM colab_reuniao11 r
                INNER JOIN colaboradores c ON r.id_colab = c.id
                LEFT JOIN usuarios u ON c.gestor = u.id
                LEFT JOIN colab_squad cs ON c.id = cs.id_colab
                LEFT JOIN squads s ON cs.id_squad = s.id
                GROUP BY 
                    r.id, 
                    r.id_colab, 
                    c.colaborador_ponto, 
                    u.nome, 
                    r.data_reuniao, 
                    r."Info_Senior", 
                    r."Realizada", 
                    r."Observacao", 
                    r.created_at
                ORDER BY r.data_reuniao DESC, c.colaborador_ponto ASC
            """)
            rows = cur.fetchall()
        conn.close()
        return pd.DataFrame(rows)
    except Exception as e:
        logger.error(f"Erro ao listar reuniões 1:1: {e}")
        return pd.DataFrame()


def criar_reuniao_11(id_colab: int, data_reuniao, info_senior: bool = False, realizada: bool = False, observacao: str = "") -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url:
        return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO colab_reuniao11 (id_colab, data_reuniao, "Info_Senior", "Realizada", "Observacao")
                VALUES (%s, %s, %s, %s, %s)
            """, (
                int(id_colab),
                data_reuniao,
                bool(info_senior),
                bool(realizada),
                observacao.strip() if observacao else None
            ))
            conn.commit()
        conn.close()
        st.cache_data.clear()
        return True, "Reunião 1:1 registrada com sucesso!"
    except Exception as e:
        if conn:
            conn.rollback()
            conn.close()
        return False, f"Erro ao criar reunião 1:1: {str(e)}"


def atualizar_reuniao_11(reuniao_id: int, id_colab: int, data_reuniao, info_senior: bool, realizada: bool, observacao: str = "") -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url:
        return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("""
                UPDATE colab_reuniao11
                SET id_colab = %s, data_reuniao = %s, "Info_Senior" = %s, "Realizada" = %s, "Observacao" = %s
                WHERE id = %s
            """, (
                int(id_colab),
                data_reuniao,
                bool(info_senior),
                bool(realizada),
                observacao.strip() if observacao else None,
                int(reuniao_id)
            ))
            conn.commit()
        conn.close()
        st.cache_data.clear()
        return True, "Reunião 1:1 atualizada com sucesso!"
    except Exception as e:
        if conn:
            conn.rollback()
            conn.close()
        return False, f"Erro ao atualizar reunião 1:1: {str(e)}"


def deletar_reuniao_11(reuniao_id: int) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url:
        return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM colab_reuniao11 WHERE id = %s", (int(reuniao_id),))
            conn.commit()
        conn.close()
        st.cache_data.clear()
        return True, "Reunião 1:1 excluída com sucesso!"
    except Exception as e:
        if conn:
            conn.rollback()
            conn.close()
        return False, f"Erro ao excluir reunião 1:1: {str(e)}"

# =============================================================================
# SQUADS
# =============================================================================
def listar_squads() -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return pd.DataFrame()
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT id, descricao, ativo FROM squads ORDER BY descricao")
            rows = cur.fetchall()
        conn.close(); return pd.DataFrame(rows)
    except Exception as e:
        logger.error(f"Erro ao listar squads: {e}"); return pd.DataFrame()

def criar_squad(descricao: str, ativo: bool = True) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM squads WHERE descricao = %s", (descricao.strip(),))
            if cur.fetchone(): conn.close(); return False, "Já existe um squad com esta descrição."
            cur.execute("INSERT INTO squads (descricao, ativo) VALUES (%s, %s)", (descricao.strip(), ativo))
            conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Squad criado com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao criar: {str(e)}"

def atualizar_squad(squad_id: int, descricao: str, ativo: bool) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM squads WHERE descricao = %s AND id != %s", (descricao.strip(), squad_id))
            if cur.fetchone(): conn.close(); return False, "Já existe outro squad com esta descrição."
            cur.execute("UPDATE squads SET descricao = %s, ativo = %s WHERE id = %s", (descricao.strip(), ativo, squad_id))
            conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Squad atualizado com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao atualizar: {str(e)}"

def deletar_squad(squad_id: int) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM squads WHERE id = %s", (squad_id,)); conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Squad removido com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao deletar: {str(e)}"

def render_squads_tab():
    if 'squad_editando' not in st.session_state: st.session_state.squad_editando = None
    if 'mostrar_form_squad' not in st.session_state: st.session_state.mostrar_form_squad = False
    if 'mostrar_confirmacao_exclusao_squad' not in st.session_state: st.session_state.mostrar_confirmacao_exclusao_squad = None
    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn2:
        if st.button("➕ Novo Squad", width="stretch", type="primary", key="btn_novo_squad"):
            st.session_state.squad_editando = None; st.session_state.mostrar_form_squad = True; st.rerun()
    df_squads = listar_squads()
    if df_squads.empty:
        st.warning("⚠️ Nenhum squad cadastrado. Clique em **➕ Novo Squad** para começar.")
    else:
        st.markdown(f"**Total de squads cadastrados:** {len(df_squads)}")
        st.dataframe(df_squads[['id', 'descricao', 'ativo']], width="stretch", hide_index=True,
            column_config={"id": st.column_config.NumberColumn("ID", width="small"), "descricao": st.column_config.TextColumn("Descrição", width="large"), "ativo": st.column_config.CheckboxColumn("Ativo")})
        st.markdown("---"); st.markdown("### Ações")
        options = [f"{row['id']} - {row['descricao']} ({'Ativo' if row['ativo'] else 'Inativo'})" for _, row in df_squads.iterrows()]
        if options:
            squad_selecionado = st.selectbox("Selecione um squad para editar ou excluir:", options=options, index=0)
            if squad_selecionado:
                squad_id = int(squad_selecionado.split(" - ")[0])
                squad_dados = df_squads[df_squads['id'] == squad_id].iloc[0].to_dict()
                col_edit, col_del = st.columns(2)
                with col_edit:
                    if st.button("✏️ Editar Squad", width="stretch", key="btn_edit_squad"):
                        st.session_state.squad_editando = squad_dados; st.session_state.mostrar_form_squad = True; st.rerun()
                with col_del:
                    if st.button("🗑️ Excluir Squad", width="stretch", type="secondary", key="btn_del_squad"):
                        st.session_state.mostrar_confirmacao_exclusao_squad = squad_id; st.rerun()
    if st.session_state.get('mostrar_confirmacao_exclusao_squad'):
        squad_id = st.session_state.mostrar_confirmacao_exclusao_squad
        squad_row = df_squads[df_squads['id'] == squad_id]
        if not squad_row.empty:
            squad_nome = squad_row.iloc[0]['descricao']
            st.warning(f"⚠️ Você está prestes a excluir o squad: **{squad_nome}**")
            col_sim, col_nao = st.columns(2)
            with col_sim:
                if st.button("✅ Sim, Excluir", width="stretch", type="primary", key="btn_sim_excluir_squad"):
                    sucesso, msg = deletar_squad(squad_id)
                    if sucesso: st.success(f"✅ {msg}"); st.session_state.mostrar_confirmacao_exclusao_squad = None; st.rerun()
                    else: st.error(f"❌ {msg}")
            with col_nao:
                if st.button("❌ Não, Cancelar", width="stretch", key="btn_nao_excluir_squad"): st.session_state.mostrar_confirmacao_exclusao_squad = None; st.rerun()
    if st.session_state.get('mostrar_form_squad'):
        st.markdown("---"); st.subheader("📝 " + ("Editar Squad" if st.session_state.squad_editando else "Novo Squad"))
        squad = st.session_state.squad_editando or {}
        with st.form("form_squad"):
            descricao = st.text_input("Descrição *", value=squad.get('descricao', ''), placeholder="Ex: Squad de Desenvolvimento Frontend")
            ativo = st.checkbox("Ativo", value=squad.get('ativo', True))
            st.markdown("---")
            col_submit, col_cancel = st.columns(2)
            with col_submit: submitted = st.form_submit_button("💾 Salvar", width="stretch", type="primary", key="btn_salvar_squad")
            with col_cancel: cancelled = st.form_submit_button("❌ Cancelar", width="stretch", key="btn_cancelar_squad")
            if cancelled: st.session_state.squad_editando = None; st.session_state.mostrar_form_squad = False; st.rerun()
            if submitted:
                if not descricao.strip(): st.error("⚠️ Descrição é obrigatória")
                else:
                    if st.session_state.squad_editando: sucesso, msg = atualizar_squad(st.session_state.squad_editando['id'], descricao, ativo)
                    else: sucesso, msg = criar_squad(descricao, ativo)
                    if sucesso: st.success(f"✅ {msg}"); st.session_state.squad_editando = None; st.session_state.mostrar_form_squad = False; st.rerun()
                    else: st.error(f"❌ {msg}")

# =============================================================================
# COORDENADORIAS
# =============================================================================
def listar_coordenadorias() -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return pd.DataFrame()
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT id, descricao, ativo FROM coordenadorias ORDER BY descricao")
            rows = cur.fetchall()
        conn.close(); return pd.DataFrame(rows)
    except Exception as e:
        logger.error(f"Erro ao listar coordenadorias: {e}"); return pd.DataFrame()

def criar_coordenadoria(descricao: str, ativo: bool = True) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM coordenadorias WHERE descricao = %s", (descricao.strip(),))
            if cur.fetchone(): conn.close(); return False, "Já existe uma coordenadoria com esta descrição."
            cur.execute("INSERT INTO coordenadorias (descricao, ativo) VALUES (%s, %s)", (descricao.strip(), ativo))
            conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Coordenadoria criada com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao criar: {str(e)}"

def atualizar_coordenadoria(coord_id: int, descricao: str, ativo: bool) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM coordenadorias WHERE descricao = %s AND id != %s", (descricao.strip(), coord_id))
            if cur.fetchone(): conn.close(); return False, "Já existe outra coordenadoria com esta descrição."
            cur.execute("UPDATE coordenadorias SET descricao = %s, ativo = %s WHERE id = %s", (descricao.strip(), ativo, coord_id))
            conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Coordenadoria atualizada com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao atualizar: {str(e)}"

def deletar_coordenadoria(coord_id: int) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM coordenadorias WHERE id = %s", (coord_id,)); conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Coordenadoria removida com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao deletar: {str(e)}"

def render_coordenadorias_tab():
    if 'coord_editando' not in st.session_state: st.session_state.coord_editando = None
    if 'mostrar_form_coord' not in st.session_state: st.session_state.mostrar_form_coord = False
    if 'mostrar_confirmacao_exclusao_coord' not in st.session_state: st.session_state.mostrar_confirmacao_exclusao_coord = None
    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn2:
        if st.button("➕ Nova Coordenadoria", width="stretch", type="primary", key="btn_nova_coord"):
            st.session_state.coord_editando = None; st.session_state.mostrar_form_coord = True; st.rerun()
    df_coords = listar_coordenadorias()
    if df_coords.empty:
        st.warning("⚠️ Nenhuma coordenadoria cadastrada. Clique em **➕ Nova Coordenadoria** para começar.")
    else:
        st.markdown(f"**Total de coordenadorias cadastradas:** {len(df_coords)}")
        st.dataframe(df_coords[['id', 'descricao', 'ativo']], width="stretch", hide_index=True,
            column_config={"id": st.column_config.NumberColumn("ID", width="small"), "descricao": st.column_config.TextColumn("Descrição", width="large"), "ativo": st.column_config.CheckboxColumn("Ativo")})
        st.markdown("---"); st.markdown("### Ações")
        options = [f"{row['id']} - {row['descricao']} ({'Ativo' if row['ativo'] else 'Inativo'})" for _, row in df_coords.iterrows()]
        if options:
            coord_selecionada = st.selectbox("Selecione uma coordenadoria para editar ou excluir:", options=options, index=0)
            if coord_selecionada:
                coord_id = int(coord_selecionada.split(" - ")[0])
                coord_dados = df_coords[df_coords['id'] == coord_id].iloc[0].to_dict()
                col_edit, col_del = st.columns(2)
                with col_edit:
                    if st.button("✏️ Editar Coordenadoria", width="stretch", key="btn_edit_coord"):
                        st.session_state.coord_editando = coord_dados; st.session_state.mostrar_form_coord = True; st.rerun()
                with col_del:
                    if st.button("🗑️ Excluir Coordenadoria", width="stretch", type="secondary", key="btn_del_coord"):
                        st.session_state.mostrar_confirmacao_exclusao_coord = coord_id; st.rerun()
    if st.session_state.get('mostrar_confirmacao_exclusao_coord'):
        coord_id = st.session_state.mostrar_confirmacao_exclusao_coord
        coord_row = df_coords[df_coords['id'] == coord_id]
        if not coord_row.empty:
            coord_nome = coord_row.iloc[0]['descricao']
            st.warning(f"⚠️ Você está prestes a excluir a coordenadoria: **{coord_nome}**")
            col_sim, col_nao = st.columns(2)
            with col_sim:
                if st.button("✅ Sim, Excluir", width="stretch", type="primary", key="btn_sim_excluir_coord"):
                    sucesso, msg = deletar_coordenadoria(coord_id)
                    if sucesso: st.success(f"✅ {msg}"); st.session_state.mostrar_confirmacao_exclusao_coord = None; st.rerun()
                    else: st.error(f"❌ {msg}")
            with col_nao:
                if st.button("❌ Não, Cancelar", width="stretch", key="btn_nao_excluir_coord"): st.session_state.mostrar_confirmacao_exclusao_coord = None; st.rerun()
    if st.session_state.get('mostrar_form_coord'):
        st.markdown("---"); st.subheader("📝 " + ("Editar Coordenadoria" if st.session_state.coord_editando else "Nova Coordenadoria"))
        coord = st.session_state.coord_editando or {}
        with st.form("form_coordenadoria"):
            descricao = st.text_input("Descrição *", value=coord.get('descricao', ''), placeholder="Ex: Coordenadoria de Tecnologia")
            ativo = st.checkbox("Ativo", value=coord.get('ativo', True))
            st.markdown("---")
            col_submit, col_cancel = st.columns(2)
            with col_submit: submitted = st.form_submit_button("💾 Salvar", width="stretch", type="primary", key="btn_salvar_coord")
            with col_cancel: cancelled = st.form_submit_button("❌ Cancelar", width="stretch", key="btn_cancelar_coord")
            if cancelled: st.session_state.coord_editando = None; st.session_state.mostrar_form_coord = False; st.rerun()
            if submitted:
                if not descricao.strip(): st.error("⚠️ Descrição é obrigatória")
                else:
                    if st.session_state.coord_editando: sucesso, msg = atualizar_coordenadoria(st.session_state.coord_editando['id'], descricao, ativo)
                    else: sucesso, msg = criar_coordenadoria(descricao, ativo)
                    if sucesso: st.success(f"✅ {msg}"); st.session_state.coord_editando = None; st.session_state.mostrar_form_coord = False; st.rerun()
                    else: st.error(f"❌ {msg}")

# =============================================================================
# COORDENADORIA_SQUAD
# =============================================================================
def listar_coordenadoria_squad() -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return pd.DataFrame()
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT cs.id, cs.id_coordenadoria, cs.id_squad, cs.created_at,
                       c.descricao AS nome_coordenadoria, s.descricao AS nome_squad
                FROM coordenadoria_squad cs
                LEFT JOIN coordenadorias c ON cs.id_coordenadoria = c.id
                LEFT JOIN squads s ON cs.id_squad = s.id
                ORDER BY c.descricao, s.descricao
            """)
            rows = cur.fetchall()
        conn.close(); return pd.DataFrame(rows)
    except Exception as e:
        logger.error(f"Erro ao listar coordenadoria_squad: {e}"); return pd.DataFrame()

def criar_coordenadoria_squad(id_coordenadoria: int, id_squad: int) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM coordenadoria_squad WHERE id_coordenadoria = %s AND id_squad = %s", (int(id_coordenadoria), int(id_squad)))
            if cur.fetchone(): conn.close(); return False, "Esta coordenadoria já está associada a este squad."
            cur.execute("INSERT INTO coordenadoria_squad (id_coordenadoria, id_squad) VALUES (%s, %s)", (int(id_coordenadoria), int(id_squad)))
            conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Associação criada com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao criar: {str(e)}"

def deletar_coordenadoria_squad(cs_id: int) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM coordenadoria_squad WHERE id = %s", (cs_id,)); conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Associação removida com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao deletar: {str(e)}"

def listar_coordenadorias_options() -> List[Dict]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return []
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT id, descricao FROM coordenadorias WHERE ativo = TRUE ORDER BY descricao")
            rows = cur.fetchall()
        conn.close(); return rows
    except Exception as e:
        logger.error(f"Erro ao buscar coordenadorias options: {e}"); return []

def listar_squads_options() -> List[Dict]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return []
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT id, descricao FROM squads WHERE ativo = TRUE ORDER BY descricao")
            rows = cur.fetchall()
        conn.close(); return rows
    except Exception as e:
        logger.error(f"Erro ao buscar squads options: {e}"); return []

def render_coordenadoria_squad_tab():
    if 'mostrar_form_cs' not in st.session_state: st.session_state.mostrar_form_cs = False
    if 'mostrar_confirmacao_exclusao_cs' not in st.session_state: st.session_state.mostrar_confirmacao_exclusao_cs = None
    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn2:
        if st.button("➕ Nova Associação", width="stretch", type="primary", key="btn_nova_assoc_cs"):
            st.session_state.mostrar_form_cs = True; st.rerun()
    df_cs = listar_coordenadoria_squad()
    if df_cs.empty:
        st.warning("⚠️ Nenhuma associação coordenadoria-squad cadastrada. Clique em **➕ Nova Associação** para começar.")
    else:
        st.markdown(f"**Total de associações:** {len(df_cs)}")
        st.dataframe(df_cs[['id', 'nome_coordenadoria', 'nome_squad', 'created_at']], width="stretch", hide_index=True,
            column_config={"id": st.column_config.NumberColumn("ID", width="small"),
                "nome_coordenadoria": st.column_config.TextColumn("Coordenadoria", width="large"),
                "nome_squad": st.column_config.TextColumn("Squad", width="large"),
                "created_at": st.column_config.DatetimeColumn("Criado em", format="DD/MM/YYYY HH:mm")})
        st.markdown("---"); st.markdown("### Ações")
        options = [f"{row['id']} - {row['nome_coordenadoria']} ↔ {row['nome_squad']}" for _, row in df_cs.iterrows()]
        if options:
            cs_selecionada = st.selectbox("Selecione uma associação para excluir:", options=options, index=0)
            if cs_selecionada:
                cs_id = int(cs_selecionada.split(" - ")[0])
                if st.button("🗑️ Excluir Associação", width="stretch", type="secondary", key="btn_del_assoc_cs"):
                    st.session_state.mostrar_confirmacao_exclusao_cs = cs_id; st.rerun()
    if st.session_state.get('mostrar_confirmacao_exclusao_cs'):
        cs_id = st.session_state.mostrar_confirmacao_exclusao_cs
        cs_row = df_cs[df_cs['id'] == cs_id]
        if not cs_row.empty:
            cs_nome = f"{cs_row.iloc[0]['nome_coordenadoria']} ↔ {cs_row.iloc[0]['nome_squad']}"
            st.warning(f"⚠️ Você está prestes a excluir a associação: **{cs_nome}**")
            col_sim, col_nao = st.columns(2)
            with col_sim:
                if st.button("✅ Sim, Excluir", width="stretch", type="primary", key="btn_sim_excluir_cs"):
                    sucesso, msg = deletar_coordenadoria_squad(cs_id)
                    if sucesso: st.success(f"✅ {msg}"); st.session_state.mostrar_confirmacao_exclusao_cs = None; st.rerun()
                    else: st.error(f"❌ {msg}")
            with col_nao:
                if st.button("❌ Não, Cancelar", width="stretch", key="btn_nao_excluir_cs"):
                    st.session_state.mostrar_confirmacao_exclusao_cs = None; st.rerun()
    if st.session_state.get('mostrar_form_cs'):
        st.markdown("---"); st.subheader("📝 Nova Associação Coordenadoria-Squad")
        coordenadorias_options = listar_coordenadorias_options()
        squads_options = listar_squads_options()
        if not coordenadorias_options: st.error("⚠️ Nenhuma coordenadoria ativa cadastrada."); return
        if not squads_options: st.error("⚠️ Nenhum squad ativo cadastrado."); return
        with st.form("form_coordenadoria_sq"):
            st.markdown("#### 🔗 Associação")
            col1, col2 = st.columns(2)
            with col1:
                coord_options_dict = {c['id']: c['descricao'] for c in coordenadorias_options}
                coord_labels = list(coord_options_dict.values()); coord_ids = list(coord_options_dict.keys())
                coord_selecionado_label = st.selectbox("Coordenadoria *", options=coord_labels, index=0)
                coord_selecionado_id = coord_ids[coord_labels.index(coord_selecionado_label)]
            with col2:
                squad_options_dict = {s['id']: s['descricao'] for s in squads_options}
                squad_labels = list(squad_options_dict.values()); squad_ids = list(squad_options_dict.keys())
                squad_selecionado_label = st.selectbox("Squad *", options=squad_labels, index=0)
                squad_selecionado_id = squad_ids[squad_labels.index(squad_selecionado_label)]
            st.markdown("---")
            col_submit, col_cancel = st.columns(2)
            with col_submit: submitted = st.form_submit_button("💾 Salvar", width="stretch", type="primary", key="btn_salvar_cs")
            with col_cancel: cancelled = st.form_submit_button("❌ Cancelar", width="stretch", key="btn_cancelar_cs")
            if cancelled: st.session_state.mostrar_form_cs = False; st.rerun()
            if submitted:
                if not coord_selecionado_id: st.error("⚠️ Coordenadoria é obrigatória")
                elif not squad_selecionado_id: st.error("⚠️ Squad é obrigatório")
                else:
                    sucesso, msg = criar_coordenadoria_squad(coord_selecionado_id, squad_selecionado_id)
                    if sucesso: st.success(f"✅ {msg}"); st.session_state.mostrar_form_cs = False; st.rerun()
                    else: st.error(f"❌ {msg}")

# =============================================================================
# PROJETOS
# =============================================================================
@st.cache_data(ttl=120, show_spinner=False)
def listar_projetos() -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return pd.DataFrame()
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'devops_projects'")
            colunas_existentes = [row['column_name'] for row in cur.fetchall()]
            colunas_select = ['id', 'name', 'is_active']
            if 'created_at' in colunas_existentes: colunas_select.append('created_at')
            elif 'criado_em' in colunas_existentes: colunas_select.append('criado_em')
            if 'updated_at' in colunas_existentes: colunas_select.append('updated_at')
            elif 'atualizado_em' in colunas_existentes: colunas_select.append('atualizado_em')
            cur.execute(f"SELECT {', '.join(colunas_select)} FROM devops_projects ORDER BY name")
            rows = cur.fetchall()
        conn.close(); return pd.DataFrame(rows)
    except Exception as e:
        logger.error(f"Erro ao listar projetos: {e}"); return pd.DataFrame()

def criar_projeto(name: str, is_active: bool = True) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM devops_projects WHERE name = %s", (name.strip(),))
            if cur.fetchone(): conn.close(); return False, "Já existe um projeto cadastrado com este nome."
            cur.execute("INSERT INTO devops_projects (name, is_active) VALUES (%s, %s)", (name.strip(), is_active))
            conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Projeto cadastrado com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao cadastrar: {str(e)}"

def atualizar_projeto(projeto_id: int, name: str, is_active: bool) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'devops_projects' AND column_name IN ('updated_at', 'atualizado_em')")
            col_timestamp = cur.fetchone()
            if col_timestamp: cur.execute(f"UPDATE devops_projects SET name = %s, is_active = %s, {col_timestamp[0]} = CURRENT_TIMESTAMP WHERE id = %s", (name.strip(), is_active, projeto_id))
            else: cur.execute("UPDATE devops_projects SET name = %s, is_active = %s WHERE id = %s", (name.strip(), is_active, projeto_id))
            conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Projeto atualizado com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao atualizar: {str(e)}"

def deletar_projeto(projeto_id: int) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM devops_projects WHERE id = %s", (projeto_id,)); conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Projeto deletado com sucesso!"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao deletar: {str(e)}"

# =============================================================================
# HELPERS
# =============================================================================
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
    if assigned_name == "Não atribuído": return assigned_name
    azure_key = assigned_name.strip().lower()
    if "<" in azure_key: azure_key = azure_key.split("<")[0].strip()
    if azure_key in azure_to_ponto: return azure_to_ponto[azure_key]
    for azure_name, ponto_name in azure_to_ponto.items():
        if azure_key in azure_name or azure_name in azure_key: return ponto_name
    palavras_azure = azure_key.split()
    for azure_name, ponto_name in azure_to_ponto.items():
        if len(set(palavras_azure) & set(azure_name.split())) >= 2: return ponto_name
    return assigned_name

# =============================================================================
# EXPORT CSV/EXCEL
# =============================================================================
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
    wb = Workbook(); ws = wb.active; ws.title = "Faturamento"
    header_font = Font(name='Montserrat', size=14, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='FF6600', end_color='FF6600', fill_type='solid')
    info_font = Font(name='Montserrat', size=11, bold=True, color='333333')
    data_font = Font(name='Montserrat', size=10)
    money_font = Font(name='Montserrat', size=10, bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    ws.merge_cells('A1:M1')
    cell_title = ws['A1']; cell_title.value = f"FATURAMENTO - {mes}/{ano}"
    cell_title.font = header_font; cell_title.fill = header_fill
    cell_title.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 35
    ws.merge_cells('A2:M2')
    cell_info = ws['A2']; cell_info.value = f"Dias Úteis: {dias_uteis/8:.0f} | Total Faturado: {format_moeda_brl(total_faturado)}"
    cell_info.font = info_font; cell_info.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 25
    headers = ["Profissional", "Cargo", "Nível", "Data de Início", "Promoção de Cargo", "Novo Colaborador", "Data de Saída", "Horas Trabalhadas", "Horas Afast Acordado", "Horas Afast Legal", "Horas Arredondadas", "Valor Profissional", "Valor Recalculado"]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_num); cell.value = header
        cell.font = Font(name='Montserrat', size=10, bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='2C5282', end_color='2C5282', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[4].height = 35
    column_widths = [35, 20, 15, 15, 20, 25, 15, 18, 22, 18, 20, 20, 20]
    for i, width in enumerate(column_widths, 1): ws.column_dimensions[get_column_letter(i)].width = width
    for row_num, (_, row) in enumerate(df_faturamento.iterrows(), 5):
        for col_idx, col_name in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=col_idx, value=row[col_name])
            if col_name in ["Valor Profissional", "Valor Recalculado"]:
                cell.font = money_font; cell.number_format = 'R$ #,##0.00'
            else: cell.font = data_font
            cell.border = thin_border
    output = io.BytesIO(); wb.save(output); output.seek(0); return output.getvalue()

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
    .b-Done { background: #d4edda; color: #155724; } .b-Resolved { background: #d1ecf1; color: #0c5460; }
    .b-In-Progress { background: #fff3cd; color: #856404; } .b-To-Do { background: #e2e3e5; color: #383d41; }
    details > summary { list-style: none; cursor: pointer; padding: 0; }
    details > summary::-webkit-details-marker { display: none; }
    details > summary::before { content: '▶'; display: inline-block; color: #FF6600; font-size: 10px; margin-right: 4px; }
    details[open] > summary::before { content: '▼'; }
    .dt { font-size: 11px; color: #666; } .sprint-path { font-size: 11px; color: #444; word-break: break-word; }
    </style>
    """
    rows = []
    for p in parents_data:
        f = p.get("fields", {})
        pid, title = p["id"], esc(f.get("System.Title", ""))
        ptype, pstate = esc(f.get("System.WorkItemType", "")), esc(f.get("System.State", ""))
        hours = format_hours(f.get("Microsoft.VSTS.Scheduling.OriginalEstimate"))
        sprint_path = esc(str(f.get("System.IterationPath", "N/A")))
        created, closed = format_date(f.get("System.CreatedDate")), format_date(f.get("Microsoft.VSTS.Common.ClosedDate"))
        proj = p.get("_project", "")
        link = f"{ADO_BASE_URL}/{urllib.parse.quote(proj, safe='')}/_workitems/edit/{pid}"
        child_ids = children_map.get(pid, [])
        seen_children = set(); unique_child_ids = []
        for cid in child_ids:
            if cid not in seen_children and cid in all_workitems:
                seen_children.add(cid); unique_child_ids.append(cid)
        children = [all_workitems[cid] for cid in unique_child_ids]
        state_class = state_to_css_class(pstate)
        parent_row = ('<tr><td colspan="10" style="padding:0;border:none;">'
            '<details><summary><table class="parent-tbl"><tr>'
            '<td style="width:3%"></td>'
            f'<td style="width:22%" class="sprint-path"><strong>{sprint_path}</strong></td>'
            f'<td style="width:6%;color:#FF6600;font-weight:700;">#{pid}</td>'
            f'<td style="width:25%" title="{title}">{title[:50]}{"..." if len(title)>50 else ""}</td>'
            f'<td style="width:10%">{ptype}</td>'
            f'<td style="width:10%"><span class="badge b-{state_class}">{pstate}</span></td>'
            f'<td style="width:7%;font-weight:600;color:#2c5282;">{hours}</td>'
            f'<td style="width:6%"><a href="{link}" target="_blank" class="lnk" onclick="event.stopPropagation()">Abrir</a></td>'
            f'<td style="width:6%" class="dt">{created}</td>'
            f'<td style="width:6%" class="dt">{closed}</td>'
            '</tr></table></summary>')
        if children:
            parent_row += '<div style="padding:0 0 0 30px;"><table class="child-tbl"><thead><tr>'
            parent_row += '<th style="width:3%"></th><th style="width:22%">Sprint</th><th style="width:6%">ID</th>'
            parent_row += '<th style="width:25%">Título</th><th style="width:10%">Tipo</th><th style="width:10%">Situação</th>'
            parent_row += '<th style="width:7%">Horas</th><th style="width:6%">Link</th>'
            parent_row += '<th style="width:6%">Dt Criação</th><th style="width:6%">Dt Término</th>'
            parent_row += '</tr></thead><tbody>'
            for c in children:
                cf = c.get("fields", {})
                cstate = esc(cf.get("System.State", "")); cid = c["id"]
                child_link = f"{ADO_BASE_URL}/{urllib.parse.quote(proj, safe='')}/_workitems/edit/{cid}"
                parent_row += '<tr><td></td>'
                parent_row += f'<td class="sprint-path">{sprint_path}</td>'
                parent_row += f'<td style="color:#666;">#{cid}</td>'
                parent_row += f'<td>{esc(cf.get("System.Title", ""))}</td>'
                parent_row += f'<td>{esc(cf.get("System.WorkItemType", ""))}</td>'
                parent_row += f'<td><span class="badge b-{state_to_css_class(cstate)}">{cstate}</span></td>'
                parent_row += f'<td>{format_hours(cf.get("Microsoft.VSTS.Scheduling.OriginalEstimate"))}</td>'
                parent_row += f'<td><a href="{child_link}" target="_blank" class="lnk">Abrir</a></td>'
                parent_row += f'<td class="dt">{format_date(cf.get("System.CreatedDate"))}</td>'
                parent_row += f'<td class="dt">{format_date(cf.get("Microsoft.VSTS.Common.ClosedDate"))}</td>'
                parent_row += '</tr>'
            parent_row += '</tbody></table></div>'
        else:
            parent_row += '<div style="padding:10px 0 10px 30px;color:#888;font-style:italic;font-size:12px;">Nenhuma atividade filha vinculada.</div>'
        parent_row += '</details></td></tr>'
        rows.append(parent_row)
    header_html = ('<table class="ado-tbl"><thead><tr>'
        '<th style="width:3%"></th><th style="width:22%">Sprint</th><th style="width:6%">ID</th>'
        '<th style="width:25%">Título</th><th style="width:10%">Tipo</th><th style="width:10%">Situação</th>'
        '<th style="width:7%">Horas</th><th style="width:6%">Link</th>'
        '<th style="width:6%">Dt Criação</th><th style="width:6%">Dt Término</th>'
        '</tr></thead><tbody>')
    footer_html = ('<tr class="ftr"><td colspan="6" style="text-align:right;">TOTAL ACUMULADO DE HORAS PREVISTAS:</td>'
        f'<td style="color:#FF6600;font-size:16px;">{total_hours:.2f} h</td>'
        '<td colspan="3"></td></tr></tbody></table>')
    full_html = css + header_html + ''.join(rows) + footer_html
    st.markdown(full_html, unsafe_allow_html=True)

# =============================================================================
# TABS DE CONTEÚDO
# =============================================================================
def render_colaborador_tab(all_workitems_colab, project_of_item, df_colaboradores, year, month_name):
    periodo_atual = f"{year}-{month_name}"
    if st.session_state.get('filtro_colab_periodo') != periodo_atual:
        st.session_state.filtro_colab_periodo = periodo_atual
        if 'filtro_colab_tab' in st.session_state: del st.session_state.filtro_colab_tab
    azure_to_ponto = {}
    for _, row in df_colaboradores.iterrows():
        azure_name = str(row["colaborador_azure"]).strip().lower()
        if "<" in azure_name: azure_name = azure_name.split("<")[0].strip()
        ponto_name = str(row["colaborador_ponto"]).strip()
        if azure_name: azure_to_ponto[azure_name] = ponto_name
    tarefas_list = []
    for item_id, item in all_workitems_colab.items():
        f = item.get("fields", {})
        if f.get("System.State") != "Done": continue
        assigned_to = f.get("System.AssignedTo")
        assigned_name = get_assigned_to_name(assigned_to)
        colaborador_ponto = encontrar_colaborador_ponto(assigned_name, azure_to_ponto)
        proj = project_of_item.get(item_id, "")
        completed_work = float(f.get("Microsoft.VSTS.Scheduling.CompletedWork") or 0)
        link = f"{ADO_BASE_URL}/{urllib.parse.quote(proj, safe='')}/_workitems/edit/{item_id}"
        tarefas_list.append({"Colaborador": colaborador_ponto, "Projeto": proj, "Sprint": str(f.get("System.IterationPath", "N/A")).split("\\")[-1],
            "ID": item_id, "Tipo": f.get("System.WorkItemType", ""), "Título": f.get("System.Title", ""), "Horas Apontadas": completed_work, "Link": link})
    if not tarefas_list:
        st.warning("⚠️ Nenhuma tarefa com State=Done encontrada."); return
    df_tarefas = pd.DataFrame(tarefas_list)
    filtro = multiselect_with_toggle("👥 Filtrar por Colaborador (Tarefas)", sorted(df_tarefas["Colaborador"].unique()), sorted(df_tarefas["Colaborador"].unique()), "filtro_colab_tab")
    if filtro: df_tarefas = df_tarefas[df_tarefas["Colaborador"].isin(filtro)]
    df_sumario = df_tarefas.groupby("Colaborador").agg({"Horas Apontadas": "sum", "ID": "count"}).reset_index()
    df_sumario.columns = ["Colaborador", "Total Horas", "Qtd Tarefas"]
    df_sumario["Média Horas/Tarefa"] = (df_sumario["Total Horas"] / df_sumario["Qtd Tarefas"]).round(2)
    df_sumario = df_sumario.sort_values("Colaborador", ascending=True)
    st.subheader("📊 Sumário por Colaborador (State: Done)")
    st.info(f"Total: **{len(df_tarefas)} tarefa(s)** | **{df_tarefas['Horas Apontadas'].sum():.2f}h**")
    st.dataframe(df_sumario, width="stretch", hide_index=True,
        column_config={"Colaborador": st.column_config.TextColumn("Colaborador", width="large"),
            "Total Horas": st.column_config.NumberColumn("Total Horas", format="%.2f h"),
            "Qtd Tarefas": st.column_config.NumberColumn("Qtd Tarefas"),
            "Média Horas/Tarefa": st.column_config.NumberColumn("Média Horas/Tarefa", format="%.2f h", help="Média de horas apontadas por tarefa concluída")})
    st.markdown("---")
    st.subheader("📈 Horas Apontadas por Colaborador")
    fig = px.bar(df_sumario, x="Colaborador", y="Total Horas", text="Total Horas", color="Colaborador", color_discrete_sequence=px.colors.qualitative.Set3)
    fig.update_layout(margin=dict(t=20, b=80, l=20, r=20), height=400, showlegend=False, xaxis_tickangle=-45)
    fig.update_traces(texttemplate='%{text:.1f}h', textposition='outside')
    st.plotly_chart(fig, width="stretch")
    st.markdown("---")
    st.subheader("📋 Detalhamento por Colaborador")
    for colaborador in sorted(df_tarefas["Colaborador"].unique()):
        df_col = df_tarefas[df_tarefas["Colaborador"] == colaborador].sort_values(["Projeto", "Sprint"])
        with st.expander(f"👤 **{colaborador}** — {len(df_col)} tarefa(s) | {df_col['Horas Apontadas'].sum():.2f}h", expanded=False):
            st.dataframe(df_col[["Projeto", "Sprint", "ID", "Tipo", "Título", "Horas Apontadas", "Link"]], width="stretch", hide_index=True,
                column_config={"Projeto": st.column_config.TextColumn("Projeto"), "Sprint": st.column_config.TextColumn("Sprint"), "ID": st.column_config.NumberColumn("ID"),
                    "Tipo": st.column_config.TextColumn("Tipo"), "Título": st.column_config.TextColumn("Título", width="large"),
                    "Horas Apontadas": st.column_config.NumberColumn("Horas", format="%.2f h"), "Link": st.column_config.LinkColumn("Link", display_text="Abrir")})

def render_horas_colaborador_tab(df_colaboradores, df_horasponto, horas_uteis_mensais, all_workitems_colab, project_of_item, year, month_name, valid_paths=None):
    periodo_atual = f"{year}-{month_name}"
    if st.session_state.get('filtro_horas_periodo') != periodo_atual:
        st.session_state.filtro_horas_periodo = periodo_atual
        if 'filtro_horas' in st.session_state: del st.session_state.filtro_horas
    colaborador_nivel = {}
    for _, row in df_colaboradores.iterrows():
        ponto_name = str(row["colaborador_ponto"]).strip()
        nivel = str(row.get("nivel", "")).strip()
        if ponto_name: colaborador_nivel[ponto_name] = nivel
    nivel_capacity = get_all_capacities()
    azure_to_ponto = {}
    for _, row in df_colaboradores.iterrows():
        azure_name = str(row["colaborador_azure"]).strip().lower()
        if "<" in azure_name: azure_name = azure_name.split("<")[0].strip()
        ponto_name = str(row["colaborador_ponto"]).strip()
        if azure_name: azure_to_ponto[azure_name] = ponto_name
    horas_azure = {}
    for item_id, item in all_workitems_colab.items():
        f = item.get("fields", {})
        if f.get("System.State") != "Done": continue
        if valid_paths and f.get("System.IterationPath", "") not in valid_paths: continue
        assigned_name = get_assigned_to_name(f.get("System.AssignedTo"))
        colaborador_ponto = encontrar_colaborador_ponto(assigned_name, azure_to_ponto)
        if colaborador_ponto != "Não atribuído":
            horas_azure[colaborador_ponto] = horas_azure.get(colaborador_ponto, 0) + float(f.get("Microsoft.VSTS.Scheduling.CompletedWork") or 0)
    lista_consolidada = []
    for _, hp_row in df_horasponto.iterrows():
        ponto_name = str(hp_row["colaborador_ponto"]).strip()
        horas_ponto = float(hp_row["horas_ponto"]) if pd.notna(hp_row.get("horas_ponto")) else 0.0
        afastamento = float(hp_row["afastamento"]) if pd.notna(hp_row.get("afastamento")) else 0.0
        folga = float(hp_row["folga"]) if pd.notna(hp_row.get("folga")) else 0.0
        nivel = colaborador_nivel.get(ponto_name, "")
        capacity = nivel_capacity.get(nivel, 0.0)
        capacity_calculado = (horas_ponto / 8.0) * capacity if (horas_ponto > 0 and capacity > 0) else 0.0
        horas_azure_colab = horas_azure.get(ponto_name, 0.0)
        diferenca = horas_azure_colab - capacity_calculado
        lista_consolidada.append({"Colaborador": ponto_name, "Ponto": horas_ponto, "Azure": horas_azure_colab,
            "Afastamento": afastamento, "Folga": folga, "Capacity": capacity_calculado, "Diferença": round(diferenca, 2)})
    df_consolidado = pd.DataFrame(lista_consolidada).sort_values("Colaborador", ascending=True)
    filtro = multiselect_with_toggle("👥 Filtrar por Colaborador (Horas)", sorted(df_consolidado["Colaborador"].unique()), sorted(df_consolidado["Colaborador"].unique()), "filtro_horas")
    if filtro: df_consolidado = df_consolidado[df_consolidado["Colaborador"].isin(filtro)]
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("🕐 Total Ponto", f"{df_consolidado['Ponto'].sum():.2f} h")
    col2.metric("💻 Total Azure", f"{df_consolidado['Azure'].sum():.2f} h")
    col3.metric("🏖️ Total Afastamento", f"{df_consolidado['Afastamento'].sum():.2f} h")
    col4.metric("🌴 Total Folga", f"{df_consolidado['Folga'].sum():.2f} h")
    col5.metric("📊 Total Capacity", f"{df_consolidado['Capacity'].sum():.2f} h")
    diff_total = df_consolidado['Diferença'].sum()
    col6.metric("⚖️ Diferença Total", f"{diff_total:+.2f} h", delta=f"{diff_total:+.2f}h",
                delta_color="normal" if abs(diff_total) < 10 else ("inverse" if diff_total < 0 else "normal"))
    st.markdown("---")
    st.subheader("🕐 Horas por Colaborador")
    st.info(f"Período: **{month_name}/{year}** | Horas úteis: **{horas_uteis_mensais:.2f}h**")
    def colorir_diferenca(val):
        if pd.isna(val): return ''
        if val > 0: return 'background-color: #d4edda; color: #155724'
        elif val < 0: return 'background-color: #f8d7da; color: #721c24'
        else: return 'background-color: #fff3cd; color: #856404'
    styler = df_consolidado.style.map(colorir_diferenca, subset=['Diferença'])
    st.dataframe(styler, width="stretch", hide_index=True,
        column_config={"Colaborador": st.column_config.TextColumn("Colaborador", width="large"),
            "Ponto": st.column_config.NumberColumn("Ponto (h)", format="%.2f h"),
            "Azure": st.column_config.NumberColumn("Azure (h)", format="%.2f h"),
            "Afastamento": st.column_config.NumberColumn("Afastamento (h)", format="%.2f h"),
            "Folga": st.column_config.NumberColumn("Folga (h)", format="%.2f h"),
            "Capacity": st.column_config.NumberColumn("Capacity (h)", format="%.2f h"),
            "Diferença": st.column_config.NumberColumn("Diferença (h)", format="%+.2f h", help="Positivo = excesso no Azure | Negativo = falta no Azure")})
    st.markdown("---")
    st.subheader("📈 Comparativo: Ponto vs Azure vs Capacity")
    df_grafico = df_consolidado.melt(id_vars=["Colaborador"], value_vars=["Ponto", "Azure", "Capacity"], var_name="Fonte", value_name="Horas")
    fig = px.bar(df_grafico, x="Colaborador", y="Horas", color="Fonte", barmode="group", color_discrete_map={"Ponto": "#28a745", "Azure": "#0078d4", "Capacity": "#FF6600"})
    fig.update_layout(margin=dict(t=20, b=80, l=20, r=20), height=450, xaxis_tickangle=-45)
    st.plotly_chart(fig, width="stretch")
    st.markdown("---")
    st.subheader("⚖️ Diferença entre Azure e Capacity por Colaborador")
    st.info("🟢 **Verde**: Horas excedentes no Azure | 🔴 **Vermelho**: Horas faltantes no Azure")
    df_diff = df_consolidado.copy().sort_values('Diferença', ascending=False)
    fig_diff = px.bar(df_diff, x="Colaborador", y="Diferença", color="Diferença", color_continuous_scale=['#dc3545', '#ffc107', '#28a745'], text="Diferença")
    fig_diff.update_layout(margin=dict(t=20, b=80, l=20, r=20), height=450, xaxis_tickangle=-45, coloraxis_showscale=False)
    fig_diff.update_traces(texttemplate='%{text:+.1f}h', textposition='outside')
    fig_diff.add_hline(y=0, line_dash="solid", line_color="black", line_width=2)
    st.plotly_chart(fig_diff, width="stretch")

def render_ponto_tab(year, month_name, month_num, tipo_usuario):
    st.subheader("🕐 Gerenciamento de Ponto Eletrônico")
    st.info(f"Período selecionado: **{month_name}/{year}**")
    if 'mostrar_importacao' not in st.session_state: st.session_state.mostrar_importacao = False
    if 'mostrar_confirmacao_limpar' not in st.session_state: st.session_state.mostrar_confirmacao_limpar = False
    if 'importacao_em_andamento' not in st.session_state: st.session_state.importacao_em_andamento = False
    df_horasponto = get_horasponto_by_periodo(year, month_num)
    pode_editar = tipo_usuario in ['Admin', 'Gestor']
    if pode_editar:
        col_btn1, col_btn2 = st.columns([1, 1])
        with col_btn1:
            if st.button("📤 Importar CSV", width="stretch", type="primary", key="btn_importar_csv_ponto"):
                st.session_state.mostrar_importacao = True; st.session_state.mostrar_confirmacao_limpar = False; st.session_state.importacao_em_andamento = False; st.rerun()
        with col_btn2:
            if not df_horasponto.empty:
                if st.button("🗑️ Limpar Período", width="stretch", type="secondary", key="btn_limpar_periodo"):
                    st.session_state.mostrar_confirmacao_limpar = True; st.session_state.mostrar_importacao = False; st.rerun()
        if st.session_state.get('mostrar_confirmacao_limpar'):
            st.warning(f"⚠️ Você está prestes a excluir **todos os registros** de {month_name}/{year}")
            st.info(f"Total de registros: **{len(df_horasponto)}**")
            col_sim, col_nao = st.columns(2)
            with col_sim:
                if st.button("✅ Sim, Excluir", width="stretch", type="primary", key="btn_sim_excluir_ponto"):
                    sucesso, msg = deletar_horasponto_by_periodo(year, month_num)
                    if sucesso: st.success(f"✅ {msg}"); st.session_state.mostrar_confirmacao_limpar = False; st.rerun()
                    else: st.error(f"❌ {msg}")
            with col_nao:
                if st.button("❌ Não, Cancelar", width="stretch", key="btn_nao_excluir_ponto"): st.session_state.mostrar_confirmacao_limpar = False; st.rerun()
        if st.session_state.get('mostrar_importacao'):
            st.markdown("---")
            st.subheader("📤 Importar Arquivo CSV de Ponto")
            st.info("O arquivo CSV deve conter as colunas: **id_colaborador**, **colaborador_ponto**, **horas_ponto**, **Folga**, **afastamento**")
            if not df_horasponto.empty:
                st.warning(f"⚠️ **Atenção:** Já existem **{len(df_horasponto)} registros** para {month_name}/{year}")
                st.info("Os dados existentes serão **substituídos** pelos dados do arquivo CSV.")
            uploaded_file = st.file_uploader("Selecione o arquivo CSV do ponto", type=['csv'], help="O arquivo deve conter as colunas: id_colaborador, colaborador_ponto, horas_ponto, Folga, afastamento")
            if uploaded_file is not None:
                try:
                    raw_bytes = uploaded_file.read(); uploaded_file.seek(0)
                    separador = ','
                    try:
                        texto_amostra = raw_bytes[:2048].decode('utf-8', errors='ignore')
                        primeira_linha = texto_amostra.split('\n')[0]
                        if primeira_linha.count(';') > primeira_linha.count(','): separador = ';'
                    except: pass
                    try: df_csv = pd.read_csv(io.BytesIO(raw_bytes), encoding='utf-8', sep=separador)
                    except UnicodeDecodeError: df_csv = pd.read_csv(io.BytesIO(raw_bytes), encoding='latin1', sep=separador)
                    df_csv.columns = [col.strip().lower().replace(' ', '_') for col in df_csv.columns]
                    mapeamento_colunas = {
                        'id_colaborador': ['id_colaborador', 'id', 'colaborador_id'],
                        'colaborador_ponto': ['colaborador_ponto', 'colaborador', 'nome', 'profissional'],
                        'horas_ponto': ['horas_ponto', 'horas', 'horas_trabalhadas'],
                        'folga': ['folga', 'folgas', 'dias_folga'],
                        'afastamento': ['afastamento', 'afast', 'afastamento_legal']
                    }
                    for coluna_padrao, variantes in mapeamento_colunas.items():
                        for col_atual in df_csv.columns:
                            if col_atual in variantes and coluna_padrao not in df_csv.columns:
                                df_csv = df_csv.rename(columns={col_atual: coluna_padrao}); break
                    colunas_necessarias = ['id_colaborador', 'colaborador_ponto', 'horas_ponto']
                    colunas_faltantes = [col for col in colunas_necessarias if col not in df_csv.columns]
                    if colunas_faltantes:
                        st.error(f"❌ Colunas faltantes no CSV: {', '.join(colunas_faltantes)}"); st.info(f"Colunas encontradas: {list(df_csv.columns)}")
                    else:
                        if 'afastamento' not in df_csv.columns: df_csv['afastamento'] = 0.0
                        if 'folga' not in df_csv.columns: df_csv['folga'] = 0.0
                        df_csv['id_colaborador'] = pd.to_numeric(df_csv['id_colaborador'], errors='coerce')
                        ids_invalidos = df_csv[df_csv['id_colaborador'].isna()]
                        if not ids_invalidos.empty:
                            st.error(f"❌ {len(ids_invalidos)} registros com id_colaborador inválido"); st.dataframe(ids_invalidos.head(10), width="stretch")
                        else:
                            df_csv['id_colaborador'] = df_csv['id_colaborador'].astype(int)
                            st.success(f"✅ Arquivo carregado: **{len(df_csv)} registros** encontrados")
                            st.info(f"Colunas detectadas: **{', '.join(df_csv.columns)}**")
                            st.dataframe(df_csv.head(10), width="stretch")
                            col_confirmar, col_cancelar = st.columns(2)
                            with col_confirmar:
                                if st.button("✅ Confirmar Importação", width="stretch", type="primary", key="btn_confirmar_importacao"):
                                    st.session_state.importacao_em_andamento = True; st.rerun()
                            with col_cancelar:
                                if st.button("❌ Cancelar", width="stretch", key="btn_cancelar_importacao"):
                                    st.session_state.mostrar_importacao = False; st.session_state.importacao_em_andamento = False; st.rerun()
                            if st.session_state.get('importacao_em_andamento'):
                                with st.spinner("🔄 Processando importação em lote..."):
                                    if not df_horasponto.empty: deletar_horasponto_by_periodo(year, month_num)
                                    registros_para_inserir = []; erros_processamento = []
                                    for idx, row in df_csv.iterrows():
                                        try:
                                            row_dict = {str(k).strip().lower(): v for k, v in row.items()}
                                            id_colab = int(row_dict.get('id_colaborador'))
                                            nome_ponto = str(row_dict.get('colaborador_ponto', '')).strip()
                                            horas = float(str(row_dict.get('horas_ponto', 0)).replace(',', '.')) if pd.notna(row_dict.get('horas_ponto')) else 0.0
                                            afastamento = float(str(row_dict.get('afastamento', 0)).replace(',', '.')) if pd.notna(row_dict.get('afastamento')) else 0.0
                                            folga = float(str(row_dict.get('folga', 0)).replace(',', '.')) if pd.notna(row_dict.get('folga')) else 0.0
                                            registros_para_inserir.append({'ano': year, 'mes': month_num, 'id_colaborador': id_colab, 'colaborador_ponto': nome_ponto, 'horas_ponto': horas, 'afastamento': afastamento, 'folga': folga})
                                        except Exception as e: erros_processamento.append(f"Linha {idx+1}: {str(e)}")
                                    inseridos, erros, err_list = inserir_horasponto_lote(registros_para_inserir)
                                    st.session_state.importacao_em_andamento = False
                                    st.markdown("---"); st.markdown("### 📊 Resultado da Importação")
                                    col1, col2, col3 = st.columns(3)
                                    col1.metric("✅ Registros Inseridos", inseridos); col2.metric("⚠️ Erros", len(erros_processamento)); col3.metric("📋 Total Processados", len(df_csv))
                                    if erros_processamento:
                                        st.error(f"⚠️ **{len(erros_processamento)} erro(s)** no processamento:")
                                        for erro in erros_processamento[:10]: st.text(f"  • {erro}")
                                    if inseridos > 0:
                                        st.success(f"🎉 **Importação concluída!** {inseridos} registros gravados com sucesso.")
                                        st.info("💡 **Importante:** Recalcule o faturamento para atualizar os valores.")
                                    st.session_state.mostrar_importacao = False; st.rerun()
                except Exception as e:
                    st.error(f"❌ Erro ao processar arquivo: {str(e)}"); st.session_state.importacao_em_andamento = False
    st.markdown("---"); st.markdown("### 📋 Registros do Período")
    df_horasponto_atualizado = get_horasponto_by_periodo(year, month_num)
    if df_horasponto_atualizado.empty:
        st.info(f"📭 Nenhum registro de ponto encontrado para **{month_name}/{year}**")
    else:
        st.markdown(f"**Total de registros:** {len(df_horasponto_atualizado)}")
        colunas_exibicao = [c for c in ['colaborador_ponto', 'cargo', 'nivel', 'horas_ponto', 'afastamento', 'folga'] if c in df_horasponto_atualizado.columns]
        st.dataframe(df_horasponto_atualizado[colunas_exibicao], width="stretch", hide_index=True,
            column_config={"colaborador_ponto": st.column_config.TextColumn("Colaborador", width="large"),
                "cargo": st.column_config.TextColumn("Cargo"), "nivel": st.column_config.TextColumn("Nível"),
                "horas_ponto": st.column_config.NumberColumn("Horas Ponto", format="%.2f h"),
                "afastamento": st.column_config.NumberColumn("Afastamento", format="%.2f h"),
                "folga": st.column_config.NumberColumn("Folga", format="%.2f h")})
        col_total1, col_total2, col_total3 = st.columns(3)
        col_total1.metric("🕐 Total de Horas", f"{df_horasponto_atualizado['horas_ponto'].sum():.2f} h")
        col_total2.metric("🏖️ Total de Afastamento", f"{df_horasponto_atualizado['afastamento'].sum():.2f} h")
        if 'folga' in df_horasponto_atualizado.columns: col_total3.metric("🌴 Total de Folga", f"{df_horasponto_atualizado['folga'].sum():.2f} h")
        st.markdown("---")
        csv_data = df_horasponto_atualizado.to_csv(index=False, encoding='utf-8')
        st.download_button("⬇️ Baixar CSV do Período", data=csv_data, file_name=f"horasponto_{year}_{month_name}.csv", mime="text/csv", width="stretch")

# =============================================================================
# ESTRUTURAÇÃO
# =============================================================================
def render_estruturacao_tab():
    tipo_usuario = st.session_state.user_info.get('cargo', 'User')
    if tipo_usuario not in ['Admin', 'Gestor']:
        st.error("🚫 Acesso negado. Apenas Administradores e Gestores podem acessar esta funcionalidade.")
        st.stop()
    st.subheader("🏗️ Estruturação")
    sub_tab_projetos, sub_tab_squads, sub_tab_coords, sub_tab_cs = st.tabs([
        "🏢 Projetos", "👥 Squads", "🏛️ Coordenadorias", "🔗 Coordenadorias x Squads"
    ])
    with sub_tab_projetos: render_projetos_tab()
    with sub_tab_squads: render_squads_tab()
    with sub_tab_coords: render_coordenadorias_tab()
    with sub_tab_cs: render_coordenadoria_squad_tab()

def render_projetos_tab():
    if 'projeto_editando' not in st.session_state: st.session_state.projeto_editando = None
    if 'mostrar_form_projeto' not in st.session_state: st.session_state.mostrar_form_projeto = False
    if 'mostrar_confirmacao_exclusao_projeto' not in st.session_state: st.session_state.mostrar_confirmacao_exclusao_projeto = None
    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn2:
        if st.button("➕ Novo Projeto", width="stretch", type="primary", key="btn_novo_projeto"):
            st.session_state.projeto_editando = None; st.session_state.mostrar_form_projeto = True; st.rerun()
    df_projetos = listar_projetos()
    if df_projetos.empty:
        st.warning("⚠️ Nenhum projeto cadastrado. Clique em **➕ Novo Projeto** para começar.")
    else:
        st.markdown(f"**Total de projetos cadastrados:** {len(df_projetos)}")
        df_exibicao = df_projetos.copy()
        for col in ['created_at', 'criado_em', 'updated_at', 'atualizado_em']:
            if col in df_exibicao.columns: df_exibicao[col] = df_exibicao[col].apply(lambda x: x.strftime("%d/%m/%Y %H:%M") if pd.notna(x) else "")
        if 'is_active' in df_exibicao.columns:
            df_exibicao['status'] = df_exibicao['is_active'].apply(lambda x: "✅ Ativo" if x else "❌ Inativo")
        colunas_exibicao = [c for c in ['id', 'name', 'status', 'created_at', 'criado_em', 'updated_at', 'atualizado_em'] if c in df_exibicao.columns]
        st.dataframe(df_exibicao[colunas_exibicao], width="stretch", hide_index=True,
            column_config={"id": st.column_config.NumberColumn("ID", width="small"), "name": st.column_config.TextColumn("Nome do Projeto", width="large"), "status": st.column_config.TextColumn("Status")})
        st.markdown("---"); st.markdown("### Ações")
        options = [f"{row['id']} - {row['name']} ({'Ativo' if row.get('is_active') else 'Inativo'})" for _, row in df_projetos.iterrows()]
        projeto_selecionado = st.selectbox("Selecione um projeto para editar ou excluir:", options=options, index=0)
        if projeto_selecionado:
            projeto_id = int(projeto_selecionado.split(" - ")[0])
            projeto_dados = df_projetos[df_projetos['id'] == projeto_id].iloc[0].to_dict()
            col_edit, col_del = st.columns(2)
            with col_edit:
                if st.button("✏️ Editar Projeto", width="stretch", key="btn_edit_projeto"): st.session_state.projeto_editando = projeto_dados; st.session_state.mostrar_form_projeto = True; st.rerun()
            with col_del:
                if st.button("🗑️ Excluir Projeto", width="stretch", type="secondary", key="btn_del_projeto"): st.session_state.mostrar_confirmacao_exclusao_projeto = projeto_id; st.rerun()
    if st.session_state.get('mostrar_confirmacao_exclusao_projeto'):
        col_id = st.session_state.mostrar_confirmacao_exclusao_projeto
        col_nome = df_projetos[df_projetos['id'] == col_id]['name'].values[0]
        st.warning(f"⚠️ Você está prestes a excluir o projeto: **{col_nome}**")
        col_sim, col_nao = st.columns(2)
        with col_sim:
            if st.button("✅ Sim, Excluir", width="stretch", type="primary", key="btn_sim_excluir_projeto"):
                sucesso, msg = deletar_projeto(col_id)
                if sucesso: st.success(f"✅ {msg}"); st.session_state.mostrar_confirmacao_exclusao_projeto = None; st.rerun()
                else: st.error(f"❌ {msg}")
        with col_nao:
            if st.button("❌ Não, Cancelar", width="stretch", key="btn_nao_excluir_projeto"): st.session_state.mostrar_confirmacao_exclusao_projeto = None; st.rerun()
    if st.session_state.get('mostrar_form_projeto'):
        st.markdown("---"); st.subheader("📝 " + ("Editar Projeto" if st.session_state.projeto_editando else "Novo Projeto"))
        projeto = st.session_state.projeto_editando or {}
        with st.form("form_projeto"):
            name = st.text_input("Nome do Projeto *", value=projeto.get('name', ''), placeholder="Ex: ASCOM - Comunicação")
            is_active = st.checkbox("Projeto Ativo", value=projeto.get('is_active', True))
            st.markdown("---")
            col_submit, col_cancel = st.columns(2)
            with col_submit: submitted = st.form_submit_button("💾 Salvar", width="stretch", type="primary", key="btn_salvar_projeto")
            with col_cancel: cancelled = st.form_submit_button("❌ Cancelar", width="stretch", key="btn_cancelar_projeto")
            if cancelled: st.session_state.projeto_editando = None; st.session_state.mostrar_form_projeto = False; st.rerun()
            if submitted:
                if not name.strip(): st.error("⚠️ Nome do projeto é obrigatório")
                else:
                    if st.session_state.projeto_editando: sucesso, msg = atualizar_projeto(st.session_state.projeto_editando['id'], name, is_active)
                    else: sucesso, msg = criar_projeto(name, is_active)
                    if sucesso: st.success(f"✅ {msg}"); st.session_state.projeto_editando = None; st.session_state.mostrar_form_projeto = False; st.rerun()
                    else: st.error(f"❌ {msg}")

# =============================================================================
# COLABORADORES (com 3 sub-abas) - COM CAMPOS MOTIVO_DEMISSAO E OBSERVACAO
# =============================================================================
def render_colaboradores_tab(df_colaboradores_filtrados=None):
    tipo_usuario = st.session_state.user_info.get('cargo', 'User')
    if tipo_usuario not in ['Admin', 'Gestor']:
        st.error("🚫 Acesso negado. Apenas Administradores e Gestores podem acessar esta funcionalidade.")
        st.stop()
    st.subheader("⚙️ Gerenciamento de Colaboradores")
    
    # 4 Sub-abas: Informações, Projetos, Squads e a nova Reuniões 1:1
    sub_tab_info, sub_tab_projetos, sub_tab_squads, sub_tab_reuniao = st.tabs([
        "👤 Informações", "🔗 Projetos", "👥 Squads", "🤝 Reuniões 1:1"
    ])
    with sub_tab_info: _render_colaboradores_info(df_colaboradores_filtrados)
    with sub_tab_projetos: _render_colab_projetos()
    with sub_tab_squads: _render_colab_squads()
    with sub_tab_reuniao: _render_reunioes_11(df_colaboradores_filtrados)

def _render_colaboradores_info(df_colaboradores_filtrados=None):
    if 'colaborador_editando' not in st.session_state: st.session_state.colaborador_editando = None
    if 'mostrar_form_colaborador' not in st.session_state: st.session_state.mostrar_form_colaborador = False
    if 'mostrar_confirmacao_exclusao' not in st.session_state: st.session_state.mostrar_confirmacao_exclusao = None
    if 'filtro_cargo_colab' not in st.session_state: st.session_state.filtro_cargo_colab = []
    if 'filtro_nivel_colab' not in st.session_state: st.session_state.filtro_nivel_colab = []
    
    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn2:
        if st.button("➕ Novo Colaborador", width="stretch", type="primary", key="btn_novo_colab"):
            st.session_state.colaborador_editando = None; st.session_state.mostrar_form_colaborador = True; st.rerun()
    
    if df_colaboradores_filtrados is not None:
        df_colaboradores = df_colaboradores_filtrados
    else:
        df_colaboradores = listar_colaboradores()
    
    render_destaques_semana(df_colaboradores)
    
    if df_colaboradores.empty:
        st.warning("⚠️ Nenhum colaborador cadastrado. Clique em **➕ Novo Colaborador** para começar.")
    else:
        st.markdown("### 🔍 Filtros")
        cargos_unicos = sorted([c for c in df_colaboradores['cargo'].dropna().unique() if str(c).strip()])
        niveis_unicos = sorted([n for n in df_colaboradores['nivel'].dropna().unique() if str(n).strip()])
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            filtro_cargo = st.multiselect("Filtrar por Cargo", options=cargos_unicos, default=st.session_state.filtro_cargo_colab, key="multiselect_filt_cargo_colab", placeholder="Todos os cargos")
        with col_f2:
            filtro_nivel = st.multiselect("Filtrar por Nível", options=niveis_unicos, default=st.session_state.filtro_nivel_colab, key="multiselect_filt_nivel_colab", placeholder="Todos os níveis")
        if filtro_cargo != st.session_state.filtro_cargo_colab: st.session_state.filtro_cargo_colab = filtro_cargo
        if filtro_nivel != st.session_state.filtro_nivel_colab: st.session_state.filtro_nivel_colab = filtro_nivel
        df_colaboradores_filtrado = df_colaboradores.copy()
        if filtro_cargo: df_colaboradores_filtrado = df_colaboradores_filtrado[df_colaboradores_filtrado['cargo'].isin(filtro_cargo)]
        if filtro_nivel: df_colaboradores_filtrado = df_colaboradores_filtrado[df_colaboradores_filtrado['nivel'].isin(filtro_nivel)]
        if filtro_cargo or filtro_nivel:
            st.info(f"📊 Exibindo **{len(df_colaboradores_filtrado)}** de **{len(df_colaboradores)}** colaboradores (filtros aplicados)")
        
        # INDICADORES
        st.markdown("### 📊 Indicadores Gerais")
        total_colaboradores = len(df_colaboradores_filtrado)
        total_ativos = len(df_colaboradores_filtrado[df_colaboradores_filtrado['status'] == 'Ativo']) if 'status' in df_colaboradores_filtrado.columns else total_colaboradores
        total_inativos = total_colaboradores - total_ativos
        col_ind1, col_ind2, col_ind3 = st.columns(3)
        with col_ind1: st.metric(label="👥 Total de Colaboradores", value=total_colaboradores)
        with col_ind2: st.metric(label="✅ Ativos", value=total_ativos, delta=f"{(total_ativos/total_colaboradores*100) if total_colaboradores > 0 else 0:.1f}%")
        with col_ind3: st.metric(label="⛔ Inativos", value=total_inativos, delta=f"-{(total_inativos/total_colaboradores*100) if total_colaboradores > 0 else 0:.1f}%", delta_color="inverse")
        st.markdown("---")
        
        # DISTRIBUIÇÃO POR EQUIPE
        st.markdown("### 👔 Distribuição por Equipe (Gestor)")
        if 'nome_gestor' in df_colaboradores_filtrado.columns:
            df_por_gestor = df_colaboradores_filtrado[df_colaboradores_filtrado['status'] == 'Ativo'].groupby('nome_gestor').size().reset_index(name='total')
            df_por_gestor = df_por_gestor.sort_values('total', ascending=False)
            if not df_por_gestor.empty:
                num_cols = min(4, len(df_por_gestor))
                cols_gestor = st.columns(num_cols)
                for idx, (_, row) in enumerate(df_por_gestor.iterrows()):
                    with cols_gestor[idx % num_cols]:
                        st.metric(label=f"👤 {row['nome_gestor'] if pd.notna(row['nome_gestor']) else 'Sem Gestor'}", value=f"{row['total']} colab.")
                fig_gestor = px.bar(df_por_gestor, x='nome_gestor', y='total', text='total', color='total', color_continuous_scale='Blues', labels={'nome_gestor': 'Gestor', 'total': 'Qtd Colaboradores'}, title='Colaboradores por Equipe')
                fig_gestor.update_traces(textposition='outside')
                fig_gestor.update_layout(margin=dict(t=40, b=80, l=20, r=20), height=350, xaxis_tickangle=-45, showlegend=False)
                st.plotly_chart(fig_gestor, width="stretch")
        st.markdown("---")
        
        # DISTRIBUIÇÃO POR SQUAD
        st.markdown("### 👥 Distribuição por Squad")
        df_squad = buscar_distribuicao_squad()
        if not df_squad.empty:
            num_cols_squad = min(4, len(df_squad))
            cols_squad = st.columns(num_cols_squad)
            for idx, (_, row) in enumerate(df_squad.iterrows()):
                with cols_squad[idx % num_cols_squad]:
                    st.metric(label=f"🏃 {row['squad'] if pd.notna(row['squad']) else 'Sem Squad'}", value=f"{row['total']} colab.")
            fig_squad = px.bar(df_squad, x='squad', y='total', text='total', color='total', color_continuous_scale='Oranges', labels={'squad': 'Squad', 'total': 'Qtd Colaboradores'}, title='Colaboradores por Squad')
            fig_squad.update_traces(textposition='outside')
            fig_squad.update_layout(margin=dict(t=40, b=80, l=20, r=20), height=350, xaxis_tickangle=-45, showlegend=False)
            st.plotly_chart(fig_squad, width="stretch")
        st.markdown("---")
        
        # DISTRIBUIÇÃO POR CARGO
        st.markdown("### 💼 Distribuição por Cargo")
        df_por_cargo = df_colaboradores_filtrado[df_colaboradores_filtrado['status'] == 'Ativo'].groupby('cargo').size().reset_index(name='total') if 'status' in df_colaboradores_filtrado.columns else df_colaboradores_filtrado.groupby('cargo').size().reset_index(name='total')
        df_por_cargo = df_por_cargo.sort_values('total', ascending=False)
        if not df_por_cargo.empty:
            num_cols_cargo = min(4, len(df_por_cargo))
            cols_cargo = st.columns(num_cols_cargo)
            for idx, (_, row) in enumerate(df_por_cargo.iterrows()):
                with cols_cargo[idx % num_cols_cargo]:
                    st.metric(label=f"🎯 {row['cargo'] if pd.notna(row['cargo']) else 'Sem Cargo'}", value=f"{row['total']} colab.")
            fig_cargo = px.pie(df_por_cargo, values='total', names='cargo', title='Distribuição por Cargo', color_discrete_sequence=px.colors.qualitative.Set2, hole=0.4)
            fig_cargo.update_traces(textposition='inside', textinfo='percent+label+value')
            fig_cargo.update_layout(margin=dict(t=40, b=20, l=20, r=20), height=400)
            st.plotly_chart(fig_cargo, width="stretch")
        st.markdown("---")
        
        # GRID
        if df_colaboradores_filtrado.empty:
            st.warning("⚠️ Nenhum colaborador encontrado com os filtros selecionados.")
        else:
            st.markdown(f"**Total de colaboradores cadastrados:** {len(df_colaboradores_filtrado)}")
            colunas_grid = [
                'id', 'colaborador_ponto', 'cargo', 'nivel', 'nome_gestor',
                'email', 'email_pessoal', 'cpf', 'telefone', 'local',
                'linkedin', 'github', 'os', 'tecnologia',
                'data_inicio', 'dt_nascimento', 'dia_aniv', 'mes_aniv',
                'promocao', 'reposicao', 'data_saida', 'status',
                'motivo_demissao', 'observacao'
            ]
            colunas_exibicao = [c for c in colunas_grid if c in df_colaboradores_filtrado.columns]
            st.dataframe(df_colaboradores_filtrado[colunas_exibicao], width="stretch", hide_index=True,
                column_config={
                    "id": st.column_config.NumberColumn("ID", width="small"),
                    "colaborador_ponto": st.column_config.TextColumn("Nome", width="large"),
                    "cargo": st.column_config.TextColumn("Cargo"),
                    "nivel": st.column_config.TextColumn("Nível"),
                    "nome_gestor": st.column_config.TextColumn("Gestor"),
                    "email": st.column_config.TextColumn("E-mail Corp."),
                    "email_pessoal": st.column_config.TextColumn("E-mail Pessoal"),
                    "cpf": st.column_config.TextColumn("CPF"),
                    "telefone": st.column_config.TextColumn("Telefone"),
                    "local": st.column_config.TextColumn("Local"),
                    "linkedin": st.column_config.LinkColumn("LinkedIn"),
                    "github": st.column_config.LinkColumn("GitHub"),
                    "os": st.column_config.TextColumn("OS"),
                    "tecnologia": st.column_config.TextColumn("Tecnologia"),
                    "data_inicio": st.column_config.DateColumn("Data Início"),
                    "dt_nascimento": st.column_config.DateColumn("Nascimento"),
                    "dia_aniv": st.column_config.NumberColumn("Dia Aniv."),
                    "mes_aniv": st.column_config.NumberColumn("Mês Aniv."),
                    "promocao": st.column_config.DateColumn("Promoção"),
                    "reposicao": st.column_config.TextColumn("Reposição"),
                    "data_saida": st.column_config.DateColumn("Data Saída"),
                    "status": st.column_config.TextColumn("Status"),
                    "motivo_demissao": st.column_config.TextColumn("Motivo Demissão"),
                    "observacao": st.column_config.TextColumn("Observação", width="large")
                })
            st.markdown("---"); st.markdown("### Ações")
            colaborador_selecionado = st.selectbox("Selecione um colaborador para editar ou excluir:", options=[f"{row['id']} - {row['colaborador_ponto']} ({row['cargo']})" for _, row in df_colaboradores_filtrado.iterrows()], index=0)
            if colaborador_selecionado:
                colaborador_id = int(colaborador_selecionado.split(" - ")[0])
                colaborador_dados = df_colaboradores_filtrado[df_colaboradores_filtrado['id'] == colaborador_id].iloc[0].to_dict()
                col_edit, col_del, col_hist = st.columns(3)
                with col_edit:
                    if st.button("✏️ Editar Colaborador", width="stretch", key="btn_edit_colab"):
                        st.session_state.colaborador_editando = colaborador_dados; st.session_state.mostrar_form_colaborador = True; st.rerun()
                with col_del:
                    if st.button("🗑️ Excluir Colaborador", width="stretch", type="secondary", key="btn_del_colab"):
                        st.session_state.mostrar_confirmacao_exclusao = colaborador_id; st.rerun()
                with col_hist:
                    if st.button("📜 Visualizar Histórico", width="stretch", type="secondary", key="btn_historico_colab"):
                        st.session_state.historico_colab_id = colaborador_id
                        st.session_state.historico_colab_nome = colaborador_dados.get('colaborador_ponto', 'Desconhecido')
                        st.rerun()
    
    if st.session_state.get('historico_colab_id'):
        mostrar_historico_modal(st.session_state.historico_colab_id, st.session_state.historico_colab_nome)
        st.session_state.historico_colab_id = None
        st.session_state.historico_colab_nome = None
    
    if st.session_state.get('mostrar_confirmacao_exclusao'):
        col_id = st.session_state.mostrar_confirmacao_exclusao
        df_todos = df_colaboradores_filtrados if df_colaboradores_filtrados is not None else listar_colaboradores()
        col_nome = df_todos[df_todos['id'] == col_id]['colaborador_ponto'].values[0]
        st.warning(f"⚠️ Você está prestes a excluir o colaborador: **{col_nome}**")
        col_sim, col_nao = st.columns(2)
        with col_sim:
            if st.button("✅ Sim, Excluir", width="stretch", type="primary", key="btn_sim_excluir_colab"):
                sucesso, msg = deletar_colaborador(col_id)
                if sucesso: st.success(f"✅ {msg}"); st.session_state.mostrar_confirmacao_exclusao = None; st.rerun()
                else: st.error(f"❌ {msg}")
        with col_nao:
            if st.button("❌ Não, Cancelar", width="stretch", key="btn_nao_excluir_colab"): st.session_state.mostrar_confirmacao_exclusao = None; st.rerun()
    
    if st.session_state.get('mostrar_form_colaborador'):
        st.markdown("---"); st.subheader("📝 " + ("Editar Colaborador" if st.session_state.colaborador_editando else "Novo Colaborador"))
        colaborador = st.session_state.colaborador_editando or {}
        gestores = listar_gestores()
        cargos_options = listar_cargos_options()
        gestores_options = {g['id']: g['nome'] for g in gestores}; gestores_options[0] = "-- Sem gestor --"
        with st.form("form_colaborador"):
            st.markdown("#### 👤 Dados Pessoais")
            col1, col2, col3 = st.columns(3)
            with col1:
                colaborador_ponto = st.text_input("Nome no Ponto *", value=colaborador.get('colaborador_ponto', ''))
                cpf = st.text_input("CPF", value=str(colaborador.get('cpf', '') or ''))
                dt_nascimento = st.date_input("Data de Nascimento", value=colaborador.get('dt_nascimento') if colaborador.get('dt_nascimento') else None, min_value=MIN_DATA, max_value=MAX_DATA, format="DD/MM/YYYY")
            with col2:
                colaborador_azure = st.text_input("Nome no Azure DevOps *", value=colaborador.get('colaborador_azure', ''))
                email_pessoal = st.text_input("E-mail Pessoal", value=colaborador.get('email_pessoal', ''))
                telefone = st.text_input("Telefone", value=colaborador.get('telefone', ''))
            with col3:
                email = st.text_input("E-mail Corporativo", value=colaborador.get('email', ''))
                local = st.text_input("Local", value=colaborador.get('local', ''))
                os_val = st.text_input("O.S", value=colaborador.get('os', ''))
            st.markdown("#### 💼 Dados Profissionais")
            col1, col2, col3 = st.columns(3)
            with col1:
                if cargos_options:
                    cargo_options_dict = {c['id']: c['descricao'] for c in cargos_options}
                    cargo_labels = list(cargo_options_dict.values())
                    cargo_ids = list(cargo_options_dict.keys())
                    cargo_atual = colaborador.get('cargo_id')
                    cargo_index = 0
                    if cargo_atual is not None:
                        try:
                            cargo_atual_int = int(cargo_atual)
                            if cargo_atual_int in cargo_ids:
                                cargo_index = cargo_ids.index(cargo_atual_int)
                        except (ValueError, TypeError):
                            pass
                    cargo_selecionado_label = st.selectbox("Cargo *", options=cargo_labels, index=cargo_index if cargo_labels else 0)
                    cargo_selecionado_id = cargo_ids[cargo_labels.index(cargo_selecionado_label)] if cargo_selecionado_label in cargo_labels else None
                else:
                    st.error("⚠️ Nenhum cargo cadastrado.")
                    cargo_selecionado_id = None
                nivel = st.selectbox("Nível *", options=["", "Senior", "Pleno", "Junior", "Portfólio"], index=["", "Senior", "Pleno", "Junior", "Portfólio"].index(colaborador.get('nivel', '')) if colaborador.get('nivel', '') in ["", "Senior", "Pleno", "Junior", "Portfólio"] else 0)
                tecnologia = st.text_input("Tecnologia Principal", value=colaborador.get('tecnologia', ''))
            with col2:
                data_inicio = st.date_input("Data de Início *", value=colaborador.get('data_inicio') if colaborador.get('data_inicio') else None, min_value=MIN_DATA, max_value=MAX_DATA, format="DD/MM/YYYY")
                data_saida = st.date_input("Data de Saída", value=colaborador.get('data_saida') if colaborador.get('data_saida') else None, min_value=MIN_DATA, max_value=MAX_DATA, format="DD/MM/YYYY")
                promocao = st.date_input("Data da Promoção", value=colaborador.get('promocao') if colaborador.get('promocao') else None, min_value=MIN_DATA, max_value=MAX_DATA, format="DD/MM/YYYY")
            with col3:
                reposicao = st.text_input("Novo Colaborador (Reposição)", value=colaborador.get('reposicao', ''))
                status = st.selectbox("Status", options=["Ativo", "Inativo", "Bloqueado"], index=["Ativo", "Inativo", "Bloqueado"].index(colaborador.get('status', 'Ativo')) if colaborador.get('status', 'Ativo') in ["Ativo", "Inativo", "Bloqueado"] else 0)
                gestor_id = colaborador.get('gestor')
                gestor_index = 0
                if gestor_id is not None:
                    try:
                        gestor_id_int = int(gestor_id)
                        if gestor_id_int in gestores_options and gestor_id_int != 0:
                            gestor_index = list(gestores_options.keys()).index(gestor_id_int)
                    except (ValueError, TypeError):
                        pass
                gestor_selecionado = st.selectbox("Gestor", options=list(gestores_options.values()), index=gestor_index)
                gestor_id_selecionado = [k for k, v in gestores_options.items() if v == gestor_selecionado][0]
            st.markdown("#### 🔗 Links e Aniversário")
            col1, col2, col3 = st.columns(3)
            with col1: linkedin = st.text_input("LinkedIn", value=colaborador.get('linkedin', ''))
            with col2: github = st.text_input("GitHub", value=colaborador.get('github', ''))
            with col3:
                dia_aniv = st.number_input("Dia do Aniversário", min_value=1, max_value=31, value=int(colaborador.get('dia_aniv') or 1))
                mes_aniv = st.number_input("Mês do Aniversário", min_value=1, max_value=12, value=int(colaborador.get('mes_aniv') or 1))
            
            # ============================================================
            # ✅ SEÇÃO NOVA: DEMISSÃO E OBSERVAÇÕES
            # ============================================================
            st.markdown("#### 📋 Demissão e Observações")
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                motivo_atual = str(colaborador.get('motivo_demissao') or '').strip()
                motivo_index = MOTIVOS_DEMISSAO.index(motivo_atual) if motivo_atual in MOTIVOS_DEMISSAO else 0
                motivo_demissao = st.selectbox(
                    "Motivo de Demissão",
                    options=MOTIVOS_DEMISSAO,
                    index=motivo_index,
                    help="Valores permitidos pela constraint do banco"
                )
            with col_d2:
                observacao = st.text_input(
                    "Observação",
                    value=str(colaborador.get('observacao') or ''),
                    max_chars=250,
                    help="Máximo 250 caracteres"
                )
            # ============================================================
            # FIM DA SEÇÃO NOVA
            # ============================================================
            
            st.markdown("---")
            col_submit, col_cancel = st.columns(2)
            with col_submit: submitted = st.form_submit_button("💾 Salvar", width="stretch", type="primary", key="btn_salvar_colab")
            with col_cancel: cancelled = st.form_submit_button("❌ Cancelar", width="stretch", key="btn_cancelar_colab")
            if cancelled: st.session_state.colaborador_editando = None; st.session_state.mostrar_form_colaborador = False; st.rerun()
            if submitted:
                erros = []
                if not colaborador_azure.strip(): erros.append("Nome no Azure DevOps é obrigatório")
                if not colaborador_ponto.strip(): erros.append("Nome no Ponto é obrigatório")
                if not cargo_selecionado_id: erros.append("Cargo é obrigatório")
                if not nivel: erros.append("Nível é obrigatório")
                if not data_inicio: erros.append("Data de Início é obrigatória")
                if motivo_demissao and motivo_demissao not in MOTIVOS_DEMISSAO:
                    erros.append("Motivo de demissão inválido")
                if erros:
                    for erro in erros: st.error(f"⚠️ {erro}")
                else:
                    dados = {
                        'colaborador_azure': colaborador_azure,
                        'colaborador_ponto': colaborador_ponto,
                        'cargo': cargo_selecionado_id,
                        'nivel': nivel,
                        'data_inicio': data_inicio if data_inicio else None,
                        'email': email,
                        'promocao': promocao if promocao else None,
                        'reposicao': reposicao,
                        'data_saida': data_saida if data_saida else None,
                        'status': status,
                        'gestor': gestor_id_selecionado,
                        'email_pessoal': email_pessoal,
                        'cpf': cpf,
                        'telefone': telefone,
                        'local': local,
                        'linkedin': linkedin,
                        'github': github,
                        'os': os_val,
                        'dt_nascimento': dt_nascimento if dt_nascimento else None,
                        'dia_aniv': dia_aniv,
                        'mes_aniv': mes_aniv,
                        'tecnologia': tecnologia,
                        'motivo_demissao': motivo_demissao if motivo_demissao else None,
                        'observacao': observacao if observacao else None
                    }
                    if st.session_state.colaborador_editando: sucesso, msg = atualizar_colaborador(st.session_state.colaborador_editando['id'], dados)
                    else: sucesso, msg = criar_colaborador(dados)
                    if sucesso: st.success(f"✅ {msg}"); st.session_state.colaborador_editando = None; st.session_state.mostrar_form_colaborador = False; st.rerun()
                    else: st.error(f"❌ {msg}")

def _render_colab_projetos():
    if 'colab_projeto_editando' not in st.session_state: st.session_state.colab_projeto_editando = None
    if 'mostrar_form_colab_projeto' not in st.session_state: st.session_state.mostrar_form_colab_projeto = False
    if 'mostrar_confirmacao_exclusao_cp' not in st.session_state: st.session_state.mostrar_confirmacao_exclusao_cp = None
    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn2:
        if st.button("➕ Nova Associação", width="stretch", type="primary", key="btn_nova_assoc_colab_proj"):
            st.session_state.colab_projeto_editando = None; st.session_state.mostrar_form_colab_projeto = True; st.rerun()
    df_colab_projetos = listar_colab_projetos()
    if df_colab_projetos.empty:
        st.warning("⚠️ Nenhuma associação colaborador-projeto cadastrada. Clique em **➕ Nova Associação** para começar.")
    else:
        st.markdown(f"**Total de associações:** {len(df_colab_projetos)}")
        colunas_grid = ['id', 'nome_colaborador', 'cargo', 'nome_projeto', 'created_at']
        colunas_exibicao = [c for c in colunas_grid if c in df_colab_projetos.columns]
        st.dataframe(df_colab_projetos[colunas_exibicao], width="stretch", hide_index=True,
            column_config={"id": st.column_config.NumberColumn("ID", width="small"),
                "nome_colaborador": st.column_config.TextColumn("Colaborador", width="large"),
                "cargo": st.column_config.TextColumn("Cargo"),
                "nome_projeto": st.column_config.TextColumn("Projeto", width="large"),
                "created_at": st.column_config.DatetimeColumn("Criado em", format="DD/MM/YYYY HH:mm")})
        st.markdown("---"); st.markdown("### Ações")
        options = [f"{row['id']} - {row['nome_colaborador']} → {row['nome_projeto']}" for _, row in df_colab_projetos.iterrows()]
        if options:
            associacao_selecionada = st.selectbox("Selecione uma associação para editar ou excluir:", options=options, index=0)
            if associacao_selecionada:
                cp_id = int(associacao_selecionada.split(" - ")[0])
                cp_dados = df_colab_projetos[df_colab_projetos['id'] == cp_id].iloc[0].to_dict()
                col_edit, col_del = st.columns(2)
                with col_edit:
                    if st.button("✏️ Editar Associação", width="stretch", key="btn_edit_assoc_cp"):
                        st.session_state.colab_projeto_editando = cp_dados; st.session_state.mostrar_form_colab_projeto = True; st.rerun()
                with col_del:
                    if st.button("🗑️ Excluir Associação", width="stretch", type="secondary", key="btn_del_assoc_cp"):
                        st.session_state.mostrar_confirmacao_exclusao_cp = cp_id; st.rerun()
    if st.session_state.get('mostrar_confirmacao_exclusao_cp'):
        cp_id = st.session_state.mostrar_confirmacao_exclusao_cp
        cp_row = df_colab_projetos[df_colab_projetos['id'] == cp_id]
        if not cp_row.empty:
            cp_nome = f"{cp_row.iloc[0]['nome_colaborador']} → {cp_row.iloc[0]['nome_projeto']}"
            st.warning(f"⚠️ Você está prestes a excluir a associação: **{cp_nome}**")
            col_sim, col_nao = st.columns(2)
            with col_sim:
                if st.button("✅ Sim, Excluir", width="stretch", type="primary", key="btn_sim_excluir_cp"):
                    sucesso, msg = deletar_colab_projeto(cp_id)
                    if sucesso: st.success(f"✅ {msg}"); st.session_state.mostrar_confirmacao_exclusao_cp = None; st.rerun()
                    else: st.error(f"❌ {msg}")
            with col_nao:
                if st.button("❌ Não, Cancelar", width="stretch", key="btn_nao_excluir_cp"): st.session_state.mostrar_confirmacao_exclusao_cp = None; st.rerun()
    if st.session_state.get('mostrar_form_colab_projeto'):
        st.markdown("---"); st.subheader("📝 " + ("Editar Associação" if st.session_state.colab_projeto_editando else "Nova Associação Colaborador-Projeto"))
        cp_dados = st.session_state.colab_projeto_editando or {}
        colaboradores_options = listar_colaboradores_options()
        projetos_options = listar_projetos_options()
        if not colaboradores_options: st.error("⚠️ Nenhum colaborador ativo cadastrado."); return
        if not projetos_options: st.error("⚠️ Nenhum projeto ativo cadastrado."); return
        with st.form("form_colab_projeto"):
            st.markdown("#### 🔗 Associação")
            col1, col2 = st.columns(2)
            with col1:
                colab_options_dict = {c['id']: f"{c['colaborador_ponto']} ({c['cargo']})" for c in colaboradores_options}
                colab_labels = list(colab_options_dict.values()); colab_ids = list(colab_options_dict.keys())
                colab_atual = cp_dados.get('id_colab'); colab_index = 0
                if colab_atual and int(colab_atual) in colab_ids: colab_index = colab_ids.index(int(colab_atual))
                colab_selecionado_label = st.selectbox("Colaborador *", options=colab_labels, index=colab_index)
                colab_selecionado_id = colab_ids[colab_labels.index(colab_selecionado_label)]
            with col2:
                proj_options_dict = {p['id']: p['name'] for p in projetos_options}
                proj_labels = list(proj_options_dict.values()); proj_ids = list(proj_options_dict.keys())
                proj_atual = cp_dados.get('id_projeto'); proj_index = 0
                if proj_atual and int(proj_atual) in proj_ids: proj_index = proj_ids.index(int(proj_atual))
                proj_selecionado_label = st.selectbox("Projeto *", options=proj_labels, index=proj_index)
                proj_selecionado_id = proj_ids[proj_labels.index(proj_selecionado_label)]
            st.markdown("---")
            col_submit, col_cancel = st.columns(2)
            with col_submit: submitted = st.form_submit_button("💾 Salvar", width="stretch", type="primary", key="btn_salvar_cp")
            with col_cancel: cancelled = st.form_submit_button("❌ Cancelar", width="stretch", key="btn_cancelar_cp")
            if cancelled: st.session_state.colab_projeto_editando = None; st.session_state.mostrar_form_colab_projeto = False; st.rerun()
            if submitted:
                if not colab_selecionado_id: st.error("⚠️ Colaborador é obrigatório")
                elif not proj_selecionado_id: st.error("⚠️ Projeto é obrigatório")
                else:
                    dados = {'id_colab': colab_selecionado_id, 'id_projeto': proj_selecionado_id}
                    if st.session_state.colab_projeto_editando: sucesso, msg = atualizar_colab_projeto(st.session_state.colab_projeto_editando['id'], dados)
                    else: sucesso, msg = criar_colab_projeto(dados)
                    if sucesso: st.success(f"✅ {msg}"); st.session_state.colab_projeto_editando = None; st.session_state.mostrar_form_colab_projeto = False; st.rerun()
                    else: st.error(f"❌ {msg}")

def _render_colab_squads():
    if 'filtro_coord_squad' not in st.session_state: st.session_state.filtro_coord_squad = []
    if 'filtro_squad_colab' not in st.session_state: st.session_state.filtro_squad_colab = []
    if 'mostrar_form_colab_squad' not in st.session_state: st.session_state.mostrar_form_colab_squad = False
    if 'mostrar_confirmacao_exclusao_cs' not in st.session_state: st.session_state.mostrar_confirmacao_exclusao_cs = None
    coordenadorias_options = listar_coordenadorias_options()
    squads_options = listar_squads_options()
    coord_labels = [c['descricao'] for c in coordenadorias_options]
    squad_labels = [s['descricao'] for s in squads_options]
    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn2:
        if st.button("➕ Nova Associação", width="stretch", type="primary", key="btn_nova_assoc_colab_squad"):
            st.session_state.mostrar_form_colab_squad = True; st.rerun()
    st.markdown("### 🔍 Filtros")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        filtro_coord = st.multiselect("Filtrar por Coordenadoria", options=coord_labels, default=st.session_state.filtro_coord_squad, key="multiselect_filt_coord_squad", placeholder="Todas as coordenadorias")
    with col_f2:
        filtro_squad = st.multiselect("Filtrar por Squad", options=squad_labels, default=st.session_state.filtro_squad_colab, key="multiselect_filt_squad_colab", placeholder="Todos os squads")
    if filtro_coord != st.session_state.filtro_coord_squad: st.session_state.filtro_coord_squad = filtro_coord
    if filtro_squad != st.session_state.filtro_squad_colab: st.session_state.filtro_squad_colab = filtro_squad
    df_colab_squad = listar_colab_squad()
    if filtro_coord: df_colab_squad = df_colab_squad[df_colab_squad['nome_coordenadoria'].isin(filtro_coord)]
    if filtro_squad: df_colab_squad = df_colab_squad[df_colab_squad['nome_squad'].isin(filtro_squad)]
    if df_colab_squad.empty:
        st.warning("⚠️ Nenhuma associação colaborador-squad encontrada com os filtros selecionados.")
    else:
        st.markdown(f"**Total de associações:** {len(df_colab_squad)}")
        colunas_grid = ['id', 'nome_colaborador', 'cargo', 'nome_coordenadoria', 'nome_squad', 'created_at']
        colunas_exibicao = [c for c in colunas_grid if c in df_colab_squad.columns]
        st.dataframe(df_colab_squad[colunas_exibicao], width="stretch", hide_index=True,
            column_config={"id": st.column_config.NumberColumn("ID", width="small"),
                "nome_colaborador": st.column_config.TextColumn("Colaborador", width="large"),
                "cargo": st.column_config.TextColumn("Cargo"),
                "nome_coordenadoria": st.column_config.TextColumn("Coordenadoria", width="large"),
                "nome_squad": st.column_config.TextColumn("Squad", width="large"),
                "created_at": st.column_config.DatetimeColumn("Criado em", format="DD/MM/YYYY HH:mm")})
        st.markdown("---"); st.markdown("### Ações")
        options = [f"{row['id']} - {row['nome_colaborador']} → {row['nome_squad']} ({row.get('nome_coordenadoria', 'N/A')})" for _, row in df_colab_squad.iterrows()]
        if options:
            associacao_selecionada = st.selectbox("Selecione uma associação para excluir:", options=options, index=0)
            if associacao_selecionada:
                cs_id = int(associacao_selecionada.split(" - ")[0])
                if st.button("🗑️ Excluir Associação", width="stretch", type="secondary", key="btn_del_assoc_colab_squad"):
                    st.session_state.mostrar_confirmacao_exclusao_cs = cs_id; st.rerun()
    if st.session_state.get('mostrar_confirmacao_exclusao_cs'):
        cs_id = st.session_state.mostrar_confirmacao_exclusao_cs
        cs_row = df_colab_squad[df_colab_squad['id'] == cs_id]
        if not cs_row.empty:
            cs_nome = f"{cs_row.iloc[0]['nome_colaborador']} → {cs_row.iloc[0]['nome_squad']}"
            st.warning(f"⚠️ Você está prestes a excluir a associação: **{cs_nome}**")
            col_sim, col_nao = st.columns(2)
            with col_sim:
                if st.button("✅ Sim, Excluir", width="stretch", type="primary", key="btn_sim_excluir_colab_squad"):
                    sucesso, msg = deletar_colab_squad(cs_id)
                    if sucesso: st.success(f"✅ {msg}"); st.session_state.mostrar_confirmacao_exclusao_cs = None; st.rerun()
                    else: st.error(f"❌ {msg}")
            with col_nao:
                if st.button("❌ Não, Cancelar", width="stretch", key="btn_nao_excluir_colab_squad"): st.session_state.mostrar_confirmacao_exclusao_cs = None; st.rerun()
    if st.session_state.get('mostrar_form_colab_squad'):
        st.markdown("---"); st.subheader("📝 Nova Associação Colaborador-Squad")
        colaboradores_options = listar_colaboradores_options()
        if not colaboradores_options: st.error("⚠️ Nenhum colaborador ativo cadastrado."); return
        if not squads_options: st.error("⚠️ Nenhum squad ativo cadastrado."); return
        with st.form("form_colab_squad"):
            st.markdown("#### 🔗 Associação")
            col1, col2 = st.columns(2)
            with col1:
                colab_options_dict = {c['id']: f"{c['colaborador_ponto']} ({c['cargo']})" for c in colaboradores_options}
                colab_labels_form = list(colab_options_dict.values()); colab_ids = list(colab_options_dict.keys())
                colab_selecionado_label = st.selectbox("Colaborador *", options=colab_labels_form, index=0)
                colab_selecionado_id = colab_ids[colab_labels_form.index(colab_selecionado_label)]
            with col2:
                squad_options_dict = {s['id']: s['descricao'] for s in squads_options}
                squad_labels_form = list(squad_options_dict.values()); squad_ids = list(squad_options_dict.keys())
                squad_selecionado_label = st.selectbox("Squad *", options=squad_labels_form, index=0)
                squad_selecionado_id = squad_ids[squad_labels_form.index(squad_selecionado_label)]
            st.markdown("---")
            col_submit, col_cancel = st.columns(2)
            with col_submit: submitted = st.form_submit_button("💾 Salvar", width="stretch", type="primary", key="btn_salvar_colab_squad")
            with col_cancel: cancelled = st.form_submit_button("❌ Cancelar", width="stretch", key="btn_cancelar_colab_squad")
            if cancelled: st.session_state.mostrar_form_colab_squad = False; st.rerun()
            if submitted:
                if not colab_selecionado_id: st.error("⚠️ Colaborador é obrigatório")
                elif not squad_selecionado_id: st.error("⚠️ Squad é obrigatório")
                else:
                    sucesso, msg = criar_colab_squad(colab_selecionado_id, squad_selecionado_id)
                    if sucesso: st.success(f"✅ {msg}"); st.session_state.mostrar_form_colab_squad = False; st.rerun()
                    else: st.error(f"❌ {msg}")

def _render_reunioes_11(df_colaboradores_filtrados=None):
    if 'reuniao_editando' not in st.session_state: st.session_state.reuniao_editando = None
    if 'mostrar_form_reuniao' not in st.session_state: st.session_state.mostrar_form_reuniao = False
    if 'mostrar_confirmacao_exclusao_reuniao' not in st.session_state: st.session_state.mostrar_confirmacao_exclusao_reuniao = None

    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn2:
        if st.button("➕ Nova Reunião 1:1", width="stretch", type="primary", key="btn_nova_reuniao_11"):
            st.session_state.reuniao_editando = None
            st.session_state.mostrar_form_reuniao = True
            st.rerun()

    df_reunioes = listar_reunioes_11()

    # Filtro cascata conforme gestores selecionados na sidebar
    if df_colaboradores_filtrados is not None and not df_colaboradores_filtrados.empty and not df_reunioes.empty:
        colabs_permitidos_ids = set(df_colaboradores_filtrados['id'].tolist())
        df_reunioes = df_reunioes[df_reunioes['id_colab'].isin(colabs_permitidos_ids)].copy()

    if df_reunioes.empty:
        st.warning("⚠️ Nenhuma reunião 1:1 encontrada. Clique em **➕ Nova Reunião 1:1** para cadastrar.")
    else:
        # =========================================================================
        # 1) PRÓXIMO CICLO (6 MESES) E 2) STATUS (COM A NOVA REGRA)
        # =========================================================================
        dt_reuniao_series = pd.to_datetime(df_reunioes['data_reuniao'], errors='coerce')
        df_reunioes['proximo_ciclo'] = dt_reuniao_series + pd.DateOffset(months=6)

        hoje = pd.Timestamp.now().normalize()

        def calcular_status(row):
            dt_reuniao = pd.to_datetime(row.get('data_reuniao'), errors='coerce')
            dt_prox = pd.to_datetime(row.get('proximo_ciclo'), errors='coerce')
            realizada = bool(row.get('Realizada', False))

            if pd.isna(dt_reuniao):
                return "-"

            dt_reuniao_norm = dt_reuniao.normalize()
            dt_prox_norm = dt_prox.normalize() if pd.notna(dt_prox) else None

            # Condição 1: data atual > próximo ciclo
            cond_ciclo_vencido = (dt_prox_norm is not None) and (hoje > dt_prox_norm)
            
            # Condição 2: data atual > data_reuniao E não foi realizada
            cond_nao_realizada_passada = (hoje > dt_reuniao_norm) and (not realizada)

            if cond_ciclo_vencido or cond_nao_realizada_passada:
                return "🔴 Atrasada"
            return "🟢 Em dia"

        df_reunioes['status_ciclo'] = df_reunioes.apply(calcular_status, axis=1)

        # Filtros de cabeçalho
        st.markdown("### 🔍 Filtros de Reuniões")
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        with col_f1:
            gestores_unicos = sorted([g for g in df_reunioes['gestor'].dropna().unique() if str(g).strip()])
            filtro_gestor = st.multiselect("Filtrar por Gestor", options=gestores_unicos, default=[], placeholder="Todos os gestores", key="filt_r11_gestor")
        with col_f2:
            status_realizada = st.selectbox("Status de Realização", options=["Todas", "Apenas Realizadas", "Apenas Pendentes"], index=0, key="filt_r11_realizada")
        with col_f3:
            filtro_status_ciclo = st.selectbox("Status", options=["Todos", "Apenas Atrasadas", "Apenas Em dia"], index=0, key="filt_r11_ciclo")
        with col_f4:
            status_senior = st.selectbox("Info Senior", options=["Todas", "Apenas Marcadas (True)", "Apenas Não Marcadas (False)"], index=0, key="filt_r11_senior")

        df_grid = df_reunioes.copy()
        if filtro_gestor:
            df_grid = df_grid[df_grid['gestor'].isin(filtro_gestor)]
        if status_realizada == "Apenas Realizadas":
            df_grid = df_grid[df_grid['Realizada'] == True]
        elif status_realizada == "Apenas Pendentes":
            df_grid = df_grid[df_grid['Realizada'] == False]

        if filtro_status_ciclo == "Apenas Atrasadas":
            df_grid = df_grid[df_grid['status_ciclo'] == "🔴 Atrasada"]
        elif filtro_status_ciclo == "Apenas Em dia":
            df_grid = df_grid[df_grid['status_ciclo'] == "🟢 Em dia"]

        if status_senior == "Apenas Marcadas (True)":
            df_grid = df_grid[df_grid['Info_Senior'] == True]
        elif status_senior == "Apenas Não Marcadas (False)":
            df_grid = df_grid[df_grid['Info_Senior'] == False]

        # Métricas de resumo
        c_m1, c_m2, c_m3, c_m4 = st.columns(4)
        c_m1.metric("Total de Reuniões", len(df_grid))
        c_m2.metric("✅ Realizadas", len(df_grid[df_grid['Realizada'] == True]))
        c_m3.metric("🔴 Atrasadas", len(df_grid[df_grid['status_ciclo'] == "🔴 Atrasada"]))
        c_m4.metric("🟢 Em dia", len(df_grid[df_grid['status_ciclo'] == "🟢 Em dia"]))
        st.markdown("---")

        # Exibição do Grid (SEM a coluna ID)
        st.dataframe(
            df_grid[['colaborador', 'gestor', 'squads', 'data_reuniao', 'proximo_ciclo', 'status_ciclo', 'Info_Senior', 'Realizada', 'observacao']],
            width="stretch",
            hide_index=True,
            column_config={
                "colaborador": st.column_config.TextColumn("Colaborador", width="large"),
                "gestor": st.column_config.TextColumn("Gestor", width="medium"),
                "squads": st.column_config.TextColumn("Squads", width="large", help="Squads vinculados separados por vírgula"),
                "data_reuniao": st.column_config.DateColumn("Data da Reunião", format="DD/MM/YYYY"),
                "proximo_ciclo": st.column_config.DateColumn("Próximo Ciclo", format="DD/MM/YYYY", help="Data sugerida 6 meses após a reunião"),
                "status_ciclo": st.column_config.TextColumn("Status", width="small", help="Atrasada se hoje > próximo ciclo ou se hoje > data da reunião e não realizada"),
                "Info_Senior": st.column_config.CheckboxColumn("Senior"),
                "Realizada": st.column_config.CheckboxColumn("Realizada"),
                "observacao": st.column_config.TextColumn(
                    "Observação",
                    width="large",
                    help="Clique duas vezes ou expanda a linha para ver o texto completo"
)
            }
        )

        st.markdown("---")
        st.markdown("### Ações")
        opcoes_reuniao = [
            f"{row['id']} - {row['colaborador']} (Data: {pd.to_datetime(row['data_reuniao']).strftime('%d/%m/%Y') if pd.notna(row['data_reuniao']) else '-'} | Status: {row['status_ciclo']})"
            for _, row in df_grid.iterrows()
        ]
        if opcoes_reuniao:
            reuniao_selecionada = st.selectbox("Selecione uma reunião para editar ou excluir:", options=opcoes_reuniao, index=0)
            if reuniao_selecionada:
                r_id = int(reuniao_selecionada.split(" - ")[0])
                r_dados = df_reunioes[df_reunioes['id'] == r_id].iloc[0].to_dict()
                col_e, col_d = st.columns(2)
                with col_e:
                    if st.button("✏️ Editar Reunião", width="stretch", key="btn_edit_r11"):
                        st.session_state.reuniao_editando = r_dados
                        st.session_state.mostrar_form_reuniao = True
                        st.rerun()
                with col_d:
                    if st.button("🗑️ Excluir Reunião", width="stretch", type="secondary", key="btn_del_r11"):
                        st.session_state.mostrar_confirmacao_exclusao_reuniao = r_id
                        st.rerun()

    # Confirmação de exclusão
    if st.session_state.get('mostrar_confirmacao_exclusao_reuniao'):
        r_del_id = st.session_state.mostrar_confirmacao_exclusao_reuniao
        st.warning(f"⚠️ Deseja realmente excluir o registro de reunião ID #{r_del_id}?")
        col_s, col_n = st.columns(2)
        with col_s:
            if st.button("✅ Sim, Excluir", width="stretch", type="primary", key="btn_sim_del_r11"):
                sucesso, msg = deletar_reuniao_11(r_del_id)
                if sucesso:
                    st.success(f"✅ {msg}")
                    st.session_state.mostrar_confirmacao_exclusao_reuniao = None
                    st.rerun()
                else:
                    st.error(f"❌ {msg}")
        with col_n:
            if st.button("❌ Não, Cancelar", width="stretch", key="btn_nao_del_r11"):
                st.session_state.mostrar_confirmacao_exclusao_reuniao = None
                st.rerun()

    # Formulário de Cadastro / Edição
    if st.session_state.get('mostrar_form_reuniao'):
        st.markdown("---")
        st.subheader("📝 " + ("Editar Reunião 1:1" if st.session_state.reuniao_editando else "Nova Reunião 1:1"))
        dados_form = st.session_state.reuniao_editando or {}
        colaboradores_opts = listar_colaboradores_options()
        if not colaboradores_opts:
            st.error("⚠️ Nenhum colaborador ativo cadastrado para associar à reunião.")
            return

        with st.form("form_reuniao_11"):
            colab_dict = {c['id']: f"{c['colaborador_ponto']} ({c['cargo']})" for c in colaboradores_opts}
            colab_ids = list(colab_dict.keys())
            colab_labels = list(colab_dict.values())
            
            id_atual = dados_form.get('id_colab')
            idx_selecionado = colab_ids.index(int(id_atual)) if id_atual and int(id_atual) in colab_ids else 0

            col_form1, col_form2 = st.columns(2)
            with col_form1:
                label_escolhido = st.selectbox("Colaborador *", options=colab_labels, index=idx_selecionado)
                id_colab_escolhido = colab_ids[colab_labels.index(label_escolhido)]
                
                data_val = dados_form.get('data_reuniao')
                if data_val and isinstance(data_val, str):
                    try: data_val = datetime.strptime(data_val, "%Y-%m-%d").date()
                    except: data_val = datetime.now().date()
                elif hasattr(data_val, 'date'):
                    data_val = data_val.date()
                elif not data_val:
                    data_val = datetime.now().date()

                data_reuniao_input = st.date_input("Data da Reunião *", value=data_val, min_value=MIN_DATA, max_value=MAX_DATA, format="DD/MM/YYYY")

            with col_form2:
                info_senior_input = st.checkbox("Info Senior (Senior)", value=bool(dados_form.get('Info_Senior', False)))
                realizada_input = st.checkbox("Realizada", value=bool(dados_form.get('Realizada', False)))

            obs_input = st.text_area("Observação", value=str(dados_form.get('observacao') or ''), max_chars=520, placeholder="Notas e alinhamentos da reunião 1:1...")

            st.markdown("---")
            col_save, col_canc = st.columns(2)
            with col_save:
                submitted = st.form_submit_button("💾 Salvar Reunião", width="stretch", type="primary")
            with col_canc:
                cancelled = st.form_submit_button("❌ Cancelar", width="stretch")

            if cancelled:
                st.session_state.reuniao_editando = None
                st.session_state.mostrar_form_reuniao = False
                st.rerun()

            if submitted:
                if not id_colab_escolhido:
                    st.error("⚠️ O colaborador é obrigatório.")
                elif not data_reuniao_input:
                    st.error("⚠️ A data da reunião é obrigatória.")
                else:
                    if st.session_state.reuniao_editando:
                        sucesso, msg = atualizar_reuniao_11(
                            st.session_state.reuniao_editando['id'],
                            id_colab_escolhido,
                            data_reuniao_input,
                            info_senior_input,
                            realizada_input,
                            obs_input
                        )
                    else:
                        sucesso, msg = criar_reuniao_11(
                            id_colab_escolhido,
                            data_reuniao_input,
                            info_senior_input,
                            realizada_input,
                            obs_input
                        )

                    if sucesso:
                        st.success(f"✅ {msg}")
                        st.session_state.reuniao_editando = None
                        st.session_state.mostrar_form_reuniao = False
                        st.rerun()
                    else:
                        st.error(f"❌ {msg}")

# =============================================================================
# FATURAMENTO
# =============================================================================
@st.cache_data(ttl=120, show_spinner=False)
def verificar_faturamento_existente(ano: int, mes: int) -> Optional[Dict]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT id, ano, mes, qt_dias_uteis, total_faturado, criado_em FROM faturamento_mensal WHERE ano = %s AND mes = %s", (ano, mes))
            result = cur.fetchone()
        conn.close(); return result
    except Exception as e:
        logger.error(f"Erro ao verificar faturamento existente: {e}"); return None

@st.cache_data(ttl=120, show_spinner=False)
def buscar_detalhes_faturamento(id_faturamento: int) -> pd.DataFrame:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return pd.DataFrame()
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT profissional, cargo, nivel, data_inicio, promocao_cargo, novo_colaborador, data_saida, horas_trabalhadas, horas_afast_acordado, horas_afast_legal, horas_arredondadas, valor_profissional, valor_recalculado FROM faturamento_detalhe WHERE id_faturamento = %s ORDER BY profissional", (id_faturamento,))
            rows = cur.fetchall()
        conn.close(); return pd.DataFrame(rows)
    except Exception as e:
        logger.error(f"Erro ao buscar detalhes do faturamento: {e}"); return pd.DataFrame()

def salvar_faturamento_no_banco(ano: int, mes: int, dias_uteis: float, total_faturado: float, df_detalhes: pd.DataFrame) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("INSERT INTO faturamento_mensal (ano, mes, qt_dias_uteis, total_faturado) VALUES (%s, %s, %s, %s) ON CONFLICT (ano, mes) DO UPDATE SET qt_dias_uteis = EXCLUDED.qt_dias_uteis, total_faturado = EXCLUDED.total_faturado, atualizado_em = CURRENT_TIMESTAMP RETURNING id", (int(ano), int(mes), float(dias_uteis), float(total_faturado)))
            id_faturamento = cur.fetchone()['id']
            cur.execute("DELETE FROM faturamento_detalhe WHERE id_faturamento = %s", (id_faturamento,))
            dados_detalhe = []
            for _, row in df_detalhes.iterrows():
                data_ini = datetime.strptime(str(row["Data de Início"]), "%d/%m/%Y").date() if row["Data de Início"] and row["Data de Início"] != "" else None
                data_saida = datetime.strptime(str(row["Data de Saída"]), "%d/%m/%Y").date() if row["Data de Saída"] and row["Data de Saída"] != "" else None
                dados_detalhe.append((int(id_faturamento), str(row["Profissional"]), str(row["Cargo"]), str(row["Nível"]), data_ini, str(row["Promoção de Cargo"]) if row["Promoção de Cargo"] else "", str(row["Novo Colaborador"]) if row["Novo Colaborador"] else "", data_saida, float(row["Horas Trabalhadas"]), float(row["Horas Afast Acordado"]), float(row["Horas Afast Legal"]), float(row["Horas Arredondadas"]), float(row["Valor Profissional"]), float(row["Valor Recalculado"])))
            query = """
                INSERT INTO faturamento_detalhe (id_faturamento, profissional, cargo, nivel, data_inicio, promocao_cargo, novo_colaborador, data_saida, horas_trabalhadas, horas_afast_acordado, horas_afast_legal, horas_arredondadas, valor_profissional, valor_recalculado)
                VALUES %s
            """
            execute_values(cur, query, dados_detalhe); conn.commit()
        conn.close(); st.cache_data.clear()
        return True, f"Faturamento salvo com sucesso! ID: {id_faturamento}"
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao salvar: {str(e)}"

def deletar_faturamento_existente(ano: int, mes: int) -> Tuple[bool, str]:
    db_url = os.getenv("NEON_DATABASE_URL")
    if not db_url: return False, "Banco de dados não configurado."
    conn = None
    try:
        conn = psycopg2.connect(db_url, connect_timeout=10)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM faturamento_mensal WHERE ano = %s AND mes = %s", (ano, mes)); conn.commit()
        conn.close(); st.cache_data.clear()
        return True, "Faturamento deletado com sucesso."
    except Exception as e:
        if conn: conn.rollback(); conn.close()
        return False, f"Erro ao deletar: {str(e)}"

def calcular_faturamento(df_colaboradores, df_horasponto, df_valores, dias_uteis):
    if 'status' in df_colaboradores.columns: df_colaboradores = df_colaboradores[df_colaboradores['status'] == 'Ativo'].copy()
    if df_colaboradores.empty: return pd.DataFrame()
    nivel_valor = {str(r["nivel"]).strip(): limpar_valor_monetario(r.get("valor")) for _, r in df_valores.iterrows() if str(r.get("nivel", "")).strip()}
    horas_por_id = {}
    for _, row in df_horasponto.iterrows():
        id_colab = row.get('id_colaborador')
        if pd.isna(id_colab): continue
        id_colab = int(id_colab)
        horas_ponto = float(row["horas_ponto"]) if pd.notna(row.get("horas_ponto")) else 0.0
        afastamento = float(row["afastamento"]) if pd.notna(row.get("afastamento")) else 0.0
        folga = float(row["folga"]) if pd.notna(row.get("folga")) else 0.0
        if id_colab in horas_por_id:
            horas_por_id[id_colab]["horas_ponto"] += horas_ponto
            horas_por_id[id_colab]["afastamento"] += afastamento
            horas_por_id[id_colab]["folga"] += folga
        else:
            horas_por_id[id_colab] = {"horas_ponto": horas_ponto, "afastamento": afastamento, "folga": folga}
    lista_faturamento = []
    for _, row in df_colaboradores.iterrows():
        colab_id = int(row["id"]); colaborador_ponto = str(row["colaborador_ponto"]).strip()
        cargo, nivel = str(row.get("cargo", "")).strip(), str(row.get("nivel", "")).strip()
        data_inicio, data_saida = row.get("data_inicio"), row.get("data_saida")
        promocao = str(row.get("promocao", "")).strip() if pd.notna(row.get("promocao")) else ""
        reposicao = str(row.get("reposicao", "")).strip() if pd.notna(row.get("reposicao")) else ""
        horas_info = horas_por_id.get(colab_id)
        if horas_info is None: horas_info = {"horas_ponto": 0.0, "afastamento": 0.0, "folga": 0.0}
        horas_trabalhadas = horas_info["horas_ponto"]; horas_afast_legal = horas_info["afastamento"]
        horas_afast_acordado = horas_info["folga"]; horas_arredondadas = horas_trabalhadas + horas_afast_acordado
        valor_profissional = nivel_valor.get(nivel, 0.0)
        valor_diario = valor_profissional / dias_uteis if dias_uteis > 0 else 0.0
        valor_recalculado = round(horas_arredondadas * valor_diario, 2)
        lista_faturamento.append({"Profissional": colaborador_ponto, "Cargo": cargo, "Nível": nivel,
            "Data de Início": data_inicio.strftime("%d/%m/%Y") if pd.notna(data_inicio) else "",
            "Promoção de Cargo": promocao, "Novo Colaborador": reposicao,
            "Data de Saída": data_saida.strftime("%d/%m/%Y") if pd.notna(data_saida) else "",
            "Horas Trabalhadas": round(horas_trabalhadas, 2), "Horas Afast Acordado": round(horas_afast_acordado, 2),
            "Horas Afast Legal": round(horas_afast_legal, 2), "Horas Arredondadas": round(horas_arredondadas, 2),
            "Valor Profissional": round(valor_profissional, 2), "Valor Recalculado": valor_recalculado})
    return pd.DataFrame(lista_faturamento)

def render_faturamento_tab(df_colaboradores, df_horasponto, df_valores, dias_uteis, year, month_name, month_num, pode_calcular=False):
    st.subheader("💰 Faturamento por Colaborador")
    st.info(f"Período: **{month_name}/{year}** | Dias úteis: **{dias_uteis/8:.0f}** ({dias_uteis:.0f}h)")
    faturamento_existente = verificar_faturamento_existente(year, month_num)
    if 'faturamento_calculado' not in st.session_state: st.session_state.faturamento_calculado = False
    if 'df_faturamento_exibido' not in st.session_state: st.session_state.df_faturamento_exibido = None
    if 'total_faturado_exibido' not in st.session_state: st.session_state.total_faturado_exibido = 0.0
    if 'mostrar_alerta_recuperar' not in st.session_state: st.session_state.mostrar_alerta_recuperar = False
    if faturamento_existente and not st.session_state.faturamento_calculado:
        st.info(f"✅ Faturamento existente encontrado para **{month_name}/{year}** (ID: {faturamento_existente['id']})")
        df_detalhes = buscar_detalhes_faturamento(faturamento_existente['id'])
        if not df_detalhes.empty:
            df_detalhes = df_detalhes.rename(columns={'profissional': 'Profissional', 'cargo': 'Cargo', 'nivel': 'Nível', 'data_inicio': 'Data de Início', 'promocao_cargo': 'Promoção de Cargo', 'novo_colaborador': 'Novo Colaborador', 'data_saida': 'Data de Saída', 'horas_trabalhadas': 'Horas Trabalhadas', 'horas_afast_acordado': 'Horas Afast Acordado', 'horas_afast_legal': 'Horas Afast Legal', 'horas_arredondadas': 'Horas Arredondadas', 'valor_profissional': 'Valor Profissional', 'valor_recalculado': 'Valor Recalculado'})
            for col in ['Data de Início', 'Data de Saída']:
                if col in df_detalhes.columns: df_detalhes[col] = df_detalhes[col].apply(lambda x: x.strftime("%d/%m/%Y") if pd.notna(x) else "")
            st.session_state.df_faturamento_exibido = df_detalhes
            st.session_state.total_faturado_exibido = faturamento_existente['total_faturado']
            st.session_state.faturamento_calculado = True; st.rerun()
    st.markdown("---")
    if pode_calcular:
        col_btn1, col_btn2 = st.columns([3, 1])
        with col_btn1:
            if st.button("🧮 Calcular Faturamento", width="stretch", type="primary", key="btn_calcular_faturamento"):
                if faturamento_existente: st.session_state.mostrar_alerta_recuperar = True; st.rerun()
                else:
                    with st.spinner("Calculando faturamento..."):
                        df_faturamento = calcular_faturamento(df_colaboradores, df_horasponto, df_valores, dias_uteis)
                        if df_faturamento.empty: st.warning("⚠️ Nenhum colaborador ativo para calcular faturamento.")
                        else:
                            total_faturado = df_faturamento["Valor Recalculado"].sum()
                            sucesso, msg = salvar_faturamento_no_banco(year, month_num, dias_uteis, total_faturado, df_faturamento)
                            if sucesso: st.session_state.df_faturamento_exibido = df_faturamento; st.session_state.total_faturado_exibido = total_faturado; st.session_state.faturamento_calculado = True; st.success(f"✅ {msg}"); st.rerun()
                            else: st.error(f"❌ {msg}")
        if st.session_state.get('mostrar_alerta_recuperar'):
            st.warning(f"⚠️ **Atenção:** Já existem dados de faturamento para **{month_name}/{year}** no banco de dados.")
            col_sim, col_nao = st.columns(2)
            with col_sim:
                if st.button("✅ Sim, Recalcular", width="stretch", type="primary", key="btn_sim_recalcular"):
                    with st.spinner("Recalculando faturamento..."):
                        sucesso_del, msg_del = deletar_faturamento_existente(year, month_num)
                        if sucesso_del:
                            df_faturamento = calcular_faturamento(df_colaboradores, df_horasponto, df_valores, dias_uteis)
                            total_faturado = df_faturamento["Valor Recalculado"].sum()
                            sucesso_save, msg_save = salvar_faturamento_no_banco(year, month_num, dias_uteis, total_faturado, df_faturamento)
                            if sucesso_save:
                                st.session_state.df_faturamento_exibido = df_faturamento; st.session_state.total_faturado_exibido = total_faturado
                                st.session_state.faturamento_calculado = True; st.session_state.mostrar_alerta_recuperar = False
                                st.success(f"✅ Faturamento recalculado e salvo! {msg_save}"); st.rerun()
                            else: st.error(f"❌ {msg_save}")
                        else: st.error(f"❌ {msg_del}")
            with col_nao:
                if st.button("❌ Não, Cancelar", width="stretch", key="btn_nao_recalcular"): st.session_state.mostrar_alerta_recuperar = False; st.rerun()
    else:
        if faturamento_existente: st.info(f"👁️ **Modo visualização** (Auditor) — Faturamento existente para **{month_name}/{year}** | Total: **{format_moeda_brl(faturamento_existente['total_faturado'])}**")
        else: st.info("👁️ **Modo visualização** (Auditor) — Nenhum faturamento calculado para este período.")
    if st.session_state.faturamento_calculado and st.session_state.df_faturamento_exibido is not None:
        df_faturamento, total_faturado = st.session_state.df_faturamento_exibido, st.session_state.total_faturado_exibido
        st.markdown("---")
        col_total1, col_total2 = st.columns([3, 1])
        with col_total1: st.markdown("### 💵 Total Faturado")
        with col_total2: st.metric(label="Total", value=format_moeda_brl(total_faturado), label_visibility="hidden")
        st.markdown("---")
        excel_data = generate_excel_faturamento(df_faturamento, month_name, year, dias_uteis, total_faturado)
        st.download_button("⬇️ Exportar Faturamento (Excel)", data=excel_data, file_name=f"faturamento_{year}_{month_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
        st.markdown("---")
        st.dataframe(df_faturamento, width="stretch", hide_index=True,
            column_config={"Profissional": st.column_config.TextColumn("Profissional", width="large"),
                "Cargo": st.column_config.TextColumn("Cargo"), "Nível": st.column_config.TextColumn("Nível"),
                "Data de Início": st.column_config.TextColumn("Data de Início"),
                "Promoção de Cargo": st.column_config.TextColumn("Promoção de Cargo"),
                "Novo Colaborador": st.column_config.TextColumn("Novo Colaborador"),
                "Data de Saída": st.column_config.TextColumn("Data de Saída"),
                "Horas Trabalhadas": st.column_config.NumberColumn("Horas Trabalhadas", format="%.2f h"),
                "Horas Afast Acordado": st.column_config.NumberColumn("Horas Afast Acordado (Folga)", format="%.2f h"),
                "Horas Afast Legal": st.column_config.NumberColumn("Horas Afast Legal", format="%.2f h"),
                "Horas Arredondadas": st.column_config.NumberColumn("Horas Arredondadas", format="%.2f h"),
                "Valor Profissional": st.column_config.NumberColumn("Valor Profissional", format="R$ %.2f"),
                "Valor Recalculado": st.column_config.NumberColumn("Valor Recalculado", format="R$ %.2f")})

def render_metricas_tab(all_workitems, project_of_item, df_colaboradores, df_horasponto, df_valores, dias_uteis, year, month_name, month_num, selected_wi_types, selected_states, parents_data, tipo_usuario='User'):
    is_auditor = (tipo_usuario == 'Auditor')
    st.subheader("📊 Métricas e Indicadores Estratégicos")
    dias_uteis_reais = dias_uteis / 8.0 if dias_uteis > 0 else 0
    st.info(f"Período: **{month_name}/{year}** | Dias úteis: **{dias_uteis_reais:.0f}** ({dias_uteis:.0f}h)")
    if is_auditor: st.info("👁️ **Modo Auditor:** Apenas métricas específicas de acompanhamento estão disponíveis.")
    faturamento_existente = verificar_faturamento_existente(year, month_num)
    total_faturado_banco = float(faturamento_existente['total_faturado']) if faturamento_existente else None
    all_capacities = get_all_capacities()
    nivel_valor = {str(r["nivel"]).strip(): limpar_valor_monetario(r.get("valor")) for _, r in df_valores.iterrows() if str(r.get("nivel", "")).strip()}
    df_colab_periodo = df_colaboradores.copy()
    azure_to_ponto, colab_nivel_map, colab_cargo_map = {}, {}, {}
    for _, row in df_colab_periodo.iterrows():
        azure_name = str(row["colaborador_azure"]).strip().lower()
        if "<" in azure_name: azure_name = azure_name.split("<")[0].strip()
        ponto_name = str(row["colaborador_ponto"]).strip()
        if azure_name: azure_to_ponto[azure_name] = ponto_name
        if ponto_name:
            colab_nivel_map[ponto_name] = str(row.get("nivel", "")).strip()
            colab_cargo_map[ponto_name] = str(row.get("cargo", "")).strip()
    horas_azure_por_colab = {}
    for item_id, item in all_workitems.items():
        f = item.get("fields", {})
        if f.get("System.State") != "Done": continue
        assigned_name = get_assigned_to_name(f.get("System.AssignedTo"))
        colaborador_ponto = encontrar_colaborador_ponto(assigned_name, azure_to_ponto)
        if colaborador_ponto != "Não atribuído" and colaborador_ponto in colab_nivel_map:
            horas_azure_por_colab[colaborador_ponto] = horas_azure_por_colab.get(colaborador_ponto, 0.0) + float(f.get("Microsoft.VSTS.Scheduling.CompletedWork") or 0.0)
    st.markdown("---"); st.markdown("### 💰 1. Métricas Financeiras e de Eficiência")
    total_capacity_geral, total_horas_ponto, total_horas_azure, total_horas_afastamento = 0.0, 0.0, 0.0, 0.0
    metricas_colab = []
    for _, hp_row in df_horasponto.iterrows():
        ponto_name = str(hp_row["colaborador_ponto"]).strip()
        if ponto_name not in colab_nivel_map: continue
        horas_ponto = float(hp_row["horas_ponto"]) if pd.notna(hp_row.get("horas_ponto")) else 0.0
        afastamento = float(hp_row["afastamento"]) if pd.notna(hp_row.get("afastamento")) else 0.0
        folga = float(hp_row["folga"]) if pd.notna(hp_row.get("folga")) else 0.0
        horas_azure = horas_azure_por_colab.get(ponto_name, 0.0)
        nivel = colab_nivel_map[ponto_name]; cargo = colab_cargo_map.get(ponto_name, "")
        capacity = all_capacities.get(nivel, 0.0)
        capacity_calculado = dias_uteis_reais * capacity if capacity > 0 else 0.0
        utilizacao = (horas_azure / capacity_calculado * 100) if capacity_calculado > 0 else 0.0
        gap_apontamento = horas_ponto - horas_azure; capacidade_ociosa = capacity_calculado - horas_azure
        horas_arredondadas = horas_ponto + folga
        metricas_colab.append({"Colaborador": ponto_name, "Cargo": cargo, "Nível": nivel,
            "Horas Ponto": horas_ponto, "Horas Azure": horas_azure, "Capacity": capacity_calculado, "Utilização %": round(utilizacao, 1),
            "Gap Apontamento": round(gap_apontamento, 2), "Capacidade Ociosa": round(capacidade_ociosa, 2),
            "Afastamento": afastamento, "Folga": folga, "Horas Arredondadas": horas_arredondadas})
        total_capacity_geral += capacity_calculado; total_horas_ponto += horas_ponto
        total_horas_azure += horas_azure; total_horas_afastamento += afastamento
    utilizacao_geral = (total_horas_azure / total_capacity_geral * 100) if total_capacity_geral > 0 else 0.0
    gap_geral = total_horas_ponto - total_horas_azure; capacidade_ociosa_geral = total_capacity_geral - total_horas_azure
    ticket_medio_hora = (total_faturado_banco / total_horas_ponto) if (total_faturado_banco and total_horas_ponto > 0) else 0.0
    if not is_auditor:
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📊 Utilização da Capacidade", f"{utilizacao_geral:.1f}%")
        col2.metric("⚠️ Gap de Apontamento", f"{gap_geral:.1f} h", help="Horas trabalhadas (Ponto) - Horas registradas no Azure")
        col3.metric("💵 Ticket Médio/Hora", format_moeda_brl(ticket_medio_hora))
        if total_faturado_banco is not None: col4.metric("💰 Total Faturado", format_moeda_brl(total_faturado_banco))
        else: col4.metric("💰 Total Faturado", "Não calculado")
        col5, col6 = st.columns(2)
        col5.metric("📉 Capacidade Ociosa", f"{capacidade_ociosa_geral:.1f} h", help="Capacity teórico - Horas Azure")
        col6.metric("🏖️ Total Afastamento", f"{total_horas_afastamento:.1f} h")
    else:
        col1, col2 = st.columns(2)
        col1.metric("⚠️ Gap de Apontamento", f"{gap_geral:.1f} h"); col2.metric("🏖️ Total Afastamento", f"{total_horas_afastamento:.1f} h")
    if metricas_colab:
        df_metricas = pd.DataFrame(metricas_colab)
        st.markdown("---"); st.subheader("📈 Utilização da Capacidade por Colaborador")
        fig_util = px.bar(df_metricas, x="Colaborador", y="Utilização %", color="Utilização %", color_continuous_scale=["#dc3545", "#ffc107", "#28a745"], text="Utilização %")
        fig_util.update_layout(margin=dict(t=20, b=80, l=20, r=20), height=400, xaxis_tickangle=-45)
        fig_util.update_traces(texttemplate='%{text:.0f}%', textposition='outside')
        fig_util.add_hline(y=85, line_dash="dash", line_color="green", annotation_text="Ideal (85%)")
        fig_util.add_hline(y=100, line_dash="dash", line_color="red", annotation_text="Máximo (100%)")
        st.plotly_chart(fig_util, width="stretch")
    st.markdown("---"); st.markdown("### 🚀 2. Métricas de Produtividade e Agile")
    lead_times, cycle_times, items_done = [], [], []
    for item_id, item in all_workitems.items():
        f = item.get("fields", {})
        if f.get("System.State") != "Done": continue
        created = parse_date(f.get("System.CreatedDate")); closed = parse_date(f.get("Microsoft.VSTS.Common.ClosedDate"))
        activated = parse_date(f.get("Microsoft.VSTS.Common.ActivatedDate"))
        if created and closed: lead_times.append((closed - created).total_seconds() / 3600)
        if activated and closed: cycle_times.append((closed - activated).total_seconds() / 3600)
        items_done.append(item)
    lead_time_medio = sum(lead_times) / len(lead_times) if lead_times else 0
    cycle_time_medio = sum(cycle_times) / len(cycle_times) if cycle_times else 0
    sprints_dict = {}
    for item in items_done:
        f = item.get("fields", {})
        sprint_name = str(f.get("System.IterationPath", "N/A")).split("\\")[-1]
        if sprint_name not in sprints_dict: sprints_dict[sprint_name] = {"qtd": 0, "horas": 0.0}
        sprints_dict[sprint_name]["qtd"] += 1; sprints_dict[sprint_name]["horas"] += float(f.get("Microsoft.VSTS.Scheduling.CompletedWork") or 0.0)
    velocidade_media_itens = sum(s["qtd"] for s in sprints_dict.values()) / len(sprints_dict) if sprints_dict else 0
    velocidade_media_horas = sum(s["horas"] for s in sprints_dict.values()) / len(sprints_dict) if sprints_dict else 0
    if not is_auditor:
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("⏱️ Lead Time Médio", f"{lead_time_medio/24:.1f} dias"); col2.metric("🔄 Cycle Time Médio", f"{cycle_time_medio/24:.1f} dias")
        col3.metric("📦 Velocidade Média", f"{velocidade_media_itens:.0f} itens/sprint"); col4.metric("⚡ Velocidade em Horas", f"{velocidade_media_horas:.1f}h/sprint")
    else:
        col1, col2 = st.columns(2)
        col1.metric("🔄 Cycle Time Médio", f"{cycle_time_medio/24:.1f} dias"); col2.metric("📦 Velocidade Média", f"{velocidade_media_itens:.0f} itens/sprint")
    st.markdown("---"); st.markdown("### 🎯 3. Distribuição de Esforço: Manutenção vs Evolução")
    categorias = {"Evolução (Novas Funcionalidades)": ["Product Backlog Item", "Melhoria"], "Manutenção Corretiva": ["Bug", "BugFix", "HotFix"], "Débito Técnico": ["Débito Técnico"], "Outros": ["Task", "Design"]}
    horas_por_categoria = {cat: 0.0 for cat in categorias.keys()}; qtd_por_categoria = {cat: 0 for cat in categorias.keys()}
    for item_id, item in all_workitems.items():
        f = item.get("fields", {}); wtype = f.get("System.WorkItemType", "")
        if wtype not in selected_wi_types: continue
        completed = float(f.get("Microsoft.VSTS.Scheduling.CompletedWork") or 0.0)
        for cat, tipos in categorias.items():
            if wtype in tipos: horas_por_categoria[cat] += completed; qtd_por_categoria[cat] += 1; break
    total_horas_esforco = sum(horas_por_categoria.values())
    if not is_auditor:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### 📊 Distribuição por Horas Apontadas")
            if total_horas_esforco > 0:
                df_esforco = pd.DataFrame([{"Categoria": cat, "Horas": round(horas_por_categoria[cat], 2), "Percentual": round(horas_por_categoria[cat]/total_horas_esforco*100, 1)} for cat in categorias.keys()])
                fig_esforco = px.pie(df_esforco, values="Horas", names="Categoria", hole=0.5, color_discrete_sequence=px.colors.qualitative.Set2)
                fig_esforco.update_traces(textposition='inside', textinfo='percent+label'); fig_esforco.update_layout(margin=dict(t=20, b=20, l=20, r=20), height=350)
                st.plotly_chart(fig_esforco, width="stretch")
            else: st.info("Sem dados de horas apontadas")
        with col2:
            st.markdown("#### 📋 Distribuição por Quantidade de Itens")
            total_qtd = sum(qtd_por_categoria.values())
            if total_qtd > 0:
                df_qtd = pd.DataFrame([{"Categoria": cat, "Quantidade": qtd_por_categoria[cat], "Percentual": round(qtd_por_categoria[cat]/total_qtd*100, 1)} for cat in categorias.keys()])
                fig_qtd = px.pie(df_qtd, values="Quantidade", names="Categoria", hole=0.5, color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_qtd.update_traces(textposition='inside', textinfo='percent+label'); fig_qtd.update_layout(margin=dict(t=20, b=20, l=20, r=20), height=350)
                st.plotly_chart(fig_qtd, width="stretch")
            else: st.info("Sem dados de quantidade")
    st.markdown("---"); st.markdown("### 👥 4. Métricas de Gestão de Pessoas")
    taxa_absenteismo = (total_horas_afastamento / total_capacity_geral * 100) if total_capacity_geral > 0 else 0.0
    st.markdown("#### 🏖️ Taxa de Absenteísmo")
    st.metric("Taxa de Afastamento", f"{taxa_absenteismo:.1f}%", help=f"Horas de afastamento ({total_horas_afastamento:.0f}h) / Capacity total ({total_capacity_geral:.0f}h)")
    if taxa_absenteismo > 15: st.error(f"🚨 Absenteísmo alto ({taxa_absenteismo:.1f}%).")
    elif taxa_absenteismo > 10: st.warning(f"⚠️ Absenteísmo moderado ({taxa_absenteismo:.1f}%)")
    else: st.success(f"✅ Absenteísmo dentro do esperado ({taxa_absenteismo:.1f}%)")
    st.markdown("---"); st.markdown("### 📋 5. Resumo Executivo do Período")
    pct_evolucao = (horas_por_categoria['Evolução (Novas Funcionalidades)'] / total_horas_esforco * 100) if total_horas_esforco > 0 else 0.0
    pct_manutencao = (horas_por_categoria['Manutenção Corretiva'] / total_horas_esforco * 100) if total_horas_esforco > 0 else 0.0
    total_faturado_exibir = total_faturado_banco if total_faturado_banco is not None else 0.0
    resumo_data = {
        "Indicador": ["Total de Colaboradores no Período", "Total de Horas Apontadas (Azure)", "Total de Horas do Ponto", "Capacidade Total da Equipe", "Taxa de Utilização", "Gap de Apontamento", "Capacidade Ociosa", "Total Faturado", "Ticket Médio por Hora", "Lead Time Médio", "Cycle Time Médio", "Velocidade (itens/sprint)", "% Manutenção/Corretivo", "% Evolução/Novas Func.", "Taxa de Absenteísmo"],
        "Valor": [f"{len(df_colab_periodo)}", f"{total_horas_azure:.1f} h", f"{total_horas_ponto:.1f} h", f"{total_capacity_geral:.1f} h", f"{utilizacao_geral:.1f}%", f"{gap_geral:.1f} h", f"{capacidade_ociosa_geral:.1f} h", format_moeda_brl(total_faturado_exibir), format_moeda_brl(ticket_medio_hora), f"{lead_time_medio/24:.1f} dias", f"{cycle_time_medio/24:.1f} dias", f"{velocidade_media_itens:.0f} itens", f"{pct_manutencao:.1f}%", f"{pct_evolucao:.1f}%", f"{taxa_absenteismo:.1f}%"],
        "Fórmula": ["Count(colaboradores)", "SUM(CompletedWork)", "SUM(horas_ponto)", "Dias_úteis × Capacity_nível", "Azure / Capacity × 100", "Ponto - Azure", "Capacity - Azure", "Horas × Valor_hora", "Faturado / Horas_ponto", "ClosedDate - CreatedDate", "ClosedDate - ActivatedDate", "Itens Done / Sprints", "Manutenção / Total × 100", "Evolução / Total × 100", "Afastamento / Capacity × 100"]}
    st.dataframe(pd.DataFrame(resumo_data), width="stretch", hide_index=True)

def render_crud_perfil():
    tipo_usuario, usuario_id = st.session_state.user_info.get('cargo', 'User'), st.session_state.user_info.get('id')
    st.markdown("---"); st.subheader("⚙️ Gerenciamento de Perfil")
    usuarios = listar_usuarios(tipo_usuario, usuario_id)
    if not usuarios: st.warning("Nenhum usuário encontrado."); return
    st.markdown("### Usuários Cadastrados")
    df_exibicao = pd.DataFrame(usuarios)
    df_exibicao['apikey'] = df_exibicao['apikey'].apply(lambda x: f"{x[:6]}...{x[-4:]}" if x and len(x) > 10 else x)
    df_exibicao['criado_em'] = pd.to_datetime(df_exibicao['criado_em']).dt.strftime('%d/%m/%Y %H:%M')
    st.dataframe(df_exibicao[['id', 'nome', 'email', 'tipo', 'status', 'apikey', 'criado_em']], width="stretch", hide_index=True)
    st.markdown("---")
    if tipo_usuario == 'Admin':
        usuario_selecionado = st.selectbox("Selecione o usuário para editar:", options=[f"{u['id']} - {u['nome']} ({u['email']})" for u in usuarios], index=0)
        usuario_id_selecionado = int(usuario_selecionado.split(" - ")[0])
        usuario_dados = next((u for u in usuarios if u['id'] == usuario_id_selecionado), None)
    else:
        usuario_dados, usuario_id_selecionado = usuarios[0], usuario_id
        st.info(f"Editando seu próprio perfil: **{usuario_dados['nome']}**")
    if not usuario_dados: return
    with st.form("form_editar_usuario"):
        nome = st.text_input("Nome Completo", value=usuario_dados.get('nome', ''))
        email = st.text_input("E-mail", value=usuario_dados.get('email', ''))
        apikey = st.text_input("API Key do Azure DevOps", value=usuario_dados.get('apikey', ''), type="password")
        senha_hash = st.text_input("Nova Senha (deixe em branco para manter)", type="password", placeholder="Mínimo 6 caracteres")
        if tipo_usuario == 'Admin':
            col1, col2 = st.columns(2)
            with col1: tipo = st.selectbox("Tipo de Usuário", options=['User', 'Admin', 'Gestor', 'Auditor'], index=['User', 'Admin', 'Gestor', 'Auditor'].index(usuario_dados.get('tipo', 'User')))
            with col2: status = st.selectbox("Status", options=['Ativo', 'Inativo', 'Bloqueado', 'Pendente'], index=['Ativo', 'Inativo', 'Bloqueado', 'Pendente'].index(usuario_dados.get('status', 'Ativo')))
        else: tipo, status = usuario_dados.get('tipo', 'User'), usuario_dados.get('status', 'Ativo')
        submitted = st.form_submit_button("💾 Salvar Alterações", width="stretch", key="btn_salvar_perfil")
        if submitted:
            dados = {'nome': nome.strip(), 'email': email.lower().strip(), 'apikey': apikey.strip()}
            if senha_hash and len(senha_hash) >= 6: dados['senha_hash'] = senha_hash
            if tipo_usuario == 'Admin': dados['tipo'], dados['status'] = tipo, status
            sucesso, msg = atualizar_usuario(usuario_id_selecionado, dados, tipo_usuario, usuario_id)
            if sucesso:
                st.success(f"✅ {msg}")
                if usuario_id_selecionado == usuario_id:
                    st.session_state.user_info['nome_completo'] = nome.strip()
                    st.session_state.user_info['email'] = email.lower().strip()
                if apikey.strip(): st.session_state.user_info['apikey'] = apikey.strip()
                st.rerun()
            else: st.error(f"❌ {msg}")

def render_login_screen():
    inject_custom_css()
    st.markdown('<div class="dotted-rings"><div class="ring ring-1"></div><div class="ring ring-2"></div><div class="ring ring-3"></div><div class="ring ring-4"></div><div class="ring ring-5"></div></div>', unsafe_allow_html=True)
    col1, col2 = st.columns([1.1, 0.9], gap="large")
    with col1: st.markdown('<div style="margin-top: 40px;"><span class="badge-qween">Qween Platform</span><h1 class="hero-title">Relatório Mensal de acompanhamento da Execução</h1><p class="hero-subtitle">Secretaria Municipal de Educação - SP</p></div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="auth-card">', unsafe_allow_html=True)
        view = st.session_state.get("login_view", "login")
        if view == "login":
            st.markdown("<h2 style='color:white; font-weight:700;'>Entrar</h2>", unsafe_allow_html=True)
            st.markdown("<p style='color:#e6e6e6; margin-bottom:30px;'>Acesse com suas credenciais institucionais</p>", unsafe_allow_html=True)
            with st.form("login_form"):
                email = st.text_input("E-mail ou Usuário", placeholder="nome@spassu.com.br")
                password = st.text_input("Senha", type="password", placeholder="••••••••")
                submitted = st.form_submit_button("Acessar Painel", width="stretch", key="btn_acessar")
                if submitted:
                    if email and password:
                        user = authenticate_user(email, password)
                        if user:
                            if user.get('status') != 'Ativo': st.error("⚠️ Usuário inativo. Contate o administrador.")
                            elif not user.get('apikey'): st.error("⚠️ Seu usuário não possui API Key do Azure DevOps cadastrada.")
                            else:
                                st.session_state.logged_in = True
                                st.session_state.user_info = {'id': user['id'], 'nome_completo': user['nome'], 'email': user['email'], 'cargo': user['tipo'], 'apikey': user['apikey']}
                                st.session_state.login_view = "login"; st.rerun()
                        else: st.error("E-mail ou senha inválidos.")
                    else: st.warning("Preencha todos os campos.")
            st.markdown("")
            col_btn1, col_btn2 = st.columns([3, 1])
            with col_btn2:
                if st.button("Esqueci minha senha", width="stretch", key="btn_esqueci_senha"): st.session_state.login_view = "forgot"; st.rerun()
            st.markdown('<div style="display:flex; align-items:center; text-align:center; margin:28px 0; color:rgba(255,255,255,0.7); font-size:0.8rem; font-weight:600;"><div style="flex:1; border-bottom:1px solid rgba(255,255,255,0.25); margin-right:12px;"></div>OU<div style="flex:1; border-bottom:1px solid rgba(255,255,255,0.25); margin-left:12px;"></div></div>', unsafe_allow_html=True)
            if st.button("Primeiro Acesso? Ativar Conta", width="stretch", key="btn_primeiro_acesso"): st.session_state.login_view = "register"; st.rerun()
        elif view == "forgot":
            st.markdown("<h2 style='color:white; font-weight:700;'>Recuperação de Senha</h2>", unsafe_allow_html=True)
            st.markdown("<p style='color:#e6e6e6; margin-bottom:30px;'>Insira seu e-mail para receber o token de recuperação</p>", unsafe_allow_html=True)
            with st.form("forgot_form"):
                email = st.text_input("E-mail Cadastrado", placeholder="seu.email@spassu.com.br")
                submitted = st.form_submit_button("Gerar Token de Recuperação", width="stretch", key="btn_gerar_token")
                if submitted:
                    if not email: st.error("⚠️ Informe o e-mail.")
                    else:
                        sucesso, msg, token = solicitar_reset_senha(email)
                        if sucesso:
                            st.success(f"✅ {msg}")
                            st.markdown(f'<div class="token-box">{token}</div>', unsafe_allow_html=True)
                            st.session_state.reset_token = token; st.session_state.reset_email = email; st.session_state.go_to_reset = True
                        else: st.error(f"❌ {msg}")
            if st.session_state.get('go_to_reset'):
                if st.button("Redefinir Senha Agora →", width="stretch", key="btn_redefinir_agora"): st.session_state.login_view = "reset_password"; st.session_state.go_to_reset = False; st.rerun()
            if st.button("← Voltar ao Login", width="stretch", key="btn_voltar_login_forgot"): st.session_state.login_view = "login"; st.session_state.go_to_reset = False; st.rerun()
        elif view == "reset_password":
            st.markdown("<h2 style='color:white; font-weight:700;'>Redefinir Senha</h2>", unsafe_allow_html=True)
            with st.form("reset_form"):
                token = st.text_input("Token de Recuperação", value=st.session_state.get('reset_token', ''))
                nova_senha = st.text_input("Nova Senha", type="password")
                confirmar_senha = st.text_input("Confirme a Nova Senha", type="password")
                submitted = st.form_submit_button("Redefinir Senha", width="stretch", key="btn_redefinir_senha")
                if submitted:
                    if not token or not nova_senha or nova_senha != confirmar_senha: st.error("⚠️ Verifique os dados informados.")
                    else:
                        sucesso, msg = redefinir_senha(token, nova_senha)
                        if sucesso: st.success(f"✅ {msg}"); st.session_state.login_view = "login"; st.rerun()
                        else: st.error(f"❌ {msg}")
            if st.button("← Voltar ao Login", width="stretch", key="btn_voltar_login_reset"): st.session_state.login_view = "login"; st.rerun()
        elif view == "register":
            st.markdown("<h2 style='color:white; font-weight:700;'>Primeiro Acesso</h2>", unsafe_allow_html=True)
            with st.form("register_form", clear_on_submit=True):
                nome = st.text_input("Nome Completo"); email = st.text_input("E-mail Corporativo")
                apikey = st.text_input("API Key do Azure DevOps")
                senha = st.text_input("Crie uma Senha", type="password")
                confirmar_senha = st.text_input("Confirme a Senha", type="password")
                submitted = st.form_submit_button("Validar e Criar Conta", width="stretch", key="btn_criar_conta")
                if submitted:
                    if not nome or not email or not apikey or not senha or senha != confirmar_senha: st.error("⚠️ Preencha os campos corretamente.")
                    else:
                        sucesso, msg = criar_usuario_com_apikey(nome, email, senha, apikey.strip())
                        if sucesso: st.success(f"✅ {msg}"); st.session_state.login_view = "login"; st.rerun()
                        else: st.error(f"❌ {msg}")
            if st.button("Já possuo senha (Entrar)", width="stretch", key="btn_ja_possuo_senha"): st.session_state.login_view = "login"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown(f'<div style="position:fixed;bottom:30px;right:40px;display:flex;align-items:center;gap:20px;z-index:100;"><img src="{SPASSU_LOGO_URL}" style="height:40px;"><div style="width:1px;height:40px;background:rgba(255,255,255,0.5);"></div><img src="{SME_LOGO_URL}" style="height:40px;"></div>', unsafe_allow_html=True)

# =============================================================================
# MAIN
# =============================================================================
def main():
    if 'logged_in' not in st.session_state: st.session_state.logged_in = False
    if 'login_view' not in st.session_state: st.session_state.login_view = "login"
    if 'show_perfil' not in st.session_state: st.session_state.show_perfil = False
    if 'historico_colab_id' not in st.session_state: st.session_state.historico_colab_id = None
    if 'historico_colab_nome' not in st.session_state: st.session_state.historico_colab_nome = None
    if not st.session_state.logged_in: render_login_screen(); return
    st.set_page_config(page_title="Dashboard de Entregas - SME Spassu", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
    st.sidebar.markdown("---")
    st.sidebar.success(f"👤 {st.session_state.user_info['nome_completo']}")
    st.sidebar.caption(f"Tipo: {st.session_state.user_info.get('cargo', 'User')}")
    if st.sidebar.button("⚙️ Perfil", width="stretch", key="btn_perfil"): st.session_state.show_perfil = True; st.rerun()
    if st.sidebar.button("🚪 Sair do Sistema", width="stretch", key="btn_sair"):
        st.session_state.logged_in = False; st.session_state.user_info = None; st.session_state.show_perfil = False; st.rerun()
    if st.session_state.get('show_perfil'):
        st.title("⚙️ Gerenciamento de Perfil"); render_crud_perfil()
        if st.button("← Voltar ao Dashboard", key="btn_voltar_dashboard"): st.session_state.show_perfil = False; st.rerun()
        return
    st.title("📊 Dashboard de Acompanhamento de Entregas")
    st.caption("Azure DevOps • SME-Spassu • Sprint Tracker")
    st.markdown("---")
    st.sidebar.header("⚙️ Configuração & Filtros")
    if st.sidebar.button("🔄 Atualizar Dados", width="stretch", type="primary", help="Recarrega todos os dados do Azure DevOps e banco de dados", key="btn_atualizar_dados"):
        st.cache_data.clear(); st.cache_resource.clear(); st.session_state.periodo_anterior = None; st.rerun()
    gestores_lista = listar_gestores_filtro()
    gestores_nomes = [g['nome'] for g in gestores_lista]
    gestor_nome_para_id = {g['nome']: g['id'] for g in gestores_lista}
    if 'filtro_gestor' not in st.session_state: st.session_state.filtro_gestor = []
    st.sidebar.markdown("**👔 Gestor**")
    gestores_selecionados_nomes = st.sidebar.multiselect(
        "Filtrar por Gestor", options=gestores_nomes,
        default=[n for n in st.session_state.filtro_gestor if n in gestores_nomes],
        key="filtro_gestor_sidebar",
        help="Selecione um ou mais gestores para filtrar os colaboradores"
    )
    if gestores_selecionados_nomes != st.session_state.filtro_gestor:
        st.session_state.filtro_gestor = gestores_selecionados_nomes
    gestores_selecionados_ids = [gestor_nome_para_id[n] for n in gestores_selecionados_nomes if n in gestor_nome_para_id]
    pat = st.session_state.user_info.get('apikey', '')
    if not pat: st.error("⚠️ API Key do Azure DevOps não cadastrada para este usuário."); st.stop()
    api_valida, msg_erro = validate_ado_api_key(pat)
    if not api_valida: st.error(f"🚫 {msg_erro}"); st.stop()
    current_year = datetime.now().year
    year = st.sidebar.selectbox("📅 Ano", options=[current_year + 1, current_year, current_year - 1], index=1 if current_year == 2026 else 0)
    month_names = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    month_name = st.sidebar.selectbox("📅 Mês", options=month_names, index=datetime.now().month - 1)
    month_num = month_names.index(month_name) + 1
    periodo_anterior = st.session_state.get('periodo_anterior', f"{year}-{month_num}")
    periodo_atual = f"{year}-{month_num}"
    if periodo_anterior != periodo_atual:
        for key in ['filtros_sprints']:
            if key in st.session_state: del st.session_state[key]
        st.session_state.periodo_anterior = periodo_atual
    all_projects = get_active_projects_from_db(pat)
    selected_projects = multiselect_with_toggle("🏢 Projetos", all_projects, all_projects, "filtros_projetos")
    if not selected_projects: st.info("Selecione ao menos um projeto."); st.stop()
    with st.spinner("🔄 Buscando sprints..."):
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = {executor.submit(fetch_all_sprints_for_project, pat, proj, year, month_num): proj for proj in selected_projects}
            sprints_by_project = {futures[f]: f.result() for f in as_completed(futures) if f.result()}
    if not sprints_by_project: st.warning(f"⚠️ Nenhuma sprint encontrada em {month_name}/{year}."); st.stop()
    sprint_options = [f"{s['path']}" for proj, sprints in sprints_by_project.items() for s in sprints]
    selected_sprints_display = multiselect_with_toggle("🏃 Sprints", sprint_options, sprint_options, "filtros_sprints")
    sprint_to_info = {}
    for proj, sprints in sprints_by_project.items():
        for s in sprints: sprint_to_info[f"{proj}|||{s['path']}"] = (proj, s["path"])
    selected_sprint_paths, paths_by_project = [], {}
    for display_name in selected_sprints_display:
        for key, (proj, path) in sprint_to_info.items():
            if path == display_name:
                selected_sprint_paths.append((proj, path)); paths_by_project.setdefault(proj, []).append(path); break
    selected_wi_types = multiselect_with_toggle("📋 Tipo de Work Item", ALL_WI_TYPES, DEFAULT_WI_TYPES, "filtros_wi_types")
    selected_states = multiselect_with_toggle("🎯 Situação", ALL_STATES, ALL_STATES, "filtros_states")
    with st.spinner("🔍 Buscando work items e hierarquias em paralelo..."):
        all_ids_by_project = {}; children_map = {}
        with ThreadPoolExecutor(max_workers=6) as executor:
            fut_ids = {executor.submit(fetch_all_workitems_ids_by_project, pat, proj, tuple(paths)): proj for proj, paths in paths_by_project.items()}
            fut_links = {executor.submit(fetch_hierarchy_links_by_project, pat, proj, tuple(paths)): proj for proj, paths in paths_by_project.items()}
            for f in as_completed(fut_ids):
                proj = fut_ids[f]; ids_res, _ = f.result()
                if ids_res: all_ids_by_project[proj] = ids_res
            for f in as_completed(fut_links):
                relations, _ = f.result()
                for rel in relations:
                    if rel and rel.get("source") and rel.get("target"):
                        src, tgt = rel["source"].get("id"), rel["target"].get("id")
                        if rel.get("rel") == "System.LinkTypes.Hierarchy-Forward": children_map.setdefault(src, []).append(tgt)
                        elif rel.get("rel") == "System.LinkTypes.Hierarchy-Reverse": children_map.setdefault(tgt, []).append(src)
        all_workitems, project_of_item = {}, {}
        with ThreadPoolExecutor(max_workers=5) as executor:
            fut_details = {executor.submit(fetch_workitems_details_batch, pat, proj, tuple(ids)): proj for proj, ids in all_ids_by_project.items() if ids}
            for f in as_completed(fut_details):
                proj = fut_details[f]; items, _ = f.result()
                for item_id, item in items.items():
                    all_workitems[item_id] = item; project_of_item[item_id] = proj
    df_colaboradores = listar_colaboradores()
    if gestores_selecionados_ids:
        df_colaboradores = df_colaboradores[df_colaboradores['gestor'].isin(gestores_selecionados_ids)].copy()
        st.sidebar.info(f"👔 Filtrando {len(df_colaboradores)} colaborador(es) dos gestores selecionados")
    colaboradores_filtrados_ids = set(df_colaboradores['id'].tolist()) if not df_colaboradores.empty else set()
    colaboradores_filtrados_azure = set()
    for _, row in df_colaboradores.iterrows():
        azure_name = str(row["colaborador_azure"]).strip().lower()
        if "<" in azure_name: azure_name = azure_name.split("<")[0].strip()
        colaboradores_filtrados_azure.add(azure_name)
        colaboradores_filtrados_azure.add(str(row["colaborador_azure"]).strip().lower())
    all_workitems_filtrados = {}
    project_of_item_filtrado = {}
    for item_id, item in all_workitems.items():
        f = item.get("fields", {})
        assigned_to = f.get("System.AssignedTo")
        if assigned_to:
            assigned_name = get_assigned_to_name(assigned_to).strip().lower()
            if "<" in assigned_name: assigned_name = assigned_name.split("<")[0].strip()
            if assigned_name in colaboradores_filtrados_azure:
                all_workitems_filtrados[item_id] = item
                project_of_item_filtrado[item_id] = project_of_item.get(item_id, "")
    parents_data = []
    valid_paths = set(p for _, p in selected_sprint_paths)
    for item_id, item in all_workitems_filtrados.items():
        f = item.get("fields", {})
        if f.get("System.WorkItemType") in selected_wi_types and f.get("System.State") in selected_states and f.get("System.IterationPath", "") in valid_paths:
            item["_project"] = project_of_item_filtrado.get(item_id, "")
            parents_data.append(item)
    df_horasponto = get_horasponto_from_db(year, month_num)
    if not df_horasponto.empty and 'id_colaborador' in df_horasponto.columns:
        df_horasponto = df_horasponto[df_horasponto['id_colaborador'].isin(colaboradores_filtrados_ids)].copy()
    horas_uteis_mensais = get_dias_uteis_from_db(year, month_num)
    df_valores = get_valores_from_db()
    tipo_usuario = st.session_state.user_info.get('cargo', 'User')
    if tipo_usuario in ['Admin', 'Gestor']:
        tab_lista, tab_colaborador, tab_horas, tab_faturamento, tab_metricas, tab_ponto, tab_estruturacao, tab_crud_colab = st.tabs([
            "📋 Lista Geral", "👥 Tarefas por Colaborador", "🕐 Horas por Colaborador", "💰 Faturamento", "📊 Métricas", "🕐 Ponto", "🏗️ Estruturação", "⚙️ Colaboradores"])
    elif tipo_usuario == 'Auditor':
        tab_lista, tab_colaborador, tab_horas, tab_faturamento, tab_metricas = st.tabs([
            "📋 Lista Geral", "👥 Tarefas por Colaborador", "🕐 Horas por Colaborador", "💰 Faturamento", "📊 Métricas"])
    else:
        tab_lista, tab_colaborador, tab_horas, tab_metricas, tab_ponto = st.tabs([
            "📋 Lista Geral", "👥 Tarefas por Colaborador", "🕐 Horas por Colaborador", "📊 Métricas", "🕐 Ponto"])
    with tab_lista:
        total_items = len(parents_data)
        total_hours = sum(float(p.get("fields", {}).get("Microsoft.VSTS.Scheduling.OriginalEstimate") or 0) for p in parents_data)
        done_resolved = sum(1 for p in parents_data if p.get("fields", {}).get("System.State") in ["Done", "Resolved"])
        in_progress_todo = sum(1 for p in parents_data if p.get("fields", {}).get("System.State") in ["In Progress", "To Do"])
        historias_nao_done = sum(1 for p in parents_data if p.get("fields", {}).get("System.WorkItemType") == "Product Backlog Item" and p.get("fields", {}).get("System.State") != "Done")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📋 Total de Itens", total_items); col2.metric("⏱️ Total de Horas", f"{total_hours:.2f} h")
        col3.metric("📊 Situação", f"✅ {done_resolved} | 🔄 {in_progress_todo}")
        col4.metric("📖 Histórias em Aberto", historias_nao_done, help="Product Backlog Items com status diferente de Done")
        st.markdown("---")
        if parents_data:
            type_counts = {}
            for p in parents_data:
                t = p.get("fields", {}).get("System.WorkItemType", "Desconhecido"); type_counts[t] = type_counts.get(t, 0) + 1
            df_types = pd.DataFrame([{"Tipo": k, "Quantidade": v} for k, v in sorted(type_counts.items(), key=lambda x: x[1], reverse=True)])
            state_counts = {}
            for p in parents_data:
                s = p.get("fields", {}).get("System.State", "Desconhecido"); state_counts[s] = state_counts.get(s, 0) + 1
            df_states = pd.DataFrame([{"Situação": k, "Quantidade": v} for k, v in sorted(state_counts.items(), key=lambda x: x[1], reverse=True)])
            state_colors = {"Done": "#28a745", "Resolved": "#17a2b8", "In Progress": "#ffc107", "To Do": "#6c757d"}
            col_chart1, col_chart2 = st.columns(2)
            with col_chart1:
                st.subheader("📊 Total por Tipo")
                fig_type = px.pie(df_types, values="Quantidade", names="Tipo", hole=0.5, color_discrete_sequence=px.colors.qualitative.Set2)
                fig_type.update_layout(margin=dict(t=20, b=20, l=20, r=20), height=350, showlegend=True)
                st.plotly_chart(fig_type, width="stretch")
            with col_chart2:
                st.subheader("📊 Total por Situação")
                fig_state = px.pie(df_states, values="Quantidade", names="Situação", hole=0.5, color="Situação", color_discrete_map=state_colors)
                fig_state.update_layout(margin=dict(t=20, b=20, l=20, r=20), height=350, showlegend=True)
                st.plotly_chart(fig_state, width="stretch")
            st.markdown("---"); st.subheader("📥 Exportar Dados"); st.markdown("**📋 Work Items Gerais**")
            col_csv1, col_csv2 = st.columns(2)
            with col_csv1:
                csv_parents = generate_csv_parents_only(parents_data, children_map, all_workitems_filtrados)
                st.download_button("⬇️ Baixar CSV (Apenas Pais)", data=csv_parents, file_name=f"sme_sprints_{year}_{month_name}_pais.csv", mime="text/csv", width="stretch")
            with col_csv2:
                csv_complete = generate_csv_complete(parents_data, children_map, all_workitems_filtrados)
                st.download_button("⬇️ Baixar CSV (Completo)", data=csv_complete, file_name=f"sme_sprints_{year}_{month_name}_completo.csv", mime="text/csv", width="stretch")
            st.markdown("---"); st.markdown("**📖 Histórias (Product Backlog Item) em Aberto**")
            st.caption("Product Backlog Items com status diferente de Done no período selecionado")
            csv_backlogs, count_backlogs = generate_csv_backlogs_nao_done(parents_data)
            col_csv3, col_csv4 = st.columns(2)
            with col_csv3:
                st.download_button(f"⬇️ Baixar CSV - Histórias em Aberto ({count_backlogs} itens)", data=csv_backlogs, file_name=f"backlogs_abertos_{year}_{month_name}.csv", mime="text/csv", width="stretch", help="Exporta Product Backlog Items com status diferente de Done")
            with col_csv4: st.info(f"📊 **{count_backlogs}** histórias em aberto encontradas no período")
            st.markdown("---"); st.markdown("**✅ Tasks Concluídas sem Horas Apontadas**")
            st.caption("Tasks com Status = Done e Completed Work = 0 ou vazio")
            csv_tasks, count_tasks = generate_csv_tasks_done_sem_horas(parents_data)
            col_csv5, col_csv6 = st.columns(2)
            with col_csv5:
                st.download_button(f"⬇️ Baixar CSV - Tasks Done sem Horas ({count_tasks} itens)", data=csv_tasks, file_name=f"tasks_done_sem_horas_{year}_{month_name}.csv", mime="text/csv", width="stretch", help="Exporta Tasks com State=Done e CompletedWork=0 ou vazio")
            with col_csv6: st.info(f"📊 **{count_tasks}** tasks concluídas sem horas apontadas")
            st.markdown("---")
            render_table_html(parents_data, children_map, all_workitems_filtrados, total_hours)
        else: st.warning("⚠️ Nenhum item após filtros.")
    with tab_colaborador:
        st.subheader("👥 Tarefas por Colaborador")
        st.info(f"Período: **{month_name}/{year}** | Apenas **State = Done**")
        if not df_colaboradores.empty: render_colaborador_tab(all_workitems_filtrados, project_of_item_filtrado, df_colaboradores, year, month_name)
        else: st.warning("⚠️ Tabela de colaboradores vazia.")
    with tab_horas:
        st.subheader("🕐 Horas por Colaborador")
        if df_colaboradores.empty: st.warning("⚠️ Tabela de colaboradores vazia.")
        elif df_horasponto.empty: st.warning(f"⚠️ Sem dados em `horasponto` para {month_name}/{year}.")
        elif horas_uteis_mensais == 0: st.warning(f"⚠️ Sem dados em `dias_uteis_mensais` para {month_name}/{year}.")
        else: render_horas_colaborador_tab(df_colaboradores, df_horasponto, horas_uteis_mensais, all_workitems_filtrados, project_of_item_filtrado, year, month_name, valid_paths)
    if tipo_usuario in ['Admin', 'Gestor', 'Auditor']:
        with tab_faturamento:
            if df_colaboradores.empty: st.warning("⚠️ Tabela de colaboradores vazia.")
            elif df_horasponto.empty: st.warning(f"⚠️ Sem dados em `horasponto` para {month_name}/{year}.")
            elif df_valores.empty: st.warning("⚠️ Tabela de valores vazia.")
            else: render_faturamento_tab(df_colaboradores, df_horasponto, df_valores, horas_uteis_mensais, year, month_name, month_num, pode_calcular=(tipo_usuario in ['Admin', 'Gestor']))
    with tab_metricas:
        render_metricas_tab(all_workitems_filtrados, project_of_item_filtrado, df_colaboradores, df_horasponto, df_valores, horas_uteis_mensais, year, month_name, month_num, selected_wi_types, selected_states, parents_data, tipo_usuario)
    if tipo_usuario != 'Auditor':
        with tab_ponto: render_ponto_tab(year, month_name, month_num, tipo_usuario)
    if tipo_usuario in ['Admin', 'Gestor']:
        with tab_estruturacao: render_estruturacao_tab()
    if tipo_usuario in ['Admin', 'Gestor']:
        with tab_crud_colab: render_colaboradores_tab(df_colaboradores_filtrados=df_colaboradores)

if __name__ == "__main__":
    main()
