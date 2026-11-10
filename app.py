import base64
from datetime import datetime
import io
import json
import os
from docx import Document
import requests
import streamlit as st
from streamlit_calendar import calendar

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Gestão Salão de Festas - Koch Cervejaria",
    page_icon="🍺",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown(
    """
    <style>
    .fc { max-width: 100% !important; font-size: 0.85rem !important; }
    .fc-toolbar { flex-direction: column !important; gap: 8px !important; align-items: center !important; }
    .fc-toolbar-title { font-size: 1.1rem !important; font-weight: bold; }
    .fc-daygrid-day-frame { min-height: 75px !important; }
    .fc-event { font-size: 0.72rem !important; padding: 2px 4px !important; border-radius: 4px !important; }
    div.stButton > button { border-radius: 6px; font-weight: 600; }
    </style>
    """,
    unsafe_allow_html=True,
)

ARQUIVO_DADOS = "reservas.json"
ESTILOS_CHOPP = ["Pilsen", "IPA", "Stout", "Bohemian Pilsener", "Premium Golden", "APA", "Vinho", "Red Ale"]

try:
    GITHUB_TOKEN = st.secrets["GITHUB_TOKEN"]
    GITHUB_REPO = st.secrets["GITHUB_REPO"]
    HEADERS = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
except Exception:
    st.error("⚠️ Erro crítico: Credenciais do GitHub não configuradas nos Secrets.")
    st.stop()


def carregar_dados_nuvem():
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{ARQUIVO_DADOS}"
    try:
        res = requests.get(url, headers=HEADERS)
        if res.status_code == 200:
            resp_json = res.json()
            dados = json.loads(base64.b64decode(resp_json["content"]).decode("utf-8"))
            for i, item in enumerate(dados):
                if not item.get("id"):
                    item["id"] = f"res-{i}-{str(item.get('cliente','sem')).replace(' ','')}"
            return dados, resp_json["sha"]
    except Exception:
        pass
    return [], None


def salvar_dados_nuvem(reservas):
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{ARQUIVO_DADOS}"
    _, sha = carregar_dados_nuvem()
    conteudo_b64 = base64.b64encode(json.dumps(reservas, ensure_ascii=False, indent=4).encode("utf-8")).decode("utf-8")
    payload = {"message": "Sincronização via app otimizado com Desconto", "content": conteudo_b64}
    if sha:
        payload["sha"] = sha
    res = requests.put(url, headers=HEADERS, json=payload)
    if res.status_code not in [200, 201]:
        st.error(f"Erro ao salvar no GitHub: {res.status_code}")
        return False
    return True


def formatar_data_br(data_str):
    try:
        p = data_str.split("-")
        return f"{p[2]}/{p[1]}/{p[0]}" if len(p) == 3 else data_str
    except Exception:
        return data_str


def calcular_totais(ev):
    v_loc = float(ev.get("valor_locacao", 0) or 0)
    v_som = float(ev.get("valor_som", 0) or 0) if ev.get("som_sn") == "Sim" else 0
    v_ret = float(ev.get("valor_retroprojetor", 0) or 0) if ev.get("retro_sn") == "Sim" else 0
    v_taxa = float(ev.get("valor_taxa_adicional", 0) or 0)
    v_ref = float(ev.get("refri_val", 0) or 0) if ev.get("refri_sn") == "Sim" else 0
    v_ag = float(ev.get("aguag_val", 0) or 0) if ev.get("aguag_sn") == "Sim" else 0
    v_as = float(ev.get("aguas_val", 0) or 0) if ev.get("aguas_sn") == "Sim" else 0
    v_suc = float(ev.get("suco_val", 0) or 0) if ev.get("suco_sn") == "Sim" else 0
    v_imp = float(ev.get("imp_val", 0) or 0) if ev.get("imp_sn") == "Sim" else 0
    v_chopp = sum(float(b.get("valor_total", 0) or 0) for b in ev.get("barris", []))
    v_desconto = float(ev.get("valor_desconto", 0) or 0)
    
    subtotal = v_loc + v_som + v_ret + v_taxa + v_ref + v_ag + v_as + v_suc + v_imp + v_chopp
    total = subtotal - v_desconto
    sinal = float(ev.get("valor_sinal", 0) or 0)
    return total, sinal, total - sinal


def gerar_documento_word(ev, data_fmt, barris, txt_opc, txt_barris):
    if not os.path.exists("modelo_contrato.docx"):
        return None
    doc = Document("modelo_contrato.docx")
    total, sinal, restante = calcular_totais(ev)
    v_desconto = float(ev.get("valor_desconto", 0) or 0)

    ctx = {
        "{{cliente}}": str(ev.get("cliente", "")),
        "{{documento}}": str(ev.get("documento", "Não informado")),
        "{{telefone}}": str(ev.get("telefone", "")),
        "{{data}}": str(data_fmt),
        "{{turno}}": str(ev.get("turno", "")),
        "{{qtd_pessoas}}": str(ev.get("qtd_pessoas", 0)),
        "{{tipo_evento}}": str(ev.get("tipo_evento", "Festividade")),
        "{{valor_locacao}}": f"R$ {float(ev.get('valor_locacao', 0)):,.2f}",
        "{{valor_desconto}}": f"R$ {v_desconto:,.2f}",
        "{{valor_sinal}}": f"R$ {sinal:,.2f}",
        "{{status_sinal}}": "Pago" if (sinal > 0 or ev.get("sinal_pago")) else "Pendente",
        "{{valor_caucao}}": f"R$ {float(ev.get('valor_caucao', 0)):,.2f}",
        "{{valor_total}}": f"R$ {total:,.2f}",
        "{{restante_pagar}}": f"R$ {restante:,.2f}",
        "{{opcionais}}": txt_opc,
        "{{chopps}}": txt_barris,
        "{{observacoes}}": str(ev.get("observacoes", "Nenhuma."))
    }

    for p in doc.paragraphs:
        for k, v in ctx.items():
            if k in p.text:
                for r in p.runs:
                    if k in r.text: r.text = r.text.replace(k, v)
                p.text = p.text.replace(k, v)

    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for k, v in ctx.items():
                    if k in cell.text: cell.text = cell.text.replace(k, v)

    io_file = io.BytesIO()
    doc.save(io_file)
    io_file.seek(0)
    return io_file


reservas, _ = carregar_dados_nuvem()

# --- MENU LATERAL ---
with st.sidebar:
    st.title("🍺 Koch Gestão")
    st.caption("Controle Profissional")
    st.divider()
    menu = st.radio("Navegação", ["📊 Dashboard & Métricas", "📅 Calendário Geral", "➕ Nova Locação", "📋 Lista & Gestão"])
    st.divider()
    filtro_mes = st.selectbox("Mês de Referência", ["Todos", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"])

eventos_filtrados = reservas if filtro_mes == "Todos" else [e for e in reservas if str(e.get("date", "")).split("-")[1] == filtro_mes]


# ==========================================
# 1. DASHBOARD & MÉTRICAS
# ==========================================
if menu == "📊 Dashboard & Métricas":
    st.title("📊 Painel de Controle Consolidado")
    
    ativos = [e for e in eventos_filtrados if e.get("status") != "Cancelado"]
    total_loc = len(eventos_filtrados)
    conf = sum(1 for e in eventos_filtrados if e.get("status") == "Confirmado")
    pend = sum(1 for e in eventos_filtrados if e.get("status") == "Pendente")
    fat_proj = sum(calcular_totais(e)[0] for e in ativos)
    sinais_rec = sum(float(e.get("valor_sinal", 0) or 0) for e in ativos if float(e.get("valor_sinal", 0) or 0) > 0 or e.get("sinal_pago"))
    litros_tot = sum(int(b.get("litros", 0) or 0) for e in ativos for b in e.get("barris", []))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Reservas", total_loc)
    c2.metric("Confirmados 🟢", conf)
    c3.metric("Em Negociação 🟡", pend)
    c4.metric("Volume Chopp", f"{litros_tot} L")

    st.divider()
    c5, c6 = st.columns(2)
