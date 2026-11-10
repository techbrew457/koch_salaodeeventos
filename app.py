import base64
from datetime import datetime, timedelta
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
ESTILOS_CHOPP = ["Pilsen", "IPA", "APA", "Bohemian", "Golden", "Stout", "Vinho", "Red Ale", "Para digitar"]
TIPOS_EVENTO = ["aniversário", "infantil", "casamento", "colação", "formatura", "churrasco", "empresa", "confraternização", "maternidade", "Outro"]

try:
    GITHUB_TOKEN = st.secrets["GITHUB_TOKEN"]
    GITHUB_REPO = st.secrets["GITHUB_REPO"]
    HEADERS = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
except Exception:
    st.error("⚠️ Erro crítico: Credenciais do GitHub não configuradas nos Secrets.")
    st.stop()


# Busca dados na API do GitHub utilizando cache para evitar lentidão
@st.cache_data(ttl=60)  # Reduzido para 1 minuto para evitar travar dados vazios
def buscar_dados_github_api():
    url = f"https://github.com{GITHUB_REPO}/contents/{ARQUIVO_DADOS}"
    try:
        res = requests.get(url, headers=HEADERS)
        if res.status_code == 200:
            resp_json = res.json()
            dados = json.loads(base64.b64decode(resp_json["content"]).decode("utf-8"))
            return dados, resp_json["sha"]
    except Exception:
        pass
    return [], None


def carregar_dados_nuvem():
    if "reservas" not in st.session_state:
        dados, sha = buscar_dados_github_api()
        for i, item in enumerate(dados):
            if not item.get("id"):
                item["id"] = f"res-{i}-{str(item.get('cliente', 'sem')).replace(' ', '')}"
        st.session_state["reservas"] = dados
        st.session_state["github_sha"] = sha
    return st.session_state["reservas"]


def salvar_dados_nuvem(novas_reservas):
    url = f"https://github.com{GITHUB_REPO}/contents/{ARQUIVO_DADOS}"
    sha = st.session_state.get("github_sha", None)
    
    conteudo_b64 = base64.b64encode(json.dumps(novas_reservas, ensure_ascii=False, indent=4).encode("utf-8")).decode("utf-8")
    payload = {"message": "Sincronização via app - Dados updated", "content": conteudo_b64}
    if sha:
        payload["sha"] = sha
        
    res = requests.put(url, headers=HEADERS, json=payload)
    if res.status_code == 200 or res.status_code == 201:
        st.session_state["github_sha"] = res.json()["content"]["sha"]
        st.session_state["reservas"] = novas_reservas
        st.cache_data.clear()  # Limpa o cache para forçar atualização das telas
        return True
    else:
        st.error(f"Erro ao salvar no GitHub: {res.status_code}")
        return False


def formatar_data_br(data_str):
    try:
        if not data_str:
            return ""
        if "-" in data_str:
            p = data_str.split("-")
            return f"{p[2]}/{p[1]}/{p[0]}" if len(p) == 3 else data_str
        return data_str
    except Exception:
        return data_str


def calcular_totais(ev):
    # Salão
    v_base = float(ev.get("salao_val", 1200.0))
    v_desc = float(ev.get("salao_desc", 0.0))
    v_salao_liq = max(0.0, v_base - v_desc)
    
    h_not_qtd = float(ev.get("h_not_qtd", 0.0))
    h_not_val_unit = float(ev.get("h_not_val_unit", 200.0))
    h_not_desc = float(ev.get("h_not_desc", 0.0))
    v_not_liq = max(0.0, (h_not_qtd * h_not_val_unit) - h_not_desc)
    
    total_salao = v_salao_liq + v_not_liq

    # Staff
    mon_qtd = float(ev.get("mon_qtd", 0.0))
    mon_val_unit = float(ev.get("mon_val_unit", 200.0))
    mon_desc = float(ev.get("mon_desc", 0.0))
    v_mon = max(0.0, mon_qtd * (mon_val_unit - mon_desc))
    
    v_chopp_cont = 0.0
    for b in ev.get("chopp_contratado", []):
        qtd = float(b.get("qtd", 0))
        preco_base = float(b.get("preco_base", 330.0))
        desc = float(b.get("desconto", 0.0))
        v_chopp_cont += max(0.0, qtd * (preco_base - desc))

    v_chopp_cons = 0.0
    for b in ev.get("chopp_consignado", []):
        qtd = float(b.get("qtd", 0))
        preco_base = float(b.get("preco_base", 330.0))
        desc = float(b.get("desconto", 0.0))
        v_chopp_cons += max(0.0, qtd * (preco_base - desc))

    cop_qtd = float(ev.get("copos_qtd", 0.0))
    v_copos = max(0.0, cop_qtd * (20.0 - float(ev.get("copos_desc", 0.0))))

    som_qtd = float(ev.get("som_qtd", 0.0))
    v_som = max(0.0, som_qtd * (100.0 - float(ev.get("som_desc", 0.0))))
    ret_qtd = float(ev.get("ret_qtd", 0.0))
    v_ret = max(0.0, ret_qtd * (100.0 - float(ev.get("ret_desc", 0.0))))

    v_beb = 0.0
    for beb_key, preco_unit in [("agua_sem", 4.0), ("agua_com", 4.0), ("refrilata", 10.0), ("sucolata", 10.0)]:
        b_info = ev.get(beb_key, {"qtd": 0, "desc": 0.0})
        bq = float(b_info.get("qtd", 0.0))
        bd = float(b_info.get("desc", 0.0))
        v_beb += max(0.0, bq * (preco_unit - bd))

    v_ext = 0.0
    for ext in ev.get("extras", []):
        eq = float(ext.get("qtd", 0.0))
        ep = float(ext.get("valor", 0.0))
        ed = float(ext.get("desconto", 0.0))
        v_ext += max(0.0, eq * (ep - ed))

    subtotal = total_salao + v_mon + v_chopp_cont + v_chopp_cons + v_copos + v_som + v_ret + v_beb + v_ext
    
    total_sinal = sum(float(p.get("valor", 0)) for p in ev.get("pagamentos_sinal", []))
    total_pendente_pago = sum(float(p.get("valor", 0)) for p in ev.get("pagamentos_pendente", []))
    
    val_caucao_estipulado = float(ev.get("valor_caucao_estipulado", 500.0))
    val_sinal_estipulado = float(ev.get("valor_sinal_estipulado", total_salao * 0.5))
    
    restante_a_pagar = max(0.0, subtotal - total_sinal - total_pendente_pago)
    
    return subtotal, val_sinal_estipulado, total_sinal, val_caucao_estipulado, restante_a_pagar


def gerar_documento_word(ev, data_fmt):
    if not os.path.exists("modelo_contrato.docx"):
        return None
    doc = Document("modelo_contrato.docx")
    subtotal, v_sinal, t_sinal, v_caucao, restante = calcular_totais(ev)

    chopps_txt = []
    for b in ev.get("chopp_contratado", []):
        if float(b.get("qtd", 0)) > 0:
            chopps_txt.append(f"- {b['qtd']} barril(is) de {b['litros']}L de Chopp {b['estilo']}")
    for b in ev.get("chopp_consignado", []):
        if float(b.get("qtd", 0)) > 0:
            chopps_txt.append(f"- Consignado: {b['qtd']} barril(is) de {b['litros']}L de Chopp {b['estilo']}")
    txt_chopps = "\n".join(chopps_txt) if chopps_txt else "Nenhum chopp extra."

    opc_txt = []
    if float(ev.get("som_qtd", 0)) > 0: opc_txt.append("- Sistema de Som contratado")
    if float(ev.get("ret_qtd", 0)) > 0: opc_txt.append("- Retroprojetor contratado")
    if float(ev.get("copos_qtd", 0)) > 0: opc_txt.append(f"- {ev.get('copos_qtd')} fardos de copos personalizados")
    txt_opc = "\n".join(opc_txt) if opc_txt else "Nenhum opcional."

    v_base_c = float(ev.get("salao_val", 1200.0))
    v_desc_c = float(ev.get("salao_desc", 0.0))
    v_salao_c = max(0.0, v_base_c - v_desc_c)
    h_not_qtd_c = float(ev.get("h_not_qtd", 0.0))
    h_not_vu_c = float(ev.get("h_not_val_unit", 200.0))
    h_not_desc_c = float(ev.get("h_not_desc", 0.0))
    v_not_c = max(0.0, (h_not_qtd_c * h_not_vu_c) - h_not_desc_c)
    total_salao_c = v_salao_c + v_not_c

    ctx = {
        "{{cliente}}": str(ev.get("cliente", "")),
        "{{documento}}": str(ev.get("documento", "Não informado")),
        "{{telefone}}": str(ev.get("whatsapp", "")),
        "{{data}}": str(data_fmt),
        "{{turno}}": f"{ev.get('turno', '')} (Início: {ev.get('hora_inicio', '10:00')})",
        "{{qtd_pessoas}}": str(ev.get("qtd_pessoas", 50)),
        "{{tipo_evento}}": str(ev.get("tipo_evento", "aniversário")),
        "{{valor_locacao}}": f"R$ {total_salao_c:,.2f}",
        "{{valor_desconto}}": f"R$ {v_desc_c:,.2f}",
        "{{valor_sinal}}": f"R$ {v_sinal:,.2f}",
        "{{status_sinal}}": "Pago" if t_sinal >= v_sinal else "Pendente",
        "{{valor_caucao}}": f"R$ {v_caucao:,.2f}",
        "{{valor_total}}": f"R$ {subtotal:,.2f}",
        "{{restante_pagar}}": f"R$ {restante:,.2f}",
        "{{opcionais}}": txt_opc,
        "{{chopps}}": txt_chopps,
        "{{observacoes}}": str(ev.get("observacoes", "Nenhuma."))
    }

    for p in doc.paragraphs:
        for k, v in ctx.items():
            if k in p.text:
                p.text = p.text.replace(k, v)

    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for k, v in ctx.items():
                        if k in p.text: 
                            p.text = p.text.replace(k, v)

    io_file = io.BytesIO()
    doc.save(io_file)
    io_file.seek(0)
    return io_file


# Carrega a base de dados
reservas = carregar_dados_nuvem()

# --- MENU LATERAL ---
with st.sidebar:
    st.title("🍺 Koch Gestão")
    st.caption("Controle Profissional - v3.3")
    st.divider()
    menu = st.radio("Navegação", ["📊 Dashboard & Métricas", "📅 Calendário Geral", "➕ Nova Locação", "📋 Lista & Gestão"])
    st.divider()
    filtro_mes = st.selectbox("Mês de Referência", ["Todos", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"])

eventos_filtrados = []
for e in reservas:
