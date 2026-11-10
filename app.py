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


def carregar_dados_nuvem():
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{ARQUIVO_DADOS}"
    try:
        res = requests.get(url, headers=HEADERS)
        if res.status_code == 200:
            resp_json = res.json()
            dados = json.loads(base64.b64decode(resp_json["content"]).decode("utf-8"))
            for i, item in enumerate(dados):
                if not item.get("id"):
                    item["id"] = f"res-{i}-{str(item.get('cliente', 'sem')).replace(' ', '')}"
            return dados, resp_json["sha"]
    except Exception:
        pass
    return [], None


def salvar_dados_nuvem(reservas):
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{ARQUIVO_DADOS}"
    _, sha = carregar_dados_nuvem()
    conteudo_b64 = base64.b64encode(json.dumps(reservas, ensure_ascii=False, indent=4).encode("utf-8")).decode("utf-8")
    payload = {"message": "Sincronização via app corrigido", "content": conteudo_b64}
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
    v_salao = max(0.0, float(ev.get("salao_val", 1200.0)) - float(ev.get("salao_desc", 0.0)))
    h_not_qtd = float(ev.get("h_not_qtd", 0.0))
    v_not = max(0.0, (h_not_qtd * 200.0) - float(ev.get("h_not_desc", 0.0)))
    
    mon_qtd = float(ev.get("mon_qtd", 0.0))
    v_mon = max(0.0, mon_qtd * (280.0 - float(ev.get("mon_desc", 0.0))))
    
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

    subtotal = v_salao + v_not + v_mon + v_chopp_cont + v_chopp_cons + v_copos + v_som + v_ret + v_beb + v_ext
    
    total_sinal = sum(float(p.get("valor", 0)) for p in ev.get("pagamentos_sinal", []))
    total_caucao = sum(float(p.get("valor", 0)) for p in ev.get("pagamentos_caucao", []))
    total_pendente_pago = sum(float(p.get("valor", 0)) for p in ev.get("pagamentos_pendente", []))
    
    val_caucao_estipulado = float(ev.get("valor_caucao_estipulado", 500.0))
    val_sinal_estipulado = float(ev.get("valor_sinal_estipulado", v_salao * 0.5))
    
    restante_a_pagar = max(0.0, subtotal - val_sinal_estipulado - total_pendente_pago)
    
    return subtotal, val_sinal_estipulado, total_sinal, val_caucao_estipulado, total_caucao, restante_a_pagar


def gerar_documento_word(ev, data_fmt):
    if not os.path.exists("modelo_contrato.docx"):
        return None
    doc = Document("modelo_contrato.docx")
    subtotal, v_sinal, t_sinal, v_caucao, t_caucao, restante = calcular_totais(ev)

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

    ctx = {
        "{{cliente}}": str(ev.get("cliente", "")),
        "{{documento}}": str(ev.get("documento", "Não informado")),
        "{{telefone}}": str(ev.get("telefone", "")),
        "{{data}}": str(data_fmt),
        "{{turno}}": f"{ev.get('turno', '')} (Início: {ev.get('hora_inicio', '10:00')})",
        "{{qtd_pessoas}}": str(ev.get("qtd_pessoas", 50)),
        "{{tipo_evento}}": str(ev.get("tipo_evento", "aniversário")),
        "{{valor_locacao}}": f"R$ {float(ev.get('salao_val', 1200.0)):,.2f}",
        "{{valor_desconto}}": f"R$ {float(ev.get('salao_desc', 0.0)):,.2f}",
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
    st.caption("Controle Profissional - v2.0")
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
    sinais_rec = sum(sum(float(p.get("valor", 0)) for p in e.get("pagamentos_sinal", [])) for e in ativos)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Reservas", total_loc)
    c2.metric("Confirmados 🟢", conf)
    c3.metric("Em Negociação 🟡", pend)
    c4.metric("Faturamento Líquido", f"R$ {fat_proj:,.2f}")

    st.divider()
    c5, c6 = st.columns(2)
    c5.metric("Sinais Recebidos", f"R$ {sinais_rec:,.2f}")


# ==========================================
# 2. CALENDÁRIO GERAL
# ==========================================
elif menu == "📅 Calendário Geral":
    st.title("📆 Calendário Interativo")
    cal_evs = []
    for ev in eventos_filtrados:
        st_ev = ev.get("status", "Pendente")
        cor, icone = ("#28a745", "✅") if st_ev == "Confirmado" else (("#dc3545", "❌") if st_ev == "Cancelado" else ("#ffc107", "⏳"))
        
        cal_evs.append({
            "title": f"{icone} {formatar_data_br(ev.get('date'))} - {ev.get('cliente')} ({ev.get('tipo_evento')})",
            "start": ev.get("date"),
            "allDay": True,
            "backgroundColor": cor,
            "borderColor": cor
        })
    calendar(events=cal_evs, options={"initialView": "dayGridMonth", "locale": "pt-br"})


# ==========================================
# 3. FORMULÁRIO COMPARTILHADO (NOVA LOCAÇÃO / EDIÇÃO)
# ==========================================
def render_formulario(ev_atual=None):
    is_edit = ev_atual is not None
    s = ev_atual or {}
    suf = s.get('id', 'new')

    st.subheader("👤 1. Dados do Cliente e Evento")
    col1, col2 = st.columns(2)
    with col1:
        cliente = st.text_input("Nome do Cliente *", value=s.get("cliente", ""), key=f"cli_{suf}")
        whatsapp = st.text_input("WhatsApp", value=s.get("whatsapp", ""), key=f"wpp_{suf}")
        telefone2 = st.text_input("Telefone 02", value=s.get("telefone2", ""), key=f"tel2_{suf}")
        documento = st.text_input("CPF / CNPJ", value=s.get("documento", ""), key=f"doc_{suf}")
        endereco = st.text_input("Endereço", value=s.get("endereco", ""), key=f"end_{suf}")
    with col2:
        tipo_evento = st.selectbox("Tipo de Evento", TIPOS_EVENTO, index=TIPOS_EVENTO.index(s.get("tipo_evento")) if s.get("tipo_evento") in TIPOS_EVENTO else 0, key=f"tp_ev_{suf}")
        
        dt_base = datetime.strptime(s.get("date", "2026-01-01"), "%Y-%m-%d") if is_edit else datetime.now()
        d1, d2, d3 = st.columns(3)
        with d1: dia = st.selectbox("Dia", list(range(1, 32)), index=dt_base.day - 1, key=f"dia_{suf}")
        with d2: mes = st.selectbox("Mês", list(range(1, 13)), index=dt_base.month - 1, format_func=lambda x: f"{x:02d}", key=f"mes_{suf}")
        with d3: ano = st.selectbox("Ano", [2025, 2026, 2027, 2028], index=[2025, 2026, 2027, 2028].index(dt_base.year) if dt_base.year in [2025, 2026, 2027, 2028] else 1, key=f"ano_{suf}")

        turnos = ["manhã", "tarde", "noite", "madrugada"]
        turno = st.selectbox("Turno", turnos, index=turnos.index(s.get("turno")) if s.get("turno") in turnos else 2, key=f"turno_{suf}")
        hora_inicio = st.text_input("Horário de Início (ex: 10:00)", value=s.get("hora_inicio", "10:00"), key=f"hr_ini_{suf}")
        qtd_pessoas = st.number_input("Quantidade de Pessoas", min_value=1, value=int(s.get("qtd_pessoas", 50)), key=f"qtd_pes_{suf}")

    st.divider()
    st.subheader("🏛️ 2. Detalhamento do Salão, Adicionais e Visitas")
    sc1, sc2, sc3 = st.columns(3)
    with sc1:
        salao_val = st.number_input("Valor Locação Base (R$)", value=float(s.get("salao_val", 1200.0)), step=50.0, key=f"sal_v_{suf}")
        salao_desc = st.number_input("Desconto Salão (R$)", value=float(s.get("salao_desc", 0.0)), step=10.0, key=f"sal_d_{suf}")
    with sc2:
        h_not_qtd = st.number_input("Horas de Adicional Noturno (Qtd)", min_value=0.0, value=float(s.get("h_not_qtd", 0.0)), step=1.0, key=f"hnot_q_{suf}")
        h_not_desc = st.number_input("Desconto Adicional Noturno (R$)", value=float(s.get("h_not_desc", 0.0)), step=10.0, key=f"hnot_d_{suf}")
    with sc3:
        higiene = st.selectbox("Produtos de Higiene e Limpeza", ["Tech", "Cliente"], index=0 if s.get("higiene") == "Tech" else 1, key=f"hig_{suf}")

    vc1, vc2, vc3, vc4 = st.columns(4)
    with vc1:
        visita_previa = st.selectbox("Visita Prévia?", ["Não", "Sim"], index=0 if s.get("visita_previa") == "Não" else 1, key=f"vprev_{suf}")
    with vc2:
        data_visita_previa = st.text_input("Data Visita Prévia (AAAA-MM-DD)", value=s.get("data_visita_previa", ""), key=f"dt_vprev_{suf}")
    with vc3:
        visita_prep = st.selectbox("Visita para Preparativos?", ["Não", "Sim"], index=0 if s.get("visita_prep") == "Não" else 1, key=f"vprep_p_{suf}")
    with vc4:
        data_visita_prep = st.text_input("Data Prep (AAAA-MM-DD)", value=s.get("data_visita_prep", ""), key=f"dt_vprep_p_{suf}")

    st.divider()
    st.subheader("👥 3. Detalhamento de Funcionários (Staff)")
    st1, st2 = st.columns(2)
    with st1:
        mon_qtd = st.number_input("Qtd de Monitores / Staff", min_value=0.0, value=float(s.get("mon_qtd", 0.0)), step=1.0, key=f"mon_q_{suf}")
    with st2:
        mon_desc = st.
