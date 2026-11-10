Aqui está o código completo do seu sistema totalmente otimizado, estruturado e pronto para uso. [1]

🛠️ O que foi otimizado e corrigido?

1. Fim do Loop Infinito no GitHub: O código antigo chamava carregar_dados_nuvem() dentro de salvar_dados_nuvem() para buscar o sha. Isso criava requisições redundantes à API e estourava o limite do GitHub. Agora, o sha é armazenado e atualizado dinamicamente no st.session_state. [1]
2. Correção no Cálculo Financeiro: A função calcular_totais subtraía erroneamente o val_sinal_estipulado em vez do sinal efetivamente pago (total_sinal) ao calcular o restante_a_pagar. Isso gerava valores errados nas telas de métricas e no contrato. Corrigido! [1]
3. Filtro de Data Robusto: O filtro por mês agora trata de forma segura valores nulos ou vazios no campo de datas. [1]
4. Performance de Renderização: O script foi limpo removendo redundâncias em inputs de formulários e loops duplicados. [1]
python
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
    url = f"https://github.com{GITHUB_REPO}/contents/{ARQUIVO_DADOS}"
    try:
        res = requests.get(url, headers=HEADERS)
        if res.status_code == 200:
            resp_json = res.json()
            dados = json.loads(base64.b64decode(resp_json["content"]).decode("utf-8"))
            for i, item in enumerate(dados):
                if not item.get("id"):
                    item["id"] = f"res-{i}-{str(item.get('cliente', 'sem')).replace(' ', '')}"
            # Armazena o SHA em cache de sessão para evitar requisições GET antes do PUT
            st.session_state["github_sha"] = resp_json["sha"]
            return dados
    except Exception:
        pass
    return []


def salvar_dados_nuvem(reservas):
    url = f"https://github.com{GITHUB_REPO}/contents/{ARQUIVO_DADOS}"
    sha = st.session_state.get("github_sha", None)
    
    conteudo_b64 = base64.b64encode(json.dumps(reservas, ensure_ascii=False, indent=4).encode("utf-8")).decode("utf-8")
    payload = {"message": "Sincronização via app - Dados atualizados", "content": conteudo_b64}
    if sha:
        payload["sha"] = sha
        
    res = requests.put(url, headers=HEADERS, json=payload)
    if res.status_code in:
        # Armazena o novo SHA retornado pela API do GitHub
        st.session_state["github_sha"] = res.json()["content"]["sha"]
        return True
    else:
        st.error(f"Erro ao salvar no GitHub: {res.status_code}")
        return False


def formatar_data_br(data_str):
    try:
        if not data_str:
            return ""
        p = data_str.split("-")
        return f"{p[2]}/{p[1]}/{p[0]}" if len(p) == 3 else data_str
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

    # Staff / Monitores
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
    total_caucao = sum(float(p.get("valor", 0)) for p in ev.get("pagamentos_caucao", []))
    total_pendente_pago = sum(float(p.get("valor", 0)) for p in ev.get("pagamentos_pendente", []))
    
    val_caucao_estipulado = float(ev.get("valor_caucao_estipulado", 500.0))
    val_sinal_estipulado = float(ev.get("valor_sinal_estipulado", total_salao * 0.5))
    
    # AJUSTE FINANCEIRO: Restante a pagar deduz o SINAL PAGO real e não o estipulado estático
    restante_a_pagar = max(0.0, subtotal - total_sinal - total_pendente_pago)
    
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
                for k, v in ctx.items():
                    if k in cell.text: 
                        cell.text = cell.text.replace(k, v)

    io_file = io.BytesIO()
    doc.save(io_file)
    io_file.seek(0)
    return io_file


# Inicialização única do Estado da Sessão (Evita loops e travamentos de carregamento)
if "reservas" not in st.session_state:
    st.session_state["reservas"] = carregar_dados_nuvem()

reservas = st.session_state["reservas"]

# --- MENU LATERAL ---
with st.sidebar:
    st.title("🍺 Koch Gestão")
    st.caption("Controle Profissional - v2.1")
    st.divider()
    menu = st.radio("Navegação", ["📊 Dashboard & Métricas", "📅 Calendário Geral", "➕ Nova Locação", "📋 Lista & Gestão"])
    st.divider()
    filtro_mes = st.selectbox("Mês de Referência", ["Todos", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"])

# Tratamento seguro contra datas nulas ou corrompidas no filtro por mês
eventos_filtrados = []
for e in reservas:
    data_str = e.get("date", "")
Use o código com cuidado.
if filtro_mes == "Todos":
eventos_filtrados.append(e)
elif data_str and len(data_str.split("-")) >= 2 and data_str.split("-")[1] == filtro_mes:
eventos_filtrados.append(e)

==========================================


1. DASHBOARD & MÉTRICAS


==========================================

if menu == "📊 Dashboard & Métricas":
st.title("📊 Painel de Controle Consolidado")
ativos = [e for e in eventos_filtrados if e.get("status") != "Cancelado"]
total_loc = len(eventos_filtrados)
conf = sum(1 for e in eventos_filtrados if e.get("status") == "Confirmado")
pend = sum(1 for e in eventos_filtrados if e.get("status") == "Pendente")
fat_proj = sum(calcular_totais(e)[0] for e in ativos)
sinais_rec = sum(sum(float(p.get("valor", 0)) for p in e.get("pagamentos_sinal", [])) for e in ativos)
pendente_receber = sum(calcular_totais(e)[5] for e in ativos)
c1, c2, c3, c4 = st.columns(4)
c1.metric("Reservas Totais", total_loc)
c2.metric("Confirmados 🟢", conf)
c3.metric("Em Negociação 🟡", pend)
c4.metric("Faturamento Estimado", f"R$ {fat_proj:,.2f}")
st.divider()
c5, c6 = st.columns(2)
c5.metric("Sinais Recebidos 💵", f"R$ {sinais_rec:,.2f}")
c6.metric("Total Pendente de Recebimento 🔴", f"R$ {pendente_receber:,.2f}")

==========================================


2. CALENDÁRIO GERAL


==========================================

elif menu == "📅 Calendário Geral":
st.title("📆 Calendário Interativo")
cal_evs = []
for ev in eventos_filtrados:
st_ev = ev.get("status", "Pendente")
cor = "#28a745" if st_ev == "Confirmado" else ("#dc3545" if st_ev == "Cancelado" else "#ffc107")
icone = "✅" if st_ev == "Confirmado" else ("❌" if st_ev == "Cancelado" else "⏳")
cal_evs.append({
"title": f"{icone} {ev.get('cliente')} ({ev.get('tipo_evento')})",
"start": ev.get("date"),
"allDay": True,
"backgroundColor": cor,
"borderColor": cor
})
calendar(events=cal_evs, options={"initialView": "dayGridMonth", "locale": "pt-br"})

==========================================


3. FORMULÁRIO COMPARTILHADO


==========================================

def render_formulario(ev_atual=None):
is_edit = ev_atual is not None
s = ev_atual or {}
suf = s.get('id', 'new')
st.subheader("👤 1. Dados do Cliente")
col1, col2 = st.columns(2)
with col1:
cliente = st.text_input("Nome do Cliente *", value=s.get("cliente", ""), key=f"cli_{suf}")
whatsapp = st.text_input("WhatsApp", value=s.get("whatsapp", ""), key=f"wpp_{suf}")
telefone2 = st.text_input("Telefone 02", value=s.get("telefone2", ""), key=f"tel2_{suf}")
with col2:
documento = st.text_input("CPF / CNPJ", value=s.get("documento", ""), key=f"doc_{suf}")
endereco = st.text_input("Endereço", value=s.get("endereco", ""), key=f"end_{suf}")
st.divider()
st.subheader("🏛️ 2. Detalhamento do Salão e do Evento")
sc_e1, sc_e2 = st.columns(2)
with sc_e1:
tipo_evento = st.selectbox("Tipo de Evento", TIPOS_EVENTO, index=TIPOS_EVENTO.index(s.get("tipo_evento")) if s.get("tipo_evento") in TIPOS_EVENTO else 0, key=f"tp_ev_{suf}")
dt_default = datetime.strptime(s.get("date", "2026-01-01"), "%Y-%m-%d") if is_edit and s.get("date") else datetime.now()
data_evento_obj = st.date_input("Data do Evento", value=dt_default, format="DD/MM/YYYY", key=f"dt_ev_{suf}")
with sc_e2:
turnos = ["manhã", "tarde", "noite", "madrugada"]
turno = st.selectbox("Turno", turnos, index=turnos.index(s.get("turno")) if s.get("turno") in turnos else 2, key=f"turno_{suf}")
lista_horas = [f"{i:02d}" for i in range(24)]
lista_minutos = ["00", "15", "30", "45"]
h_init, m_init = s.get("hora_inicio", "10:00").split(":") if ":" in s.get("hora_inicio", "") else ("10", "00")
col_h1, col_h2 = st.columns(2)
with col_h1:
hora_sel = st.selectbox("Hora Início", lista_horas, index=lista_horas.index(h_init) if h_init in lista_horas else 10, key=f"hr_sel_{suf}")
with col_h2:
min_sel = st.selectbox("Minuto Início", lista_minutos, index=lista_minutos.index(m_init) if m_init in lista_minutos else 0, key=f"min_sel_{suf}")
hora_inicio = f"{hora_sel}:{min_sel}"
qtd_pessoas = st.number_input("Quantidade de Pessoas", min_value=1, value=int(s.get("qtd_pessoas", 50)), key=f"qtd_pes_{suf}")
st.markdown("---")
sc1, sc2, sc2_val, sc3 = st.columns(4)
with sc1:
salao_val = st.number_input("Valor Locação Base (R$)", value=float(s.get("salao_val", 1200.0)), step=50.0, key=f"sal_v_{suf}")
salao_desc = st.number_input("Desconto Salão (R$)", value=float(s.get("salao_desc", 0.0)), step=10.0, key=f"sal_d_{suf}")
with sc2:
h_not_qtd = st.number_input("Qtd Horas Adicional Noturno", min_value=0, value=int(s.get("h_not_qtd", 0)), step=1, key=f"hnot_q_{suf}")
h_not_desc = st.number_input("Desconto Adicional Noturno (R$)", value=float(s.get("h_not_desc", 0.0)), step=10.0, key=f"hnot_d_{suf}")
with sc2_val:
h_not_val_unit = st.number_input("Valor Hora Adicional (R$)", value=float(s.get("h_not_val_unit", 200.0)), step=10.0, key=f"hnot_vu_{suf}")
with sc3:
higiene = st.selectbox("Produtos de Higiene e Limpeza", ["Tech", "Cliente"], index=0 if s.get("higiene") == "Tech" else 1, key=f"hig_{suf}")
val_salao_liq = max(0.0, salao_val - salao_desc)
val_not_liq = max(0.0, (float(h_not_qtd) * h_not_val_unit) - h_not_desc)
st.metric(label="🏷️ Valor Final do Salão calculado", value=f"R$ {(val_salao_liq + val_not_liq):,.2f}")
st.markdown("---")
vc1, vc2 = st.columns(2)
with vc1:
visita_previa = st.selectbox("Visita Prévia?", ["Não", "Sim"], index=0 if s.get("visita_previa") != "Sim" else 1, key=f"vprev_{suf}")
dt_vp_default = datetime.strptime(s.get("data_visita_previa", "2026-01-01"), "%Y-%m-%d") if s.get("data_visita_previa") else datetime.now()
data_visita_previa_obj = st.date_input("Data Visita Prévia", value=dt_vp_default, format="DD/MM/YYYY", key=f"dt_vprev_{suf}")
vp_h, vp_m = s.get("hora_visita_previa", "14:00").split(":") if ":" in s.get("hora_visita_previa", "") else ("14", "00")
vph_col, vpm_col = st.columns(2)
with vph_col: vp_h_sel = st.selectbox("Hora (Visita)", lista_horas, index=lista_horas.index(vp_h) if vp_h in lista_horas else 14, key=f"vph_sel_{suf}")
with vpm_col: vp_m_sel = st.selectbox("Minuto (Visita)", lista_minutos, index=lista_minutos.index(vp_m) if vp_m in lista_minutos else 0, key=f"vpm_sel_{suf}")
hora_visita_previa = f"{vp_h_sel}:{vp_m_sel}"
with vc2:
visita_prep = st.selectbox("Visita para Preparativos?", ["Não", "Sim"], index=0 if s.get("visita_prep") != "Sim" else 1, key=f"vprep_p_{suf}")
dt_vpr_default = datetime.strptime(s.get("data_visita_prep", "2026-01-01"), "%Y-%m-%d") if s.get("data_visita_prep") else datetime.now()
data_visita_prep_obj = st.date_input("Data Prep", value=dt_vpr_default, format="DD/MM/YYYY", key=f"dt_vprep_p_{suf}")
prep_h, prep_m = s.get("hora_visita_prep", "09:00").split(":") if ":" in s.get("hora_visita_prep", "") else ("09", "00")
preph_col, prepm_col = st.columns(2)
with preph_col: prep_h_sel = st.selectbox("Hora (Prep)", lista_horas, index=lista_horas.index(prep_h) if prep_h in lista_horas else 9, key=f"preph_sel_{suf}")
with prepm_col: prep_m_sel = st.selectbox("Minuto (Prep)", lista_minutos, index=lista_minutos.index(prep_m) if prep_m in lista_minutos else 0, key=f"prepm_sel_{suf}")
hora_visita_prep = f"{prep_h_sel}:{prep_m_sel}"
st.divider()
st.subheader("👥 3. Detalhamento de Funcionários (Staff)")
st1, st2, st3 = st.columns(3)
with st1: mon_qtd = st.number_input("Qtd de Monitores", min_value=0.0, value=float(s.get("mon_qtd", 0.0)), step=1.0, key=f"mon_q_{suf}")
with st2: mon_val_unit = st.number_input("Valor Unitário Staff", value=float(s.get("mon_val_unit", 200.0)), step=10.0, key=f"mon_vu_{suf}")
with st3: mon_desc = st.number_input("Desconto por Staff", value=float(s.get("mon_desc", 0.0)), step=10.0, key=f"mon_d_{suf}")
st.divider()
st.subheader("🍺 4. Detalhamento de Chopp")
st.markdown("##### 📌 Chopp Contratado")
chopp_cont_data = s.get("chopp_contratado", [])
novo_chopp_cont = []
estilos_principais = ["Pilsen", "IPA", "APA", "Bohemian"]
cc_cols = st.columns(4)
for idx, estilo in enumerate(estilos_principais):
b_exist = next((x for x in chopp_cont_data if x.get("estilo") == estilo), {"qtd": 0, "litros": 30, "desconto": 0.0})
with cc_cols[idx]:
st.markdown(f"{estilo}")
q = st.number_input("Qtd Barris", min_value=0, value=int(b_exist.get("qtd", 0)), key=f"cc_q_{estilo}{suf}")
lit = st.selectbox("Litros", [30, 50], index=0 if b_exist.get("litros", 30) == 30 else 1, key=f"cc_l{estilo}{suf}")
desc = st.number_input("Desc. R$", value=float(b_exist.get("desconto", 0.0)), key=f"cc_d{estilo}_{suf}")
if q > 0:
novo_chopp_cont.append({"estilo": estilo, "qtd": q, "litros": lit, "preco_base": 330.0 if estilo == "Pilsen" else 360.0, "desconto": desc})
st.markdown("##### 📌 Chopp Consignado")
chopp_cons_data = s.get("chopp_consignado", [])
novo_chopp_cons = []
ccon_cols = st.columns(4)
for idx, estilo in enumerate(estilos_principais):
b_exist = next((x for x in chopp_cons_data if x.get("estilo") == estilo), {"qtd": 0, "litros": 30, "desconto": 0.0})
with ccon_cols[idx]:
st.markdown(f"Consig. {estilo}")
q = st.number_input("Qtd Consig.", min_value=0, value=int(b_exist.get("qtd", 0)), key=f"con_q_{estilo}{suf}")
lit = st.selectbox("Litros Consig.", [30, 50], index=0 if b_exist.get("litros", 30) == 30 else 1, key=f"con_l{estilo}{suf}")
desc = st.number_input("Desc. Consig. R$", value=float(b_exist.get("desconto", 0.0)), key=f"con_d{estilo}_{suf}")
if q > 0:
novo_chopp_cons.append({"estilo": estilo, "qtd": q, "litros": lit, "preco_base": 330.0 if estilo == "Pilsen" else 360.0, "desconto": desc})
st.divider()
st.subheader("🥤 5. Opcionais e Bebidas")
cop1, cop2 = st.columns(2)
with cop1:
copos_qtd = st.number_input("Fardos de Copos (R$ 20/fardo)", min_value=0.0, value=float(s.get("copos_qtd", 0.0)), key=f"cop_q_{suf}")
copos_desc = st.number_input("Desconto fardo (R$)", value=float(s.get("copos_desc", 0.0)), key=f"cop_d_{suf}")
with cop2:
som_qtd = st.number_input("Som (1 = Sim, 0 = Não)", min_value=0.0, max_value=1.0, value=float(s.get("som_qtd", 0.0)), key=f"som_q_{suf}")
som_desc = st.number_input("Desconto Som", value=float(s.get("som_desc", 0.0)), key=f"som_d_{suf}")
ret_qtd = st.number_input("Retroprojetor (1 = Sim, 0 = Não)", min_value=0.0, max_value=1.0, value=float(s.get("ret_qtd", 0.0)), key=f"ret_q_{suf}")
ret_desc = st.number_input("Desconto Retro", value=float(s.get("ret_desc", 0.0)), key=f"ret_d_{suf}")
st.markdown("##### 💧 Bebidas em Geral")
beb_cols = st.columns(4)
bebidas_inputs = {}
beb_config = [("agua_sem", "Água Sem Gás", 4.0), ("agua_com", "Água Com Gás", 4.0), ("refrilata", "Refri Lata", 10.0), ("sucolata", "Suco Lata", 10.0)]
for i, (b_key, b_label, def_p) in enumerate(beb_config):
b_val = s.get(b_key, {"text": "0", "qtd": 0, "desc": 0.0})
with beb_cols[i]:
st.markdown(f"{b_label}")
bq = st.number_input("Quantidade", min_value=0, value=int(b_val.get("qtd", 0)), key=f"bq_{b_key}{suf}")
bd = st.number_input("Desconto R$", value=float(b_val.get("desc", 0.0)), key=f"bd{b_key}_{suf}")
bebidas_inputs[b_key] = {"qtd": bq, "desc": bd}
st.divider()
st.subheader("💰 6. Financeiro e Pagamentos")
fc1, fc2 = st.columns(2)
with fc1: val_sinal_estipulado = st.number_input("Sinal Estipulado (R$)", value=float(s.get("valor_sinal_estipulado", 600.0)), key=f"vsinal_est_{suf}")
with fc2: val_caucao_estipulado = st.number_input("Caução Estipulado (R$)", value=float(s.get("valor_caucao_estipulado", 500.0)), key=f"vcaucao_est_{suf}")
def render_bloco_pagamentos(titulo, chave_lista):
st.markdown(f"{titulo}")
pagamentos_atuais = s.get(chave_lista, [])
novos_pags = []
for i in range(3):
p_ex = pagamentos_atuais[i] if i < len(pagamentos_atuais) else {"valor": 0.0, "data": "", "metodo": "PIX"}
try:
dt_p_default = datetime.strptime(p_ex.get("data", ""), "%Y-%m-%d").date()
except Exception:
dt_p_default = datetime.now().date()
pc1, pc2, pc3 = st.columns(3)
with pc1: v = st.number_input("Valor (R$)", value=float(p_ex.get("valor", 0.0)), key=f"{chave_lista}v{i}{suf}")
with pc2: dt_obj = st.date_input("Data", value=dt_p_default, format="DD/MM/YYYY", key=f"{chave_lista}dt{i}{suf}")
with pc3: met = st.selectbox("Método", ["PIX", "débito", "crédito", "dinheiro"], index=["PIX", "débito", "crédito", "dinheiro"].index(p_ex.get("metodo", "PIX")) if p_ex.get("metodo") in ["PIX", "débito", "crédito", "dinheiro"] else 0, key=f"{chave_lista}m{i}_{suf}")
if v > 0:
novos_pags.append({"valor": v, "data": dt_obj.strftime("%Y-%m-%d"), "metodo": met})
return novos_pags
pag_sinal = render_bloco_pagamentos("Sinal", "pagamentos_sinal")
pag_caucao = render_bloco_pagamentos("Caução", "pagamentos_caucao")
pag_pendente = render_bloco_pagamentos("Saldo Final", "pagamentos_pendente")
st.divider()
status = st.selectbox("Status Geral", ["Pendente", "Confirmado", "Cancelado"], index=["Pendente", "Confirmado", "Cancelado"].index(s.get("status", "Pendente")), key=f"status_{suf}")
observacoes = st.text_area("Observações", value=s.get("observacoes", ""), key=f"obs_{suf}")
return {
"cliente": cliente, "whatsapp": whatsapp, "telefone2": telefone2, "documento": documento, "endereco": endereco,
"tipo_evento": tipo_evento, "date": data_evento_obj.strftime("%Y-%m-%d"), "turno": turno, "hora_inicio": hora_inicio,
"qtd_pessoas": int(qtd_pessoas), "salao_val": float(salao_val), "salao_desc": float(salao_desc),
"h_not_qtd": float(h_not_qtd), "h_not_desc": float(h_not_desc), "h_not_val_unit": float(h_not_val_unit),
"higiene": higiene, "visita_previa": visita_previa, "data_visita_previa": data_visita_previa_obj.strftime("%Y-%m-%d"), "hora_visita_previa": hora_visita_previa,
"visita_prep": visita_prep, "data_visita_prep": data_visita_prep_obj.strftime("%Y-%m-%d"), "hora_visita_prep": hora_visita_prep,
"mon_qtd": float(mon_qtd), "mon_val_unit": float(mon_val_unit), "mon_desc": float(mon_desc),
"chopp_contratado": novo_chopp_cont, "chopp_consignado": novo_chopp_cons,
"copos_qtd": float(copos_qtd), "copos_desc": float(copos_desc), "som_qtd": float(som_qtd), "som_desc": float(som_desc), "ret_qtd": float(ret_qtd), "ret_desc": float(ret_desc),
"agua_sem": bebidas_inputs["agua_sem"], "agua_com": bebidas_inputs["agua_com"], "refrilata": bebidas_inputs["refrilata"], "sucolata": bebidas_inputs["sucolata"],
"extras": [], "valor_sinal_estipulado": float(val_sinal_estipulado), "valor_caucao_estipulado": float(val_caucao_estipulado),
"pagamentos_sinal": pag_sinal, "pagamentos_caucao": pag_caucao, "pagamentos_pendente": pag_pendente, "status": status, "observacoes": observacoes
}
if menu == "➕ Nova Locação":
st.title("📝 Registrar Nova Locação")
with st.form("form_nova_locacao", clear_on_submit=True):
dados_form = render_formulario()
if st.form_submit_button("💾 Salvar Locação na Nuvem"):
if not dados_form["cliente"].strip():
st.warning("⚠️ O nome do cliente é obrigatório.")
else:
dados_form["id"] = str(int(datetime.now().timestamp()))
reservas.append(dados_form)
if salvar_dados_nuvem(reservas):
st.success("🎉 Locação registrada com sucesso!")
st.rerun()

==========================================


4. LISTA & GESTÃO


==========================================

elif menu == "📋 Lista & Gestão":
st.title("📋 Gerenciamento de Locações")
if eventos_filtrados:
c_f1, c_f2 = st.columns(2)
with c_f1: f_status = st.selectbox("Filtrar por Status", ["Todos", "Confirmado", "Pendente", "Cancelado"])
with c_f2: f_busca = st.text_input("🔍 Buscar por Cliente", "")
lista = [e for e in eventos_filtrados if (f_status == "Todos" or e.get("status") == f_status) and (f_busca.lower() in e.get("cliente", "").lower())]
hoje = datetime.now().date()
for ev in sorted(lista, key=lambda x: str(x.get("date", ""))):
ev_id = ev.get("id")
st_ev = ev.get("status", "Pendente")
emoji = "✅" if st_ev == "Confirmado" else ("❌" if st_ev == "Cancelado" else "⏳")
dt_fmt = formatar_data_br(ev.get("date"))
subtotal, v_sinal, t_sinal, v_caucao, t_caucao, restante = calcular_totais(ev)
with st.container(border=True):
# Alertas Inteligentes de Operação
if ev.get("date"):
try:
dt_evento = datetime.strptime(ev.get("date"), "%Y-%m-%d").date()
if hoje == (dt_evento - timedelta(days=1)):
st.error("🚨 ALERTA DE VÉSPERA: Checar o salão, evento amanhã!")
if restante > 0 and hoje == (dt_evento - timedelta(days=3)):
st.warning("⚠️ FINANCEIRO PENDENTE: Faltam 3 dias para o evento e restam valores a pagar!")
except Exception: pass
col_i1, col_i2 = st.columns([3, 1])
with col_i1:
st.markdown(f"{emoji} Data: {dt_fmt} - {ev.get('cliente')}")
st.caption(f"Turno: {ev.get('turno')} | Início: {ev.get('hora_inicio')} | Pessoas: {ev.get('qtd_pessoas')} | Status: {st_ev}")
whatsapp = "".join(filter(str.isdigit, ev.get("whatsapp", "")))
if whatsapp: st.markdown(f"💬 Abrir WhatsApp")
st.markdown(f"💰 Total: R$ {subtotal:,.2f} | 💵 Sinal Pago: R$ {t_sinal:,.2f} | 🔴 Falta: R$ {restante:,.2f}")
with col_i2:
btn_edit = st.button("✏️ Editar", key=f"ed_{ev_id}", use_container_width=True)
btn_docx = st.button("📄 Contrato", key=f"ct_{ev_id}", use_container_width=True)
btn_del = st.button("🗑️ Excluir", key=f"dl_{ev_id}", use_container_width=True)
if btn_del:
reservas = [r for r in reservas if r.get("id") != ev_id]
if salvar_dados_nuvem(reservas):
st.success("Excluído!")
st.rerun()
if btn_docx:
arq = gerar_documento_word(ev, dt_fmt)
if arq:
st.download_button("📥 Baixar Contrato Word", arq, file_name=f"Contrato_{ev.get('cliente','').replace(' ','')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dl_doc{ev_id}")
else:
st.warning("⚠️ Modelo de contrato não encontrado localmente.")
if btn_edit:
st.session_state[f"edit_mode_{ev_id}"] = not st.session_state.get(f"edit_mode_{ev_id}", False)
if st.session_state.get(f"edit_mode_{ev_id}", False):
st.divider()
with st.form(f"form_edit_{ev_id}"):
novos_dados = render_formulario(ev)
if st.form_submit_button("💾 Salvar Alterações"):
if novos_dados["cliente"].strip():
novos_dados["id"] = ev_id
reservas = [novos_dados if r.get("id") == ev_id else r for r in reservas]
if salvar_dados_nuvem(reservas):
st.session_state[f"edit_mode_{ev_id}"] = False
st.success("✅ Atualizado com sucesso!")
st.rerun()
else:
st.info("Nenhuma locação cadastrada.")


