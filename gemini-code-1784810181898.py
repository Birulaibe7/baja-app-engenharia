import streamlit as st
import pandas as pd
from datetime import datetime
import os
import tempfile
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from fpdf import FPDF

# ==========================================
# CONFIGURAÇÕES DA PÁGINA E BANCO DE DADOS
# ==========================================
st.set_page_config(page_title="Central de Engenharia - Baja", page_icon="⚙️", layout="wide")

EXCEL_FALHAS = "historico_baja.xlsx"
STATUS_LIST = ["🔴 Ocorrência Registrada", "🟡 Em Análise", "🔵 Fila de Usinagem", "🟢 Validado / Fechado"]
COLUNAS = [
    "ID", "Status", "Data_Ocorrencia", "Componente", "Area", "Responsavel", 
    "Contexto", "Piloto", "Atividade", "Sintomas", "Inspecao", "Drive_Link", 
    "Porque1", "Porque2", "Porque3", "Porque4", "Porque5", 
    "Acao", "Modificacoes", "Resultado", "Parecer"
]

def inicializar_bd():
    if not os.path.exists(EXCEL_FALHAS):
        df = pd.DataFrame(columns=COLUNAS)
        df.to_excel(EXCEL_FALHAS, index=False)

inicializar_bd()

def carregar_dados():
    return pd.read_excel(EXCEL_FALHAS)

def salvar_dados(df):
    df.to_excel(EXCEL_FALHAS, index=False)

# ==========================================
# FUNÇÕES DE GERADORES DE RELATÓRIO
# ==========================================
def processar_imagem_temp(uploaded_file):
    if uploaded_file is None:
        return None
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        img = Image.open(uploaded_file)
        img.save(tmp.name)
        return tmp.name

def gerar_word(dados, img_paths):
    doc = Document()
    doc.add_heading(f"RAF - RELATÓRIO DE ANÁLISE DE FALHA ({dados['ID']})", level=1)
    
    doc.add_heading("1. Identificação", level=2)
    doc.add_paragraph(f"Status Atual: {dados['Status']}")
    doc.add_paragraph(f"Componente: {dados['Componente']} | Área: {dados['Area']}")
    doc.add_paragraph(f"Data: {dados['Data_Ocorrencia']} | Responsável: {dados['Responsavel']}")
    doc.add_paragraph(f"Link CAD (Drive): {dados['Drive_Link']}")

    doc.add_heading("2. Registro do Fato", level=2)
    doc.add_paragraph(f"Contexto: {dados['Contexto']}")
    doc.add_paragraph(f"Piloto: {dados['Piloto']} | Atividade: {dados['Atividade']}")
    doc.add_paragraph(f"Sintomas: {dados['Sintomas']}")
    if img_paths.get('quebra'):
        doc.add_picture(img_paths['quebra'], width=Inches(2.5)) # Imagem de ~6cm

    doc.add_heading("3. Análise da Causa Raiz", level=2)
    doc.add_paragraph(f"Inspeção Visual: {dados['Inspecao']}")
    for i in range(1, 6):
        doc.add_paragraph(f"{i}. {dados[f'Porque{i}']}")

    doc.add_heading("4. Solução e Validação", level=2)
    doc.add_paragraph(f"Ação Escolhida: {dados['Acao']}")
    doc.add_paragraph(f"Modificações: {dados['Modificacoes']}")
    if img_paths.get('nova'):
        doc.add_picture(img_paths['nova'], width=Inches(2.5)) # Imagem de ~6cm
    
    doc.add_paragraph(f"Resultado: {dados['Resultado']}")
    doc.add_paragraph(f"Parecer Final: {dados['Parecer']}")

    tmp_path = tempfile.mktemp(suffix=".docx")
    doc.save(tmp_path)
    return tmp_path

def gerar_pdf(dados, img_paths):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, f"RAF - RELATÓRIO DE ANALISE DE FALHA ({dados['ID']})", ln=True, align='C')
    
    pdf.set_font("Arial", size=11)
    def add_linha(titulo, texto):
        pdf.set_font("Arial", 'B', 11)
        pdf.cell(0, 8, f"{titulo}: ", ln=False)
        pdf.set_font("Arial", '', 11)
        pdf.multi_cell(0, 8, str(texto))

    pdf.ln(5)
    add_linha("Status", dados['Status'])
    add_linha("Componente", dados['Componente'])
    add_linha("Data", dados['Data_Ocorrencia'])
    add_linha("Link CAD", dados['Drive_Link'])
    
    pdf.ln(5)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, "2. O Fato", ln=True)
    add_linha("Contexto", dados['Contexto'])
    if img_paths.get('quebra'):
        pdf.image(img_paths['quebra'], w=60) # Imagem de 60mm (6cm)
        pdf.ln(2)

    pdf.ln(5)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, "3. Analise de Causa Raiz", ln=True)
    add_linha("Inspecao Visual", dados['Inspecao'])
    add_linha("Causa Raiz", dados['Porque5'])

    pdf.ln(5)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, "4. Solucao", ln=True)
    add_linha("Acao", dados['Acao'])
    if img_paths.get('nova'):
        pdf.image(img_paths['nova'], w=60)
        pdf.ln(2)
    add_linha("Parecer Final", dados['Parecer'])

    tmp_path = tempfile.mktemp(suffix=".pdf")
    pdf.output(tmp_path)
    return tmp_path

# ==========================================
# INTERFACE DO USUÁRIO (ABAS)
# ==========================================
st.title("⚙️ Central de Engenharia & Oficina - Baja")

aba1, aba2, aba3 = st.tabs(["📝 Kanban (Lançar/Editar RAF)", "🗜️ Fila de Usinagem", "🗂️ Exportar e Relatórios"])

# --- ABA 1: KANBAN E EDIÇÃO CONTÍNUA ---
with aba1:
    df = carregar_dados()
    opcoes_raf = ["✨ Abrir NOVO Registro"] + df["ID"].tolist()
    raf_selecionado = st.selectbox("Selecione a Ação:", opcoes_raf)

    if raf_selecionado != "✨ Abrir NOVO Registro":
        dados_atuais = df[df["ID"] == raf_selecionado].iloc[0].to_dict()
        st.info(f"Editando o registro: **{raf_selecionado}**")
    else:
        novo_id = f"RAF-{datetime.now().year}-{str(len(df)+1).zfill(3)}"
        dados_atuais = {col: "" for col in COLUNAS}
        dados_atuais["ID"] = novo_id
        dados_atuais["Status"] = STATUS_LIST[0]
        st.success(f"Criando novo registro: **{novo_id}**")

    with st.form("form_kanban"):
        col1, col2 = st.columns(2)
        status_idx = STATUS_LIST.index(dados_atuais["Status"]) if dados_atuais["Status"] in STATUS_LIST else 0
        
        status_novo = col1.selectbox("Status do RAF (Kanban)", STATUS_LIST, index=status_idx)
        drive_link = col2.text_input("🔗 Link da Pasta no Drive (CAD 3D/2D)", value=str(dados_atuais["Drive_Link"]))
        
        st.divider()
        st.subheader("Dados da Ocorrência")
        c1, c2, c3 = st.columns(3)
        comp = c1.text_input("Componente", value=str(dados_atuais["Componente"]))
        area = c2.selectbox("Área", ["Suspensão", "Powertrain", "Chassi", "Freios", "Elétrica"], index=0)
        data = c3.date_input("Data da Ocorrência")
        resp = st.text_input("Responsável Técnico", value=str(dados_atuais["Responsavel"]))
        contexto = st.text_area("Contexto da Quebra", value=str(dados_atuais["Contexto"]))
        
        st.divider()
        st.subheader("Análise e Solução")
        insp = st.text_input("Inspeção Macro", value=str(dados_atuais["Inspecao"]))
        pq5 = st.text_input("Causa Raiz (5º Porquê)", value=str(dados_atuais["Porque5"]))
        acao = st.text_area("Ação Escolhida / Solução", value=str(dados_atuais["Acao"]))
        
        st.info("💡 Como o banco é leve, as fotos devem ser enviadas apenas na Aba 3 na hora de gerar o PDF final.")
        
        submit = st.form_submit_button("Salvar no Banco de Dados", use_container_width=True)

    if submit:
        nova_linha = dados_atuais.copy()
        nova_linha.update({
            "Status": status_novo, "Drive_Link": drive_link, "Componente": comp, 
            "Area": area, "Data_Ocorrencia": data.strftime("%d/%m/%Y"), "Responsavel": resp, 
            "Contexto": contexto, "Inspecao": insp, "Porque5": pq5, "Acao": acao
        })

        if raf_selecionado == "✨ Abrir NOVO Registro":
            df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
        else:
            # Lógica blindada para atualizar apenas a linha correta sem apagar dados
            idx_update = df[df["ID"] == raf_selecionado].index[0]
            for coluna, valor in nova_linha.items():
                df.at[idx_update, coluna] = valor
        
        salvar_dados(df)
        st.success("✅ Dados salvos com sucesso! A base principal foi atualizada.")
        st.rerun()

# --- ABA 2: FILA DE USINAGEM E HISTÓRICO ---
with aba2:
    st.subheader("🗜️ Fila de Produção / Manufatura")
    df = carregar_dados()
    fila = df[df["Status"] == "🔵 Fila de Usinagem"]
    
    if fila.empty:
        st.success("A fila de usinagem está vazia! Nenhuma peça pendente.")
    else:
        for idx, row in fila.iterrows():
            with st.expander(f"🛠️ {row['Componente']} (RAF: {row['ID']})"):
                st.write(f"**Área:** {row['Area']} | **Responsável:** {row['Responsavel']}")
                st.write(f"**O que fazer (Ação):** {row['Acao']}")
                st.markdown(f"🔗 [Acessar Desenho no Google Drive]({row['Drive_Link']})")
                
                if st.button(f"✅ Marcar '{row['Componente']}' como Validado/Usinado", key=f"btn_{idx}"):
                    df.at[idx, "Status"] = "🟢 Validado / Fechado"
                    salvar_dados(df)
                    st.rerun()

    st.divider()
    st.subheader("🏁 Histórico de Peças Concluídas (Validadas)")
    concluidos = df[df["Status"] == "🟢 Validado / Fechado"]
    
    if concluidos.empty:
        st.info("Nenhuma RAF foi finalizada ainda.")
    else:
        # Mostra uma tabela limpa apenas com o resumo dos concluídos
        st.dataframe(
            concluidos[["ID", "Componente", "Area", "Responsavel", "Data_Ocorrencia"]], 
            hide_index=True, 
            use_container_width=True
        )
# --- ABA 3: GERADOR DE RELATÓRIOS E EXPORTAÇÃO ---
with aba3:
    st.subheader("📄 Gerar Documento Final do RAF")
    df = carregar_dados()
    
    raf_gerar = st.selectbox("Escolha o RAF para gerar documento:", df["ID"].tolist())
    
    col_img1, col_img2 = st.columns(2)
    foto_quebra = col_img1.file_uploader("Foto da Quebra", type=["jpg", "png"])
    foto_nova = col_img2.file_uploader("Foto da Nova Peça", type=["jpg", "png"])
    
    col_btn1, col_btn2 = st.columns(2)
    if col_btn1.button("📄 Gerar PDF"):
        dados = df[df["ID"] == raf_gerar].iloc[0].to_dict()
        paths = {'quebra': processar_imagem_temp(foto_quebra), 'nova': processar_imagem_temp(foto_nova)}
        pdf_path = gerar_pdf(dados, paths)
        with open(pdf_path, "rb") as f:
            st.download_button("📥 Baixar PDF Final", f, file_name=f"{raf_gerar}.pdf", mime="application/pdf")
            
    if col_btn2.button("📝 Gerar Word (.docx)"):
        dados = df[df["ID"] == raf_gerar].iloc[0].to_dict()
        paths = {'quebra': processar_imagem_temp(foto_quebra), 'nova': processar_imagem_temp(foto_nova)}
        word_path = gerar_word(dados, paths)
        with open(word_path, "rb") as f:
            st.download_button("📥 Baixar Word Final", f, file_name=f"{raf_gerar}.docx")

    st.divider()
    st.subheader("📊 Exportar Base de Dados para Dashboard (Power BI, Looker, etc.)")
    st.info("Baixe a base leve e atualizada para alimentar seu software de gráficos.")
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(label="📥 Baixar Base (CSV)", data=csv, file_name="base_dashboard_baja.csv", mime="text/csv")
