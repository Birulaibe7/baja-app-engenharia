import tempfile
import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional

import streamlit as st
import pandas as pd
import plotly.express as px
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from streamlit.runtime.uploaded_file_manager import UploadedFile

# ==========================================
# CONFIGURAÇÕES E CONSTANTES
# ==========================================
class Config:
    DATA_DIR = Path("data")
    EXCEL_FALHAS = DATA_DIR / "historico_falhas_baja.xlsx"
    EXCEL_CHECKLISTS = DATA_DIR / "historico_checklists.xlsx"
    
    AREAS = [
        "Suspensão e Direção", "Design e Estrutura", "Elétrica e Eletrônica",
        "Powertrain", "Freios", "Gestão", "Marketing", "Geral"
    ]
    ATIVIDADES = ["Treino/Oficina", "Validação de Projeto", "Competição"]

# Configuração de Logs
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# ==========================================
# CAMADA DE DADOS (MODEL)
# ==========================================
def init_database() -> None:
    """Garante que o diretório de dados e os arquivos Excel existam."""
    Config.DATA_DIR.mkdir(exist_ok=True)
    
    if not Config.EXCEL_FALHAS.exists():
        df_falhas = pd.DataFrame(columns=[
            "Codigo", "Data_Ocorrencia", "Data_Fechamento", "Responsavel", "Area",
            "Componente", "CAD_Software", "Contexto", "Piloto", "Tipo_Atividade",
            "Sintomas", "Inspecao_Visual", "Porque1", "Porque2", "Porque3", "Porque4", "Porque5",
            "Acao_Tecnica", "Modificacoes", "Nova_Versao_CAD", "Resultado_Teste", "Parecer_Final"
        ])
        df_falhas.to_excel(Config.EXCEL_FALHAS, index=False)
        
    if not Config.EXCEL_CHECKLISTS.exists():
        df_check = pd.DataFrame(columns=[
            "Data_Hora", "Tipo_Inspecao", "Responsavel", "Status_Final", "Itens_Nok"
        ])
        df_check.to_excel(Config.EXCEL_CHECKLISTS, index=False)

def save_to_excel(file_path: Path, new_data: Dict[str, Any]) -> None:
    """Salva um novo registro em um arquivo Excel existente de forma segura."""
    try:
        df_existente = pd.read_excel(file_path)
        df_novo = pd.concat([df_existente, pd.DataFrame([new_data])], ignore_index=True)
        df_novo.to_excel(file_path, index=False)
    except Exception as e:
        logging.error(f"Erro ao salvar no Excel {file_path}: {e}")
        st.error("Erro interno ao acessar o banco de dados. Contate a engenharia de software.")

# ==========================================
# CAMADA DE SERVIÇOS (WORD GENERATION)
# ==========================================
def _add_image_to_doc(doc: Document, image_file: UploadedFile, label: str) -> None:
    """Processa e adiciona uma imagem ao documento usando arquivos temporários."""
    if not image_file:
        return
        
    doc.add_paragraph(label)
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            img = Image.open(image_file)
            img.save(tmp.name)
            doc.add_picture(tmp.name, width=Inches(5))
        Path(tmp.name).unlink() # Deleta após o uso
    except Exception as e:
        logging.error(f"Erro ao processar imagem {label}: {e}")
        doc.add_paragraph("[Erro ao carregar a imagem anexa]")

def generate_word_report(dados: Dict[str, Any], img_quebra: Optional[UploadedFile], 
                         img_cad: Optional[UploadedFile], img_depois: Optional[UploadedFile]) -> Path:
    """Gera o documento Word formatado e retorna o caminho do arquivo temporário."""
    doc = Document()
    
    # Estilos de Cabeçalho
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("RAF - RELATÓRIO DE ANÁLISE DE FALHA")
    run_title.font.name = 'Arial'
    run_title.font.size = Pt(18)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    
    # 1. Identificação (Tabela)
    doc.add_heading('1. Identificação do Sistema', level=2)
    t1 = doc.add_table(rows=3, cols=2)
    t1.style = 'Table Grid'
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    t1.rows[0].cells[0].text = f"Código: {dados.get('Codigo', '')}"
    t1.rows[0].cells[1].text = f"Data: {dados.get('Data_Ocorrencia', '')}"
    t1.rows[1].cells[0].text = f"Responsável: {dados.get('Responsavel', '')}"
    t1.rows[1].cells[1].text = f"Área: {dados.get('Area', '')}"
    t1.rows[2].cells[0].text = f"Componente: {dados.get('Componente', '')}"
    t1.rows[2].cells[1].text = f"CAD: {dados.get('CAD_Software', '')}"

    # 2. Fato e Sintomas
    doc.add_heading('2. Registro do Fato', level=2)
    doc.add_paragraph(f"Contexto: {dados.get('Contexto', '')}")
    doc.add_paragraph(f"Piloto: {dados.get('Piloto', '')} | Atividade: {dados.get('Tipo_Atividade', '')}")
    doc.add_paragraph(f"Sintomas: {dados.get('Sintomas', '')}")
    _add_image_to_doc(doc, img_quebra, "Evidência Visual da Falha (No veículo):")

    # 3. Análise (5 Porquês)
    doc.add_heading('3. Análise Técnica e Causa Raiz', level=2)
    doc.add_paragraph(f"Inspeção Macroscópica: {dados.get('Inspecao_Visual', '')}")
    _add_image_to_doc(doc, img_cad, "Análise Geométrica / Concentradores de Tensão:")
    
    doc.add_heading('Método dos 5 Porquês:', level=3)
    for i in range(1, 6):
        doc.add_paragraph(f"{i}. {dados.get(f'Porque{i}', '')}")

    # 4. Solução
    doc.add_heading('4. Solução Implementada e Validação', level=2)
    doc.add_paragraph(f"Ação Escolhida: {dados.get('Acao_Tecnica', '')}")
    doc.add_paragraph(f"Modificações: {dados.get('Modificacoes', '')}")
    _add_image_to_doc(doc, img_depois, "Componente Final / Modificado:")
    doc.add_paragraph(f"Resultado do Teste: {dados.get('Resultado_Teste', '')}")
    
    p_final = doc.add_paragraph()
    run_final = p_final.add_run(f"PARECER FINAL: {dados.get('Parecer_Final', '')}")
    run_final.font.bold = True

    # Salva em um arquivo temporário seguro
    tmp_path = Path(tempfile.gettempdir()) / f"RAF_{dados.get('Codigo', '000')}.docx"
    doc.save(tmp_path)
    return tmp_path

# ==========================================
# CAMADA DE INTERFACE (VIEW)
# ==========================================
def render_tab_raf() -> None:
    st.header("Registro de Análise de Falha (RAF)")
    
    with st.form("form_raf", clear_on_submit=False):
        col1, col2, col3 = st.columns(3)
        codigo = col1.text_input("Código do Relatório", "RAF-2026-001")
        responsavel = col1.text_input("Responsável Técnico")
        data_ocorr = col2.date_input("Data da Ocorrência")
        area = col2.selectbox("Área Responsável", Config.AREAS)
        componente = col3.text_input("Componente Específico")
        cad_sw = col3.text_input("Software CAD Usado", "SolidWorks")

        st.subheader("2. O Fato e Sintomas")
        contexto = st.text_area("Contexto Detalhado da Quebra")
        col_f1, col_f2 = st.columns(2)
        piloto = col_f1.text_input("Piloto / Operador")
        sintomas = col_f1.text_input("Sintomas da Falha")
        tipo_atv = col_f2.selectbox("Tipo de Atividade", Config.ATIVIDADES)
        foto_quebra = col_f2.file_uploader("Foto da Peça Quebrada", type=["png", "jpg", "jpeg"])

        st.subheader("3. Análise da Causa Raiz (5 Porquês)")
        inspecao = st.text_area("Inspeção Macroscópica da Superfície")
        foto_cad = st.file_uploader("Print da Região Crítica no CAD", type=["png", "jpg", "jpeg"])
        
        p1 = st.text_input("1. Por que falhou mecanicamente?")
        p2 = st.text_input("2. Por que ocorreu essa condição?")
        p3 = st.text_input("3. Por que o esforço agiu assim?")
        p4 = st.text_input("4. Por que o projeto não mitigou?")
        p5 = st.text_input("5. Por que não foi previsto? (Causa Raiz)")

        st.subheader("4. Solução Final & Validação")
        acao_final = st.text_area("Ação Técnica Escolhida & Justificativa")
        modificacoes = st.text_input("Modificações de Material/Usinagem")
        nova_v_cad = st.text_input("Nome/Versão do Novo Arquivo CAD")
        foto_depois = st.file_uploader("Foto da Nova Peça", type=["png", "jpg", "jpeg"])
        resultado_teste = st.text_area("Resultado dos Testes de Campo")
        parecer = st.radio("Parecer Final", ["APROVADO PARA COMPETIÇÃO", "REJEITADO"])

        submit = st.form_submit_button("🚀 Gerar Relatório e Salvar", use_container_width=True)

    if submit:
        if not codigo or not componente:
            st.warning("⚠️ Código e Componente são campos obrigatórios.")
            return

        dados_dict = {
            "Codigo": codigo, "Data_Ocorrencia": str(data_ocorr), 
            "Data_Fechamento": str(datetime.now().date()), "Responsavel": responsavel, 
            "Area": area, "Componente": componente, "CAD_Software": cad_sw,
            "Contexto": contexto, "Piloto": piloto, "Tipo_Atividade": tipo_atv, 
            "Sintomas": sintomas, "Inspecao_Visual": inspecao, "Porque1": p1, 
            "Porque2": p2, "Porque3": p3, "Porque4": p4, "Porque5": p5,
            "Acao_Tecnica": acao_final, "Modificacoes": modificacoes, 
            "Nova_Versao_CAD": nova_v_cad, "Resultado_Teste": resultado_teste, 
            "Parecer_Final": parecer
        }
        
        with st.spinner("Processando dados e gerando documento..."):
            save_to_excel(Config.EXCEL_FALHAS, dados_dict)
            arq_word_path = generate_word_report(dados_dict, foto_quebra, foto_cad, foto_depois)
        
        st.success(f"✅ Falha {codigo} registrada com sucesso!")
        
        with open(arq_word_path, "rb") as file:
            st.download_button(
                label=f"📥 Baixar Documento Oficial ({arq_word_path.name})",
                data=file,
                file_name=arq_word_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )

def render_tab_dashboard() -> None:
    st.header("📊 Dashboard de Confiabilidade")
    try:
        df_f = pd.read_excel(Config.EXCEL_FALHAS)
    except Exception:
        st.error("Erro ao ler banco de dados de falhas.")
        return

    if df_f.empty:
        st.info("Nenhuma falha cadastrada para gerar métricas.")
        return

    m1, m2, m3 = st.columns(3)
    m1.metric("Total de Falhas", len(df_f))
    area_mode = df_f["Area"].mode()[0] if not df_f["Area"].empty else "N/A"
    m2.metric("Sub-sistema Crítico", area_mode)
    
    taxa_aprov = (len(df_f[df_f["Parecer_Final"] == "APROVADO PARA COMPETIÇÃO"]) / len(df_f)) * 100
    m3.metric("Taxa de Aprovação", f"{taxa_aprov:.1f}%")

    col_g1, col_g2 = st.columns(2)
    with col_g1:
        fig_area = px.bar(df_f, x="Area", title="Ocorrências por Sub-sistema", color="Area")
        st.plotly_chart(fig_area, use_container_width=True)
    with col_g2:
        fig_atv = px.pie(df_f, names="Tipo_Atividade", title="Falhas por Atividade", hole=0.4)
        st.plotly_chart(fig_atv, use_container_width=True)

def render_tab_search() -> None:
    st.header("🔍 Busca de Histórico (Base de Conhecimento)")
    termo = st.text_input("Digite uma palavra-chave (Ex: trinca, manga, 4340):")
    
    if termo:
        try:
            df_f = pd.read_excel(Config.EXCEL_FALHAS)
            mask = (
                df_f["Componente"].astype(str).str.contains(termo, case=False, na=False) |
                df_f["Contexto"].astype(str).str.contains(termo, case=False, na=False) |
                df_f["Acao_Tecnica"].astype(str).str.contains(termo, case=False, na=False)
            )
            resultados = df_f[mask]
            
            st.caption(f"Encontrados {len(resultados)} registro(s).")
            for _, row in resultados.iterrows():
                with st.expander(f"📌 {row['Codigo']} - {row['Componente']} ({row['Area']})"):
                    st.markdown(f"**Contexto:** {row['Contexto']}")
                    st.markdown(f"**Causa Raiz:** {row['Porque5']}")
                    st.markdown(f"**Solução:** {row['Acao_Tecnica']}")
        except Exception:
            st.error("Erro ao realizar a busca.")

def render_tab_checklist() -> None:
    st.header("📋 Checklist de Pista")
    with st.form("form_checklist"):
        col1, col2 = st.columns(2)
        c_resp = col1.text_input("Responsável")
        c_tipo = col2.selectbox("Sessão", Config.ATIVIDADES)
        
        st.divider()
        chk = [
            st.checkbox("Cinto de segurança e fixações Ok"),
            st.checkbox("Pressão do freio sem vazamentos"),
            st.checkbox("Torque dos parafusos de suspensão e roda"),
            st.checkbox("Nível de fluidos (Óleo/Freio)"),
            st.checkbox("Chave geral (Kill-switch) funcional")
        ]
        
        if st.form_submit_button("Finalizar Vistoria", use_container_width=True):
            status = "LIBERADO PARA PISTA" if all(chk) else "RETER NO BOX"
            novo_c = {
                "Data_Hora": str(datetime.now().strftime("%Y-%m-%d %H:%M")),
                "Tipo_Inspecao": c_tipo,
                "Responsavel": c_resp,
                "Status_Final": status,
                "Itens_Nok": len(chk) - sum(chk)
            }
            save_to_excel(Config.EXCEL_CHECKLISTS, novo_c)
            
            if status == "LIBERADO PARA PISTA":
                st.success("🟢 CARRO LIBERADO!")
                st.balloons()
            else:
                st.error("🔴 CARRO RETIDO! Verifique pendências.")

# ==========================================
# EXECUÇÃO PRINCIPAL
# ==========================================
def main() -> None:
    st.set_page_config(page_title="Central de Engenharia - Baja SAE", page_icon="🏎️", layout="wide")
    init_database()
    
    st.title("🏎️ Central de Engenharia & Oficina - Baja")
    tabs = st.tabs(["📝 Registrar RAF", "📊 Dashboard", "🔍 Base de Conhecimento", "📋 Checklist"])
    
    with tabs[0]: render_tab_raf()
    with tabs[1]: render_tab_dashboard()
    with tabs[2]: render_tab_search()
    with tabs[3]: render_tab_checklist()

if __name__ == "__main__":
    main()