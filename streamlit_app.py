import streamlit as st
from docx import Document
import tempfile

# Função para preencher o documento de termo de compromisso
def preencher_termo(nome_aluno, caminho_modelo):
    # Abre o modelo do documento
    doc = Document(caminho_modelo)
    
    # Substitui as informações do aluno
    for paragrafo in doc.paragraphs:
        if 'NOME_DO_ALUNO' in paragrafo.text:
            paragrafo.text = paragrafo.text.replace('NOME_DO_ALUNO', nome_aluno)
    
    # Salva o documento temporariamente
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp_file.name)
    return temp_file.name

# Caminho fixo para o modelo de documento
caminho_modelo = r"G:/Meu Drive/Repositório Git/Unirb/Painel de Documentos/Modelo_de_Estagio.docx"

# Caminho para a logo da Rede UNIRB
# caminho_logo = r"G:/Meu Drive/Repositório Git/Unirb/Painel de Documentos/logo_rede_unirb.png"

# Exibe a logo no topo da aplicação
# st.image(caminho_logo, width=200)  # Ajuste o valor de width conforme necessário

# Interface do Streamlit
st.title('Gerador de documentação dos Estágio')

# Formulário para entrada de dados do aluno
st.subheader("Dados do Aluno")
nome_aluno = st.text_input("Nome do Aluno")

# Botão para gerar o documento
if st.button("Gerar Documento"):
    if nome_aluno:
        # Preenche o termo de compromisso
        nome_arquivo = preencher_termo(nome_aluno, caminho_modelo)
        
        # Exibe o link para download do documento
        with open(nome_arquivo, "rb") as file:
            st.download_button(label=f"Download Termo de Compromisso - {nome_aluno}", data=file, file_name=f"Termo_Compromisso_{nome_aluno.replace(' ', '_')}.docx")
    else:
        st.warning("Por favor, insira o nome do aluno.")


