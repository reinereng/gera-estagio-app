import streamlit as st
from docx import Document
import tempfile
from datetime import datetime, timedelta
import numpy as np
import pandas as pd

from PIL import Image
import requests
from io import BytesIO

# URL da imagem no GitHub
image_url = "https://raw.githubusercontent.com/reinereng/gera-estagio-app/main/imagens/logo_rede_unirb.jpg"

# Carrega a imagem diretamente da URL
response = requests.get(image_url)
img = Image.open(BytesIO(response.content))

# Caminho das DOCUMENTAÇÕES E IMAGENS

# Função para preencher o documento de termo de compromisso
def preencher_termo(nome_aluno, documento):
    dados_alunos = str(nome_aluno) + ", RG: " + str(rg_aluno) + ", CPF: " + str(cpf_aluno)
    dados_alunos = dados_alunos + ", estudante com matrícula " + str(matricula) + ", residente " + endereco_aluno
    
    dados_empresa = nome_empresa + " com sede e foro na Cidade de " + cidade_empresa + ", " + uf_empresa + ", estabelecida no endereço: "
    dados_empresa = dados_empresa +  endereco_empresa + ", " + bairro_empresa + ", CEP: " + cep_empresa + ", cadastrada no CNPJ: " 
    dados_empresa = dados_empresa + cnpj_empresa + ", neste ato representada por " + representante + ", CPF: " + cpf_empresa
 
    horasestagio = str(total_horas_estagio) + " horas, com jornada de estágio de " + str(horas_por_dia) + " horas por dia e carga horária máxima de 30 horas/semanais"
    
    supervisor_texto = supervisor + ", " + cargosupervisor + ", com registro no conselho profissional sob o número "  + str(conselho_empresa)

    # Usar o documento já carregado
    doc = documento
    
    # Substitui as informações do aluno
    for paragrafo in doc.paragraphs:
        for run in paragrafo.runs:  # Cada 'run' é uma sequência de caracteres com formatação específica  
            if '<DADOS_EMPRESA>' in run.text:
                run.text = run.text.replace('<DADOS_EMPRESA>', dados_empresa)   
                run.bold = True
            if '<DADOS_ALUNO>' in run.text:
                run.text = run.text.replace('<DADOS_ALUNO>', dados_alunos)   
                run.bold = True
            if '<CURSO_ALUNO>' in run.text:
                run.text = run.text.replace('<CURSO_ALUNO>', curso_aluno.upper()) 
                run.bold = True         
            if '<SEMESTRE>' in run.text:
                run.text = run.text.replace('<SEMESTRE>', str(semestre))   
                run.bold = True           
            if '<DATAIN>' in run.text:
                run.text = run.text.replace('<DATAIN>', data_inicio_str)   
                run.bold = True                    
            if '<DATAFIM>' in run.text:
                run.text = run.text.replace('<DATAFIM>', data_termino_str)   
                run.bold = True         
            if '<HORAS>' in run.text:
                run.text = run.text.replace('<HORAS>', horasestagio)   
                run.bold = True                                
            if '<PROFESSOR>' in run.text:
                run.text = run.text.replace('<PROFESSOR>', professor)   
                run.bold = True                                
            if '<SUPERVISOR>' in run.text:
                run.text = run.text.replace('<SUPERVISOR>', supervisor_texto)   
                run.bold = True                                
            if '<ALUNO>' in run.text:
                run.text = run.text.replace('<ALUNO>', nome_aluno)   
                run.bold = True                                
            if '<EMPRESA>' in run.text:
                run.text = run.text.replace('<EMPRESA>', nome_empresa.upper())   
                run.bold = True       
            if '<DONOEMPRESA>' in run.text:
                run.text = run.text.replace('<DONOEMPRESA>', representante)   
                run.bold = True                   
            if '<HOJE>' in run.text:
                run.text = run.text.replace('<HOJE>', datetime.now().strftime('%d/%m/%Y'))   
                run.bold = True      
                   
    # Salva o documento temporariamente
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp_file.name)
    return temp_file.name


# ----------------------- STREAMLIT ------------------------------------------
# ----------------------------------------------------------------------------
# Exibe a logo no topo da aplicação
st.image(img, width=300)  # Ajuste o valor de width conforme necessário

# Interface do Streamlit
st.title('Gerador de documentação para Estágio Supervisionado')

# Lista de opções de IES
opcoes_ies = [
    "Centro Universitário Unirb",
    "Centro Universitário Unirb Alagoinhas",
    "Faculdade Diplomata"
]

# Exibe a lista suspensa para escolha da IES
ies_escolhida = st.selectbox("Escolha a IES", opcoes_ies)

# Exibe a opção escolhida
st.write(f"Você selecionou: {ies_escolhida}")

if ies_escolhida == "Centro Universitário Unirb":
    #DOCUMENTOS
    # URL do documento no GitHub para o termo de compromisso
    doc_url_termo = "https://raw.githubusercontent.com/reinereng/gera-estagio-app/main/modelos/Modelo_Termo_Compromisso_Centro.docx"
    response_termo = requests.get(doc_url_termo)
    caminho_termo = Document(BytesIO(response_termo.content))

    # URL do documento no GitHub para o termo de convênio
    doc_url_conv = "https://raw.githubusercontent.com/reinereng/gera-estagio-app/main/modelos/Modelo_Termo_Convenio_Centro.docx"
    response_conv = requests.get(doc_url_conv)
    caminho_conv = Document(BytesIO(response_conv.content))

elif ies_escolhida == "Centro Universitário Unirb Alagoinhas":
    #DOCUMENTOS
    caminho_termo = "C:/Users/Reiner/Painel de Documentos/Modelo_Termo_Compromisso_Centro_Alagoinhas.docx"
    caminho_conv = "C:/Users/Reiner/Painel de Documentos/Modelo_Termo_Convenio_Centro_Alagoinhas.docx"

elif ies_escolhida == "Faculdade Diplomata":
    #DOCUMENTOS
    caminho_termo = "C:/Users/Reiner/Painel de Documentos/Modelo_Termo_Compromisso_Diplomata.docx"
    caminho_conv = "C:/Users/Reiner/Painel de Documentos/Modelo_Termo_Convenio_Diplomata.docx"


dados_ies = "AMERICA EDUCACIONAL S.A - CENTRO UNIVERSITÁRIO UNIRB - SALVADOR, situado à Av. Tamburugy, 474 - Patamares, Salvador - BA, CEP: 41680-440, tel: (71) 3368-8300 e-mail: unirb@unirb.edu.br inscrita no CNPJ 28.844.791/0001-91 representada neste ato, por Carlos Joel Pereira, CPF: 159.659.615-53"

# Lista de opções de IES
opcoes_Professor = [
    "Reiner Requião de Souza",
    "Francianne Oliveira",
    "Rejane da Costa"]

# Exibe a lista suspensa para escolha da IES
professor_escolhido = st.selectbox("Escolha a IES", opcoes_Professor)

if professor_escolhido == "Reiner Requião de Souza":
    professor = "Reiner Requião de Souza, portador do CPF n.º 009.893.855-07 e RG nº 07584711-65 SSP/BA"
if professor_escolhido == "Rejane da Costa":
    professor = "Rejane da Costa, portador do CPF n.º 006.411.315-93"
if professor_escolhido == "Francianne Oliveira":
    professor = "Francianne Oliveira, portador do CPF n.º NÃO TEM, RG nº "


# Inicializa o estado da sessão para armazenar os arquivos gerados e o DataFrame
if 'arquivos_gerados' not in st.session_state:
    st.session_state.arquivos_gerados = {}

if 'df' not in st.session_state:
    st.session_state.df = None

# Escolha do método de entrada
opcao_entrada = st.radio(
    "Como você gostaria de inserir os dados?",
    ("Manual", "Planilha")
)

if opcao_entrada == "Planilha":
    # Upload de planilha
    arquivo_planilha = st.file_uploader("Carregue a planilha de entrada", type=["xlsx"])
    
    if arquivo_planilha is not None:
        # # Leitura dos dados da planilha
        # df = pd.read_excel(arquivo_planilha)
        # st.write("Dados carregados da planilha:", df)
        # # Use 'dados_planilha' para preencher automaticamente os campos necessários

        # Limpa o DataFrame antigo no estado da sessão
        st.session_state.df = pd.read_excel(arquivo_planilha)
        st.write("Dados carregados da planilha:", st.session_state.df)

        # Faz uma cópia local do DataFrame
        df = st.session_state.df.copy()
        
        # Itera sobre cada linha do DataFrame
        for i in range(len(df)):
            # Dados do Aluno
            nome_aluno = df["Nome do Aluno"][i]
            curso_aluno = df["Curso do Aluno"][i]
            rg_aluno = df["RG do Aluno"][i]
            semestre = df["Semestre Atual do Aluno"][i]
            cpf_aluno = df["CPF do Aluno"][i]
            matricula = df["Matrícula"][i]
            endereco_aluno = df["Endereço Completo do Aluno"][i]
            
            # Dados do Estágio
            total_horas_estagio = df["Total de Horas do Estágio"][i]
            horas_por_dia = df["Horas de Estágio por Dia"][i]

                
            # Converte data_inicio para datetime com tratamento de erros
            data_inicio = pd.to_datetime(df["Data de Início do Estágio"][i], format='%d/%m/%Y', errors='coerce')
            
            # Converte para o tipo datetime64[D] necessário para np.busday_offset
            if pd.notnull(data_inicio):
                data_inicio_np = data_inicio.to_numpy().astype('datetime64[D]')
            else:
                print(f"Data de início inválida para o índice {i}.")
                continue

            # Verifica se a data de término é nula ou precisa ser calculada
            data_termino = df["Data de Término do Estágio"][i]
            if pd.isnull(data_termino) or data_termino == '':
                dias_necessarios = np.ceil(total_horas_estagio / horas_por_dia)
                data_termino_np = np.busday_offset(data_inicio_np, dias_necessarios-1, roll='forward')
            else:
                # Converte data_termino para datetime
                data_termino = pd.to_datetime(data_termino, format='%d/%m/%Y', errors='coerce')
                data_termino_np = data_termino.to_numpy().astype('datetime64[D]')

            # Garante que ambas as datas estejam no formato datetime antes de usar strftime
            if pd.notnull(data_inicio) and pd.notnull(data_termino_np):
                data_inicio_str = data_inicio.strftime('%d/%m/%Y')
                data_termino_str = pd.Timestamp(data_termino_np).strftime('%d/%m/%Y')

                # Processa ou exibe os dados como necessário
                print(f"Data de Início: {data_inicio_str}, Data de Término: {data_termino_str}")
            else:
                print("Erro: Datas inválidas encontradas.")    
                
            
            # Dados da Empresa
            nome_empresa = df["Nome da Empresa"][i]
            representante = df["Representante da Empresa"][i]
            supervisor = df["Supervisor"][i]
            cargosupervisor = df["Cargo do Supervisor"][i]
            cnpj_empresa = df["CNPJ da Empresa"][i]
            cpf_empresa = df["CPF do Representante"][i]
            conselho_empresa = df["Registro do Conselho do Supervisor"][i]
            endereco_empresa = df["Endereço da Empresa"][i]
            bairro_empresa = df["Bairro da Empresa"][i]
            cidade_empresa = df["Cidade da Empresa"][i]
            cep_empresa = df["CEP da Empresa"][i]
            uf_empresa = df["UF da Empresa"][i]

            # Inicializa o estado da sessão para armazenar os arquivos gerados
            if nome_aluno not in st.session_state.arquivos_gerados:
                st.session_state.arquivos_gerados[nome_aluno] = {'termo': None, 'convenio': None}

            if nome_aluno:
                # Preenche o termo de compromisso
                nome_arquivo_termo = preencher_termo(nome_aluno, caminho_termo)
                
                # Preenche o termo de convênio
                nome_arquivo_convenio = preencher_termo(nome_aluno, caminho_conv)
                
                # Armazena os caminhos dos arquivos no estado da sessão
                st.session_state.arquivos_gerados[nome_aluno]['termo'] = nome_arquivo_termo
                st.session_state.arquivos_gerados[nome_aluno]['convenio'] = nome_arquivo_convenio

            # Verifica se os arquivos foram gerados e exibe os botões de download
            if st.session_state.arquivos_gerados[nome_aluno]['termo'] and st.session_state.arquivos_gerados[nome_aluno]['convenio']:
                with open(st.session_state.arquivos_gerados[nome_aluno]['termo'], "rb") as file_termo:
                    st.download_button(
                        label=f"Download Termo de Compromisso - {nome_aluno}", 
                        data=file_termo, 
                        file_name=f"Termo_Compromisso_{nome_aluno.replace(' ', '_')}.docx"
                    )
                
                with open(st.session_state.arquivos_gerados[nome_aluno]['convenio'], "rb") as file_convenio:
                    st.download_button(
                        label=f"Download Termo de Convênio - {nome_aluno}", 
                        data=file_convenio, 
                        file_name=f"Termo_Convenio_{nome_aluno.replace(' ', '_')}.docx"
                    )

else:
    st.subheader("Dados do Aluno")
    # Cria três colunas para os campos do formulário
    c1, c2, c3 = st.columns(3)
    with c1:
        nome_aluno = st.text_input("Nome do Aluno")
        curso_aluno = st.text_input("Curso do Aluno")
    with c2:
        rg_aluno = st.text_input("RG do Aluno")
        semestre = st.text_input("Semestre Atual do Aluno")    
    with c3:
        cpf_aluno = st.text_input("CPF do Aluno")
        matricula = st.text_input("Matrícula")    

    endereco_aluno = st.text_input("Endereço Completo do Aluno")    

    st.subheader("Dados do Estágio")
    # Cria duas colunas para os campos do formulário
    col1, col2 = st.columns(2)
    with col1:
        total_horas_estagio = st.number_input("Total de Horas do Estágio", min_value=1, value=120)

    with col2:
        horas_por_dia = st.number_input("Horas de Estágio por Dia", min_value=1, max_value=6, value=4)# Calcula a data de término 20 dias úteis após a data de início

    # Cria duas colunas para os campos do formulário
    col1a, col2a = st.columns(2)
    with col1a:
        data_inicio = st.date_input("Data de Início do Estágio")
        # Calcula o número de dias úteis necessários para completar o estágio
        dias_necessarios = np.ceil(total_horas_estagio / horas_por_dia)

        # Exibe o número de dias úteis calculados
        st.write(f"Número de dias úteis necessários: {int(dias_necessarios)}")

        data_termino_sugerida = np.busday_offset(data_inicio, dias_necessarios-1, roll='forward')

        # Converte a data de término sugerida para o formato datetime do Python
        data_termino_sugerida = datetime.strptime(str(data_termino_sugerida), '%Y-%m-%d')

    with col2a:
        data_termino = st.date_input("Data de Término do Estágio", value=data_termino_sugerida)

    # Formatar as datas para o formato "DD/MM/YYYY"
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_termino_str = data_termino.strftime('%d/%m/%Y')

    # Calcula o número de dias úteis entre as duas datas
    dias_uteis = np.busday_count(data_inicio, data_termino)

    # Exibe o número de dias úteis
    st.write(f"Total de dias úteis alocados: {dias_uteis}")

    st.subheader("Dados da Empresa")    

    # Cria duas colunas para os campos do formulário
    col6, col7 = st.columns(2)
    with col6:
        nome_empresa = st.text_input("Nome da Empresa")
        representante = st.text_input("Representante da Empresa")    
        supervisor = st.text_input("Supervisor")
        cargosupervisor = st.text_input("Cargo do Supervisor")
    with col7:
        cnpj_empresa = st.text_input("CNPJ da Empresa")    
        cpf_empresa = st.text_input("CPF do Representante")
        conselho_empresa = st.text_input("Registro do Conselho do Supervisor") 

    # Cria três colunas para os campos do formulário
    col3, col4, col5 = st.columns(3)

    # Campos de entrada para os dados do aluno
    with col3:
        endereco_empresa = st.text_input("Endereço da Empresa")
        bairro_empresa = st.text_input("Bairro da Empresa")
    with col4:
        cidade_empresa = st.text_input("Cidade da Empresa")
        cep_empresa = st.text_input("CEP da Empresa")
    with col5:
        uf_empresa = st.text_input("UF da Empresa")
        
    # Botão para gerar o documento
    if st.button("Gerar Documento"):
        
        if nome_aluno:
            # Inicializa o estado da sessão para armazenar os arquivos gerados
            nome_arquivo_termo = preencher_termo(nome_aluno, caminho_termo)
            
            # Preenche o termo de convênio
            nome_arquivo_convenio = preencher_termo(nome_aluno, caminho_conv)
            
            # Armazena os caminhos dos arquivos no estado da sessão
            st.session_state.arquivos_gerados[nome_aluno]['termo'] = nome_arquivo_termo
            st.session_state.arquivos_gerados[nome_aluno]['convenio'] = nome_arquivo_convenio

            # Verifica se os arquivos foram gerados e exibe os botões de download
            if st.session_state.arquivos_gerados[nome_aluno]['termo'] and st.session_state.arquivos_gerados[nome_aluno]['convenio']:
                with open(st.session_state.arquivos_gerados[nome_aluno]['termo'], "rb") as file_termo:
                    st.download_button(
                        label=f"Download Termo de Compromisso - {nome_aluno}", 
                        data=file_termo, 
                        file_name=f"Termo_Compromisso_{nome_aluno.replace(' ', '_')}.docx"
                    )
                
                with open(st.session_state.arquivos_gerados[nome_aluno]['convenio'], "rb") as file_convenio:
                    st.download_button(
                        label=f"Download Termo de Convênio - {nome_aluno}", 
                        data=file_convenio, 
                        file_name=f"Termo_Convenio_{nome_aluno.replace(' ', '_')}.docx"
                    )        
        else:
            st.warning("Por favor, insira o nome do aluno.")

