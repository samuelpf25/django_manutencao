# última edição 24/01/2024
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import streamlit as st
import pandas as pd
from datetime import date, datetime
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

# ocultar menu
hide_streamlit_style = """
<meta http-equiv="Content-Language" content="pt-br">
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)
# fim ocultar menu

# 1) DECLARAÇÃO DE VARIÁVEIS GLOBAIS ####################################################################################
scope = ['https://spreadsheets.google.com/feeds']
k = "456"
creds = ServiceAccountCredentials.from_json_keyfile_name("controle.json", scope)

cliente = gspread.authorize(creds)

# sheet = cliente.open("Ciente Limpeza").sheet1 # Open the spreadhseet

sheet = cliente.open_by_url(
    'https://docs.google.com/spreadsheets/d/1Pf9OLpWmrmVtIBjQyxxZEwnFmHrf8jgeZ6kB9z2_F9g/edit#gid=0').get_worksheet(
    0)  # https://docs.google.com/spreadsheets/d/1PhJXFOKdEAjcILQCDyJ-couaDM6EWBUXM1GVh-3gZWM/edit#gid=96577098

dados = sheet.get_all_records()  # Get a list of all records

df = pd.DataFrame(dados)
df = df.astype(str)

datando = []
n_solicitacao = []
nome = []
telefone = []
predio = []
sala = []
data = []
observacao = []
os = []
obsemail = []
obsinterna = []
stat = []
d_agend = []
h_agend = []
horas = ['08:00', '08:30', '09:00', '09:30', '10:00', '10:30', '11:00', '11:30', '12:00', '12:30', '13:00', '13:30',
         '14:00', '14:30', '15:00', '15:30', '16:00', '16:30', '17:00', '17:30', '18:00', '18:30', '19:00', '19:30']

# 2) padroes #####################################################################################################
padrao = '<p style="font-family:Courier; color:Blue; font-size: 16px;">'
infor = '<p style="font-family:Courier; color:Green; font-size: 16px;">'
alerta = '<p style="font-family:Courier; color:Red; font-size: 17px;">'
titulo = '<p style="font-family:Courier; color:Blue; font-size: 20px;">'
cabecalho = '<div id="logo" class="span8 small"><a title="Universidade Federal do Tocantins"><img src="https://ww2.uft.edu.br/images/template/brasao.png" alt="Universidade Federal do Tocantins"><span class="portal-title-1"></span><h1 class="portal-title corto">Universidade Federal do Tocantins</h1><span class="portal-description">COINFRA - MANUTENÇÃO PREDIAL</span></a></div>'


# @st.cache
# def carrega_todos(status,indice,os,obsemail,obsinterna):
#     status = st.selectbox('Selecione o status:', status, index=indice)
#     os = st.text_input('Número da OS:', value=os[n])
#     obs_email = st.text_area('Observação para o Usuário:', value=obsemail[n])
#     obs_interna = st.text_area('Observação Interna:', value=obsinterna[n])
#     s = st.text_input("Senha:", value="", type="password")  # , type="password"
#     return status,os,obs_email,obs_interna,s

# 3) FUNÇÕES GLOBAIS #############################################################################################


def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return str(len(str_list) + 1)


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'})
    worksheet.set_column('A:A', None, format1)
    writer.close()
    processed_data = output.getvalue()
    return processed_data


# 4) PÁGINAS ####################################################################################


st.sidebar.title('Gestão Manutenção Predial')
a = k
# pg=st.sidebar.selectbox('Selecione a Página',['Solicitações em Aberto','Solicitações a Finalizar','Consulta'])
pg = st.sidebar.radio('', ['Edição individual', 'Edição em Lote', 'Consulta'])
status = ['', 'Todas Ativas', 'Indeferido', 'OS Aberta', 'Pendente de Material', 'Pendente Solicitante',
          'Pendente Outros', 'Atendida', 'Cancelada', 'Atendida Parcialmente', 'Material Solicitado', 'Ignorar',
          'Observação', 'Reitoria', 'Material Disponível', 'Agendamento', 'Agendado', 'Em Andamento', 'Programada']
status_todos = ['', 'OS Aberta', 'Pendente de Material', 'Pendente Solicitante', 'Pendente Outros',
                'Material Solicitado', 'Observação', 'Material Disponível', 'Agendamento', 'Agendado', 'Em Andamento',
                'Programada']
if (pg == 'Edição individual'):
    # PÁGINA EDIÇÃO INDIVIDUAL ******************************************************************************************
    st.markdown(cabecalho, unsafe_allow_html=True)
    st.subheader(pg)
    # cabeçalho

    col1, col2 = st.columns(2)
    filtrando = col1.multiselect('Selecione o Status para Filtrar', status)
    # print(filtrando)
    filtra_os = col2.text_input('Filtrar OS:', value='')

    for dic in df.index:
        if (filtrando == ['Todas Ativas']):
            filtrando = status_todos
        if filtra_os != '':
            if df['Status'][dic] in filtrando and df['Área de Manutenção'][dic] != '' and str(
                    df['Ordem de Serviço'][dic]) == str(filtra_os):
                # print(df['Código da UFT'][dic])
                n_solicitacao.append(df['Código da UFT'][dic])
                nome.append(df['Nome do solicitante'][dic])
                telefone.append(df['Telefone'][dic])
                predio.append(df['Prédio'][dic])
                sala.append(df['Sala/Local'][dic])
                data.append(df['DATASOL'][dic])

                observacao.append(df['Descrição sucinta'][dic])
                os.append(df['Ordem de Serviço'][dic])
                obsemail.append(df['Observação p/ Solicitante'][dic])
                obsinterna.append(df['Observação Interna'][dic])
                stat.append(df['Status'][dic])
                d_agend.append(df['data_agendamento'][dic])
                h_agend.append(df['hora_agendamento'][dic])
        else:

            if df['Status'][dic] in filtrando and df['Área de Manutenção'][dic] != '':
                # print(df['Código da UFT'])
                n_solicitacao.append(df['Código da UFT'][dic])
                nome.append(df['Nome do solicitante'][dic])
                telefone.append(df['Telefone'][dic])
                predio.append(df['Prédio'][dic])
                sala.append(df['Sala/Local'][dic])
                data.append(df['DATASOL'][dic])
                observacao.append(df['Descrição sucinta'][dic])
                os.append(df['Ordem de Serviço'][dic])
                obsemail.append(df['Observação p/ Solicitante'][dic])
                obsinterna.append(df['Observação Interna'][dic])
                stat.append(df['Status'][dic])
                d_agend.append(df['data_agendamento'][dic])
                h_agend.append(df['hora_agendamento'][dic])

    if len(n_solicitacao) > 1 and (filtra_os != ''):
        st.markdown(
            alerta + f'<Strong><i>Foram encontradas {len(n_solicitacao)} Ordens de Serviço com este mesmo número, selecione abaixo a solicitação correspondente:</i></Strong></p>',
            unsafe_allow_html=True)
    selecionado = st.selectbox('Nº da solicitação:', n_solicitacao)

    if (len(n_solicitacao) > 0):
        n = n_solicitacao.index(selecionado)

        # apresentar dados da solicitação
        st.markdown(titulo + '<b>Dados da Solicitação</b></p>', unsafe_allow_html=True)
        # st.text('<p style="font-family:Courier; color:Blue; font-size: 20px;">Nome: '+ nome[n]+'</p>',unsafe_allow_html=True)

        st.markdown(padrao + '<b>Nome</b>: ' + str(nome[n]) + '</p>', unsafe_allow_html=True)
        st.markdown(padrao + '<b>Telefone</b>: ' + str(telefone[n]) + '</p>', unsafe_allow_html=True)
        st.markdown(padrao + '<b>Prédio</b>: ' + str(predio[n]) + '</p>', unsafe_allow_html=True)
        st.markdown(padrao + '<b>Sala</b>: ' + str(sala[n]) + '</p>', unsafe_allow_html=True)
        st.markdown(alerta + '<b>Data</b>: ' + str(data[n]) + '</p>', unsafe_allow_html=True)
        st.markdown(padrao + '<b>Descrição</b>: ' + observacao[n] + '</p>', unsafe_allow_html=True)

        # st.markdown(padrao + '<b>Data de Agendamento</b>: ' + d_agend[n] + '</p>', unsafe_allow_html=True)
        # st.markdown(padrao + '<b>Hora de Agendamento</b>: ' + h_agend[n] + '</p>', unsafe_allow_html=True)

        celula = sheet.find(str(n_solicitacao[n]))
        # procurando status equivalente na lista
        indice = 0
        cont = 0
        numero = ""
        for j in status:
            if j == stat[n]:
                indice = cont
                numero = j
            cont = cont + 1

        cont = 0
        ind_ag = 0
        if (h_agend[n] != ''):

            for j in horas:
                if j == h_agend[n]:
                    ind_ag = cont
                cont = cont + 1
        # histórico

        with st.expander("Histórico da OS"):
            bot = st.button("Carregar Histórico")
            if bot == True:
                with st.spinner('Carregando dados...'):
                    sheet.update_acell('AC1', selecionado)  # Numero UFT
                    # df = pd.DataFrame(dados)
                    # df = df.astype(str)
                    st.dataframe(df[['DATA_HIST', 'HORA_HIST', 'STATUS_HIST', 'OBS_HIST']])
                # st.success('Dados carregados!')
        with st.form(key='my_form'):
            status = st.selectbox('Selecione o status:', status, index=indice)
            os = st.text_input('Número da OS:', value=os[n])
            obs_email = st.text_area('Observação para o Usuário:', value=obsemail[n])
            obs_interna = st.text_area('Observação Interna:', value=obsinterna[n])

            with st.expander("Agendamento"):
                d = '01/01/2022'
                print('Data Agendamento registrada: ' + d_agend[n])
                if (d_agend[n] != ''):
                    d = d_agend[n]
                else:
                    # st.text('OS sem agendamento registrado ou com data de agendamento anterior a hoje!')
                    st.markdown(
                        alerta + '<b>OS sem agendamento registrado ou com data de agendamento anterior a hoje!</b></p>',
                        unsafe_allow_html=True)
                d = d.replace('/', '-')
                data_ag = datetime.strptime(d, '%d-%m-%Y')
                if (data_ag == ''):
                    data_ag = datetime.strptime("01-01-2022", '%d-%m-%Y')
                data_agendamento = st.date_input('Data de Agendamento', value=data_ag)
                hora_agendamento = st.selectbox('Hora de agendamento:', horas, index=ind_ag)
            s = st.text_input("Senha:", value="", type="password")  # , type="password"

            botao = st.form_submit_button('Registrar')

        if (botao == True and s == a):
            if (sheet.cell(celula.row, 20).value == n_solicitacao[n] and sheet.cell(celula.row,
                                                                                    20).value != 0 and sheet.cell(
                    celula.row, 20).value != ''):
                with st.spinner('Registrando dados...Aguarde!'):
                    st.markdown(infor + '<b>Registro efetuado!</b></p>', unsafe_allow_html=True)

                    sheet.update_acell('W' + str(celula.row), status)  # Status
                    sheet.update_acell('V' + str(celula.row), os)  # os
                    sheet.update_acell('X' + str(celula.row), obs_email)  # obs_email
                    sheet.update_acell('Y' + str(celula.row), obs_interna)  # obs_interna
                    # sheet.update_acell('R' + str(celula.row), '')  # apagar Sim para enviar e-mail
                    print(os)
                    # print(data_agendamento)
                    data_agend = str(data_agendamento.year) + '-' + str(data_agendamento.month) + '-' + str(
                        data_agendamento.day)
                    data_ver = str(data_ag.year) + '-' + str(data_ag.month) + '-' + str(data_ag.day)
                    # print(data_ver)
                    if (data_agend != data_ver):
                        data = data_agendamento
                        data_formatada = str(data.day) + '/' + str(data.month) + '/' + str(data.year)
                        sheet.update_acell('AM' + str(celula.row), data_formatada)
                        sheet.update_acell('AN' + str(celula.row), hora_agendamento)
                    # sheet.update_acell('P'+str(celula.row),status) #Status
                    # sheet.update_acell('O' + str(celula.row), os)  # os
                    # sheet.update_acell('S' + str(celula.row), obsemail)  # obs_email
                    # sheet.update_acell('AA' + str(celula.row), obsinterna)  # obs_interna
                st.success('Registro efetuado!')
            else:
                st.error('Código de OS inválido!')
        elif (botao == True and s != a):
            st.markdown(alerta + '<b>Senha incorreta!</b></p>', unsafe_allow_html=True)
    else:
        st.markdown(infor + '<b>Não há itens na condição ' + pg + '</b></p>', unsafe_allow_html=True)

elif pg == 'Edição em Lote':

    # PÁGINA EDIÇÃO EM LOTE  ******************************************************************************************
    st.markdown(cabecalho, unsafe_allow_html=True)

    st.subheader(pg)
    col1, col2 = st.columns(2)
    filtrando = col1.multiselect('Selecione o Status para Filtrar', status)
    if (filtrando == ['Todas Ativas']):
        filtrando = status_todos
    os_gerais = []
    for dic in df.index:
        if (filtrando != ''):
            if df['Ordem de Serviço'][dic] != '' and df['Status'][dic] in filtrando:
                os_gerais.append(df['Ordem de Serviço'][dic])
        else:
            if df['Ordem de Serviço'] != '':
                os_gerais.append(df['Ordem de Serviço'][dic])

    filtra_os = col2.multiselect('Filtrar OS:', os_gerais)
    for dic in df.index:
        print(str(df['Ordem de Serviço'][dic]))
        # print(str(filtra_os))
        if filtra_os != '':
            if df['Status'][dic] in filtrando and df['Área de Manutenção'][dic] != '' and (
                    str(df['Ordem de Serviço'][dic]) in filtra_os) and (str(df['Ordem de Serviço'][dic]) != '') and (
                    str(df['Ordem de Serviço'][dic]) != 0):
                # print(df['Código da UFT'][dic])
                n_solicitacao.append(df['Código da UFT'][dic])
                nome.append(df['Nome do solicitante'][dic])
                telefone.append(df['Telefone'][dic])
                predio.append(df['Prédio'][dic])
                sala.append(df['Sala/Local'][dic])
                data.append(df['DATASOL'][dic])
                observacao.append(df['Descrição sucinta'][dic])
                os.append(df['Ordem de Serviço'][dic])
                obsemail.append(df['Observação p/ Solicitante'][dic])
                obsinterna.append(df['Observação Interna'][dic])
                stat.append(df['Status'][dic])
        else:
            if df['Status'][dic] in filtrando and df['Área de Manutenção'][dic] != '':
                # print(df['Código da UFT'][dic])
                n_solicitacao.append(df['Código da UFT'][dic])
                nome.append(df['Nome do solicitante'][dic])
                telefone.append(df['Telefone'][dic])
                predio.append(df['Prédio'][dic])
                sala.append(df['Sala/Local'][dic])
                data.append(df['DATASOL'][dic])
                observacao.append(df['Descrição sucinta'][dic])
                os.append(df['Ordem de Serviço'][dic])
                obsemail.append(df['Observação p/ Solicitante'][dic])
                obsinterna.append(df['Observação Interna'][dic])
                stat.append(df['Status'][dic])
    # if len(n_solicitacao)>1:
    #    st.markdown(alerta + f'<Strong><i>Foram encontradas {len(n_solicitacao)} Ordens de Serviço com este mesmo número, exclua da lista abaixo a solicitação que não for correspondente a que queira editar:</i></Strong></p>',unsafe_allow_html=True)

    selecionado = st.multiselect('Nº da solicitação:', n_solicitacao, n_solicitacao)
    filtro = selecionado
    dados1 = df[['Área de Manutenção', 'Prédio', 'DATASOL', 'Ordem de Serviço', 'Status', 'Código da UFT']]
    filtrar = dados1['Código da UFT'].isin(filtro)
    # print(dados1[filtrar]['Ordem de Serviço'].value_counts())
    # print(dados1[filtrar]['Ordem de Serviço'].value_counts().values)
    lista_repetidos = list(dados1[filtrar]['Ordem de Serviço'].value_counts().values)
    repeticao = 0
    for repetido in lista_repetidos:
        #    valor = dados1['Ordem de Serviço'].value_counts()
        if int(repetido) > 1:
            st.markdown(
                alerta + f'<Strong><i>Foram encontradas Ordens de Serviço com números repetidos, exclua da lista a solicitação que não for correspondente a que queira editar</i></Strong></p>',
                unsafe_allow_html=True)
            repeticao = 1
            break
    st.dataframe(dados1[filtrar].head())
    # selecionado=n_solicitacao
    # print(nome[n_solicitacao.index(selecionado)])
    if (1 > 0):  # len(n_solicitacao)

        # procurando status equivalente na lista
        with st.form(key='my_form'):
            status = st.selectbox('Selecione o status:', status)
            obs_email = st.text_area('Observação para o Usuário:', value='')
            obs_interna = st.text_area('Observação Interna:', value='')
            with st.expander("Agendamento"):
                data_ag = datetime.strptime("01-01-2022", '%d-%m-%Y')
                data_agendamento = st.date_input('Data de Agendamento', value=data_ag)
                hora_agendamento = st.selectbox('Hora de agendamento:', horas)
            s = st.text_input("Senha:", value="", type="password")  # , type="password"
            botao = st.form_submit_button('Registrar')
            efetuado = 0
        if (botao == True and s == a):
            with st.spinner('Registrando dados...Aguarde!'):
                for selecionado_i in selecionado:
                    celula = sheet.find(str(selecionado_i))
                    # sheet.update_acell('P' + str(celula.row), status)  # Status
                    # print(sheet.cell(celula.row, 20).value)
                    # print(repeticao)
                    # print(selecionado_i)
                    if (sheet.cell(celula.row, 20).value == selecionado_i and sheet.cell(celula.row,
                                                                                         20).value != '' and repeticao == 0):
                        efetuado = 1
                        sheet.update_acell('W' + str(celula.row), status)  # Status
                        # sheet.update_acell('R' + str(celula.row), '')  # apagar Sim para enviar e-mail
                        if (obsemail != ''):
                            # sheet.update_acell('S' + str(celula.row), obsemail)  # obs_email
                            sheet.update_acell('X' + str(celula.row), obs_email)  # obs_email
                        if (obsinterna != ''):
                            # sheet.update_acell('AA' + str(celula.row), obsinterna)  # obs_interna
                            sheet.update_acell('Y' + str(celula.row), obs_interna)  # obs_interna

                        data_agend = str(data_agendamento.year) + '-' + str(data_agendamento.month) + '-' + str(
                            data_agendamento.day)

                        if (data_agend != '2022-1-1'):
                            data = data_agendamento
                            data_formatada = str(data.day) + '/' + str(data.month) + '/' + str(data.year)
                            sheet.update_acell('AM' + str(celula.row), data_formatada)
                            sheet.update_acell('AN' + str(celula.row), hora_agendamento)

                            # st.markdown(infor+'<b>Registro efetuado!</b></p>',unsafe_allow_html=True)
            if (efetuado == 1):
                st.success('Registro efetuado!')
            elif (efetuado == 0 and repeticao == 1):
                st.error('Remova as OS com números repetidos!')
            elif (efetuado == 1 and repeticao == 1):
                st.error('Dados parcialmente cadastrados! As OS com números repetidos não foram registradas!')
        elif (botao == True and s != a):
            st.markdown(alerta + '<b>Senha incorreta!</b></p>', unsafe_allow_html=True)
    else:
        st.markdown(infor + '<b>Não há itens na condição ' + pg + '</b></p>', unsafe_allow_html=True)
elif pg == 'Consulta':

    # PÁGINA DE CONSULTA ************************************************************************************************

    # dados = sheet.get_all_records()  # Get a list of all records
    # df = pd.DataFrame(dados)
    n_solicitacao.append('')
    nome.append('')
    telefone.append('')
    predio.append('')
    sala.append('')
    data.append('')
    observacao.append('')
    os.append('')
    obsemail.append('')
    obsinterna.append('')
    stat.append('')

    for dic in df.index:
        if df['Prédio'][dic] != '':
            # print(df['Código da UFT'][dic])
            n_solicitacao.append(df['Código da UFT'][dic])
            nome.append(df['Nome do solicitante'][dic])
            telefone.append(df['Telefone'][dic])
            predio.append(df['Prédio'][dic])
            sala.append(df['Sala/Local'][dic])
            data.append(df['DATASOL'][dic])
            observacao.append(df['Descrição sucinta'][dic])
            os.append(df['Ordem de Serviço'][dic])
            obsemail.append(df['Observação p/ Solicitante'][dic])
            obsinterna.append(df['Observação Interna'][dic])
            stat.append(df['Status'][dic])

    st.markdown(cabecalho, unsafe_allow_html=True)
    st.subheader(pg)
    titulos = ['Carimbo de data/hora', 'Endereço de e-mail', 'Nome do solicitante', 'Área de Manutenção',
               'Descrição sucinta', 'Prédio', 'Sala/Local', 'Telefone', 'DATASOL', 'Ordem de Serviço', 'Status',
               'Observação p/ Solicitante', 'Observação Interna', 'Código da UFT']
    with st.form(key='form1'):
        texto = st.text_input('Busca por argumento em qualquer lugar: ')
        col1, col2 = st.columns(2)
        col3, col4 = st.columns(2)
        col5, col6 = st.columns(2)
        filtrar = []

        dados = df[titulos]
        valor = data
        valor = list(dict.fromkeys(valor))  # removendo valores duplicados
        valor = sorted(valor)  # ordenando lista de string
        filtro_data = col1.multiselect('Filtrar por Data:', valor)
        if (len(filtro_data) > 0):
            if (len(filtrar) > 0):
                filtrar = filtrar & dados['DATASOL'].isin(filtro_data)
            else:
                filtrar = dados['DATASOL'].isin(filtro_data)

        valor = os
        valor = list(dict.fromkeys(valor))  # removendo valores duplicados
        valor = sorted(valor)  # ordenando lista de string
        filtro_os = col2.multiselect('Filtrar por Ordem de Serviço:', valor)
        if (len(filtro_os) > 0):
            if (len(filtrar) > 0):
                filtrar = filtrar & dados['Ordem de Serviço'].isin(filtro_os)
            else:
                filtrar = dados['Ordem de Serviço'].isin(filtro_os)
        valor = nome
        valor = list(dict.fromkeys(valor))  # removendo valores duplicados
        valor = sorted(valor)  # ordenando lista de string
        filtro_solicitante = col3.multiselect('Filtrar por Nome do Solicitante:', valor)
        if (len(filtro_solicitante) > 0):
            # filtro_solicitante=valor
            if (len(filtrar) > 0):
                filtrar = filtrar & dados['Nome do solicitante'].isin(filtro_solicitante)
            else:
                filtrar = dados['Nome do solicitante'].isin(filtro_solicitante)

        valor = stat
        valor = list(dict.fromkeys(valor))  # removendo valores duplicados
        valor = sorted(valor)  # ordenando lista de string
        filtro_status = col4.multiselect('Filtrar por Status:', valor)
        if (len(filtro_status) > 0):
            if (len(filtrar) > 0):
                filtrar = filtrar & dados['Status'].isin(filtro_status)
            else:
                filtrar = dados['Status'].isin(filtro_status)

        valor = predio
        valor = list(dict.fromkeys(valor))  # removendo valores duplicados
        valor = sorted(valor)  # ordenando lista de string
        filtro_predio = col5.multiselect('Filtrar por Prédio:', valor)
        if (len(filtro_predio) > 0):
            if (len(filtrar) > 0):
                filtrar = filtrar & dados['Prédio'].isin(filtro_predio)
            else:
                filtrar = dados['Prédio'].isin(filtro_predio)

        valor = sala
        valor = list(dict.fromkeys(valor))  # removendo valores duplicados
        valor = sorted(valor)  # ordenando lista de string
        filtro_sala = col6.multiselect('Filtrar por Sala/Local:', valor)
        if (len(filtro_sala) > 0):
            if (len(filtrar) > 0):
                filtrar = filtrar & dados['Sala/Local'].isin(filtro_sala)
            else:
                filtrar = dados['Sala/Local'].isin(filtro_sala)
        # st.text('Colunas Gráfico')
        # col1,col2=st.columns(2)
        # coluna1 = col1.selectbox('Coluna 1: ',titulos)
        # coluna2 =col2.selectbox('Coluna 2: ', titulos)
        # print(filtro_predio)
        btn1 = st.form_submit_button('Filtrar')
        if (len(filtrar) == 0):
            filtrar = titulos
    if (btn1 == True):
        # dados=df[titulos]
        # filtrar=dados[titulo_coluna].isin([filtro])
        # print(filtrar)
        # if(len(filtrar)>0):
        if (texto != ''):
            
            tit_plan = ['Nome do solicitante','Endereço de e-mail','Carimbo de data/hora','Área de Manutenção','Prédio','Sala/Local','Telefone','Ordem de Serviço','Status','Observação p/ Solicitante','Observação Interna','Descrição sucinta']
            for coluna in tit_plan:
                try:
                    #filtrar = filtrar + list(filter(lambda x: any(substring in x for substring in [texto]), [coluna]))
                    #filtrar = filtrar & dados[coluna].str.contains(texto, na=False)
                    filtrar = filtrar & dados[filtrar][dados[coluna].str.contains(texto, na=False)]
                except:
                    print('pulou')                    
            dad1 = dados[filtrar]#[dados['Descrição sucinta'].str.contains(texto, na=False)]
            #dad2 = dados[filtrar][dados['Carimbo de data/hora'].str.contains(texto, na=False)]
            dad = dad1
        else:
            dad = dados[filtrar]
        st.dataframe(dad)  # dados[filtrar].head()
        df_xlsx = to_excel(dad)
        st.download_button(label='📥 Baixar Resultado do Filtro em Excel', data=df_xlsx,
                           file_name='filtro_planilha.xlsx')
        # dados_graf=pd.DataFrame(dados[filtrar],columns=[coluna1,coluna2])
        # fig = px.bar(dados_graf, x=coluna1, y=coluna2, barmode='group', height=400)
        # st.plotly_chart(fig)
        # plost.line_chart(dados_graf, coluna1, coluna2)

        # else:
        #    st.dataframe(df[titulos])
    else:
        st.dataframe(df[titulos])


