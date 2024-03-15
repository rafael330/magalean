import gspread as gs
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import matplotlib.pyplot as plt
import streamlit as st 

st.set_page_config(layout='wide')
st.title('PENDÊNCIAS e OPORTUNIDADES')

#Tabulação dos Magaleans registrados e pendentes de 2ª aprovação - PLANILHA MAGALEAN AUTOMÁTICO
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credenciais = ServiceAccountCredentials.from_json_keyfile_name('jsonkey_raf.json', scope)
client = gs.authorize(credenciais)

planilha = client.open_by_key('1cGS2ODt2vsyPwF3UBa_geSH0Jp8Mb1XGcbRpKfw3NYg')
aba = planilha.worksheet('Visualização Aprovação')
dados = aba.get_all_records()
df1 = pd.DataFrame(dados)

#Filtro dos Magalean pendentes de 2ª aprovação
df1 = df1[['ID Melhoria','Status 2ª Aprovação','2ª Aprovação']]
df1 = df1[df1['Status 2ª Aprovação'] == 'Aguardando aprovação ENG']

#Tabulação dos Magaleans registrados e pendentes de 2ª aprovação - PLANILHA BASE MAGALEAN
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credenciais = ServiceAccountCredentials.from_json_keyfile_name('jsonkey_raf.json', scope)
client = gs.authorize(credenciais)

planilha = client.open_by_key('1bbnKFSaz6dYwW8KQsxhaocLlUx9_N9g_yMaCf9HUt4Q')
aba = planilha.worksheet('Respostas ao formulário 1')
dados = aba.get_all_records()
df2 = pd.DataFrame(dados)   

#Filtro dos Magalean pendentes de 2ª aprovação
df2 = df2[['ID Melhoria','Nome','Área','Unidade','ANO','Mês','Status Atual']]
df2 = df2[df2['Status Atual'] == 'Aguardando aprovação ENG']

#Mesclando as informações com base no ID da melhoria
df3 = pd.merge(df1,df2, on='ID Melhoria')
df3 = df3[['ID Melhoria','Nome','Área','Unidade','ANO','Mês','2ª Aprovação']]
contagem1 = df3['ID Melhoria'].count()

#Tabulação dos Magaleans registrados e pendentes de aprovação do KPO - PLANILHA MAGALEAN AUTOMÁTICO
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credenciais = ServiceAccountCredentials.from_json_keyfile_name('jsonkey_raf.json', scope)
client = gs.authorize(credenciais)

planilha = client.open_by_key('1cGS2ODt2vsyPwF3UBa_geSH0Jp8Mb1XGcbRpKfw3NYg')
aba = planilha.worksheet('Visualização Aprovação')
dados = aba.get_all_records()
df4 = pd.DataFrame(dados)

#Filtro dos Magalean pendentes de aprovação KPO
df4 = df4[['ID Melhoria','Status Aprovação KPO']]
df4 = df4[df4['Status Aprovação KPO'] == 'Aguardando Aprovação KPO']

#Tabulação dos Magaleans registrados e pendentes de aprovação do KPO - PLANILHA BASE MAGALEAN
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credenciais = ServiceAccountCredentials.from_json_keyfile_name('jsonkey_raf.json', scope)
client = gs.authorize(credenciais)

planilha = client.open_by_key('1bbnKFSaz6dYwW8KQsxhaocLlUx9_N9g_yMaCf9HUt4Q')
aba = planilha.worksheet('Respostas ao formulário 1')
dados = aba.get_all_records()
df5 = pd.DataFrame(dados)   

df5 = df5[['ID Melhoria','Nome','Área','Unidade','ANO','Mês','Status Atual']]

#Mesclando as informações com base no ID da melhoria
df6 = pd.merge(df4,df5, on='ID Melhoria')
df6 = df6[['ID Melhoria','Nome','Área','Unidade','ANO','Mês']]
contagem2 = df6['ID Melhoria'].count()

# Configurações iniciais
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credenciais = ServiceAccountCredentials.from_json_keyfile_name('jsonkey_raf.json', scope)
client = gs.authorize(credenciais)

# Carregando os dados
planilha = client.open_by_key('13l2btUKJ93DLqYjWh1Oe3xbFucm5g_lZr4IQUd9i8B4')
aba = planilha.worksheet('Yellow')
dados = aba.get_all_records()
df_op = pd.DataFrame(dados)

# Preparando os dados
df_op['Ano'] = df_op['Ano'].astype(str)
df_op = df_op[df_op['Ano'].isin(['2023', '2024'])]
df_op = df_op[df_op['Status'].isin(['Falta Kaizen', 'Falta Magalean'])]
df_op = df_op[df_op['Nome'] != '']

# Agrupando os dados
agrupado = df_op.groupby(['Filial'])['ID'].count().reset_index(name='Oportunidades')

#Preparando interface do Streamlit
tab1, tab2 = st.tabs(['Magaleans','Certificados'])

with tab1:
    def style_table(df):
        return df.style.set_table_styles(
            [{'selector': 'th', 'props': [('background-color', '#8A2BE2'), ('color', 'white'), ('text-align', 'center')]}]
        ).set_properties(**{'text-align': 'center'}).hide()

    st.header('Magaleans pendentes de 2ª aprovação')  
    styled_df = style_table(df3)
    st.write(styled_df.to_html(escape=False), unsafe_allow_html=True)

    st.header('Magaleans pendentes aprovação de KPO')    
    styled_df = style_table(df6)
    st.write(styled_df.to_html(escape=False), unsafe_allow_html=True)   

with tab2:    
    col1, col2 = st.columns(2)    
    with col1:
        filiais = agrupado['Filial'].unique().tolist()
        box_filial = st.selectbox('Qual a filial?:',filiais, index=None, placeholder='Selecione a filial')    
        df_filtrado = agrupado[agrupado['Filial'] == box_filial]
        df_select = df_op[['Filial','ID','Nome','Status']]    
        df_select = df_select[df_select['Filial'] == box_filial]
    
        def style_table(df):
            return df.style.set_table_styles(
                [{'selector': 'th', 'props': [('background-color', '#8A2BE2'), ('color', 'white'), ('text-align', 'center')]}] #Linhas para estilizar a planilha exibida
            ).set_properties(**{'text-align': 'center'}).hide()  
        
        st.header('Oportunidade de certificação')
        styled_df = style_table(df_filtrado)
        st.write(styled_df.to_html(escape=False), unsafe_allow_html=True)
        
        st.header('Distribuição por filial')
        st.bar_chart(agrupado, x='Filial', y='Oportunidades', color=["#20B2AA"])
    
    with col2:
        def style_table(df):
            return df.style.set_table_styles(
                [{'selector': 'th', 'props': [('background-color', '#8A2BE2'), ('color', 'white'), ('text-align', 'center')]}] #Linhas para estilizar a planilha exibida
            ).set_properties(**{'text-align': 'center'}).hide()  
        
        st.header('ID Oportunidades')
        styled_df = style_table(df_select)
        st.write(styled_df.to_html(escape=False), unsafe_allow_html=True)
    

def emestudo():
    
    #Criar imagem para juntar as duas tabelas
    #1ª Tabela
    def tabela1():
        axs[0].axis('tight')
        axs[0].axis('off')
        axs[0].set_title('Magaleans pendentes de 2ª aprovação', fontweight='bold', pad=1) #Título da tabela
        table1 = axs[0].table(cellText=df3.values, colLabels=df3.columns, cellLoc='center', loc='center', colColours=['#f2f2f2']*len(df3.columns))
        table1.auto_set_font_size(False)
        table1.set_fontsize(10)

    #2ª Tabela
    def tabela2():
        axs[1].axis('tight')
        axs[1].axis('off')
        axs[1].set_title('Magaleans pendentes de Aprovação do KPO', fontweight='bold', pad=1) #Título da tabela
        table1 = axs[1].table(cellText=df6.values, colLabels=df6.columns, cellLoc='center', loc='center', colColours=['#f2f2f2']*len(df6.columns))
        table1.auto_set_font_size(False)
        table1.set_fontsize(10)

    #Condição para decidir de que forma essas tabelas serão enviadas:
    if contagem1 > 0 and contagem2 > 0:
        fig, axs = plt.subplots(2,1,figsize=(20,5))
        tabela1()
        tabela2()
        fig.subplots_adjust(top=0.25)
        plt.tight_layout()
        plt.savefig('magaleans.png',bbox_inches='tight') #Salvar a imagem dentro do diretório
        #plt.show()
    elif contagem1 > 0 and contagem2 < 1:
        fig, ax = plt.subplots(figsize=(20,5))
        ax.axis('tight')
        ax.axis('off')
        ax.set_title('Magaleans pendentes de 2ª aprovação', fontweight='bold', pad=1) #Título da tabela
        table1 = ax.table(cellText=df3.values, colLabels=df3.columns, cellLoc='center', loc='center', colColours=['#f2f2f2']*len(df3.columns))
        table1.auto_set_font_size(False)
        table1.set_fontsize(10)
        fig.subplots_adjust(top=0.25)
        plt.tight_layout()
        plt.savefig('magaleans.png',bbox_inches='tight') #Salvar a imagem dentro do diretório
        #plt.show()
    elif contagem1 < 1 and contagem2 > 0:
        fig, ax = plt.subplots(figsize=(20,5))
        ax.axis('tight')
        ax.axis('off')
        ax.set_title('Magaleans pendentes de Aprovação do KPO', fontweight='bold', pad=1) #Título da tabela
        table1 = ax.table(cellText=df6.values, colLabels=df6.columns, cellLoc='center', loc='center', colColours=['#f2f2f2']*len(df6.columns))
        table1.auto_set_font_size(False)
        table1.set_fontsize(10)
        fig.subplots_adjust(top=0.25)
        plt.tight_layout()
        plt.savefig('magaleans.png',bbox_inches='tight') #Salvar a imagem dentro do diretório
        #plt.show()

    from selenium import webdriver
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.by import By
    from time import sleep
    import os

    #Incluindo o "@" na frente do nome do KPO da 2ª aprovação para utiliza-lo na mensagem padrão
    df1['2ª Aprovação'] = '@' + df1['2ª Aprovação'].astype(str)
    df6['Nome'] = '@' + df6['Nome'].astype(str)

    if contagem1 > 0 or contagem2 > 0:
        
        chrome_options = webdriver.ChromeOptions()
        path_to_user_data_dir = os.path.join(os.getcwd(), 'ChromeProfile')
        chrome_options.add_argument(f"user-data-dir={path_to_user_data_dir}")
        navegador = webdriver.Chrome(options=chrome_options)
        navegador.get('https://web.whatsapp.com/')
        sleep(10)

        #Mensagem com os nomes dos KPOs e separa com virgulas
        mensagem = ','.join(df1['2ª Aprovação'])
        mensagem1 = ','.join(df6['Nome'])

        while len(navegador.find_elements(By.ID,'side'))<1:
            sleep(5)

        navegador.find_element(By.XPATH,'//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]/p').send_keys('Rafael Rocha', Keys.ENTER)
        sleep(2)
        navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/div/span').click()
        sleep(2)
        navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[2]/li/div/input').send_keys(r'C:\Users\raa_rocha\Desktop\boot_magalean\magaleans.png')
        sleep(2)
        navegador.find_element(By.XPATH,'//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div').click()
        sleep(2)
        if contagem1 > 0:
            navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(f'*Atenção para as análises (2ª aprovações) pendentes:* {mensagem}', Keys.ENTER)
            sleep(2)
        if contagem2 > 0:
            navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(f'*Atenção análises prendentes:* {mensagem1}', Keys.ENTER)
            sleep(2)
        navegador.quit()