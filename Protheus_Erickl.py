import time
import pandas as pd
import numpy as np
from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC  
import chromedriver_autoinstaller
import undetected_chromedriver as uc
#chromedriver_autoinstaller.install() 
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pyautogui as pg

# import requests
# criar arquivo de log para seguir a partir da última linha

# Definindo funções para esperar carregamento da página
def esperar_id(elemento):
    WebDriverWait(nav, 20).until(EC.presence_of_element_located((By.ID, elemento)))

def esperar_xpath(elemento):
    WebDriverWait(nav, 20).until(EC.presence_of_element_located((By.XPATH, elemento)))

dtype_dict = {
    'CNPJ': str,  # Define a coluna 'CNPJ' como texto (str)
    'CNPJ FORNECEDOR': str,  
}

# Extraindo DataFrame pelo caminho do arquivo
caminho = r"C:\Users\erickmedeiros\Desktop\REINF_2023.xlsx"
df = pd.read_excel(caminho, dtype= dtype_dict)
# Cópia para salvar as NFs não registradas
df_copy = df.copy()

###### EMPRESA ESCOLHIDA ('SAS','SAE','IS','ATECH','ARCO')######
empresa_escolhida = 'SAE'

# Extraindo mês anterior (competência atual) e formatando datas de emissão e digitação
#mes_anterior = (datetime.now() - relativedelta(months=1)).strftime('%m-%Y')
mes_anterior = '10-2023'
df['DT.EMISSÃO'] = df['DT.EMISSÃO'].dt.strftime('%d-%m-%Y')
df['DT.DIGITAÇÃO'] = df['DT.DIGITAÇÃO'].dt.strftime('%d-%m-%Y')
df['MÊS DE EMISSÃO'] = df['MÊS DE EMISSÃO'].dt.strftime('%m-%Y')
df['MÊS DE DIGITAÇÃO'] = df['MÊS DE DIGITAÇÃO'].dt.strftime('%m-%Y')
 

# Excluindo tudo que não foi digitado no mês anterior (TESTAR)
df = df.drop(df[df['MÊS DE DIGITAÇÃO'] != mes_anterior].index)
df = df.drop(df[df['MÊS DE EMISSÃO'] != mes_anterior].index)
df = df.drop(df[df['EMPRESA'] != empresa_escolhida].index)
# Resetando os índices das linhas
df = df.reset_index(drop=True)

# Criando coluna vazia para sinalizar de a nota foi preenchida ou não
df['LANÇADA'] = np.nan

# Corrigindo CNPJ
#df['CNPJ'] = df['CNPJ'].astype(str)
#df['CNPJ FORNECEDOR'] = df['CNPJ FORNECEDOR'].astype(str)

def format_cnpj(CNPJ):
    # Verifica se o CNPJ tem 14 dígitos e adiciona um 0 no início
    if len(CNPJ) == 13:
        CNPJ = '0' + CNPJ
    return CNPJ

df['CNPJ'] = df['CNPJ'].apply(format_cnpj)
df['CNPJ FORNECEDOR'] = df['CNPJ FORNECEDOR'].apply(format_cnpj)

# Agrupando filial por empresas  
cnpj_por_empresas = df.groupby('EMPRESA')['CNPJ'].unique().to_dict()
# Agrupando fornecedores por filial em um dicionário
fornecedores_por_filial = df.groupby('CNPJ')['CNPJ FORNECEDOR'].unique().to_dict()
# Agrupando fornecedores por mês
#fornecedores_por_mes = df.groupby('MÊS DE EMISSÃO')['CNPJ FORNECEDOR'].unique().to_dict()

# Lista de fornecedores por serviço
limpeza= ['69039154000193', '02936838000117', '10987099000110', '78570397000144', '17874775000199', '26316214000165', '26951054000126', '01711083000190']
vigilancia = ['04808914000134']
transporte = ['12158137000158', '15225033000107', '11820004000132']

### INÍCIO DA AUTOMATIZAÇÃO ##
#1. Entrar no eCAC 
nav = webdriver.Chrome()
nav.get('https://cav.receita.fazenda.gov.br/autenticacao/')
nav.maximize_window()
#2. Clica em "Logar por gov.br"
esperar_id("login-dados-certificado")
nav.find_element(By.ID, "login-dados-certificado").click()

#3. Clica em "Logar por certificado digital" (Escolher manualmente)
esperar_id("cert-digital")
nav.find_element(By.ID, "cert-digital").click()

'''
time.sleep(3)
coord = pg.locateCenterOnScreen("cert digital login.jpg", confidence= 0.8)  
pg.leftClick(x= coord[0], y= coord[1], duration=1)

imagens_certificado = {'ATECH': r"C:\Users\erickmedeiros\Desktop\.py\imagens\Etapa de Login por Cert Digital\ATECH.png",
                       'ARCO' : r"C:\Users\erickmedeiros\Desktop\.py\imagens\Etapa de Login por Cert Digital\ARCO.png",
                       'IS'   : r"C:\Users\erickmedeiros\Desktop\.py\imagens\Etapa de Login por Cert Digital\IS.png",
                       'SAS'  : r"C:\Users\erickmedeiros\Desktop\.py\imagens\Etapa de Login por Cert Digital\CBE.png",
                       'SAE'  : r"C:\Users\erickmedeiros\Desktop\.py\imagens\Etapa de Login por Cert Digital\SAE.png",
                       'EI'   : r"C:\Users\erickmedeiros\Desktop\.py\imagens\Etapa de Login por Cert Digital\EI.png"}
time.sleep(2)
caminho_imagem = imagens_certificado[empresa_escolhida]

certificado_encontrado = False
while not certificado_encontrado: #Looping para encontrar o cert digital correspondente
    try:
        coord = pg.locateCenterOnScreen(caminho_imagem, confidence= 0.8)  
        pg.leftClick(x= coord[0], y= coord[1], duration=1)
        certificado_encontrado = True
        print('Certificado encontrado')
    except Exception as e:
        print(f'Erro: {e}, tentanto novamente')
        pg.press('down', presses= 1)
        time.sleep(1)'''
        

coord  = pg.locateCenterOnScreen(r"C:\Users\erickmedeiros\Desktop\.py\imagens\Etapa de Login por Cert Digital\OK.png", confidence= 0.8)
pg.leftClick(x= coord[0], y= coord[1], duration=1)

time.sleep(2)

cont = 0 

#4. Ir para REINF
nav.get('https://www3.cav.receita.fazenda.gov.br/reinfweb/#/2010/lista')
time.sleep(2)

for cnpj in set(df['CNPJ']): # For A: para cada CNPJ da empresa
    #if cnpj in cnpj_por_empresas[empresa_escolhida]: # certificar que o CNPJ específico é da empresa escolhida para rodar a automatização
        print(f'CNPJ Empresa: {cnpj}')
        for cnpj_forn in set(df['CNPJ FORNECEDOR']): # For B: para cada fornecedor
            # Verificar se houveram NFs desse fornecedor na filial específica do For A & se o fornecedor do For B teve NF emitidas na competência atual
            if cnpj_forn in fornecedores_por_filial[cnpj]:
                i = 0 #contador para escolher a opção certa de incluir serviço
                print(f"------------------preenchendo fornecedor: {cnpj_forn}----------------------")
                time.sleep(2)
                nav.get('https://www3.cav.receita.fazenda.gov.br/reinfweb/#/2010/lista')
                #5. Clicar em "Incluir" e preencher competência
                time.sleep(1)
                esperar_xpath('/html/body/app-root/div/div[3]/app-evento2010-lista-pesquisa/div/button')
                nav.find_element(By.XPATH, '/html/body/app-root/div/div[3]/app-evento2010-lista-pesquisa/div/button').click() 
                esperar_id("periodo_apuracao")
                nav.find_element(By.ID, "periodo_apuracao").send_keys(str(mes_anterior).replace('-','/'))
                
                # Informações do tomador de serviços
                esperar_id('tipo_inscricao_estabelecimento')
                nav.find_element(By.ID, 'tipo_inscricao_estabelecimento').send_keys('1')

                nav.find_element(By.XPATH, '/html/body/app-root/div/div[3]/app-evento2010-formulario/app-reinf-versao-leiaute/app-evento2010-inclusao-chave/form/div[1]/fieldset/div/div[1]/app-reinf-campo-formulario/div/div[2]/input').send_keys(cnpj)
                # Informações do prestador de serviço
                esperar_id('cnpj_prestador')
                nav.find_element(By.ID, 'cnpj_prestador').send_keys(cnpj_forn)
                # Clicar 'Continuar'
                nav.find_element(By.XPATH, '/html/body/app-root/div/div[3]/app-evento2010-formulario/app-reinf-versao-leiaute/app-evento2010-inclusao-chave/form/div[2]/button[1]').click() 
                esperar_xpath('//*[@id="indicativo_obra"]')
                nav.find_element(By.XPATH, '//*[@id="indicativo_obra"]').send_keys('0')
                nav.find_element(By.ID, 'indicativo_cprb').send_keys('0')
                for j, nf in enumerate(df['NUMERO']): # For C: para cada NF da tabela original
                    # Verificar se 
                    if (df['CNPJ'][j] == cnpj) and (df['CNPJ FORNECEDOR'][j] == cnpj_forn):
                        i+=1
                        print(f"    preenchendo nota: {nf} {i}")
                        #INFO NF
                        nav.find_element(By.XPATH, '/html/body/app-root/div/div[3]/app-evento2010-formulario/app-reinf-versao-leiaute/form/fieldset[3]/app-reinf-linha-titulo-inclusao[1]/div/div/div/span').click()
                        esperar_id('serie')
                        nav.find_element(By.ID, 'serie').send_keys('1')
                        nav.find_element(By.ID, 'numero_documento').send_keys(nf)
                        nav.find_element(By.ID, 'data_emissao_nf').send_keys(str(df['DT.EMISSÃO'][j]).replace('-','/'))
                        nav.find_element(By.ID, 'valor_bruto').send_keys('{:.2f}'.format(df['VALOR'][j]))
                        nav.find_element(By.XPATH, '/html/body/app-root/div/div[3]/app-evento2010-formulario/app-evento2010-modal-nfs/app-reinf-versao-leiaute/app-reinf-modal/div/div/div[3]/div/button[1]').click()
                        
                        # INFO SERVIÇO
                        # Escolhendo serviço com base no fornecedor
                        if i == 1:
                            id_servico = '/html/body/app-root/div/div[3]/app-evento2010-formulario/app-reinf-versao-leiaute/form/fieldset[3]/app-reinf-collapse/details/div/app-reinf-linha-titulo-inclusao/div/div/div/span'
                        else:
                            id_servico = f'/html/body/app-root/div/div[3]/app-evento2010-formulario/app-reinf-versao-leiaute/form/fieldset[3]/app-reinf-collapse[{i}]/details/div/app-reinf-linha-titulo-inclusao/div/div/div/span'
                        esperar_xpath(id_servico)
                        nav.find_element(By.XPATH, id_servico).click()
                        esperar_xpath('//*[@id="tipo_servico"]')
                        time.sleep(1)
                        if cnpj_forn in limpeza:
                            nav.find_element(By.XPATH,'//*[@id="tipo_servico"]').send_keys('100000001 - Limpeza, conservação ou zeladoria')
                        elif cnpj_forn in vigilancia:
                            nav.find_element(By.XPATH,'//*[@id="tipo_servico"]').send_keys('100000002 - Vigilância ou segurança')
                        elif cnpj_forn in transporte:
                            nav.find_element(By.XPATH,'//*[@id="tipo_servico"]').send_keys('100000024 - Operação de transporte de passageiros')
                        else:
                            nav.find_element(By.XPATH,'//*[@id="tipo_servico"]').send_keys('100000031 - Trabalho temporário na forma da Lei nº 6.019, de janeiro de 1974')
                        
                        nav.find_element(By.ID, 'valor_base_ret').send_keys('{:.2f}'.format(df['INSS'][j]/0.11))
                        nav.find_element(By.ID, 'valor_retencao').send_keys('{:.2f}'.format(df['INSS'][j]))
                        nav.find_element(By.XPATH, '/html/body/app-root/div/div[3]/app-evento2010-formulario/app-evento2010-modal-info-tp-serv/app-reinf-versao-leiaute/app-reinf-modal/div/div/div[3]/div/button[1]').click()
                    
                        df.loc[j, 'LANÇADA'] = 'Sim'
                        cont += 1
                        print(cont)

                # Salvar NFs do fornecedor (como rascunho por enquanto)
                esperar_xpath('/html/body/app-root/div/div[3]/app-evento2010-formulario/app-reinf-versao-leiaute/form/app-reinf-botoes-formulario/div/div/button[1]')
                nav.find_element(By.XPATH, '/html/body/app-root/div/div[3]/app-evento2010-formulario/app-reinf-versao-leiaute/form/app-reinf-botoes-formulario/div/div/button[1]').click()
                print('Salvar Rascunho')
                
                '''if cont == 1:
                    esperar_xpath('/html/body/app-root/div/div[3]/app-evento2010-formulario/app-reinf-versao-leiaute/form/app-reinf-botoes-formulario/div/div/button[2]')
                    nav.find_element(By.XPATH, '/html/body/app-root/div/div[3]/app-evento2010-formulario/app-reinf-versao-leiaute/form/app-reinf-botoes-formulario/div/div/button[2]').click()
                    time.sleep(3)
                    esperar_xpath('/html/body/app-root/div/app-reinf-mensagens-alerta/div/div/a')
                    nav.find_element(By.XPATH, '/html/body/app-root/div/app-reinf-mensagens-alerta/div/div/a').click()
                    time.sleep(3)
                    nav.switch_to.window(nav.window_handles[1])
                    esperar_xpath('//*[@id="details-button"]')
                    nav.find_element(By.XPATH, '//*[@id="details-button"]').click()
                    esperar_xpath('//*[@id="proceed-link"]')
                    nav.find_element(By.XPATH, '//*[@id="proceed-link"]').click()
                    time.sleep(3)
                    nav.close()
                    nav.switch_to.window(nav.window_handles[0])
                    time.sleep(1)
                time.sleep(1)
                esperar_xpath('/html/body/app-root/div/div[3]/app-evento2010-formulario/app-reinf-versao-leiaute/form/app-reinf-botoes-formulario/div/div/button[2]')
                nav.find_element(By.XPATH, '/html/body/app-root/div/div[3]/app-evento2010-formulario/app-reinf-versao-leiaute/form/app-reinf-botoes-formulario/div/div/button[2]').click()
                time.sleep(5)
                esperar_xpath('/html/body/app-root/div/div[3]/app-evento4020-totalizador/div/button[2]')
                print('Concluir e Enviar')'''
                time.sleep(2)
nav.quit()
                        

# Imprimir um Excel com as não preenchidas (espera-se que sejam apenas as de fora da competência)
df.to_excel(f'REINF_não_preenchidas {empresa_escolhida} {mes_anterior}.xlsx', index=False)
