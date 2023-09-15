from playwright.sync_api import sync_playwright
import pandas as pd

TENTATIVAS = 2

TEMPO_DE_ESPERA_REDUZIDO = 5
TEMPO_DE_ESPERA_PADRAO = 10
TEMPO_DE_ESPERA_ESTENDIDO = 15

VINCULAR_ESTUDANTE_PELO_CPF = 1
DEFINIR_LOCALIZACAO_DIFERENCIADA_PELO_CPF = 2
AJUSTAR_3ANO_COCOMITANTE_PELO_CPF = 3
AJUSTAR_TURMA_PELO_IDENTIFICADOR = 4
DEFINIR_LOCALIZACAO_DIFERENCIADA_PELO_IDENTIFICADOR = 5
DEFINIR_MUNICIPIO_DE_NASCIMENTO_PELO_IDENTIFICADOR = 6
EXTRAIR_DADOS_PESSOAIS_PELO_IDENTIFICADOR = 7

lista_de_opcoes = [
    VINCULAR_ESTUDANTE_PELO_CPF, 
    DEFINIR_LOCALIZACAO_DIFERENCIADA_PELO_CPF,
    AJUSTAR_3ANO_COCOMITANTE_PELO_CPF,
    AJUSTAR_TURMA_PELO_IDENTIFICADOR,
    DEFINIR_LOCALIZACAO_DIFERENCIADA_PELO_IDENTIFICADOR,
    DEFINIR_MUNICIPIO_DE_NASCIMENTO_PELO_IDENTIFICADOR,
    EXTRAIR_DADOS_PESSOAIS_PELO_IDENTIFICADOR
]

sisacad_xlsx = 'estudantes_sisacad.xlsx'
censo_xlsx = 'estudantes_censo.xlsx'

def desenhar_tela_inicial():
    print('\n#####################################################')
    print('######## ROBÔ AUXILIAR NO CENSO ESCOLAR 2022 ########')
    print('#####################################################')

def escolher_opcao_de_execucao():
    while True:
        print('\n######## MENU DE OPÇÕES ########')
        print('1: Vincular estudante pelo CPF.')
        print('2: Definir localização diferenciada, caso necessário, consultando pelo CPF.')
        print('3: Ajustar estudante do 3º ano, no ano anterior, para subsequente, pelo CPF.')
        print('4: Ajustar turma do CENSO pelo identificador.')
        print('5: Definir localização diferenciada, caso necessário, consultando pelo identificador.')
        print('6: Ajustar município de nascimento, consultando pelo identificador.')
        print('7: Extrair dados pessoais, consultando pelo identificador.')
        try:
            opcao = int(input('Digite o número da opção desejada e em seguida tecle [ENTER]: '))
        except BaseException:
            print('Opção inválida. Tente novamente.')
        else:
            if opcao in lista_de_opcoes:
                break
            print('Opção inválida. Tente novamente.')
    return opcao

def verificar_necessidade_de_vincular(sisacad_xlsx):
    df = pd.read_excel(sisacad_xlsx, dtype=str)
    if df.loc[df['STATUS_VINCULO'] == '0'].shape[0] > 0:
        df = None
        return True
    df = None
    return False
    
def proximo_estudante_para_vincular(sisacad_xlsx):
    df = pd.read_excel(sisacad_xlsx, dtype=str)
    estudante = {
        'index': None,
        'cpf': None,
        'turma': None,
        'status': None
    }
    for i in df.index:
        if df.loc[i, 'STATUS_VINCULO'] == '0':
            estudante['index'] = i
            estudante['cpf'] = str(df.loc[i, 'CPF']).replace('"','').zfill(11)
            estudante['turma'] = str(df.loc[i, 'TURMA'])
            estudante['status'] = str(df.loc[i, 'STATUS_VINCULO'])
            break
    df = None
    return estudante

def salvar_status_do_estudante_em_relacao_ao_vinculo(sisacad_xlsx, estudante):
    if estudante['index'] != None:
        df = pd.read_excel(sisacad_xlsx, dtype=str)
        df.loc[estudante['index'], 'STATUS_VINCULO'] = estudante['status']
        df.to_excel(sisacad_xlsx, index=False)
        df = None
        return True
    return False

def acessar_pagina_inicial(page):
    page.goto('http://censobasico.inep.gov.br/censobasico/')

def acessar_pagina_de_cadastro(page):
    page.goto('http://censobasico.inep.gov.br/censobasico/#/aluno/pesquisar')

def preencher_cpf(page, cpf):
    page.locator('#cpf').fill(cpf)

def clicar_em_pesquisar(page):
    page.locator('button:has-text("Pesquisar")').click()

def clicar_em_vincular(page):
    page.locator('button:has-text("Vincular")').click()

def selecionar_turma(page, turma):
    page.locator('#nomeTurma').select_option(label=turma)

def clicar_em_enviar(page):
    page.locator('button:has-text("Enviar")').click()

def recarregar_pagina(page):
    page.reload()

def definir_tempo_de_espera(page, segundos):
    page.set_default_timeout(segundos * 1000)

def esperar(page, segundos):
    page.wait_for_timeout(segundos * 1000)

def desvincular(page):
    page.locator('strong:has-text("Vínculo") >> nth=0').click()
    page.locator('button:has-text("Desvincular")').click()

def verificar_se_estudante_nao_foi_localizado(page):
    try:
        page.wait_for_selector('legend:has-text("Nenhum registro encontrado. Realize uma nova pesquisa utilizando outros campos, como nome e data de nascimento ou nome e filiação")')
    except BaseException:
        return False
    else:
        return True

# VINCULAR_ESTUDANTE_PELO_CPF
def executar_vinculacao_de_estudantes_no_censo():

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        
        tentativa_cont = 1
        tentativa_cpf = None
        
        try:
            definir_tempo_de_espera(page, TEMPO_DE_ESPERA_ESTENDIDO)
            acessar_pagina_inicial(page)
        except BaseException:
            pass

        #Aguardar login humano
        input('Após fazer login, tecle [Enter]: ')

        while verificar_necessidade_de_vincular(sisacad_xlsx):

            estudante = proximo_estudante_para_vincular(sisacad_xlsx)
            
            if tentativa_cpf != estudante['cpf']:
                tentativa_cpf = estudante['cpf']
                tentativa_cont = 1
            
            print(f'''{estudante['cpf']} - Tentantiva {tentativa_cont}/{TENTATIVAS}''')

            try:

                estudante['status'] = '2'
                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_ESTENDIDO)
                acessar_pagina_de_cadastro(page)

                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_PADRAO)
                preencher_cpf(page, estudante['cpf'])
                clicar_em_pesquisar(page)

                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_REDUZIDO)

                try:
                    desvincular(page)
                except BaseException:
                    pass
                else:
                    definir_tempo_de_espera(page, TEMPO_DE_ESPERA_PADRAO)
                    preencher_cpf(page, estudante['cpf'])
                    clicar_em_pesquisar(page)
                    definir_tempo_de_espera(page, TEMPO_DE_ESPERA_REDUZIDO)
                
                estudante['status'] = '3'
                clicar_em_vincular(page)

                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_PADRAO)

                estudante['status'] = '4'
                selecionar_turma(page, estudante['turma'])

                estudante['status'] = '5'
                clicar_em_enviar(page)

                estudante['status'] = '1'
        
            except BaseException:
                if estudante['status'] == '2':
                    print(f'Falha no acesso ao site. Aguardando {TEMPO_DE_ESPERA_PADRAO} segundos para tentar novamente...')
                    try:
                        esperar(page, TEMPO_DE_ESPERA_PADRAO)
                        recarregar_pagina(page)
                    except BaseException:
                        break
                    else:
                        continue

                if not (estudante['status'] == '3' and verificar_se_estudante_nao_foi_localizado(page)) and tentativa_cont < TENTATIVAS:
                    tentativa_cont += 1
                    continue
            else:
                esperar(page, TEMPO_DE_ESPERA_REDUZIDO)
            
            print(f'''{estudante['cpf']} - Status: {estudante['status']}''')
            salvar_status_do_estudante_em_relacao_ao_vinculo(sisacad_xlsx, estudante)

        if not verificar_necessidade_de_vincular(sisacad_xlsx):
            print('####### TRABALHO CONCLUÍDO #######')
        else:
            print('!!!!!!! AINDA HÁ ESTUDANTES PARA CADASTRAR !!!!!!!')

        try:
            esperar(page, TEMPO_DE_ESPERA_ESTENDIDO)
        except BaseException:
            pass

        try:
            context.close()
            browser.close()
        except BaseException:
            pass

def verificar_necessidade_de_definir_localizacao(sisacad_xlsx):
    df = pd.read_excel(sisacad_xlsx, dtype=str)
    if df.loc[(df['STATUS_LOCALIZACAO'] == '0') & (df['STATUS_VINCULO'] == '1')].shape[0] > 0 :
        df = None
        return True
    df = None
    return False

def proximo_estudante_para_definir_localizacao(sisacad_xlsx):
    df = pd.read_excel(sisacad_xlsx, dtype=str)
    estudante = {
        'index': None,
        'cpf': None,
        'status': None,
        'status_loc': None
    }
    for i in df.index:
        if df.loc[i, 'STATUS_VINCULO'] == '1' and df.loc[i, 'STATUS_LOCALIZACAO'] == '0':
            estudante['index'] = i
            estudante['cpf'] = str(df.loc[i, 'CPF']).replace('"','').zfill(11)
            estudante['status'] = str(df.loc[i, 'STATUS_VINCULO'])
            estudante['status_loc'] = str(df.loc[i, 'STATUS_LOCALIZACAO'])
            break
    df = None
    return estudante

def salvar_status_do_estudante_em_relacao_a_localizacao(sisacad_xlsx, estudante):
    if estudante['index'] != None:
        df = pd.read_excel(sisacad_xlsx, dtype=str)
        df.loc[estudante['index'], 'STATUS_LOCALIZACAO'] = estudante['status_loc']
        df.to_excel(sisacad_xlsx, index=False)
        df = None
        return True
    return False

def clicar_em_editar_dados_pessoais(page):
    page.locator('button:has-text("Editar dados pessoais")').click()

def aguardar_select_definir_localizacao_diferenciada(page):
    page.wait_for_selector('#idLocalizacaoDiferenciada')

def verificar_se_select_definir_localizacao_diferenciada_esta_vazio(page):
    value = page.evaluate(''' () => { return document.querySelector('#idLocalizacaoDiferenciada').value } ''')
    return value == ''

def selecionar_localizacao_diferenciada(page):
    page.locator('#idLocalizacaoDiferenciada').select_option(label='Não está em área de localização diferenciada')
    
# DEFINIR_LOCALIZACAO_DIFERENCIADA_PELO_CPF
def executar_definicao_de_localizacao_diferenciada_de_estudantes_no_censo():

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        
        tentativa_cont = 1
        tentativa_cpf = None
        
        try:
            definir_tempo_de_espera(page, TEMPO_DE_ESPERA_ESTENDIDO)
            acessar_pagina_inicial(page)
        except BaseException:
            pass

        #Aguardar login humano
        input('\nApós fazer login, tecle [Enter]: ')

        while verificar_necessidade_de_definir_localizacao(sisacad_xlsx):

            estudante = proximo_estudante_para_definir_localizacao(sisacad_xlsx)
            
            if tentativa_cpf != estudante['cpf']:
                tentativa_cpf = estudante['cpf']
                tentativa_cont = 1
            
            print(f'''{estudante['cpf']} - Tentantiva {tentativa_cont}/{TENTATIVAS}''')

            try:

                estudante['status_loc'] = '2'
                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_ESTENDIDO)
                acessar_pagina_de_cadastro(page)

                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_PADRAO)
                preencher_cpf(page, estudante['cpf'])
                clicar_em_pesquisar(page)
                
                estudante['status_loc'] = '3'
                clicar_em_editar_dados_pessoais(page)

                estudante['status_loc'] = '4'
                aguardar_select_definir_localizacao_diferenciada(page)

                estudante['status_loc'] = '5'
                if verificar_se_select_definir_localizacao_diferenciada_esta_vazio(page):
                    estudante['status_loc'] = '6'
                    selecionar_localizacao_diferenciada(page)

                estudante['status_loc'] = '7'
                clicar_em_enviar(page)

                estudante['status_loc'] = '1'
        
            except BaseException:
                if estudante['status_loc'] == '2':
                    print(f'Falha no acesso ao site. Aguardando {TEMPO_DE_ESPERA_PADRAO} segundos para tentar novamente...')
                    try:
                        esperar(page, TEMPO_DE_ESPERA_PADRAO)
                        recarregar_pagina(page)
                    except BaseException:
                        break
                    else:
                        continue

                if tentativa_cont < TENTATIVAS:
                    tentativa_cont += 1
                    continue
            else:
                esperar(page, TEMPO_DE_ESPERA_REDUZIDO)
            
            print(f'''{estudante['cpf']} - Status: {estudante['status_loc']}''')
            salvar_status_do_estudante_em_relacao_a_localizacao(sisacad_xlsx, estudante)
        
        if not verificar_necessidade_de_definir_localizacao(sisacad_xlsx):
            print('####### TRABALHO CONCLUÍDO #######')
        else:
            print('!!!!!!! AINDA HÁ ESTUDANTES PARA DEFINIR LOCALIZAÇÃO !!!!!!!')

        try:
            esperar(page, TEMPO_DE_ESPERA_ESTENDIDO)
        except BaseException:
            pass

        try:
            context.close()
            browser.close()
        except BaseException:
            pass

def verificar_necessidade_de_ajustar_para_subsequente(censo_xlsx):
    df = pd.read_excel(censo_xlsx, dtype=str)
    if df.loc[(df['STATUS_3ANO'] == '0')].shape[0] > 0 :
        df = None
        return True
    df = None
    return False

def proximo_estudante_para_ajustar_para_subsequente(censo_xlsx):
    df = pd.read_excel(censo_xlsx, dtype=str)
    estudante = {
        'index': None,
        'identificador': None,
        'turma': None,
        'status_3ano': None
    }
    for i in df.index:
        if df.loc[i, 'STATUS_3ANO'] == '0':
            estudante['index'] = i
            estudante['identificador'] = str(df.loc[i, 'IDENTIFICADOR']).replace(' ','')
            estudante['turma'] = str(df.loc[i, 'TURMA'])
            estudante['status_3ano'] = str(df.loc[i, 'STATUS_3ANO'])
            break
    df = None
    return estudante

def salvar_status_do_estudante_em_relacao_ao_3ano(censo_xlsx, estudante):
    if estudante['index'] != None:
        df = pd.read_excel(censo_xlsx, dtype=str)
        df.loc[estudante['index'], 'STATUS_3ANO'] = estudante['status_3ano']
        df.to_excel(censo_xlsx, index=False)
        df = None
        return True
    return False

def preencher_identificador(page, identificador):
    page.locator('#input_identificacaoUnica').fill(identificador)

def clicar_em_visualizar_vinculos_ano_anterior(page):
    page.locator('text=Visualizar vínculos ano anterior').click()

def clicar_no_segundo_vinculo(page):
    page.locator('strong:has-text("Vínculo") >> nth=1').click()

def verificar_se_estudante_era_turma_terminal(page):
    try:
        page.wait_for_selector('text=Ensino médio - 3ª Série')
    except BaseException:
        pass
    else:
        return True
    try:
        page.wait_for_selector('text=EJA - ensino médio')
    except BaseException:
        pass
    else:
        return True
    return False


# AJUSTAR_3ANO_COCOMITANTE_PELO_CPF
def executar_ajuste_estudante_3ano_e_eja_cocomitante_para_subsequente():

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        
        tentativa_cont = 1
        tentativa_identificador = None
        
        try:
            definir_tempo_de_espera(page, TEMPO_DE_ESPERA_ESTENDIDO)
            acessar_pagina_inicial(page)
        except BaseException:
            pass

        #Aguardar login humano
        input('\nApós fazer login, tecle [Enter]: ')

        while verificar_necessidade_de_ajustar_para_subsequente(censo_xlsx):

            estudante = proximo_estudante_para_ajustar_para_subsequente(censo_xlsx)
            
            if tentativa_identificador != estudante['identificador']:
                tentativa_identificador = estudante['identificador']
                tentativa_cont = 1
            
            print(f'''{estudante['identificador']} - Tentantiva {tentativa_cont}/{TENTATIVAS}''')

            try:

                estudante['status_3ano'] = '2'
                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_ESTENDIDO)
                acessar_pagina_de_cadastro(page)

                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_PADRAO)
                #preencher_cpf(page, estudante['cpf'])
                preencher_identificador(page, estudante['identificador'])
                clicar_em_pesquisar(page)
                
                estudante['status_3ano'] = '3'
                #clicar_em_editar_dados_pessoais(page)
                clicar_em_visualizar_vinculos_ano_anterior(page)

                estudante['status_3ano'] = '4'
                #aguardar_select_definir_localizacao_diferenciada(page)
                clicar_no_segundo_vinculo(page)

                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_REDUZIDO)

                estudante['status_3ano'] = '5'
                if verificar_se_estudante_era_turma_terminal(page):
                    estudante['status_3ano'] = '6'
                    #desvincular(page)

                    try:
                        desvincular(page)
                    except BaseException:
                        pass
                    else:
                        definir_tempo_de_espera(page, TEMPO_DE_ESPERA_PADRAO)
                        preencher_identificador(page, estudante['identificador'])
                        clicar_em_pesquisar(page)
                        definir_tempo_de_espera(page, TEMPO_DE_ESPERA_REDUZIDO)

                estudante['status_3ano'] = '7'
                clicar_em_vincular(page)

                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_PADRAO)

                estudante['status_3ano'] = '8'
                selecionar_turma(page, estudante['turma'])

                estudante['status_3ano'] = '9'
                clicar_em_enviar(page)

                estudante['status_3ano'] = '1'
        
            except BaseException:
                if estudante['status_3ano'] == '2':
                    print(f'Falha no acesso ao site. Aguardando {TEMPO_DE_ESPERA_PADRAO} segundos para tentar novamente...')
                    try:
                        esperar(page, TEMPO_DE_ESPERA_PADRAO)
                        recarregar_pagina(page)
                    except BaseException:
                        break
                    else:
                        continue

                if tentativa_cont < TENTATIVAS:
                    tentativa_cont += 1
                    continue
            else:
                esperar(page, TEMPO_DE_ESPERA_REDUZIDO)
            
            print(f'''{estudante['identificador']} - Status: {estudante['status_3ano']}''')
            salvar_status_do_estudante_em_relacao_ao_3ano(censo_xlsx, estudante)
        
        if not verificar_necessidade_de_ajustar_para_subsequente(censo_xlsx):
            print('####### TRABALHO CONCLUÍDO #######')
        else:
            print('!!!!!!! AINDA HÁ ESTUDANTES PARA AJUSTAR PARA SUBSEQUENTE !!!!!!!')

        try:
            esperar(page, TEMPO_DE_ESPERA_ESTENDIDO)
        except BaseException:
            pass

        try:
            context.close()
            browser.close()
        except BaseException:
            pass

def verificar_necessidade_de_ajustar_turma_pelo_identificador(censo_xlsx):
    df = pd.read_excel(censo_xlsx, dtype=str)
    if df.loc[(df['STATUS_TURMA'] == '0')].shape[0] > 0 :
        df = None
        return True
    df = None
    return False

def proximo_estudante_para_ajustar_turma_pelo_identificador(censo_xlsx):
    df = pd.read_excel(censo_xlsx, dtype=str)
    estudante = {
        'index': None,
        'identificador': None,
        'turma': None,
        'status_turma': None
    }
    for i in df.index:
        if df.loc[i, 'STATUS_TURMA'] == '0':
            estudante['index'] = i
            estudante['identificador'] = str(df.loc[i, 'IDENTIFICADOR']).replace(' ','')
            estudante['turma'] = str(df.loc[i, 'TURMA'])
            estudante['status_turma'] = str(df.loc[i, 'STATUS_TURMA'])
            break
    df = None
    return estudante

def salvar_status_do_estudante_em_relacao_ao_ajuste_de_turma(censo_xlsx, estudante):
    if estudante['index'] != None:
        df = pd.read_excel(censo_xlsx, dtype=str)
        df.loc[estudante['index'], 'STATUS_TURMA'] = estudante['status_turma']
        df.to_excel(censo_xlsx, index=False)
        df = None
        return True
    return False

#  AJUSTAR_TURMA_PELO_IDENTIFICADOR
def executar_ajuste_de_turma_pelo_identificador():

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        
        tentativa_cont = 1
        tentativa_identificador = None
        
        try:
            definir_tempo_de_espera(page, TEMPO_DE_ESPERA_ESTENDIDO)
            acessar_pagina_inicial(page)
        except BaseException:
            pass

        #Aguardar login humano
        input('\nApós fazer login, tecle [Enter]: ')

        while verificar_necessidade_de_ajustar_turma_pelo_identificador(censo_xlsx):

            estudante = proximo_estudante_para_ajustar_turma_pelo_identificador(censo_xlsx)
            
            if tentativa_identificador != estudante['identificador']:
                tentativa_identificador = estudante['identificador']
                tentativa_cont = 1
            
            print(f'''{estudante['identificador']} - Tentantiva {tentativa_cont}/{TENTATIVAS}''')

            try:

                estudante['status_turma'] = '2'
                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_ESTENDIDO)
                acessar_pagina_de_cadastro(page)

                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_PADRAO)
                preencher_identificador(page, estudante['identificador'])
                clicar_em_pesquisar(page)
                
                #estudante['status_turma'] = '3'
                #clicar_em_visualizar_vinculos_ano_anterior(page)

                #estudante['status_turma'] = '4'
                #clicar_no_segundo_vinculo(page)

                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_REDUZIDO)

                estudante['status_turma'] = '5'
                try:
                    desvincular(page)
                except BaseException:
                    pass
                else:
                    definir_tempo_de_espera(page, TEMPO_DE_ESPERA_PADRAO)
                    preencher_identificador(page, estudante['identificador'])
                    clicar_em_pesquisar(page)
                    definir_tempo_de_espera(page, TEMPO_DE_ESPERA_REDUZIDO)

                estudante['status_turma'] = '6'
                clicar_em_vincular(page)

                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_PADRAO)

                estudante['status_turma'] = '7'
                selecionar_turma(page, estudante['turma'])

                estudante['status_turma'] = '8'
                clicar_em_enviar(page)

                estudante['status_turma'] = '1'
        
            except BaseException:
                if estudante['status_turma'] == '2':
                    print(f'Falha no acesso ao site. Aguardando {TEMPO_DE_ESPERA_PADRAO} segundos para tentar novamente...')
                    try:
                        esperar(page, TEMPO_DE_ESPERA_PADRAO)
                        recarregar_pagina(page)
                    except BaseException:
                        break
                    else:
                        continue

                if tentativa_cont < TENTATIVAS:
                    tentativa_cont += 1
                    continue
            else:
                esperar(page, TEMPO_DE_ESPERA_REDUZIDO)
            
            print(f'''{estudante['identificador']} - Status: {estudante['status_turma']}''')
            salvar_status_do_estudante_em_relacao_ao_ajuste_de_turma(censo_xlsx, estudante)
        
        if not verificar_necessidade_de_ajustar_turma_pelo_identificador(censo_xlsx):
            print('####### TRABALHO CONCLUÍDO #######')
        else:
            print('!!!!!!! AINDA HÁ ESTUDANTES PARA AJUSTAR TURMA NA LISTA !!!!!!!')

        try:
            esperar(page, TEMPO_DE_ESPERA_ESTENDIDO)
        except BaseException:
            pass

        try:
            context.close()
            browser.close()
        except BaseException:
            pass

def verificar_necessidade_de_definir_localizacao_pelo_identificador(censo_xlsx):
    df = pd.read_excel(censo_xlsx, dtype=str)
    if df.loc[(df['STATUS_LOCALIZACAO'] == '0') & (df['STATUS_TURMA'] == '1')].shape[0] > 0 :
        df = None
        return True
    df = None
    return False

def proximo_estudante_para_definir_localizacao_pelo_identificador(censo_xlsx):
    df = pd.read_excel(censo_xlsx, dtype=str)
    estudante = {
        'index': None,
        'identificador': None,
        'status_turma': None,
        'status_loc': None
    }
    for i in df.index:
        if df.loc[i, 'STATUS_TURMA'] == '1' and df.loc[i, 'STATUS_LOCALIZACAO'] == '0':
            estudante['index'] = i
            estudante['identificador'] = str(df.loc[i, 'IDENTIFICADOR'])
            estudante['status_turma'] = str(df.loc[i, 'STATUS_TURMA'])
            estudante['status_loc'] = str(df.loc[i, 'STATUS_LOCALIZACAO'])
            break
    df = None
    return estudante

def salvar_status_do_estudante_em_relacao_a_localizacao_pelo_identificador(censo_xlsx, estudante):
    if estudante['index'] != None:
        df = pd.read_excel(censo_xlsx, dtype=str)
        df.loc[estudante['index'], 'STATUS_LOCALIZACAO'] = estudante['status_loc']
        df.to_excel(censo_xlsx, index=False)
        df = None
        return True
    return False

# DEFINIR_LOCALIZACAO_DIFERENCIADA_PELO_IDENTIFICADOR
def executar_definicao_de_localizacao_diferenciada_de_estudantes_no_censo_pelo_identificador():

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        
        tentativa_cont = 1
        tentativa_identificador = None
        
        try:
            definir_tempo_de_espera(page, TEMPO_DE_ESPERA_ESTENDIDO)
            acessar_pagina_inicial(page)
        except BaseException:
            pass

        #Aguardar login humano
        input('\nApós fazer login, tecle [Enter]: ')

        while verificar_necessidade_de_definir_localizacao_pelo_identificador(censo_xlsx):

            estudante = proximo_estudante_para_definir_localizacao_pelo_identificador(censo_xlsx)
            
            if tentativa_identificador != estudante['identificador']:
                tentativa_identificador = estudante['identificador']
                tentativa_cont = 1
            
            print(f'''{estudante['identificador']} - Tentantiva {tentativa_cont}/{TENTATIVAS}''')

            try:

                estudante['status_loc'] = '2'
                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_ESTENDIDO)
                acessar_pagina_de_cadastro(page)

                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_PADRAO)
                preencher_identificador(page, estudante['identificador'])
                clicar_em_pesquisar(page)
                
                estudante['status_loc'] = '3'
                clicar_em_editar_dados_pessoais(page)

                estudante['status_loc'] = '4'
                aguardar_select_definir_localizacao_diferenciada(page)

                estudante['status_loc'] = '5'
                if verificar_se_select_definir_localizacao_diferenciada_esta_vazio(page):
                    estudante['status_loc'] = '6'
                    selecionar_localizacao_diferenciada(page)

                esperar(page, TEMPO_DE_ESPERA_REDUZIDO)

                estudante['status_loc'] = '7'
                clicar_em_enviar(page)

                estudante['status_loc'] = '1'
        
            except BaseException:
                if estudante['status_loc'] == '2':
                    print(f'Falha no acesso ao site. Aguardando {TEMPO_DE_ESPERA_PADRAO} segundos para tentar novamente...')
                    try:
                        esperar(page, TEMPO_DE_ESPERA_PADRAO)
                        recarregar_pagina(page)
                    except BaseException:
                        break
                    else:
                        continue

                if tentativa_cont < TENTATIVAS:
                    tentativa_cont += 1
                    continue
            else:
                esperar(page, TEMPO_DE_ESPERA_REDUZIDO)
            
            print(f'''{estudante['identificador']} - Status: {estudante['status_loc']}''')
            salvar_status_do_estudante_em_relacao_a_localizacao_pelo_identificador(censo_xlsx, estudante)
        
        if not verificar_necessidade_de_definir_localizacao_pelo_identificador(censo_xlsx):
            print('####### TRABALHO CONCLUÍDO #######')
        else:
            print('!!!!!!! AINDA HÁ ESTUDANTES PARA DEFINIR LOCALIZAÇÃO !!!!!!!')

        try:
            esperar(page, TEMPO_DE_ESPERA_ESTENDIDO)
        except BaseException:
            pass

        try:
            context.close()
            browser.close()
        except BaseException:
            pass

def verificar_necessidade_de_ajuste_no_municipio_pelo_identificador(censo_xlsx):
    df = pd.read_excel(censo_xlsx, dtype=str)
    if df.loc[(df['STATUS_MUNICIPIO'] == '0')].shape[0] > 0 :
        df = None
        return True
    df = None
    return False

def proximo_estudante_para_ajustar_municipio_pelo_identificador(censo_xlsx):
    df = pd.read_excel(censo_xlsx, dtype=str)
    estudante = {
        'index': None,
        'identificador': None,
        'municipio': None,
        'status_municipio': None
    }
    for i in df.index:
        if df.loc[i, 'STATUS_MUNICIPIO'] == '0':
            estudante['index'] = i
            estudante['identificador'] = str(df.loc[i, 'IDENTIFICADOR'])
            estudante['municipio'] = str(df.loc[i, 'MUNICIPIO'])
            estudante['status_municipio'] = str(df.loc[i, 'STATUS_MUNICIPIO'])
            break
    df = None
    return estudante

def salvar_status_do_estudante_em_relacao_ao_municipio(censo_xlsx, estudante):
    if estudante['index'] != None:
        df = pd.read_excel(censo_xlsx, dtype=str)
        df.loc[estudante['index'], 'STATUS_MUNICIPIO'] = estudante['status_municipio']
        df.to_excel(censo_xlsx, index=False)
        df = None
        return True
    return False

def clicar_em_editar_identificacao(page):
    page.locator('button:has-text("Editar identificação")').click()

def aguardar_select_municipio_nascimento(page):
    page.wait_for_selector('#municipioNascimento')

def selecionar_municipio_nascimento(page, estudante):
    page.locator('#municipioNascimento').select_option(label=estudante['municipio'])

# DEFINIR_MUNICIPIO_DE_NASCIMENTO_PELO_IDENTIFICADOR
def executar_ajuste_no_municipio_pelo_identificador():

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        
        tentativa_cont = 1
        tentativa_identificador = None
        
        try:
            definir_tempo_de_espera(page, TEMPO_DE_ESPERA_ESTENDIDO)
            acessar_pagina_inicial(page)
        except BaseException:
            pass

        #Aguardar login humano
        input('\nApós fazer login, tecle [Enter]: ')

        while verificar_necessidade_de_ajuste_no_municipio_pelo_identificador(censo_xlsx):

            estudante = proximo_estudante_para_ajustar_municipio_pelo_identificador(censo_xlsx)
            
            if tentativa_identificador != estudante['identificador']:
                tentativa_identificador = estudante['identificador']
                tentativa_cont = 1
            
            print(f'''{estudante['identificador']} - Tentantiva {tentativa_cont}/{TENTATIVAS}''')

            try:

                estudante['status_loc'] = '2'
                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_ESTENDIDO)
                acessar_pagina_de_cadastro(page)

                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_PADRAO)
                preencher_identificador(page, estudante['identificador'])
                clicar_em_pesquisar(page)
                
                estudante['status_loc'] = '3'
                clicar_em_editar_identificacao(page)

                estudante['status_loc'] = '4'
                aguardar_select_municipio_nascimento(page)

                estudante['status_loc'] = '5'
                selecionar_municipio_nascimento(page, estudante)

                esperar(page, TEMPO_DE_ESPERA_REDUZIDO)

                estudante['status_loc'] = '6'
                clicar_em_enviar(page)

                estudante['status_loc'] = '1'
        
            except BaseException:
                if estudante['status_loc'] == '2':
                    print(f'Falha no acesso ao site. Aguardando {TEMPO_DE_ESPERA_PADRAO} segundos para tentar novamente...')
                    try:
                        esperar(page, TEMPO_DE_ESPERA_PADRAO)
                        recarregar_pagina(page)
                    except BaseException:
                        break
                    else:
                        continue

                if tentativa_cont < TENTATIVAS:
                    tentativa_cont += 1
                    continue
            else:
                esperar(page, TEMPO_DE_ESPERA_REDUZIDO)
            
            print(f'''{estudante['identificador']} - Status: {estudante['status_loc']}''')
            salvar_status_do_estudante_em_relacao_ao_municipio(censo_xlsx, estudante)
        
        if not verificar_necessidade_de_ajuste_no_municipio_pelo_identificador(censo_xlsx):
            print('####### TRABALHO CONCLUÍDO #######')
        else:
            print('!!!!!!! AINDA HÁ ESTUDANTES PARA AJUSTAR O MUNICÍPIO DE NASCIMENTO !!!!!!!')

        try:
            esperar(page, TEMPO_DE_ESPERA_ESTENDIDO)
        except BaseException:
            pass

        try:
            context.close()
            browser.close()
        except BaseException:
            pass

def verificar_necessidade_de_extracao_de_dados_pessoais_pelo_identificador(censo_xlsx):
    df = pd.read_excel(censo_xlsx, dtype=str)
    if df.loc[(df['STATUS_DADOS_PESSOAIS'] == '0')].shape[0] > 0 :
        df = None
        return True
    df = None
    return False

def proximo_estudante_para_extracao_de_dados_pessoais_pelo_identificador(censo_xlsx):
    df = pd.read_excel(censo_xlsx, dtype=str)
    estudante = {
        'index': None,
        'identificador': None,
        'nome': None,
        'cpf': None,
        'data_de_nascimento': None,
        'estado_de_nascimento': None,
        'municipio_de_nascimento': None,
        'filiacao1': None,
        'filiacao2': None,
        'status_dados_pessoais': None
    }
    for i in df.index:
        if df.loc[i, 'STATUS_DADOS_PESSOAIS'] == '0':
            estudante['index'] = i
            estudante['identificador'] = str(df.loc[i, 'IDENTIFICADOR'])
            estudante['nome'] = str(df.loc[i, 'NOME'])
            estudante['cpf'] = str(df.loc[i, 'CPF'])
            estudante['data_de_nascimento'] = str(df.loc[i, 'DATA_DE_NASCIMENTO'])
            estudante['estado_de_nascimento'] = str(df.loc[i, 'ESTADO_DE_NASCIMENTO'])
            estudante['municipio_de_nascimento'] = str(df.loc[i, 'MUNICIPIO_DE_NASCIMENTO'])
            estudante['filiacao1'] = str(df.loc[i, 'FILIACAO1'])
            estudante['filiacao2'] = str(df.loc[i, 'FILIACAO2'])
            estudante['status_dados_pessoais'] = str(df.loc[i, 'STATUS_DADOS_PESSOAIS'])
            break
    df = None
    return estudante

def salvar_dados_pessoais_do_estudante(censo_xlsx, estudante):
    if estudante['index'] != None:
        df = pd.read_excel(censo_xlsx, dtype=str)
        df.loc[estudante['index'], 'NOME'] = estudante['nome']
        df.loc[estudante['index'], 'CPF'] = estudante['cpf']
        df.loc[estudante['index'], 'DATA_DE_NASCIMENTO'] = estudante['data_de_nascimento']
        df.loc[estudante['index'], 'ESTADO_DE_NASCIMENTO'] = estudante['estado_de_nascimento']
        df.loc[estudante['index'], 'MUNICIPIO_DE_NASCIMENTO'] = estudante['municipio_de_nascimento']
        df.loc[estudante['index'], 'FILIACAO1'] = estudante['filiacao1']
        df.loc[estudante['index'], 'FILIACAO2'] = estudante['filiacao2']
        df.loc[estudante['index'], 'STATUS_DADOS_PESSOAIS'] = estudante['status_dados_pessoais']
        df.to_excel(censo_xlsx, index=False)
        df = None
        return True
    return False

def extrair_dados_pessoais(page):

    texto = page.evaluate(''' () => {
                          
        return document.querySelectorAll('div.well.well-sm')[3].innerText;
                          
    } ''')

    lista = texto.split('\n')

    proximo = len(lista) - 1

    while (proximo > 0):
        lista.pop(proximo)
        proximo -= 2

    return lista

# EXTRAIR_DADOS_PESSOAIS_PELO_IDENTIFICADOR
def executar_extracao_de_dados_pessoais_pelo_identificador():

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        
        tentativa_cont = 1
        tentativa_identificador = None
        
        try:
            definir_tempo_de_espera(page, TEMPO_DE_ESPERA_ESTENDIDO)
            acessar_pagina_inicial(page)
        except BaseException:
            pass

        #Aguardar login humano
        input('\nApós fazer login, tecle [Enter]: ')

        while verificar_necessidade_de_extracao_de_dados_pessoais_pelo_identificador(censo_xlsx):

            estudante = proximo_estudante_para_extracao_de_dados_pessoais_pelo_identificador(censo_xlsx)
            
            if tentativa_identificador != estudante['identificador']:
                tentativa_identificador = estudante['identificador']
                tentativa_cont = 1
            
            print(f'''{estudante['identificador']} - Tentantiva {tentativa_cont}/{TENTATIVAS}''')

            try:

                estudante['status_dados_pessoais'] = '2'

                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_ESTENDIDO)
                acessar_pagina_de_cadastro(page)

                definir_tempo_de_espera(page, TEMPO_DE_ESPERA_PADRAO)
                preencher_identificador(page, estudante['identificador'])
                clicar_em_pesquisar(page)

                estudante['status_dados_pessoais'] = '3'

                esperar(page, TEMPO_DE_ESPERA_REDUZIDO)

                lista_de_dados_pessoais = extrair_dados_pessoais(page)

                estudante['status_dados_pessoais'] = '4'

                estudante['nome'] = lista_de_dados_pessoais[0]
                estudante['cpf'] = lista_de_dados_pessoais[1]
                estudante['data_de_nascimento'] = lista_de_dados_pessoais[2]
                estudante['estado_de_nascimento']  = lista_de_dados_pessoais[3]
                estudante['municipio_de_nascimento']  = lista_de_dados_pessoais[4]
                estudante['filiacao1'] = lista_de_dados_pessoais[5]
                estudante['filiacao2'] = lista_de_dados_pessoais[6]
                
                estudante['status_dados_pessoais'] = '1'
        
            except BaseException:
                if estudante['status_dados_pessoais'] == '2':
                    print(f'Falha no acesso ao site. Aguardando {TEMPO_DE_ESPERA_PADRAO} segundos para tentar novamente...')
                    try:
                        esperar(page, TEMPO_DE_ESPERA_PADRAO)
                        recarregar_pagina(page)
                    except BaseException:
                        break
                    else:
                        continue

                if tentativa_cont < TENTATIVAS:
                    tentativa_cont += 1
                    continue
            else:
                esperar(page, TEMPO_DE_ESPERA_REDUZIDO)
            
            print(f'''{estudante['identificador']} - Status: {estudante['status_dados_pessoais']}''')
            salvar_dados_pessoais_do_estudante(censo_xlsx, estudante)
        
        if not verificar_necessidade_de_extracao_de_dados_pessoais_pelo_identificador(censo_xlsx):
            print('####### TRABALHO CONCLUÍDO #######')
        else:
            print('!!!!!!! AINDA HÁ ESTUDANTES PARA EXTRAÇÃO DE DADOS PESSOAIS !!!!!!!')

        try:
            esperar(page, TEMPO_DE_ESPERA_ESTENDIDO)
        except BaseException:
            pass

        try:
            context.close()
            browser.close()
        except BaseException:
            pass

# Carregamento da aplicação
if __name__ == '__main__':

    desenhar_tela_inicial()
    opcao = escolher_opcao_de_execucao()

    if opcao == VINCULAR_ESTUDANTE_PELO_CPF:
        executar_vinculacao_de_estudantes_no_censo() 

    elif opcao == DEFINIR_LOCALIZACAO_DIFERENCIADA_PELO_CPF:
        executar_definicao_de_localizacao_diferenciada_de_estudantes_no_censo()

    elif opcao == AJUSTAR_3ANO_COCOMITANTE_PELO_CPF:
        executar_ajuste_estudante_3ano_e_eja_cocomitante_para_subsequente()

    elif opcao == AJUSTAR_TURMA_PELO_IDENTIFICADOR:
        executar_ajuste_de_turma_pelo_identificador()

    elif opcao == DEFINIR_LOCALIZACAO_DIFERENCIADA_PELO_IDENTIFICADOR:
        executar_definicao_de_localizacao_diferenciada_de_estudantes_no_censo_pelo_identificador()
    
    elif opcao == DEFINIR_MUNICIPIO_DE_NASCIMENTO_PELO_IDENTIFICADOR:
        executar_ajuste_no_municipio_pelo_identificador()

    elif opcao == EXTRAIR_DADOS_PESSOAIS_PELO_IDENTIFICADOR:
        executar_extracao_de_dados_pessoais_pelo_identificador()
    
