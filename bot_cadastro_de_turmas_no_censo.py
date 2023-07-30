from playwright.sync_api import sync_playwright
import pandas as pd

class Bot:

    def __init__(self):

        self.__URL_PRINCIPAL = 'https://censobasico.inep.gov.br/'
        self.__URL_CADASTRO_DE_TURMA = 'https://censobasico.inep.gov.br/censobasico/#/turma/cadastrar'
        self.__PLANILHA = 'turmas.xlsx'
        self.__index = None
        self.__nome_da_turma = None
        self.__codigo_do_curso = None
        self.__etapa = None

    def __acessar_pagina_inicial(self, page):

        page.goto(self.__URL_PRINCIPAL)

    def __verificar_se_ha_turmas_para_cadastrar(self):

        try:

            df = pd.read_excel(self.__PLANILHA, dtype=str)

            quant_linhas = df.loc[df['STATUS'] == '0'].shape[0]

        except Exception:

            quant_linhas = 0

        if quant_linhas > 0:

            return True
        
        return False
    
    def __selecionar_proxima_turma_para_cadastrar(self):

        df = pd.read_excel(self.__PLANILHA, dtype=str)
        df_aux = df.loc[df['STATUS'] == '0'].reset_index()

        self.__index = df_aux.loc[0, 'index']
        self.__nome_da_turma = df_aux.loc[0, 'NOME_DA_TURMA']
        self.__codigo_do_curso = df_aux.loc[0, 'CODIGO_DO_CURSO']
        self.__etapa = df_aux.loc[0, 'ETAPA']

        df = None
        df_aux = None

    def __salvar_status_do_cadastro_da_turma(self, status):

        df = pd.read_excel(self.__PLANILHA, dtype=str)

        df.loc[self.__index, 'STATUS'] = str(status)

        df.to_excel(self.__PLANILHA, index=False)

        df = None
    
    def __verificar_se_navegador_esta_funcional(self, page):

        try:

            page.title()

        except BaseException:

            return False
        
        return True


    def __cadastrar_turma(self, page):

        try:

            status = 22

            # Acessar página de cadastro de turma
            page.goto(self.__URL_CADASTRO_DE_TURMA)

            status = 21

            # Preencher o nome da turma
            page.locator('#input_nomeTurma').fill(self.__nome_da_turma)

            status = 20

            # Selecionar o tipo
            page.locator('#idtipoMediacaoDidaticoPedagogica').select_option(label='Educação a distância - EAD')

            status = 19

            # esperar 3 segundos
            page.wait_for_timeout(3 * 1000)

            # Marcar a escolarização
            page.evaluate(''' () => { 
        
                document.querySelector('#checkbox_idTipoAtendimento0').click();  

            } ''')

            status = 18

            # esperar 3 segundos
            page.wait_for_timeout(3 * 1000)

            status = 17

            # Marcar não se aplica na estrutura curricular
            page.evaluate(''' () => { 
        
                document.querySelector('#checkbox_idEstruturaCurricular2').click();  

            } ''')

            status = 16

            # esperar 3 segundos
            page.wait_for_timeout(3 * 1000)

            status = 15

            # Selecionar a modalidade
            page.locator('#idModalidadeAno1').select_option(label='Educação profissional')

            status = 14

            # Selecionar o filtro da etapa
            page.locator('#filtroEtapa').select_option(label='Educação Profissional Técnica de Nível Médio')

            status = 13

            # Selecionar a etapa
            if (self.__etapa.upper() == 'C'):
                page.locator('#idEtapa').select_option(label='Curso técnico  - concomitante')
            elif (self.__etapa.upper() == 'S'):
                page.locator('#idEtapa').select_option(label='Curso técnico  - subsequente')

            status = 12

            # esperar 3 segundos
            page.wait_for_timeout(3 * 1000)

            status = 11

            # Clicar na confirmação da etapa
            page.locator('i.fa.fa-check-circle.fa-2x.text-success.btn.btn-xs').click()

            status = 10

            # esperar 3 segundos
            page.wait_for_timeout(3 * 1000)

            status = 9

            # Selecionar código do curso
            page.locator('#idCurso').select_option(self.__codigo_do_curso)

            status = 8

            # esperar 3 segundos
            page.wait_for_timeout(3 * 1000)

            status = 7

            # Clicar em módulos
            page.evaluate(''' () => { 
        
                document.querySelector('#checkbox_idFormaOrganizacaoTurma3').click();  

            } ''')

            status = 6

            # esperar 3 segundos
            page.wait_for_timeout(3 * 1000)

            status = 5

            # Marcar áreas do conhecimento profissionalizantes
            page.evaluate(''' () => { 
        
                document.querySelector('#checkbox_outraAreas0').click();  

            } ''')

            status = 4

            # esperar 3 segundos
            page.wait_for_timeout(3 * 1000)

            status = 3

            # Clicar no botão enviar
            page.evaluate(''' () => { 
        
                document.querySelector('button[type="submit"]').click();  

            } ''')

            status = 2

            # esperar 10 segundos
            page.wait_for_timeout(10 * 1000)

            status = 1

        except BaseException:

            pass

        finally:

            self.__salvar_status_do_cadastro_da_turma(status)

        self.__index = None
        self.__nome_da_turma = None
        self.__codigo_do_curso = None
        self.__etapa = None

    def run():

        print('\nAutomação Auxiliar no Cadastro de Turmas no Censo Escolar')

        with sync_playwright() as p:

            browser = p.webkit.launch(headless=False)
            context = browser.new_context()
            context.clear_cookies()
            page = context.new_page()
            bot = Bot()

            falha_critica = False

            try:

                bot.__acessar_pagina_inicial(page)

            except BaseException:

                pass

            input('Faça login no CENSO ESCOLAR, depois tecle [ENTER] para continuar...')
                
            while (bot.__verificar_se_ha_turmas_para_cadastrar() and not falha_critica):
                
                bot.__selecionar_proxima_turma_para_cadastrar()
                
                print(f'Trabalhando na linha nº {bot.__index + 2} da planilha.')

                bot.__cadastrar_turma(page)

                falha_critica = not bot.__verificar_se_navegador_esta_funcional(page)

                if (falha_critica): 
                    
                    print('Houve uma falha crítica.')

            try:

                if not falha_critica:
                    page.wait_for_timeout(60 * 1000)
                browser.close()

            except BaseException:

                pass
            
            print('Automação finalizada.')

if __name__ == '__main__':

    Bot.run()