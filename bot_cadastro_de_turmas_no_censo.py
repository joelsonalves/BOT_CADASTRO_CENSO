from playwright.sync_api import sync_playwright
import pandas as pd

class Bot:

    def __init__(self):

        self.__URL_PRINCIPAL = 'https://censobasico.inep.gov.br/'
        self.__URL_CADASTRO_DE_TURMA = 'https://censobasico.inep.gov.br/censobasico/#/turma/cadastrar'
        self.__PLANILHA = 'planilha_de_cadastro_de_turmas_no_censo.xlsx'
        self.__index = None
        self.__nome_da_turma = None
        self.__codigo_do_curso = None

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

            status = 16

            # Acessar página de cadastro de turma
            page.goto(self.__URL_CADASTRO_DE_TURMA)

            status = 15

            # Preencher o nome da turma
            page.locator('#input_nomeTurma').fill(self.__nome_da_turma)

            status = 14

            # Selecionar o tipo
            page.locator('#idtipoMediacaoDidaticoPedagogica').select_option(label='Educação a distância - EAD')

            status = 13

            # Marcar a escolarização
            page.locator('#checkbox_idTipoAtendimento0').click()

            status = 12

            # espearar carregamento
            page.wait_for_selector('#idModalidadeAno1')

            status = 11

            # Selecionar a modalidade
            page.locator('#idModalidadeAno1').select_option(label='Educação profissional')

            status = 10

            # Selecionar o filtro da etapa
            page.locator('#filtroEtapa').select_option(label='Educação Profissional Técnica de Nível Médio - Integrada')

            status = 9

            # Selecionar a etapa
            page.locator('#idEtapa').select_option(label='Curso técnico integrado (ensino médio integrado) não seriada')

            status = 8

            # espearar carregamento
            page.wait_for_selector('i.fa.fa-check-circle.fa-2x.text-success.btn.btn-xs')

            status = 7

            # Clicar na confirmação da etapa
            page.locator('i.fa.fa-check-circle.fa-2x.text-success.btn.btn-xs').click()

            status = 6

            # Selecionar código do curso
            page.locator('#idCurso').select_option(label=f'{self.__codigo_do_curso}')

            status = 5

            # Clicar em módulos
            page.locator('#checkbox_idFormaOrganizacaoTurma3').click()

            status = 4

            # Clicar em outras áreas de conhecimento
            page.locator('#checkbox_outraAreas0').click()

            status = 3

            # Clicar no botão enviar
            page.locator('button[type="submit"]')

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

    def run():

        print('\nAutomação Auxiliar no Cadastro de Turmas no Censo Escolar')

        with sync_playwright() as p:

            browser = p.chromium.launch(headless=False)
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
                
                print(f'Trabalhando na linha nº {bot.__index + 1} da planilha.')

                bot.__cadastrar_turma(page)

                falha_critica = not bot.__verificar_se_navegador_esta_funcional(page)

                if (falha_critica): 
                    
                    print('Houve uma falha crítica.')

            try:

                browser.close()

            except BaseException:

                pass
            
            print('Automação finalizada.')

if __name__ == '__main__':

    Bot.run()