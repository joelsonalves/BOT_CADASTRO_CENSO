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

            status = 19

            # Acessar página de cadastro de turma
            print('Abrir a página de cadastro de turma.')
            page.goto(self.__URL_CADASTRO_DE_TURMA)

            status = 18

            # Preencher o nome da turma
            print('Preencher o nome da turma.')
            page.locator('#input_nomeTurma').fill(self.__nome_da_turma)

            status = 17

            # Selecionar o tipo
            print('Selecionar o tipo.')
            page.locator('#idtipoMediacaoDidaticoPedagogica').select_option(label='Educação a distância - EAD')

            status = 16

            # Marcar a escolarização
            print('Marcar a escolarização.')
            page.locator('#checkbox_idTipoAtendimento0').click()

            status = 15

            # espearar carregamento
            print('Aguardar o seletor de modalidade ficar disponível.')
            page.wait_for_selector('#idModalidadeAno1')

            status = 14

            # Selecionar a modalidade
            print('Selecionar a modalidade.')
            page.locator('#idModalidadeAno1').select_option(label='Educação profissional')

            status = 13

            # Selecionar o filtro da etapa
            print('Selecionar o filtro da etapa.')
            page.locator('#filtroEtapa').select_option(label='Educação Profissional Técnica de Nível Médio - Integrada')

            status = 12

            # Selecionar a etapa
            print('Selecionar a etapa.')
            page.locator('#idEtapa').select_option(label='Curso técnico integrado (ensino médio integrado) não seriada')

            status = 11

            # esperar carregamento
            print('Aguardar o botão de confirmação da etapa ficar disponível.')
            page.wait_for_selector('i.fa.fa-check-circle.fa-2x.text-success.btn.btn-xs')

            status = 10

            # Clicar na confirmação da etapa
            print('Clicar no botão de confirmação da etapa.')
            page.locator('i.fa.fa-check-circle.fa-2x.text-success.btn.btn-xs').click()

            status = 9

            # esperar carregamento
            print('Aguardar o seletor do código do curso ficar disponível.')
            page.wait_for_selector('#idCurso')

            status = 8

            # Selecionar código do curso
            print(f'Selecionar o código do curso: {self.__codigo_do_curso}')
            page.locator('#idCurso').select_option(self.__codigo_do_curso)

            status = 7

            # Clicar em módulos
            print('Marcar a caixa de Módulos.')
            page.locator('#checkbox_idFormaOrganizacaoTurma3').click()

            status = 6

            # esperar carregamento
            print('Aguardar a caixa da área ficar disponível.')
            page.wait_for_selector('#checkbox_outraAreas0')

            status = 5

            # Clicar na área
            print('Marcar a área.')
            page.locator('#checkbox_outraAreas0').click()

            status = 4

            # esperar carregamento
            print('Aguardar o botão Enviar ficar disponível.')
            page.wait_for_selector('button:has-text("Enviar")')

            status = 3

            # Clicar no botão enviar
            print('Clicar no botão Enviar.')
            page.evaluate(''' () => { 
        
                document.querySelector('button[type="submit"]');  

            } ''')

            status = 2

            # esperar 10 segundos
            for i in range(10):
                print(f'Esperando {i + 1} segundo(s).')
                page.wait_for_timeout(1000)

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

                if not falha_critica:
                    page.wait_for_timeout(60 * 1000)
                browser.close()

            except BaseException:

                pass
            
            print('Automação finalizada.')

if __name__ == '__main__':

    Bot.run()