import pandas as pd #Para ler arquivos e transformar em dataframes
from selenium import webdriver #O navegador - antigo
from webdriver_manager.chrome import ChromeDriverManager #O navegador -novo
from selenium.webdriver.chrome.service import Service #Para chamar o navegador
from selenium.webdriver.common.by import By #Achar os elementos
from selenium.webdriver.common.keys import Keys #Para digitar no teclado na web
from selenium.webdriver.support.ui import WebDriverWait #Foi necessário para o captcha
from selenium.webdriver.support import expected_conditions as EC #Foi necessário para o captcha
from selenium.webdriver.common.action_chains import ActionChains #Para clicar e navegar em submenus
from selenium.webdriver.support.select import Select #Para preencher droplists
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import time #Para colocar wait
from datetime import datetime
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter.ttk import Combobox


chromedrive_path = ''
chrome = ''
df = '' #dataframe que sera utilizado quando importar a planilha
nome_do_arquivo = ''

def iniciar_driver():
    global chromedrive_path
    global chrome
    # chromedrive_path = 'C:\\Users\Max\Google Drive\Max\CWMC\chromedriver.exe'
    # chrome = webdriver.Chrome(executable_path=chromedrive_path)
    s=Service(ChromeDriverManager().install())
    chrome = webdriver.Chrome(service=s)


def logar():
    user_max = tela_principal.nome.get()
    pass_max = tela_principal.senha.get()
    url_nfp = 'https://www.nfp.fazenda.sp.gov.br'
    chrome.get(url_nfp)
    elemento_texto_user = WebDriverWait(chrome, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="UserName"]')))
    elemento_texto_pass = WebDriverWait(chrome, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Password"]')))
    elemento_texto_user.send_keys(user_max)
    elemento_texto_pass.send_keys(pass_max)
    
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False) 
    WebDriverWait(chrome, 10).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR,"iframe[name^='a-'][src^='https://www.google.com/recaptcha/api2/anchor?']")))
    WebDriverWait(chrome, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[@id='recaptcha-anchor']"))).click()
            
    tela_principal.inserirResult('Verifique ou resolva o Captha na tela do navegador e clique em Ok...')
    captcha_validado = messagebox.askokcancel(title = "Resolver Captcha", message = 'Resolva o Captcha na tela do navegador e clique em Ok...')
    if captcha_validado == True:
        time.sleep(1)
        chrome.switch_to.default_content()
        try:
            elemento_botao_login = WebDriverWait(chrome, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnLogin"]')))
            elemento_botao_login.click()
            time.sleep(1)
            return True
        except TimeoutException:
            WebDriverWait(chrome, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menuSuperior"]/ul/li[4]/a')))
            return True

    else:
        tela_principal.inserirResult('Você cancelou a resolução do Captcha.')
        return False

def navegar():
    acao = ActionChains(chrome)
    elemento_ent = WebDriverWait(chrome, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menuSuperior"]/ul/li[4]/a')))
    acao.move_to_element(elemento_ent).perform()
    elemento_cad_cupom = WebDriverWait(chrome, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menuSuperior:submenu:12"]/li[1]/a')))
    acao.move_to_element(elemento_cad_cupom).click().perform()

    elemento_pagina_bt_ok = WebDriverWait(chrome, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_ConteudoPagina_btnOk"]')))
    elemento_pagina_bt_ok.click()
    
    elemento_caixa = Select(WebDriverWait(chrome, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ddlEntidadeFilantropica"]'))))
    cod_projeto = tela_principal.switcher_projeto()
    elemento_caixa.select_by_value(cod_projeto)
    #<option value="90107">LIGA DAS SENHORAS CATOLICAS DE SAO PAULO</option>
    #<option value="32523">PROJETO CASULO</option>
    
    elemento_pagina_bt_nova_nf = WebDriverWait(chrome, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_ConteudoPagina_btnNovaNota"]')))
    elemento_pagina_bt_nova_nf.click()
    time.sleep(1)
    
    chrome.find_element(By.TAG_NAME, value='body').send_keys(Keys.ESCAPE)


def lancar_notas():    
    for i, j in df.iterrows():
        
        t = i + 1
        statusTotal = f'Processando registro {t} de {statusTotal}...'
        tela_principal.lblTotal['text'] = statusTotal

        if j.Status == 'Processado':
            continue
        else:
            try:
                #elemento_num_nf = chrome.find_element(By.XPATH, '//*[contains(@title, "Digite ou Utilize")]')
                elemento_num_nf = WebDriverWait(chrome, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[contains(@title, "Digite ou Utilize")]')))
                elemento_num_nf.send_keys(Keys.DELETE)
                elemento_num_nf.send_keys(Keys.DELETE)
                elemento_num_nf.send_keys(str(j.Num))
                #elemento_salva_nf = chrome.find_element(By.XPATH, '//*[@id="btnSalvarNota"]')
                elemento_salva_nf = WebDriverWait(chrome, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnSalvarNota"]')))
                #time.sleep(1)
                elemento_salva_nf.click()
                time.sleep(1)
                time_now = datetime.now()
                time_now_formated = time_now.strftime("%Y-%m-%d %H:%M:%S")
                logExec = open('logExe.txt', 'a')
                try:
                    elemento_sucesso = chrome.find_element(By.XPATH, '//*[@id="lblInfo"]')
                    if 'registrada com sucesso' in elemento_sucesso.text:
                        elem_sucess_msg_sucess = time_now_formated + ' | Sucesso! | ' + 'Msg: ' + elemento_sucesso.text + ' | ' + str(j.Num)
                        print(elem_sucess_msg_sucess, file = logExec)
                        df.at[i, 'Msg']= elemento_sucesso.text
                        tela_principal.inserirResult(elem_sucess_msg_sucess)
                    else:
                        elem_sucess_msg_erro = time_now_formated + ' | Erro! | ' + 'Msg: ' + elemento_sucesso.text + ' | ' + str(j.Num) 
                        print(elem_sucess_msg_erro, file = logExec)
                        df.at[i, 'Msg']= elemento_sucesso.text
                        tela_principal.inserirResult(elem_sucess_msg_erro)
                except NoSuchElementException:
                    try:
                        elemento_erro = chrome.find_element(By.XPATH, '//*[@id="lblErro"]')
                        elem_erro_msg_erro = time_now_formated + ' | Erro! | ' + 'Msg: ' + elemento_erro.text + ' | ' + str(j.Num)
                        print(elem_erro_msg_erro, file = logExec)
                        df.at[i, 'Msg']= elemento_erro.text
                        tela_principal.inserirResult(elem_erro_msg_erro)
                    except:
                        elem_erro_msg_gen = time_now_formated + ' | Erro! | ' + ' Não foi possível cadastrar, erro não mapeado. | ' + str(j.Num)
                        print(elem_erro_msg_gen, file = logExec)
                        df.at[i, 'Msg']= elemento_erro.text
                        tela_principal.inserirResult(elem_erro_msg_gen)
                df.at[i, 'Status']= 'Processado'
                
            except:
                erro_fora_tela = 'Faltam registros a serem processados, execute novamente o mesmo arquivo!'
                print(erro_fora_tela, file = logExec)
                tela_principal.inserirResult(erro_fora_tela)
                break
                

    df.to_excel(nome_do_arquivo, index=False)
    tela_principal.inserirResult('Fim. Resultados dos processamentos inseridos na planilha em:')
    tela_principal.inserirResult(nome_do_arquivo)
    
    print("Fim!!!!!!!")
    chrome.quit()


class Tela:
    def __init__(self, master=None):
        root.title('NFP - Lançamento automatizado')

        self.fontePadrao = ("Arial", "11")

        self.containerTitulos = Frame(master)
        self.containerTitulos["pady"] = 10
        self.containerTitulos.pack()

        self.containerBotoesPlanilhas = Frame(master)
        self.containerBotoesPlanilhas["pady"] = 20
        self.containerBotoesPlanilhas["padx"] = 20
        self.containerBotoesPlanilhas.pack()

        #self.fr_containerTitLogin = LabelFrame(self.containerTotal, text='ae', pady=5)
        #self.fr_containerTitLogin.pack(expand=YES, fill=BOTH)

        self.containerTitulosLogin = Frame(master)
        self.containerTitulosLogin["pady"] = 10
        self.containerTitulosLogin.pack()

        self.containerUser = Frame(master)
        self.containerUser["padx"] = 20
        self.containerUser.pack()

        self.containerSenha = Frame(master)
        self.containerSenha["padx"] = 20
        self.containerSenha.pack()

        self.containerMenuProjeto = Frame(master)
        self.containerMenuProjeto["pady"] = 10
        self.containerMenuProjeto.pack()

        self.containerIniciar = Frame(master)
        self.containerIniciar["pady"] = 10
        self.containerIniciar.pack()

        self.containerTotal = Frame(master)
        self.containerTotal.pack(expand=NO, fill=X)
        self.containerTotal["padx"] = 10

        self.containerResult = Frame(master)
        self.containerResult.pack()
        self.containerResult["padx"] = 10
        
        self.containerFim = Frame(master)
        self.containerFim.pack(pady= 5)

        self.titulo = Label(self.containerTitulos, text="Liga Solidária", font = ("Arial", "12", "bold"))
        self.titulo2 = Label(self.containerTitulos, text="- Lançamento automatizado de NFP -", font = self.fontePadrao)
        self.titulo3 = Label(self.containerTitulosLogin, text="Informe seu nome de usuário e senha do site:" + "\n" "https://www.nfp.fazenda.sp.gov.br", font= self.fontePadrao)
        self.titulo.pack()
        self.titulo2.pack()
        self.titulo3["pady"] = 10
        self.titulo3.pack()
        
        self.userLabel = Label(self.containerUser,text="Usuário", font=self.fontePadrao)
        self.userLabel.pack(side=LEFT,)

        self.nome = Entry(self.containerUser)
        self.nome.insert(0, '')
        self.nome["width"] = 30
        self.nome["font"] = self.fontePadrao
        self.nome.pack(side=LEFT)
        
        self.senhaLabel = Label(self.containerSenha, text="Senha", font=self.fontePadrao)
        self.senhaLabel.pack(side=LEFT)

        self.senha = Entry(self.containerSenha)
        self.senha.insert(0, '')
        self.senha["width"] = 30
        self.senha["font"] = self.fontePadrao
        self.senha["show"] = "*"
        self.senha.pack(side=LEFT)

        self.btBaixar_modelo = Button(self.containerBotoesPlanilhas)
        self.btBaixar_modelo["text"] = "Baixar Planilha Modelo"
        self.btBaixar_modelo["font"] = self.fontePadrao
        self.btBaixar_modelo["width"] = 24
        self.btBaixar_modelo["command"] = self.baixarModelo
        self.btBaixar_modelo.pack(side=LEFT)
        
        self.btImportarPlan = Button(self.containerBotoesPlanilhas)
        self.btImportarPlan["text"] = "Importar Planilha de Notas"
        self.btImportarPlan["font"] = self.fontePadrao
        self.btImportarPlan["width"] = 24
        self.btImportarPlan["command"] = self.importarPlan
        self.btImportarPlan.pack(side=RIGHT)

        dict_projetos = ['-- Selecione o projeto --', 'Projeto Liga Solidária', 'Projeto Casulo']
        self.menuProjetos = StringVar()
        self.menuProjetos = Combobox(self.containerMenuProjeto, textvariable=self.menuProjetos, state='readonly')
        self.menuProjetos['values'] = dict_projetos
        self.menuProjetos.current(0)
        self.menuProjetos["font"] = self.fontePadrao
        self.menuProjetos.pack()

        self.btLogin = Button(self.containerIniciar)
        self.btLogin["text"] = "Login"
        self.btLogin["font"] = self.fontePadrao
        self.btLogin["width"] = 24
        self.btLogin["command"] = self.iniciar
        self.btLogin.pack()

        self.lblTotal = Label(self.containerTotal, text="Total de registros importados: ", font= ('Arial', '9'))
        self.lblTotal.pack(side = LEFT, expand = NO, fill = NONE )

        #Definindo a lista de resultado com barra de rolagem
        scroll_bary = Scrollbar(self.containerResult)
        scroll_bary.pack(side = RIGHT, fill = Y)

        scroll_barx = Scrollbar(self.containerResult, orient=HORIZONTAL)
        scroll_barx.pack(side = BOTTOM, fill = X)

        self.listaResult = Listbox(self.containerResult, yscrollcommand= scroll_bary.set, xscrollcommand= scroll_barx.set, width=90)
        
        self.listaResult.pack( side = LEFT, fill = BOTH ) 
        scroll_bary.config( command = self.listaResult.yview ) 
        scroll_barx.config( command = self.listaResult.xview ) 

    def inserirResult(self, result):
        self.listaResult.insert(END, result)
        self.listaResult.update()
        self.listaResult.see(END)
        
    #Método baixar modelo de planilha
    def baixarModelo(self):
        dict_modelo = {'Num':["98765432109876543210987654321098765432101234"],
        'Status': ["Será preenchido pelo sistema"],
        'Msg':["Será preenchido pelo sistema"]}
        df_modelo = pd.DataFrame(dict_modelo)
        caminho_salvar = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="modelo_planilha_nfp.xlsx", title="Salvar Planilha Modelo", filetypes=[('Xlsx', '.xlsx'), ('All Files', '*')])
        
        if caminho_salvar == '':
            pass
        else:
            # Create a Pandas Excel writer using XlsxWriter as the engine.
            writer = pd.ExcelWriter(caminho_salvar, engine='xlsxwriter')
            # Convert the dataframe to an XlsxWriter Excel object.
            df_modelo.to_excel(writer, sheet_name='Planilha_Modelo', index=False)
            # Get the xlsxwriter workbook and worksheet objects.
            workbook  = writer.book
            worksheet = writer.sheets['Planilha_Modelo']
            # Add some cell formats.
            format1 = workbook.add_format({'num_format': 49})
            worksheet.set_column(0, 2, 46, format1)
            writer.save()

            
            self.listaResult.insert(END, 'Planilha modelo salva em: ' + caminho_salvar)
        

    #Método importar planilha com as notas
    def importarPlan(self):
        global df
        global nome_do_arquivo
        nome_do_arquivo = filedialog.askopenfilename()
        print(nome_do_arquivo)
        df = pd.read_excel(nome_do_arquivo)
        #print(df)
        
        validacao = True
        for i, j in df.iterrows():
            if len(j.Num) == 44:
                continue
            else:
                msg_erro = 'Linha ' + str(i+2) + ' da planilha não possui 44 caracteres.'
                messagebox.showerror(title='Erro na validação dos dados', message=msg_erro)
                validacao = False
                break
        
        if validacao:
            tela_principal.inserirResult("Planilha importada: " + nome_do_arquivo)
        else:
            tela_principal.inserirResult(msg_erro)
        
        statusTotal = len(df.index)
        self.lblTotal['text'] = f'Total de registros importados: {statusTotal}.'
        self.lblTotal['font'] = ("Arial", "11", 'bold')

    def switcher_projeto(x):
        x = tela_principal.menuProjetos.get()
        switcher = {
            'Projeto Liga Solidária': '90107' ,
            'Projeto Casulo': '32523'
        }
        return switcher.get(x)
        #<option value="90107">LIGA DAS SENHORAS CATOLICAS DE SAO PAULO</option>
        #<option value="32523">PROJETO CASULO</option>

    def validar_premissas(self):
        if not self.nome.get() or not self.senha.get():
            messagebox.showerror(title='Erro', message='Necessário preencher usuário e senha.')
            return False
        elif len(df) == 0:
            messagebox.showerror(title='Erro', message='Necessário importar um arquivo antes de iniciar')
            return False
        elif tela_principal.menuProjetos.get() == '-- Selecione o projeto --':
            messagebox.showerror(title='Erro', message='Necessário selecionar um projeto antes de iniciar')
            return False
        else:
            return True

    def iniciar(self):
        if self.validar_premissas():
            iniciar_driver()
            if logar():
                navegar()
                lancar_notas()
            print("Fim!!!!!!!")
            #chrome.quit()


root = Tk()
tela_principal = Tela(root)
root.mainloop()

