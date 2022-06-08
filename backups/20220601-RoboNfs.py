import pandas as pd #Para ler arquivos e transdormar em dataframes
from selenium import webdriver #O navegador
from selenium.webdriver.common.by import By #Achar os elementos
from selenium.webdriver.common.keys import Keys #Para digitar no teclado na web
from selenium.webdriver.support.ui import WebDriverWait #Foi necessário para o captcha
from selenium.webdriver.support import expected_conditions as EC #Foi necessário para o captcha
from selenium.webdriver.common.action_chains import ActionChains #Para clicar e navegar em submenus
from selenium.webdriver.support.select import Select #Para preencher droplists
from selenium.common.exceptions import NoSuchElementException
import time #Para colocar wait
from datetime import date
from datetime import datetime
from tkinter import *
#from selenium.webdriver.chrome.options import Options
#from selenium.webdriver.chrome.service import Service
#from webdriver_manager.chrome import ChromeDriverManager

def principal():
    nome_do_arquivo = 'nfs.xlsx'
    url_nfp = 'https://www.nfp.fazenda.sp.gov.br'
    user_max = ''
    pass_max = ''
    df = pd.read_excel(nome_do_arquivo)
    #print(df)
    #serie = df.squeeze('columns')

    chromedrive_path = 'C:\\Users\Max\Google Drive\Max\CWMC\chromedriver.exe'
    chrome = webdriver.Chrome(executable_path=chromedrive_path)

    #options = Options()
    #chrome = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    chrome.get(url_nfp)
    time.sleep(3)
    elemento_texto_user = chrome.find_element(By.XPATH, '//*[@id="UserName"]')
    elemento_texto_pass = chrome.find_element(By.XPATH, '//*[@id="Password"]')
    elemento_texto_user.send_keys(user_max)
    elemento_texto_pass.send_keys(pass_max)

    #Para clicar no recaptha
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False) 
    WebDriverWait(chrome, 10).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR,"iframe[name^='a-'][src^='https://www.google.com/recaptcha/api2/anchor?']")))
    WebDriverWait(chrome, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[@id='recaptcha-anchor']"))).click()

    time.sleep(5)
    chrome.switch_to.default_content()
    time.sleep(1)
    elemento_botao_login = chrome.find_element(By.XPATH, '//*[@id="btnLogin"]')
    elemento_botao_login.click()
    time.sleep(1)

    acao = ActionChains(chrome)
    elemento_ent = chrome.find_element(By.XPATH, '//*[@id="menuSuperior"]/ul/li[4]/a')
    acao.move_to_element(elemento_ent).perform()
    elemento_cad_cupom = chrome.find_element(By.XPATH, '//*[@id="menuSuperior:submenu:12"]/li[1]/a')
    acao.move_to_element(elemento_cad_cupom).click().perform()

    chrome.find_element(By.XPATH,'//*[@id="ctl00_ConteudoPagina_btnOk"]').click()
    time.sleep(1)

    elemento_caixa = Select(chrome.find_element(By.XPATH,'//*[@id="ddlEntidadeFilantropica"]'))
    elemento_caixa.select_by_value('90107')
    #<option value="90107">LIGA DAS SENHORAS CATOLICAS DE SAO PAULO</option>
    #<option value="32523">PROJETO CASULO</option>


    time.sleep(1)
    chrome.find_element(By.XPATH,'//*[@id="ctl00_ConteudoPagina_btnNovaNota"]').click()
    time.sleep(1)

    chrome.find_element(By.TAG_NAME, value='body').send_keys(Keys.ESCAPE)

    #now = datetime.now()
    #dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    #print("date and time =", dt_string)

    for i, j in df.iterrows():
        if j.Status == 'Processado':
            continue
        else:
            try:
                elemento_num_nf = chrome.find_element(By.XPATH, '//*[contains(@title, "Digite ou Utilize")]')
                elemento_num_nf.send_keys(str(j.Num))
                elemento_salva_nf = chrome.find_element(By.XPATH, '//*[@id="btnSalvarNota"]')
                time.sleep(1)
                elemento_salva_nf.click()
                time.sleep(1)
                time_now = datetime.now()
                time_now_formated = time_now.strftime("%Y-%m-%d %H:%M:%S")
                logExec = open('logExe.txt', 'a')
                try:
                    elemento_sucesso = chrome.find_element(By.XPATH, '//*[@id="lblInfo"]')
                    if 'registrada com sucesso' in elemento_sucesso.text:
                        print(time_now_formated + ' | Sucesso! | ' + 'Msg: ' + elemento_sucesso.text + ' | ' + str(j.Num), file = logExec)
                    else:
                        print(time_now_formated + ' | Erro! | ' + 'Msg: ' + elemento_sucesso.text + ' | ' + str(j.Num) , file = logExec)
                except NoSuchElementException:
                    try:
                        elemento_erro = chrome.find_element(By.XPATH, '//*[@id="lblErro"]')
                        print(time_now_formated + ' | Erro! | ' + 'Msg: ' + elemento_erro.text + ' | ' + str(j.Num), file = logExec)
                    except:
                        print(time_now_formated + ' | Erro! | ' + ' Não foi possível cadastrar, erro não mapeado. | ' + str(j.Num), file = logExec)
                df.at[i, 'Status']= 'Processado'
                
                if 'elemento_sucesso' in locals():
                    df.at[i, 'Msg']= elemento_sucesso.text
                elif 'elemento_erro' in locals():
                    df.at[i, 'Msg']= elemento_erro.text
                else:
                    pass
            except NoSuchElementException:
                break
            

    df.to_excel('nfs.xlsx', index=False)

    print("Fim!!!!!!!")
    chrome.quit()

def login(user_max, pass_max):
    nome_do_arquivo = 'nfs.xlsx'
    url_nfp = 'https://www.nfp.fazenda.sp.gov.br'
    #user_max = ''
    #pass_max = ''
    df = pd.read_excel(nome_do_arquivo)
    #print(df)
    #serie = df.squeeze('columns')

    chromedrive_path = 'C:\\Users\Max\Google Drive\Max\CWMC\chromedriver.exe'
    chrome = webdriver.Chrome(executable_path=chromedrive_path)

    #options = Options()
    #chrome = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    chrome.get(url_nfp)
    time.sleep(3)
    elemento_texto_user = chrome.find_element(By.XPATH, '//*[@id="UserName"]')
    elemento_texto_pass = chrome.find_element(By.XPATH, '//*[@id="Password"]')
    elemento_texto_user.send_keys(user_max)
    elemento_texto_pass.send_keys(pass_max)

#principal()


class Tela:
    def __init__(self, master=None):
        root.title('NFP - Lançamento automatizado')

        self.fontePadrao = ("Arial", "10")
        self.containerTitulos = Frame(master)
        self.containerTitulos["pady"] = 10
        self.containerTitulos.pack()

        self.containerBotoesPlanilhas = Frame(master)
        self.containerBotoesPlanilhas["pady"] = 20
        self.containerBotoesPlanilhas["padx"] = 20
        self.containerBotoesPlanilhas.pack()

        self.fontePadrao = ("Arial", "10")
        self.containerTitulosLogin = Frame(master)
        self.containerTitulosLogin["pady"] = 10
        self.containerTitulosLogin.pack()

        self.containerNome = Frame(master)
        self.containerNome["padx"] = 20
        self.containerNome.pack()

        self.containerSenha = Frame(master)
        self.containerSenha["padx"] = 20
        self.containerSenha.pack()

        self.containerMenuProjeto = Frame(master)
        self.containerMenuProjeto["pady"] = 10
        self.containerMenuProjeto.pack()

        self.containerIniciar = Frame(master)
        self.containerIniciar["pady"] = 10
        self.containerIniciar.pack()

        self.containerResult = Frame(master)
        self.containerResult.pack()

        self.titulo = Label(self.containerTitulos, text="Liga Solidária")
        self.titulo2 = Label(self.containerTitulos, text="- Lançamento automatizado de NFP -")
        self.titulo3 = Label(self.containerTitulosLogin, text="Informe seu nome de usuário e senha do site:" + "\n" "https://www.nfp.fazenda.sp.gov.br")
        self.titulo["font"] = ("Arial", "10", "bold")
        self.titulo.pack()
        self.titulo2.pack()
        self.titulo3["pady"] = 10
        self.titulo3.pack()
        
        self.nomeLabel = Label(self.containerNome,text="Nome", font=self.fontePadrao)
        self.nomeLabel.pack(side=LEFT)

        self.nome = Entry(self.containerNome)
        self.nome["width"] = 30
        self.nome["font"] = self.fontePadrao
        self.nome.pack(side=LEFT)

        self.senhaLabel = Label(self.containerSenha, text="Senha", font=self.fontePadrao)
        self.senhaLabel.pack(side=LEFT)

        self.senha = Entry(self.containerSenha)
        self.senha["width"] = 30
        self.senha["font"] = self.fontePadrao
        self.senha["show"] = "*"
        self.senha.pack(side=LEFT)

        self.btBaixar_modelo = Button(self.containerBotoesPlanilhas)
        self.btBaixar_modelo["text"] = "Baixar Planilha Modelo"
        self.btBaixar_modelo["font"] = ("Calibri", "8")
        self.btBaixar_modelo["width"] = 24
        self.btBaixar_modelo["command"] = self.baixarModelo
        self.btBaixar_modelo.pack(side=LEFT)
        
        self.btImportarPlan = Button(self.containerBotoesPlanilhas)
        self.btImportarPlan["text"] = "Importar Planilha de Notas"
        self.btImportarPlan["font"] = ("Calibri", "8")
        self.btImportarPlan["width"] = 24
        self.btImportarPlan["command"] = self.importarPlan
        self.btImportarPlan.pack(side=RIGHT)

        lista_projetos = ["Liga", "Casulo"]
        variable = StringVar(root)
        variable.set(lista_projetos[0]) # valor default
        self.menuProjetos = OptionMenu(self.containerMenuProjeto, variable, *lista_projetos)
        self.menuProjetos["font"] = ("Calibri", "8")
        self.menuProjetos.pack()

        self.btIniciar = Button(self.containerIniciar)
        self.btIniciar["text"] = "INICIAR"
        self.btIniciar["font"] = ("Calibri", "8")
        self.btIniciar["width"] = 24
        self.btIniciar["command"] = self.iniciar
        self.btIniciar.pack()

        #Definindo a lista de resultado com barra de rolagem
        scroll_bar = Scrollbar(self.containerResult)
        scroll_bar.pack(side = RIGHT, fill = Y)

        self.listaResult = Listbox(self.containerResult, yscrollcommand= scroll_bar.set, width=50)
        
        # for line in range(1, 26): 
        #     listaResult.insert(END, "Geeks " + str(line)) 
        
        self.listaResult.pack( side = LEFT, fill = BOTH ) 
        scroll_bar.config( command = self.listaResult.yview ) 


    #Método baixar modelo de planilha
    def baixarModelo(self):
        dict_modelo = {'Num':["Cole aqui o número da nota"],
        'Status': ["Será preenchido pelo sistema"],
        'Msg':["Será preenchido pelo sistema"]}
        
        df_modelo = pd.DataFrame(dict_modelo)

        df_modelo.to_excel('modelo_planilha_nfp.xlsx',index=False)

        self.listaResult.insert(END, "Planilha modelo salva!")
        

    #Método importar planilha com as notas
    def importarPlan(self):
        pass
        #implementar

    def iniciar(self):
        self.user_max = self.nome.get()
        self.pass_max = self.senha.get()
        login(self.user_max, self.pass_max)


    #Método verificar senha
    # def verificaSenha(self):
    #     usuario = self.nome.get()
    #     senha = self.senha.get()
    #     if usuario == "usuariodevmedia" and senha == "dev":
    #         self.mensagem["text"] = "Autenticado"
    #     else:
    #         self.mensagem["text"] = "Erro na autenticação"


root = Tk()
Tela(root)

root.mainloop()