# -*- coding: utf-8 -*-

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import datetime as dt
import numpy as np


if __name__ == "__main__":

    today = dt.date.today().strftime("%d.%m.%y")

    # WEBDRIVER
    # PATH = 'C:/Program Files (x86)/msedgedriver.exe'
    PATH = 'IEDriverServer.exe'
    #driver = webdriver.Ie(PATH)
    driver = webdriver.Chrome()

    url_forms = r'https://forms.office.com/___'
    driver.get(url_forms)

    #driver.maximize_window()

    #advanced = driver.find_element_by_id("details-button")
    #advanced.send_keys(Keys.RETURN)

    #proceed = driver.find_element_by_id("proceed-link")
    #proceed.send_keys(Keys.RETURN)

    #%%
    wdw = WebDriverWait(driver, 10, poll_frequency=2)
    wdw_short = WebDriverWait(driver, 6, poll_frequency=2)

    #%%LOGIN
    try:
        login = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "usernameField")))
    except:
        driver.quit()
    login = driver.find_element_by_id("usernameField")
    login.send_keys("99004806")
    password = driver.find_element_by_id("passwordField")
    password.send_keys("-")
    enter_login = driver.find_element_by_id("SubmitButton")
    enter_login.click()


    #%%LANGUAGE

    #driver.switch_to.window(driver.window_handles[0])
    #wdw.until(EC.presence_of_element_located((By.XPATH,'//*[@id="PageLayoutRN"]/table[1]/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/table/tbody/tr/td[2]/a'))).click()
    #driver.implicitly_wait(5)
    #driver.find_element(By.ID,'CurrentLanguage').send_keys('Brasilian Portuguese')
    #driver.find_element(By.ID,'Apply').click()
    #driver.find_element(By.XPATH ,'//*[@id="PageLayoutRN"]/table[1]/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/a').click()



    #%%NAVIGATION
    pre_cadastro_solicitante = driver.find_element_by_link_text("Pré-Cadastro de Itens - Solicitante")
    pre_cadastro_solicitante.click()
    sumario = driver.find_element_by_id("N149")
    sumario.click()


    #%%CREATE A DATAFRAME TO SAVE SLCT
    solict = pd.DataFrame(columns=['Item', 'Tipo do Item', 'Numero da Solicitacao', 'Numero do Formulario'])


    #%%
    i = 0


    #%%LOOP
    while i < len(base):

        #%Adicionar Solicitacao
        driver.switch_to.window(driver.window_handles[0])
        wdw.until(EC.presence_of_element_located((By.ID, "AddSolicitacao"))).click()
        
        
        #%TIPO   
        driver.switch_to.window(driver.window_handles[0])
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "ItemType"))).send_keys(base.Tipo[i])
        
        
        #%TIPO DO ITEM
        driver.implicitly_wait(5)
        driver.switch_to.window(driver.window_handles[0])
        driver.find_element_by_id("Dctipoitemusuario").send_keys(base.Tipo_do_Item[i], Keys.TAB)
        try:
            wdw_short.until(EC.new_window_is_opened(driver.window_handles))
            driver.switch_to.window(driver.window_handles[-1])
            wdw_short.until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "html > frameset > frame")))
            wdw.until(EC.presence_of_element_located((By.XPATH,'/html/body/span/div[1]/div[2]/form/table[2]/tbody/tr[4]/td/table/tbody/tr/td/div/div[3]/span[1]/table[2]/tbody/tr[2]/td[1]/input'))).click()
            wdw.until(EC.presence_of_element_located((By.XPATH,'/html/body/span/div[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[3]/table/tbody/tr/td[4]/button'))).click()
        except:
            pass
        
        
        #%CATALOGO
        time.sleep(4)
        driver.switch_to.window(driver.window_handles[0])
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "DcCatalogo"))).send_keys(base.Catalogo[i], Keys.TAB)
        try:
            wdw_short.until(EC.new_window_is_opened(driver.window_handles))
            driver.switch_to.window(driver.window_handles[-1])
            #wdw_short.until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "html > frameset > frame")))
            driver.switch_to.frame(0)
            wdw.until(EC.presence_of_element_located((By.XPATH,'/html/body/span/div[1]/div[2]/form/table[2]/tbody/tr[4]/td/table/tbody/tr/td/div/div[3]/span[1]/table[2]/tbody/tr[2]/td[1]/input'))).click()
            wdw.until(EC.presence_of_element_located((By.XPATH,'/html/body/span/div[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[3]/table/tbody/tr/td[4]/button'))).click()
        except:
            pass
        
        
        #%UNIDADE DE MEDIDA
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[0])
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID,"DcPrimaryUom"))).send_keys(base.Unidade_de_Medida[i],Keys.TAB)
        
        
        #%PREÇO
        time.sleep(2)
        driver.switch_to.window(driver.window_handles[0])
        driver.find_element_by_id("ListPricePerUnit").send_keys(str(base.Preco[i]))
        
        
        #%SOLICITANTE
        time.sleep(2)
        driver.switch_to.window(driver.window_handles[0])
        driver.find_element_by_xpath('//*[@id="Nomeusuario__xc_0"]/a/img').click()
        wdw.until(EC.new_window_is_opened(driver.window_handles))
        driver.switch_to.window(driver.window_handles[-1])
        driver.switch_to.frame(0)
        driver.find_element_by_id("categoryChoice").send_keys("Email")
        driver.find_element_by_xpath('//*[@id="_LOVResFrm"]/table[1]/tbody/tr[4]/td/table/tbody/tr/td/div/div/input').clear()
        driver.find_element_by_xpath('//*[@id="_LOVResFrm"]/table[1]/tbody/tr[4]/td/table/tbody/tr/td/div/div/input').send_keys(base.Solicitante[i])
        driver.find_element_by_xpath('//*[@id="_LOVResFrm"]/table[1]/tbody/tr[4]/td/table/tbody/tr/td/div/div/button').click()
        driver.switch_to.window(driver.window_handles[-1])
        driver.switch_to.frame(0)
        wdw.until(EC.presence_of_element_located((By.ID,'N1:N8:0'))).click()
        wdw.until(EC.presence_of_element_located((By.XPATH,'/html/body/span/div[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[3]/table/tbody/tr/td[4]/button'))).click()
        
        
        #%ESPECIALIDADE
        time.sleep(4)
        driver.switch_to.window(driver.window_handles[0])
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID,"Dcespecialidade"))).send_keys(base.Especialidade[i], Keys.TAB)
        try:
            wdw_short.until(EC.new_window_is_opened(driver.window_handles))
            driver.switch_to.window(driver.window_handles[-1])
            driver.switch_to.frame(0)
            wdw.until(EC.presence_of_element_located((By.XPATH,'/html/body/span/div[1]/div[2]/form/table[2]/tbody/tr[4]/td/table/tbody/tr/td/div/div[3]/span[1]/table[2]/tbody/tr[2]/td[1]/input'))).click()
            wdw.until(EC.presence_of_element_located((By.XPATH,'/html/body/span/div[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[3]/table/tbody/tr/td[4]/button'))).click()
        except:
            pass
        '''
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[0])
        driver.find_element_by_xpath('//*[@id="Dcespecialidade__xc_0"]/a/img').click()
        time.sleep(2)
        driver.switch_to.window(driver.window_handles[-1])
        driver.switch_to.frame(0)
        time.sleep(2)
        wdw.until(EC.presence_of_element_located((By.ID,'categoryChoice'))).send_keys("Especialidade")
        #driver.find_element_by_id("categoryChoice").send_keys("Especialidade")
        driver.find_element_by_xpath('//*[@id="_LOVResFrm"]/table[1]/tbody/tr[4]/td/table/tbody/tr/td/div/div/input').clear()
        driver.find_element_by_xpath('//*[@id="_LOVResFrm"]/table[1]/tbody/tr[4]/td/table/tbody/tr/td/div/div/input').send_keys(base.Especialidade[i])
        driver.find_element_by_xpath('//*[@id="_LOVResFrm"]/table[1]/tbody/tr[4]/td/table/tbody/tr/td/div/div/button').click()
        driver.switch_to.window(driver.window_handles[-1])
        driver.switch_to.frame(0)
        wdw.until(EC.presence_of_element_located((By.ID,'N1:N8:0'))).click()
        wdw.until(EC.presence_of_element_located((By.XPATH,'/html/body/span/div[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[3]/table/tbody/tr/td[4]/button'))).click()
        '''    
        
        #%SUBESPECIALIDADE
        time.sleep(4)
        driver.switch_to.window(driver.window_handles[0])
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID,"Dcsubespecialidade"))).send_keys(base.Subespecialidade[i], Keys.TAB)
        try:
            wdw_short.until(EC.new_window_is_opened(driver.window_handles))
            driver.switch_to.window(driver.window_handles[-1])
            driver.switch_to.frame(0)
            wdw.until(EC.presence_of_element_located((By.XPATH,'/html/body/span/div[1]/div[2]/form/table[2]/tbody/tr[4]/td/table/tbody/tr/td/div/div[3]/span[1]/table[2]/tbody/tr[2]/td[1]/input'))).click()
            wdw.until(EC.presence_of_element_located((By.XPATH,'/html/body/span/div[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[3]/table/tbody/tr/td[4]/button'))).click()
        except:
            pass
        '''
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[0])
        driver.find_element_by_xpath('//*[@id="Dcsubespecialidade__xc_0"]/a/img').click()
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(2)
        driver.switch_to.frame(0)
        time.sleep(2)
        wdw.until(EC.presence_of_element_located((By.ID,'categoryChoice'))).send_keys("Cod. Subespecialidade")
        #driver.find_element_by_id("categoryChoice").send_keys("Cod. Subespecialidade")
        driver.find_element_by_xpath('//*[@id="_LOVResFrm"]/table[1]/tbody/tr[4]/td/table/tbody/tr/td/div/div/input').clear()
        driver.find_element_by_xpath('//*[@id="_LOVResFrm"]/table[1]/tbody/tr[4]/td/table/tbody/tr/td/div/div/input').send_keys(base.Subespecialidade[i])
        driver.find_element_by_xpath('//*[@id="_LOVResFrm"]/table[1]/tbody/tr[4]/td/table/tbody/tr/td/div/div/button').click()
        driver.switch_to.window(driver.window_handles[-1])
        driver.switch_to.frame(0)
        driver.find_element_by_id("N1:N8:0").click()
        driver.find_element_by_xpath('//*[@id="SubespecialidadeLovRN"]/div[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[3]/table/tbody/tr/td[4]/button').click()
        '''
        
        #%CATEGORIA
        time.sleep(4)
        driver.switch_to.window(driver.window_handles[0])
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID,"Dccategoria"))).send_keys(base.Categoria[i], Keys.TAB)
        try:
            wdw_short.until(EC.new_window_is_opened(driver.window_handles))
            driver.switch_to.window(driver.window_handles[-1])
            driver.switch_to.frame(0)
            wdw.until(EC.presence_of_element_located((By.XPATH,'/html/body/span/div[1]/div[2]/form/table[2]/tbody/tr[4]/td/table/tbody/tr/td/div/div[3]/span[1]/table[2]/tbody/tr[2]/td[1]/input'))).click()
            wdw.until(EC.presence_of_element_located((By.XPATH,'/html/body/span/div[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[3]/table/tbody/tr/td[4]/button'))).click()
        except:
            pass
        '''
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[0])
        driver.find_element_by_xpath('//*[@id="Dccategoria__xc_0"]/a/img').click()
        time.sleep(2)
        driver.switch_to.window(driver.window_handles[-1])
        driver.switch_to.frame(0)
        time.sleep(2)
        wdw.until(EC.presence_of_element_located((By.XPATH,'/html/body/span/div[1]/div[2]/form/table[1]/tbody/tr[4]/td/table/tbody/tr/td/div/div/select'))).send_keys("Cod. Categoria")
        #driver.find_element_by_id("categoryChoice").send_keys("Cod. Categoria")
        driver.find_element_by_xpath('//*[@id="_LOVResFrm"]/table[1]/tbody/tr[4]/td/table/tbody/tr/td/div/div/input').clear()
        driver.find_element_by_xpath('//*[@id="_LOVResFrm"]/table[1]/tbody/tr[4]/td/table/tbody/tr/td/div/div/input').send_keys(base.Categoria[i])
        driver.find_element_by_xpath('//*[@id="_LOVResFrm"]/table[1]/tbody/tr[4]/td/table/tbody/tr/td/div/div/button').click()
        driver.switch_to.window(driver.window_handles[-1])
        driver.switch_to.frame(0)
        driver.find_element_by_id("N1:N8:0").click()
        driver.find_element_by_xpath('//*[@id="CategoriaInsertLovRN"]/div[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[3]/table/tbody/tr/td[4]/button').click()
        '''
        
        #%SUBESPEC. CTBL
        if not base.Tipo_do_Item[i] == "ATIVO":
            time.sleep(2)
            driver.switch_to.window(driver.window_handles[0])
            driver.find_element_by_id("CdSubEsplCntb").send_keys(base.Subespecialidade_Contabil[i], Keys.TAB)
            time.sleep(3)
            driver.switch_to.window(driver.window_handles[0])
        
        
        #%CRITICIDADE MANUT
        if base.Especialidade[i] != "MATERIAIS INDIRETOS" or base.Especialidade[i] != "ENERGIA":
            time.sleep(2)
            driver.switch_to.window(driver.window_handles[0]) 
            driver.find_element_by_id("CtcdManu").send_keys(base.Criticidade_Manutencao[i])

            #%LOCAL INSPECAO
            driver.switch_to.window(driver.window_handles[0]) 
            driver.find_element_by_id("LocalInsp").send_keys(base.Local_de_Inspecao[i])

        
        #%FABRICANTE
        driver.switch_to.window(driver.window_handles[0])
        wdw.until(EC.presence_of_element_located((By.ID,'NmFabr'))).send_keys(base.Fabricante[i])   
        
        
        #%PART NUMBER
        driver.switch_to.window(driver.window_handles[0])
        wdw.until(EC.presence_of_element_located((By.ID,'PartNumber'))).send_keys(base.Part_Number[i],Keys.TAB)


        #%PART BOLETIM
        driver.switch_to.window(driver.window_handles[0])
        wdw.until(EC.presence_of_element_located((By.ID,'PartBoletim'))).send_keys(base.Part_Boletim[i])   
    
        
        #%MODELO
        driver.switch_to.window(driver.window_handles[0]) 
        wdw.until(EC.presence_of_element_located((By.ID, "ModlItem"))).send_keys(base.Modelo[i])
        
        
        #%MATERIAL
        driver.switch_to.window(driver.window_handles[0]) 
        wdw.until(EC.presence_of_element_located((By.ID,'MatrItemPtb'))).send_keys(base.Material_de_Composicao[i])

        
        #%DESCRICAO CURTA
        driver.switch_to.window(driver.window_handles[0]) 
        driver.find_element_by_id("DescriptionPtb").send_keys(base.Descricao_Curta[i])
        
        
        #%APLICACAO
        driver.switch_to.window(driver.window_handles[0])
        wdw.until(EC.presence_of_element_located((By.ID,'AplcItemPtb'))).send_keys(base.Aplicacao[i])
        
        
        #%CARAC TECNICAS
        driver.switch_to.window(driver.window_handles[0]) 
        driver.find_element_by_id("CrctAdilItemPtb").send_keys(base.Caract_Tecnicas[i])
        
        
        #%FORMULARIO
        driver.switch_to.window(driver.window_handles[0]) 
        wdw.until(EC.presence_of_element_located((By.ID,'CrctAdilItemUs'))).send_keys(str(base.Formulario[i]))


        #%INVESTIMENTO
        if not base.Especialidade[i] == "MATERIAIS INDIRETOS":
            if base.Investimento[i] == 'Sim':
                driver.switch_to.window(driver.window_handles[0])
                driver.find_element(By.XPATH,'/html/body/form/span[2]/div/div[3]/div/div[2]/table[7]/tbody/tr[4]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr/td[6]/table/tbody/tr[4]/td[3]/span/input').click()
                driver.implicitly_wait(5)
        

        #%PESO
        driver.switch_to.window(driver.window_handles[0]) 
        driver.find_element_by_id("DcWeightUom").send_keys(base.Peso_Unidade[i],Keys.TAB)
        try:
            wdw_short.until(EC.new_window_is_opened(driver.window_handles))
            driver.switch_to.window(driver.window_handles[-1])
            #wdw_short.until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "html > frameset > frame")))
            driver.switch_to.frame(0)
            driver.find_element(By.ID,'N1:N8:0').click()
            driver.find_element(By.XPATH,'//*[@id="UOMInsertLovRN"]/div[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[3]/table/tbody/tr/td[4]/button').click()
        except:
            pass
        
        #%PESO VALOR
        time.sleep(2)
        driver.switch_to.window(driver.window_handles[0])
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID,"UnitWeight"))).send_keys(base.Peso[i])
        
        
        #%VOLUME
        driver.switch_to.window(driver.window_handles[0]) 
        driver.find_element_by_id("DcVolumeUom").send_keys(base.Volume_Unidade[i],Keys.TAB)
        try:
            wdw_short.until(EC.new_window_is_opened(driver.window_handles))
            driver.switch_to.window(driver.window_handles[-1])
            #wdw_short.until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "html > frameset > frame")))
            driver.switch_to.frame(0)
            driver.find_element(By.ID,'N1:N8:0').click()
            driver.find_element(By.XPATH,'//*[@id="UOMInsertLovRN"]/div[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[3]/table/tbody/tr/td[4]/button').click()
        except:
            pass
        
        
        #%VOLUME VALOR
        time.sleep(2)
        driver.switch_to.window(driver.window_handles[0])
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "UnitVolume"))).send_keys(str(base.Volume[i]))
        
        
        #%DIMENSOES
        driver.switch_to.window(driver.window_handles[0]) 
        driver.find_element_by_id("DcDimensionUom").send_keys(base.Dimensoes_Unidade[i],Keys.TAB)
        try:
            wdw_short.until(EC.new_window_is_opened(driver.window_handles))
            driver.switch_to.window(driver.window_handles[-1])
            driver.switch_to.frame(0)
            driver.find_element(By.ID,'N1:N8:0').click()
            driver.find_element(By.XPATH,'//*[@id="UOMInsertLovRN"]/div[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[3]/table/tbody/tr/td[4]/button').click()
        except:
            pass
        
        
        #%DIMESOES VALOR
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[0])
        wdw.until(EC.presence_of_element_located((By.ID,"UnitLength"))).send_keys(str(base.Comprimento[i]))
        wdw.until(EC.presence_of_element_located((By.ID,"UnitWidth"))).send_keys(str(base.Largura[i]))
        #driver.find_element_by_id("UnitWidth").send_keys(str(base.Largura[i]))
        wdw.until(EC.presence_of_element_located((By.ID,"UnitHeight"))).send_keys(str(base.Altura[i]))
        #driver.find_element_by_id("UnitHeight").send_keys(str(base.Altura[i]))
        
        
        #%ANEXO
        driver.find_element_by_id("ATTACH_/mrs/oracle/apps/mrsc/ssc/webui/SolicitacaoSolicitantePG.attachItem_AddAttachmentButton").click()
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID,"FileInput_oafileUpload"))).send_keys("//w2kjf649.mrs.com.br/gg_pcm$/GER. DE PCM MATERIAIS\COMPARTILHADOS/COORD. PCM MATERIAIS/1.Material Rodante/14.Cadastro/GDR - Cadastro Itens PCM/Controle Pré Cadastro/2021/"+str(base.Arquivo[i]))
        #driver.find_element(By.XPATH,'/html/body/form/span[2]/div/div[6]/div/table/tbody/tr[1]/td[2]/table/tbody/tr/td[4]/button').click()
        #WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID,"FileInput_oafileUpload"))).send_keys("C:/Users/xande/Desktop/57031_USO_GERAL_PR_ME_VERDE_AMAZONAS.pdf")
        driver.find_element_by_id("Okay_uixr").click()

        #%SOLICT NUMBER
        solict.loc[i] = [base.loc[i,'Descricao_Curta']] + [base.loc[i,'Tipo_do_Item']] +[driver.find_element_by_id("IdSlctItem").text] + [base.loc[i,'Formulario']]
            
        #%SAVE
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID,"Salvar_uixr"))).click()
        i += 1
        
        
    #%%QUIT AND SAVE
    #driver.quit()
    solict.to_excel(today + "_SOLICITACOES.xlsx")

