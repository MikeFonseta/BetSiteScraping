from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl
import sys

  
def analisi():
    #Apertura foglio Excel già esistente
    wb = openpyxl.load_workbook("Scraping.xlsx")

    #Caricamento file driver broswer (in questo caso Chrome con file per Mac)
    driver = webdriver.Chrome('chromedriver')
    #Impostazini grandezza browser
    driver.set_window_size(1200, 700)
    #Indirizzamento alla pagina web
    driver.get("https://www.eurobet.it/it/scommesse/#!/calcio/ing-premier-league/")

    #Tempo di attesa per caricare la pagina
    sleep(5)

    #Ricerca dei pulsanti (Più giocate, Novità, Goal ecc)
    buttons = driver.find_elements(By.CLASS_NAME,'sportfilter-item')
    x=0
    #Controllo che ci sia filtro ARBITRO
    while buttons[x].find_element(By.TAG_NAME,'a').text != "ARBITRO":
        x+=1
        pass
        
    #Click pulsante ARBITRO
    buttons[x].find_element(By.TAG_NAME,'a').click()
    #Tempo di attesa per caricare la pagina
    sleep(3)

    #Ricerca dei pulsanti (U/O CARTELLINI ecc..)
    buttons = driver.find_elements(By.CLASS_NAME, 'sportfilter-item')
    x=-1
    #Controllo che ci sia filtro U/O CARTELLINI
    while len(buttons)>0 and buttons[x].find_element(By.TAG_NAME,'a').text != "U/O CARTELLINI" and x<len(buttons):
        x+=1
        pass
        
    #Se il ciclo termina senza essere andato avanti o senza aver trovato il pulsante U/O CARTELLINI esci dal programma
    if(x >= len(buttons) or x==-1):
        print("Quote EUROBET non disponibili")
        sys.exit()

    #Click pulsante U/O CARTELLINI trovato nel ciclo while
    buttons[x].find_element(By.TAG_NAME,'a').click()
    #Tempo di attesa per caricare la pagina
    sleep(3)

    #Codice utilizzato per scorrere tutta la pagina lentamente e permettere la visualizzazione di tutte le partite
    height = driver.execute_script("return document.body.scrollHeight")
    for scrol in range(100,height,100):
        driver.execute_script(f"window.scrollTo(0,{scrol})")
        sleep(0.1)
    sleep(2)

    #Ricerca partite
    teams = driver.find_elements(By.CLASS_NAME, 'box-row-event')

    #Lista utilizzata per aggiungere i valori quote
    result = []
    x = 0
    #Analisi per ogni singola partita
    for match in teams: 
        #Ricerca riga delle quote
        quoteMatch = match.find_elements(By.CLASS_NAME, 'riga-quota')
        #Ricerca nome squadre
        name = match.find_element(By.CLASS_NAME, 'event-name.prematch-name')
        #Analisi per ogni singola quota
        for quota in quoteMatch:
            quoteMatchSingolo = quota.find_elements(By.CLASS_NAME, 'containerQuota')      
            
            i = 0
            quotaValue = ""
            quote = []
            #Analisi per ogni elemento della quota singola
            for quotaSingola in quoteMatchSingolo:
                if i == 0:
                    #Ricerca tipo quota ES: 0.5, 1.5, 5.5
                    quotaValue = quotaSingola.find_element(By.CLASS_NAME, 'info_aggiuntiva')
                else:
                    #Ricerca valore quota
                    quota = quotaSingola.find_element(By.CLASS_NAME, 'quota')   
                    quote.append(quota.text)    
                i+=1
            #Aggiunta dei valori alla lista
            result.append([str(name.text).replace("\n","-").split("-"),quotaValue.text,quote])
    
    #Creazione foglio nel file excel già aperto
    sheet = wb.create_sheet(title = 'EUROBET')
    #Inserimento nel foglio dei seguenti valori
    sheet['F3'] = "Quota"
    sheet['G3'] = "UNDER Eurobet"
    sheet['H3'] = "OVER Eurobet"

    #Inserimento di nomi squadre, tipo quote, valori
    for x in range(0, len(result)):
        sheet['D'+str(5+x)] = result[x][0][0]
        sheet['E'+str(5+x)] = result[x][0][1]    
        sheet['F'+str(5+x)] = result[x][1]
        sheet['G'+str(5+x)] = result[x][2][0]
        sheet['H'+str(5+x)] = result[x][2][1]
    
    print("Done Eurobet")
    #Salvataggio foglio excel
    wb.save("Scraping.xlsx")
    #Chiusura driver
    driver.close()

