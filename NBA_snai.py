from time import sleep
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.edge.service import Service

CATEGORIA = 'U/O PUNTI GIOC. INCL. EV. OT'


def analisi():
    
    #Apertura foglio Excel già esistente
    data = []
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    #Impostazini grandezza browser
    driver.set_window_size(1200, 700)
    #Indirizzamento alla pagina web
    driver.maximize_window()
    driver.get('https://www.snai.it/sport/BASKET/NBA')

    sleep(5)

    
    buttons = driver.find_elements(By.CLASS_NAME, 'btn.btn-default.col-xs-3.btn-xs.ellipsis.ng-binding.ng-scope')
    
    x = 0

    while x<len(buttons) and buttons[x].text != 'GIOCATORE':
        x+=1
    
    if x>=len(buttons):
        print("Categoria giocatore non trovata.")
        sys.exit()
    else:
        buttons[x].click()


    

    #(By.CLASS_NAME, '2.btn.btn-default.col-xs-4.btn-xs.ellipsis.ng-binding.ng-scope')

    tab = driver.find_element(By.CLASS_NAME, 'btn-group.width-100.tipo-scommessa-prematch')

    buttons = tab.find_elements(By.TAG_NAME, 'button')
    x=0

    # while x<len(buttons) and buttons[x].text != 'U/O PUNTI+RIMB+ASSIST INCL. EV. OT':
    #     x+=1
    
    while x<len(buttons) and buttons[x].text != CATEGORIA:
        x+=1
    

    if x>=len(buttons):
        print(CATEGORIA + " non trovata.")
        sys.exit()
    else:
        buttons[x].click()

    sleep(3)

    height = driver.execute_script("return document.body.scrollHeight")
    for scrol in range(100,height,100):
        driver.execute_script(f"window.scrollTo(0,{scrol})")
        sleep(0.1)
    sleep(2)
    
    matches = driver.find_elements(By.CLASS_NAME, 'col-xs-12.nopaddingLeftRight.whiteOneMargin')

    for z in range(0,len(matches)):
        text = matches[z].find_element(By.CLASS_NAME, 'nopaddingLeftRight.matchDescriptionFirstCol.footballWidthFirstCol').text
        
        time = text[0:5]
        match = text[6:len(text)]
        #data.append([text[0:5],text[6:len(text)],{}])

        switchFieldPlayers = matches[z].find_elements(By.CLASS_NAME, 'switch-fieldPlayers')

        for y in range(0,len(switchFieldPlayers)):

            clickPlayers = switchFieldPlayers[y].find_elements(By.CLASS_NAME, 'ng-scope')
            
            for x in range(0,len(clickPlayers)):        
                matches = driver.find_elements(By.CLASS_NAME, 'col-xs-12.nopaddingLeftRight.whiteOneMargin')
                switchFieldPlayer = matches[z].find_elements(By.CLASS_NAME, 'switch-fieldPlayers')[y]
                clickPlayers = switchFieldPlayer.find_elements(By.CLASS_NAME, 'ng-scope')
                label = clickPlayers[x].find_element(By.TAG_NAME, 'label')
                
                switchfieldScoresFifty = matches[z].find_element(By.CLASS_NAME, 'ng-scope.switch-fieldScoresFifty')
                quote = switchfieldScoresFifty.find_elements(By.TAG_NAME, 'label')
                quoteValue = []
                quoteValue.append(label.text)
                t = 0
                temp = []
                for quota in quote:
                    if(t == 0):
                        temp = [quota.text.split("\n")]
                    elif(t == 1):
                        temp.append(quota.text.split("\n"))
                        t = 0
                        data.append([time,match,label.text,temp[0][0],temp[0][1]])
                        data.append([time,match,label.text,temp[1][0],temp[1][1]])
                        temp = []
                    t+=1


                if x+1 < len(clickPlayers):
                    driver.execute_script("arguments[0].click();", clickPlayers[x+1].find_element(By.TAG_NAME, 'input'))
                sleep(3)
    
    pd.DataFrame(data,columns=['Ora', 'Partita', 'Giocatore','Tipo Quota','Quota']).to_excel('Scraping.xlsx',index=False)

    print("Done")
    #Salvataggio foglio excel

    sleep(5)
    driver.close()


analisi()