import undetected_chromedriver as uc
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import Keys, ActionChains
from selenium.webdriver.support.select import Select
from pyvirtualdisplay import Display
from datetime import datetime,timedelta
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail
import configparser
import os
import time
import re

def inicia_navegador():

    # Configuracio del navegador
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.6167.140 Safari/537.36"
    driver = uc.Chrome(useAutomationExtension=False,use_subprocess=False,binary_location='./chromedriver')
    driver.implicitly_wait(30)
    return(driver)

def obrir_ieduca(driver, usuari, contrasenya):
    driver.get('https://inslacetania.ieduca.com/')

    driver.find_element(By.ID,'username_email').send_keys(usuari)
    driver.find_element(By.ID,'password').send_keys(contrasenya)
    driver.find_element(By.ID,'btn-login').click()

    # Click al captcha de CloudFlare... que ha d'estar maximitzat per funcionar ok.
    # TODO Calcular el 1 de setembre de cada any automàticament.
    time.sleep(3)
    ActionChains(driver).send_keys(Keys.TAB).perform()
    ActionChains(driver).send_keys(Keys.SPACE).perform()
    time.sleep(3)
    return(driver)

def obtenir_faltes(driver):
    any_actual =  datetime.now().year
    driver.get('https://inslacetania.ieduca.com/index.php?seccio=154')
    driver.find_element(By.NAME,'data1').clear()
    driver.find_element(By.NAME,'data1').send_keys(str(any_actual)+'-09-01')
    Select(driver.find_element(By.NAME,'tipusf')).select_by_visible_text('Falta sense justificar')
    driver.find_element(By.NAME,'filtrar').click()

    taula = driver.find_element(By.CLASS_NAME,'taula')
    rows = taula.find_elements(By.TAG_NAME,'tr')

    # CARREGAR LES FALTES EN UN Diccionari de Faltes
    faltes = {}
    fila = []
    for row in rows:
        cols = row.find_elements(By.TAG_NAME,'td')
        for col in cols:
            fila.append(col.text)
        if len(fila) == 4:
            parts_nom = fila[1].strip().split(',')
            nom = parts_nom[1].strip()
            cognoms = parts_nom[0].strip()
            nom_alumne =f"{nom} {cognoms}"

            text_faltes =''.join(fila[3:])
            # Patró per trobar les hores i dates
            patro_faltes = r'(\d{2}:\d{2}-\d{2}:\d{2}) .*? (\d{2}/\d{2}/\d{4})'
            alumne_faltes = re.findall(patro_faltes, text_faltes)
            ## Convertim el resultat en un format més llegible
            faltes[nom_alumne] = [f"{hora} {data}" for hora, data in alumne_faltes]
        fila = []
    return(faltes)

def obtenir_incidencies(driver):
    # CARREGAR FALTES LLEUS PER FALTES D'ASSISTÈNCIA

    driver.get('https://inslacetania.ieduca.com/index.php?seccio=208')

    taula = driver.find_element(By.CLASS_NAME,'taula')
    rows = taula.find_elements(By.TAG_NAME,'tr')

    # CARREGAR LES FALTES EN UN Diccionari de Faltes
    incidencies = {}
    fila = []
    for row in rows:
        cols = row.find_elements(By.TAG_NAME,'td')
        for col in cols:
            fila.append(col.text)
        if len(fila) == 4 and "Faltes injustificades" in fila[3]:
            nom_alumne = re.search(r'^(.*?)CFP',fila[3]).group(1).strip()
            data_incidencia = fila[1].split('\n')[0]
            if nom_alumne not in incidencies:
                incidencies[nom_alumne] = data_incidencia
        fila = []
    return(incidencies)

def obtenir_faltes_desde_incidencia(faltes, incidencies):
    # CALCUL DE FALTES
    resum = {}
    any_actual = datetime.now().year

    inici_curs = f'01/09/{any_actual}'

    for alumne in faltes:
        totals = 0
        setmana = 0
        incidencia = 0
        for data_incidencia in faltes[alumne]:
            totals+=1
            if alumne not in incidencies:
                incidencies[alumne] = inici_curs
            if datetime.strptime(data_incidencia.split()[1], '%d/%m/%Y') > datetime.strptime(incidencies[alumne],'%d/%m/%Y'):
                incidencia+=1
            if datetime.strptime(data_incidencia.split()[1], '%d/%m/%Y') > datetime.now() - timedelta(days=7):
                setmana+=1
        resum[alumne]=[incidencia,setmana,totals]
    return(resum)

def obtenir_faltes_llindar(faltes, llindar):
    llistat = {alumne : valors for alumne, valors in faltes.items() if valors[0] >= int(llindar)} 
    return(llistat)

def notificacio_linux(faltes):
    # Notificació per missatge al sistema
    text = "======================================"
    for alumne, valors in faltes.items():
        text += f'\n<b>{alumne}</b>\t<b>I: {valors[0]}</b>\tS: {valors[1]}\tT: {valors[2]}'
    text += '\n=====================================\n\n <b>LLEGENDA</b> \n <b>I</b>: Faltes des de última Incidència <b>S</b>: Faltes última setmana <b>T</b>: Faltes totals'
    os.system(f'notify-send -u critical -t 100000 "ALUMNES FALTES INJUSTIFICADES PER SOBRE LLINDAR" "{text}"')

def informe_html(titol,faltes):
    # Taula HTML de faltes
    html = f"""
    <h2>{titol}</h2>
    <table border="1" style="border-collapse: collapse; width: 100%;">
    <thead>
        <tr>
            <th>Alumne</th>
            <th>Faltes des de l'última incidència</th>
            <th>Faltes dels últims 7 dies</th>
            <th>Faltes des de principi de curs</th>
        </tr>
    </thead>
    <tbody>
    """

    # Afegir les files
    for nom, valors in faltes.items():
        html += f"""
        <tr>
            <td>{nom}</td>
            <td>{valors[0]}</td>
            <td>{valors[1]}</td>
            <td>{valors[2]}</td>
          </tr>
        """

    # Tancar la taula
    html += """
        </tbody>
    </table>
    """
    return(html)


def notificacio_arxiu(cos_html, arxiu='informe_faltes.html'):
    data_actual = datetime.now()
    html = f"""
    <html>
        <hear>
        <meta charset="UTF-8">
        <title>INFORME FALTES INJUSTIFICADES</title>
        </head>
        <body>
            <h1>INFORME DE FALTES INJUSTIFICADES ({data_actual})</h1>
    """
    html = html + cos_html
    html = html + """
        </body>
    </html>
    """

    with open(arxiu, "w", encoding="utf-8") as fitxer:
        fitxer.write(html)


def notificacio_email(apikey,origen,desti,html):
    message = Mail(
            from_email = origen,
            to_emails = desti,
            subject = 'Informe de faltes injustificades',
            html_content = html)
    sg = SendGridAPIClient(apikey)
    response = sg.send(message)

def main():
    # Carregar configuracio
    configuracio = configparser.ConfigParser()
    configuracio.read('./TAC.cfg')

    # Aquestes 2 linies son per executar Selenium en un diplay virtual, de forma que no interfereix amb la pantalla normal.  Cal tancar el display en acabar-ho tot.
    display = Display(visible=0, size=(800, 600))
    display.start()

    navegador = inicia_navegador()

    navegador = obrir_ieduca(navegador,configuracio['iEduca']['usuari_ieduca'],configuracio['iEduca']['contrasenya_ieduca'])
    faltes = obtenir_faltes_desde_incidencia(obtenir_faltes(navegador), obtenir_incidencies(navegador))
    faltes_sobre_llindar = obtenir_faltes_llindar(faltes,configuracio['general']['llindar'])

    if configuracio.has_option('notificacions','linux'):
        if configuracio['notificacions']['linux'] == 'True':
          notificacio_linux(faltes_sobre_llindar)

    if configuracio.has_option('email','sendgrid_apikey') and configuracio.has_option('email','sendgrid_email') and configuracio.has_option('email','email_desti'):
        notificacio_email(configuracio['email']['sendgrid_apikey'],configuracio['email']['sendgrid_email'],configuracio['email']['email_desti'],informe_html('FALTES SOBRE LLINDAR', faltes_sobre_llindar)+informe_html('FALTES COMPLERT', faltes))

    if configuracio.has_option('arxiu','html'):
        if configuracio['arxiu']['html'] == 'True':
            notificacio_arxiu(informe_html('FALTES SOBRE LLINDAR', faltes_sobre_llindar)+informe_html('FALTES COMPERT', faltes))


    # Tanquem display virtual i driver de Selenium
    navegador.quit()
    display.stop()

if __name__ == "__main__":
    main()
