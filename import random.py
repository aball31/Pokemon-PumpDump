import random
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import time
import os
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import re
from tenacity import retry, stop_after_attempt, wait_exponential, wait_fixed
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from selenium.common.exceptions import NoSuchElementException

# Configurazione colori e stili
COLORI = {
    'info': '\033[94m',     # Blu
    'success': '\033[92m',  # Verde
    'warning': '\033[93m',  # Giallo
    'error': '\033[91m',    # Rosso
    'reset': '\033[0m',     # Reset
    'divisore': '\033[90m┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅┅\033[0m'
}

# Configurazione Chrome
chrome_options = Options()
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--window-size=1920x1080")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-software-rasterizer")
chrome_options.add_argument("--disable-webgl")
chrome_options.add_argument("--use-angle=swiftshader")
chrome_options.add_argument("--disable-features=Vulkan")
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

# Aggiungi scelta modalità headless
headless = input("Vuoi usare la modalità headless? (s/n): ").lower() == 's'
if headless:
    chrome_options.add_argument("--headless=new")
    print(f"{COLORI['success']}🚀 Modalità headless attivata{COLORI['reset']}")
else:
    print("👀 Modalità con browser visibile")

# Percorso del driver Chrome per Windows
CHROME_DRIVER_PATH = './chromedriver.exe'  # Percorso relativo

print(f"Percorso driver: {CHROME_DRIVER_PATH}")
print(f"File esiste? {os.path.isfile(CHROME_DRIVER_PATH)}")

# Configurazione selettori aggiornati
SELECTORI = {
    'edizione': 'body > div.absolute.inset-0.w-full.h-2\\/5.bg-blend-overlay > div.flex.flex-col.w-full.px-2.pb-6.mx-auto.overflow-x-hidden.lg\\:flex-nowrap.lg\\:flex-col.sm\\:px-6.xl\\:px-8.max-w-screen-2xl > div.flex.flex-col.w-full.pb-6.mt-5.xl\\:flex-row.xl\\:mt-12 > div > div > main > div:nth-child(1) > div > div:nth-child(3) > p',
    'titolo': 'body > div.absolute.inset-0.w-full.h-2\\/5.bg-blend-overlay > div.flex.flex-col.w-full.px-2.pb-6.mx-auto.overflow-x-hidden.lg\\:flex-nowrap.lg\\:flex-col.sm\\:px-6.xl\\:px-8.max-w-screen-2xl > div.flex.flex-col.w-full.pb-6.mt-5.xl\\:flex-row.xl\\:mt-12 > div > div > main > div:nth-child(1) > div > h3',
    'numero': 'body > div.absolute.inset-0.w-full.h-2\\/5.bg-blend-overlay > div.flex.flex-col.w-full.px-2.pb-6.mx-auto.overflow-x-hidden.lg\\:flex-nowrap.lg\\:flex-col.sm\\:px-6.xl\\:px-8.max-w-screen-2xl > div.flex.flex-col.w-full.pb-6.mt-5.xl\\:flex-row.xl\\:mt-12 > div > div > main > div:nth-child(1) > div > div:nth-child(4) > div > p:nth-child(2)',
    'prezzo': 'body > div.absolute.inset-0.w-full.h-2\\/5.bg-blend-overlay > div.flex.flex-col.w-full.px-2.pb-6.mx-auto.overflow-x-hidden.lg\\:flex-nowrap.lg\\:flex-col.sm\\:px-6.xl\\:px-8.max-w-screen-2xl > div.flex.flex-col.w-full.pb-6.mt-5.xl\\:flex-row.xl\\:mt-12 > div > div > main > div:nth-child(1) > div > div.flex.justify-between.mt-1 > h2'
}

def trova_selettore_effettivo(driver, selectors):
    """Verifica presenza selettori direttamente via Selenium"""
    for selector in selectors:
        elements = driver.find_elements(By.CSS_SELECTOR, selector)
        if elements:
            return selector
    return None

@retry(stop=stop_after_attempt(3), wait=wait_fixed(2))
def scarica_dati():
    try:
        num_carte = int(input("Quante carte vuoi analizzare? "))
        driver = webdriver.Chrome(service=Service(CHROME_DRIVER_PATH), options=chrome_options)
        
        if not headless:
            driver.maximize_window()  # Solo in modalità visibile
            
        driver.get('https://app.getcollectr.com')
        
        # Gestione cookie
        try:
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/button[1]/p"))
            ).click()
            print(f"{COLORI['success']}✅ Cookie accettati!{COLORI['reset']}")
        except:
            print(f"{COLORI['warning']}⚠️  Nessun popup cookie trovato{COLORI['reset']}")

        prodotti = []
        current_card = 1
        tentativi_falliti = 0

        while len(prodotti) < num_carte:
            try:
                # Costruzione dinamica dei selettori
                base_xpath = f"/html/body/div[1]/div[3]/div[3]/div/div/main/div[{current_card}]"
                css_base = f"div.flex.flex-col.w-full.pb-6.mt-5.xl\\:flex-row.xl\\:mt-12 > div > div > main > div:nth-child({current_card})"
                
                # Mappatura dinamica dei selettori
                selettori = {
                    'Edizione': (By.CSS_SELECTOR, f"{css_base} > div > div:nth-child(3) > p"),
                    'Titolo': (By.CSS_SELECTOR, f"{css_base} > div > h3"),
                    'Numero': (By.XPATH, f"{base_xpath}/div/div[3]/div/p[2]"),
                    'Prezzo': (By.XPATH, f"{base_xpath}/div/div[4]/h2")
                }

                # Estrazione con gestione errori
                card_data = {campo: 'N/A' for campo in selettori}
                for campo, (by, selector) in selettori.items():
                    try:
                        elemento = WebDriverWait(driver, 3).until(
                            EC.visibility_of_element_located((by, selector))
                        )
                        card_data[campo] = elemento.text.strip()
                    except Exception as e:
                        print(f"Campo {campo} non trovato: {str(e)}")
                
                # Scroll dinamico
                driver.execute_script(f"""
                    document.evaluate('{base_xpath}', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null)
                        .singleNodeValue.scrollIntoView({{behavior: 'smooth', block: 'center'}});
                """)
                time.sleep(random.uniform(0.3, 0.7))
                
                if all(val != 'N/A' for val in card_data.values()):
                    prodotti.append(card_data)
                    print(f"{COLORI['divisore']}")
                    print(f"{COLORI['info']}📦  Estrazione carta {current_card}/{num_carte}{COLORI['reset']}")
                    print(f"{COLORI['info']}Edizione: {COLORI['reset']}{card_data['Edizione']}")
                    print(f"{COLORI['info']}Titolo: {COLORI['reset']}{card_data['Titolo']}")
                    print(f"{COLORI['info']}Numero: {COLORI['reset']}{card_data['Numero']}")
                    print(f"{COLORI['success']}💵  Prezzo: {card_data['Prezzo']}{COLORI['reset']}")
                    print(COLORI['divisore'])
                else:
                    print(f"⚠️ Problemi con la carta {current_card}, tentativo scroll...")
                    driver.execute_script("window.scrollBy(0, 600)")
                    time.sleep(2)
                
                current_card += 1
                tentativi_falliti = 0

            except NoSuchElementException as e:
                print(f"⚠️ Elemento {current_card} non trovato, tentativo scroll...")
                tentativi_falliti += 1
                
                # Scroll di emergenza
                driver.execute_script("window.scrollBy(0, 500)")
                time.sleep(2)
                
                if tentativi_falliti > 3:
                    print("🚨 Troppi tentativi falliti, interruzione")
                    break

            except Exception as e:
                print(f"🔥 Errore grave: {str(e)}")
                break

        # Salvataggio Excel
        file_path = f'carte_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        df = pd.DataFrame(prodotti)
        
        # Sovrascrittura file
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, index=False, sheet_name='DatiCarte')

        print(f"\n{COLORI['success']}🎉  File salvato: {file_path}{COLORI['reset']}")
        print(f"{COLORI['info']}📊  Carte salvate: {len(prodotti)}/{num_carte}{COLORI['reset']}")
        print(COLORI['divisore'])

    except Exception as e:
        print(f"{COLORI['error']}🔥 Errore critico: {str(e)}{COLORI['reset']}")
    finally:
        if 'driver' in locals():
            print(f"\n{COLORI['divisore']}")
            print(f"{COLORI['info']}🛑 Chiusura browser in corso...{COLORI['reset']}")
            driver.quit()
            print(f"{COLORI['success']}✅ Browser chiuso correttamente{COLORI['reset']}")
            print(COLORI['divisore'])

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1))
def estrai_elemento(container, selectors, campo=''):
    """Estrazione con fallback a regex per i prezzi"""
    for selector in selectors:
        try:
            # Prova con XPath se inizia con //
            if selector.startswith('//'):
                element = container.xpath(selector)
            else:
                element = container.select_one(selector)
            
            if element and element.text.strip():
                testo = element.text.strip()
                
                # Pulizia aggiuntiva per i prezzi
                if campo == 'prezzo':
                    return pulisci_prezzo(testo)
                return testo
        except Exception as e:
            continue
    
    # Fallback a regex per i prezzi
    if campo == 'prezzo':
        testo_container = container.get_text()
        prezzo = trova_prezzo_testuale(testo_container)
        if prezzo:
            return prezzo
    
    raise ValueError(f"Nessun selettore valido: {selectors}")

def pulisci_prezzo(testo):
    """Normalizzazione formati prezzi"""
    testo = testo.replace(',', '.').replace(' ', '').replace('€', '')
    return f"€{float(testo):.2f}"

def trova_prezzo_testuale(testo):
    """Ricerca prezzi via regex avanzata"""
    pattern = r'''
        €?                    # Simbolo valuta opzionale
        \s*                   # Spazi opzionali
        (\d{1,3}(?:\.\d{3})*  # Formato con punti/migliaia
        (?:,\d{2})?           # Decimali con virgola
        |
        \d+                   # Formato senza decimali
        (?:\.\d{2})?)         # Decimali con punto
    '''
    match = re.search(pattern, testo, re.VERBOSE | re.IGNORECASE)
    return f"€{match.group(0)}" if match else None

# Schedulazione giornaliera
if __name__ == "__main__":
    try:
        scarica_dati()
    except KeyboardInterrupt:
        print(f"\n{COLORI['error']}🚨 Interruzione manuale rilevata!{COLORI['reset']}")
    finally:
        print(f"\n{COLORI['divisore']}")
        print(f"{COLORI['success']}🏁 Programma terminato{COLORI['reset']}")
        print(COLORI['divisore'])