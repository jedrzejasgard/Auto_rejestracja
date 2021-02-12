"""
Program na podstawie podanego ID klienta z Vendo zakłada karte klienta oraz wysyła maila do handlowca oraz klienta
"""
import configparser
from email.header import decode_header
from vendoasg.vendoasg import Vendo
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from time import sleep


config = configparser.ConfigParser()
config.read('auto_rejestracja.ini')

l = config.get('evolve','user')
had = config.get('evolve','pass')


# połączenie z bazą vendo
vendoApi = Vendo(config.get('vendo','vendo_API_port'))
vendoApi.logInApi(config.get('vendo','logInApi_user'),config.get('vendo','logInApi_pass'))
vendoApi.loginUser(config.get('vendo','loginUser_user'),config.get('vendo','loginUser_pass'))

def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

def opiekun(vendo_id):
  zapytanie_klient_vendo = vendoApi.getJson(
      '/json/reply/CRM_Klienci_KlientRozszerzony',
      {"Token":vendoApi.USER_TOKEN,"Model":
      {"ID":vendo_id}})
  return zapytanie_klient_vendo['Wynik']['Rekordy'][0]['Klient']['PracownikOpiekun']['Kod']


def update_kod(vendo_id,login):
  update_kod_klienta_vendo = vendoApi.getJson(
      '/json/reply/CRM_Klienci_Aktualizuj',
      {"Token":vendoApi.USER_TOKEN,"Model":{"NowyKod": login
      ,"ID":vendo_id}})
  print(update_kod_klienta_vendo)
  #return update_kod_klienta_vendo['Wynik']['Rekordy']


def wyslij_maila_z_h(mail_klinta,mail_opiekuna,jezyk_maila):
  password = config.get('rejestracja','pass')
  mail_sender = 'rejestracja@asgard.gifts'
  mail_reciver = mail_klinta
  #mail_reciver = 'j.pawlewski@asgard.gifts'
  mail_opiekuna = [f'{mail_opiekuna}','j.pawlewski@asgard.gifts']
  #mail_opiekuna = ['j.pawlewski@asgard.gifts']
  msg = MIMEMultipart()
  msg['Subject'] = "Aktywacja konta - www.asgard.gifts"
  msg['From'] = mail_sender
  msg['To'] = mail_reciver
  if jezyk_maila == 'pl':
    html = f"""\
        <html>
        <head>
      
          <meta http-equiv="content-type" content="text/html; charset=UTF-8">
        </head>
        <body>
          <table width="800" cellspacing="2" cellpadding="2" border="0"
          align="center">
              <tbody>
                  <title>Aktywacja konta - www.asgard.gifts</title>
                  <tr>
                      <td>
                          <br>
                          <br>
                          <br>
                          <br>
                          <br>
                          <br>
                    <img src="https://asgard.gifts/www/stopki/tworzenie_kont/re_register_01.jpg" width="800" height="113"><br>
                  </td>
                </tr>
                <tr>
                  <td valign="middle" align="left"><font color="#273a65"><br>
                      <big>&nbsp;&nbsp;<font color="#000066">&nbsp; Drogi
                          Kliencie,<br>
                          <br>
                        </font> <font color="#000066"> &nbsp;&nbsp;&nbsp;
                          Twoje konto w serwisie <a moz-do-not-send="true"
                            class="moz-txt-link-abbreviated"
                            href="https://www.asgard.gifts">www.asgard.gifts</a>
                          zostało utworzone i <u>jest już aktywne.</u><br>
                          <br>
                          &nbsp;&nbsp;&nbsp; Od teraz prosimy logować się za
                          pomocą poniższego loginu.<br>
                          <br>
                        </font> <font color="#000066"> &nbsp;&nbsp;&nbsp;
                          Login: </font></big></font> <a
                      class="moz-txt-link-abbreviated"
                      href="mailto:{mail_klinta}">{mail_klinta}
                    </a> <font color="#273a65"><big><font color="#000066"><br>
                          &nbsp;&nbsp;&nbsp; Hasło: Wzwert1x<br>
                        </font> <font color="#000066"> <small>&nbsp;&nbsp;&nbsp;
      
                            <small>( Po pierwszym logowaniu prosimy zmienić
                              hasło. Opcja zmiany hasła dostępna jest w prawym
                              górnym rogu strony. )</small></small></font></big></font><font
                      color="#000066"> </font> <br>
                    <br>
                  </td>
                </tr>
                <tr>
                  <td valign="top"><img src="https://asgard.gifts/www/stopki/tworzenie_kont/re_register_03.jpg" width="800"
                      height="502"></td>
                </tr>
              </tbody>
            </table>
            <br>
          </div>
          <br>
          <br>
          <p>
            <meta http-equiv="content-type" content="text/html; charset=UTF-8">
          </p>
        </body>
      </html>
      """
  else:
    html = f"""
        <html>
          <head>

            <meta http-equiv="content-type" content="text/html; charset=UTF-8">
          </head>
          <body>
            <p> </p>
            <p>
              <meta http-equiv="content-type" content="text/html; charset=UTF-8">
              <title>Account activation - www.asgard.gifts</title>
              <br>
              <div class="moz-forward-container">
                <meta http-equiv="content-type" content="text/html;
                  charset=UTF-8">
                <br>
                <br>
                <br>
                <table width="800" cellspacing="2" cellpadding="2" border="0"
                  align="center">
                  <tbody>
                    <tr>
                      <td valign="top"><img src="https://asgard.gifts/www/stopki/tworzenie_kont/re_register_ang_01.jpg"
                          width="800" height="113"><br>
                      </td>
                    </tr>
                    <tr>
                      <td valign="middle" align="left"><font color="#273a65"><br>
                        </font><font color="#273a65"><br>
                          <big>&nbsp;&nbsp;<font color="#000066">&nbsp; Dear Sir
                              or Madam,<br>
                              <br>
                            </font> <font color="#000066"> &nbsp;&nbsp;&nbsp;
                              Your account in <a moz-do-not-send="true"
                                class="moz-txt-link-abbreviated"
                                href="https://www.asgard.gifts">www.asgard.gifts</a>
                              has been created and is <u>active</u>.<br>
                              <br>
                              &nbsp;&nbsp;&nbsp; From now on you may log in
                              using given credentials.<br>
                              <br>
                            </font> <font color="#000066"> &nbsp;&nbsp;&nbsp;
                              User: <a moz-do-not-send="true"
                                class="moz-txt-link-abbreviated" href="">{mail_klinta}</a></font></big></font>
                        <font color="#273a65"><big><font color="#000066"><br>
                              &nbsp;&nbsp;&nbsp; Password:&nbsp; Wzwert1x</font></big></font><font
                          color="#273a65"><big><font color="#000066"><br>
                            </font></big></font><font color="#273a65"><big><font
                              color="#000066"><br>
                            </font> <font color="#000066"> <small>&nbsp;&nbsp;&nbsp;

                                <small>( After first loging in please change
                                  given password to your own. You may change the
                                  password in upper right corner of the<br>
                                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; website.
                                  )</small></small></font></big></font><font
                          color="#000066"> </font> <br>
                        <br>
                      </td>
                    </tr>
                    <tr>
                      <td valign="top"><img src="https://asgard.gifts/www/stopki/tworzenie_kont/re_register_ang_03.jpg"
                          width="800" height="502"></td>
                    </tr>
                  </tbody>
                </table>
                <br>
              </div>
              <br>
            </p>
          </body>
        </html>
        """
  part = MIMEText(html, 'html')
  msg.attach(part)  
  # Send the message via local SMTP server.
  mail = smtplib.SMTP_SSL('smtp.asgard.gifts', 465)
  mail.login(mail_sender, password)
  odbiorcy = [mail_reciver]+mail_opiekuna
  mail.sendmail(mail_sender, odbiorcy, msg.as_string())
  mail.quit()
  print('Mail wysłany z haslem')

def wyslij_maila_bez_h(mail_klinta,mail_opiekuna,jezyk_maila):
  password = config.get('rejestracja','pass')
  mail_sender = 'rejestracja@asgard.gifts'
  mail_reciver = mail_klinta
  #mail_reciver = 'j.pawlewski@asgard.gifts'
  mail_opiekuna = [f'{mail_opiekuna}','j.pawlewski@asgard.gifts']
  #mail_opiekuna = ['j.pawlewski@asgard.gifts']
  msg = MIMEMultipart()
  msg['Subject'] = "Aktywacja konta - www.asgard.gifts"
  msg['From'] = mail_sender
  msg['To'] = mail_reciver
  if jezyk_maila == 'pl':
    html = f"""\
        <html>
        <head>

          <meta http-equiv="content-type" content="text/html; charset=UTF-8">
        </head>
        <body>
          <p> </p>
          <p>
            <meta http-equiv="content-type" content="text/html; charset=UTF-8">
          </p>
          <p>
            <title>Aktywacja konta - www.asgard.gifts</title>
            <br>
          </p>
          <div class="moz-forward-container"> <br>
            <br>
            <table width="800" cellspacing="2" cellpadding="2" border="0"
              align="center">
              <tbody>
                <tr>
                  <td valign="top"><img src="https://asgard.gifts/www/stopki/tworzenie_kont/re_register_01.jpg" width="800"
                      height="113"><br>
                  </td>
                </tr>
                <tr>
                  <td valign="middle" align="left"><font color="#273a65"><br>
                      <big>&nbsp;&nbsp;<font color="#000066">&nbsp; Drogi
                          Kliencie,<br>
                          <br>
                        </font> <font color="#000066"> &nbsp;&nbsp;&nbsp;
                          Twoje konto w serwisie <a moz-do-not-send="true"
                            class="moz-txt-link-abbreviated"
                            href="https://www.asgard.gifts">www.asgard.gifts</a>
                          zostało utworzone i <u>jest już aktywne.</u><br>
                          <br>
                          &nbsp;&nbsp;&nbsp; Od teraz prosimy logować się za
                          pomocą poniższego loginu.<br>
                          <br>
                        </font> <font color="#000066"> &nbsp;&nbsp;&nbsp;
                          Login: </font></big></font><a
                      class="moz-txt-link-abbreviated"
                      href="mailto:{mail_klinta}">{mail_klinta}
                      <br>
                      <br>
                    </a><font <br=""> <br>
                    </font></td>
                </tr>
                <tr>
                  <td valign="top"><img src="https://asgard.gifts/www/stopki/tworzenie_kont/re_register_03.jpg" width="800"
                      height="502"></td>
                </tr>
              </tbody>
            </table>
            <br>
          </div>
          <br>
          <br>
          <p>
            <meta http-equiv="content-type" content="text/html; charset=UTF-8">
          </p>
        </body>
      </html>
      """
  else:
    html = f"""
        <html>
        <head>

          <meta http-equiv="content-type" content="text/html; charset=UTF-8">
        </head>
        <body>
          <p> </p>
          <p>
            <meta http-equiv="content-type" content="text/html; charset=UTF-8">
            <title>Account activation - www.asgard.gifts</title>
            <br>
            <div class="moz-forward-container">
              <meta http-equiv="content-type" content="text/html;
                charset=UTF-8">
              <br>
              <br>
              <br>
              <table width="800" cellspacing="2" cellpadding="2" border="0"
                align="center">
                <tbody>
                  <tr>
                    <td valign="top"><img src="https://asgard.gifts/www/stopki/tworzenie_kont/re_register_ang_01.jpg"
                        width="800" height="113"><br>
                    </td>
                  </tr>
                  <tr>
                    <td valign="middle" align="left"><font color="#273a65"><br>
                      </font><font color="#273a65"><br>
                        <big>&nbsp;&nbsp;<font color="#000066">&nbsp; Dear Sir
                            or Madam,<br>
                            <br>
                          </font> <font color="#000066"> &nbsp;&nbsp;&nbsp;
                            Your account in <a moz-do-not-send="true"
                              class="moz-txt-link-abbreviated"
                              href="https://www.asgard.gifts">www.asgard.gifts</a>
                            has been created and is <u>active</u>.<br>
                            <br>
                            &nbsp;&nbsp;&nbsp; From now on you may log in
                            using given credentials.<br>
                            <br>
                          </font> <font color="#000066"> &nbsp;&nbsp;&nbsp;
                            User: <a moz-do-not-send="true"
                              class="moz-txt-link-abbreviated" href="">{mail_klinta}</a></font></big></font>
                      <font color="#000066"> </font> <br>
                      <br>
                    </td>
                  </tr>
                  <tr>
                    <td valign="top"><img src="https://asgard.gifts/www/stopki/tworzenie_kont/re_register_ang_03.jpg"
                        width="800" height="502"></td>
                  </tr>
                </tbody>
              </table>
              <br>
            </div>
            <br>
          </p>
        </body>
      </html>
        """
  part = MIMEText(html, 'html')
  msg.attach(part)  
  # Send the message via local SMTP server.
  mail = smtplib.SMTP_SSL('smtp.asgard.gifts', 465)
  mail.login(mail_sender, password)
  odbiorcy = [mail_reciver]+mail_opiekuna
  mail.sendmail(mail_sender, odbiorcy, msg.as_string())
  mail.quit()
  print('Mail wysłany bez hasla')

def wpisanie_do_bazy_evolve(vendo_id,login,imie,nazwisko):
  #dodaj sprawdzenie czy dany mail jest już zajęty
  jezyk = ''
  typ_maila = ''
  h = config.get('evolve','defolt_pass')
  evolve_url = 'https://asgard.gifts/admin/'
  klienci_lista_url = 'https://asgard.gifts/admin/userList'
  dodaj_klienta_url = 'https://asgard.gifts/admin/userEdit/0'
  chrome = webdriver.Chrome(r'C:\Users\asgard_48\Documents\chromedriver_win32\chromedriver.exe')
  chrome.get(evolve_url)
  chrome.maximize_window()
  chrome.find_element_by_name('uLogin').send_keys(l)
  chrome.find_element_by_name('uPasswd').send_keys(had)
  chrome.find_element_by_xpath('/html/body/div[2]/div[1]/form/button').click()
  chrome.get(klienci_lista_url)
  input_box = chrome.find_element_by_xpath('/html/body/div[2]/section[2]/div[4]/div/div/div/div/div[3]/div/table/thead/tr[2]/td[3]/input')
  for character in login:
    actions = ActionChains(chrome)
    actions.move_to_element(input_box)
    actions.click()
    actions.send_keys(Keys.END)
    actions.send_keys(character)
    actions.perform()
  brak = 'znalazlem klienta'
  try:
    brak = chrome.find_element_by_class_name('no_result').text
  except:
    print(brak)
  if brak == 'BRAK WYNIKÓW':
    chrome.get(dodaj_klienta_url)
    sleep(1)
    chrome.find_element_by_id('uLogin').send_keys(login)
    chrome.find_element_by_id('vendoID').send_keys(vendo_id)
    chrome.find_element_by_id('uPasswd').send_keys(h)
    chrome.find_element_by_id('uPassCnfrm').send_keys(h)
    chrome.find_element_by_id('uFName').send_keys(imie)
    chrome.find_element_by_id('uLName').send_keys(nazwisko)
    chrome.find_element_by_id('uEmail').send_keys(login)
    chrome.find_element_by_id('uContactEmail').send_keys(login)
    chrome.find_element_by_xpath('/html/body/div[2]/section[2]/div[4]/div/div/form/div[2]/div[2]/div/div/button[1]').click()
    chrome.find_element_by_xpath('/html/body/div[2]/section[2]/div[5]/div/div/form[1]/div[2]/div[2]/div/div/button[4]').click()
    chrome.find_element_by_id('uLockShop').click()
    jezyk = Select(chrome.find_element_by_id('uLang'))
    jezyk = jezyk.first_selected_option.text
    typ_maila = 'haslo'
    chrome.find_element_by_name('SaveClose').click()
    print('Klient wpisany do evolve')
  else:
    chrome.find_element_by_xpath('/html/body/div[2]/section[2]/div[4]/div/div/div/div/div[3]/div/table/tbody/tr/td[3]/a').click()
    wpisane_id = chrome.find_element_by_id('vendoID').get_property('value')
    if wpisane_id == '0':
      chrome.find_element_by_id('vendoID').clear()
      chrome.find_element_by_id('vendoID').send_keys(vendo_id)
      chrome.find_element_by_id('active').click()
      chrome.find_element_by_xpath('/html/body/div[2]/section[2]/div[4]/div/div/form/div[2]/div[2]/div/div/button[1]').click()
      chrome.find_element_by_xpath('/html/body/div[2]/section[2]/div[5]/div/div/form[1]/div[2]/div[2]/div/div/button[4]').click()
      chrome.find_element_by_id('uLockShop').click()
      jezyk = Select(chrome.find_element_by_id('uLang'))
      jezyk = jezyk.first_selected_option.text
      chrome.find_element_by_name('SaveClose').click()
      typ_maila = 'bez_h'
      print('Klient wpisany do evolve')
    else:
      print('Klient o takim loginie jest już wpisany')
      pass

    
  
  
  #try:
  #  chrome.find_element_by_xpath('/html/body/div[2]/section[2]/div[4]/div/div/div/div/div[3]/div/table/tbody[1]/tr/td[3]/a').click()
  #except:
  #  chrome.get(dodaj_klienta_url)
  
  chrome.close()
  return [jezyk, typ_maila]


def zmiana_KOD(login,vendo_id):
  login = str(login).strip().upper()
  vendo_id = int(vendo_id)
  #dodaj sprawdzenie czy dany mail jest już zajęty
  zmien_kod_klienta = vendoApi.getJson(
      '/json/reply/CRM_Klienci_Aktualizuj',
      {"Token":vendoApi.USER_TOKEN,"Model":
      {"NowyKod": login,"ID":vendo_id}})
  print(zmien_kod_klienta)


def dodaj_klienta(login,vendo_id,imie,nazwisko):
  opiekuna = list(str(opiekun(vendo_id)))
  koniec = ''.join(opiekuna[1:])
  mail_opiekuna = f'{opiekuna[0]}.{koniec}@asgard.gifts'
  print(mail_opiekuna)
  #update_kod(vendo_id, login)
  jaki_maila = wpisanie_do_bazy_evolve(vendo_id,login,imie,nazwisko)
  if jaki_maila[1] == 'haslo':
    wyslij_maila_z_h(login,mail_opiekuna,jaki_maila[0])
  elif jaki_maila[1] == 'bez_h':
    wyslij_maila_bez_h(login,mail_opiekuna,jaki_maila[0])
  else:
    print('problem z mailami')

if __name__ == "__main__":
  try:
    dodaj_klienta(login,vendo_id,imie,nazwisko)
  except:
    login = input('Login: ')
    vendo_id = input('Vendo ID: ')
    imie = input('Imie: ')
    nazwisko = input('Nazwisko: ')
    dodaj_klienta(login,vendo_id,imie,nazwisko)
    zmiana_KOD(login,vendo_id)
  
    ###Wyślij maila z errorem do JP