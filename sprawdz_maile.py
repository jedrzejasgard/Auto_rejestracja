import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import re
from vendoasg.vendoasg import Vendo
from selenium import webdriver
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import random
from googletrans import Translator
from dodaj_klienta_do_evolve import dodaj_klienta

# połączenie z bazą vendo
vendoApi = Vendo("http://vendo.asgard.pl:5560")
vendoApi.logInApi("esklep","e12345")
vendoApi.loginUser("jpawlewski","jp12345")

# account credentials
username = "rejestracja@asgard.gifts"
password = r"@Asg%rej189#%"
#selenium data

l = 'j.pawlewski'
h = 'Jj123456'
evolve_url = 'https://asgard.gifts/admin/'
dodaj_klienta_url = r'https://asgard.gifts/admin/userEdit/0'


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)


def wyslij_maila_obsluga(mail_klinta,problem):
  password = r"@Asg%rej189#%"
  mail_sender = 'rejestracja@asgard.gifts'
  mail_reciver = 'j.pawlewski@asgard.gifts'
  msg = MIMEMultipart()
  msg['Subject'] = f"Problem z klientem {mail_klinta}"
  msg['From'] = mail_sender
  msg['To'] = mail_reciver
  wiadomosc = f'Klient {mail_klinta} ma taki {problem} '
  part = MIMEText(wiadomosc, 'text')
  msg.attach(part)  
  # Send the message via local SMTP server.
  mail = smtplib.SMTP_SSL('smtp.asgard.gifts', 465)
  mail.login(mail_sender, password)
  mail.sendmail(mail_sender, mail_reciver, msg.as_string())
  mail.quit()
  print('Mail wysłany')

def wyslij_maila_do_handlowca_haslo_PL(mail_klinta,vendo_id,mail_handlowca):
  password = r"@Asg%rej189#%"
  mail_sender = 'rejestracja@asgard.gifts'
  #do testów jestem ja ustawiony
  #mail_reciver = mail_handlowca
  mail_reciver = 'j.pawlewski@asgard.gifts'
  msg = MIMEMultipart()
  msg['Subject'] = f"Problem z klientem {mail_klinta} o ID {vendo_id}"
  msg['From'] = mail_sender
  msg['To'] = mail_reciver
  wiadomosc = f'Klient {mail_klinta} o ID: {vendo_id} już posiada konto pod tym loginem skontaktuj się z opiekunem '
  part = MIMEText(wiadomosc, 'text')
  msg.attach(part)  
  # Send the message via local SMTP server.
  mail = smtplib.SMTP_SSL('smtp.asgard.gifts', 465)
  mail.login(mail_sender, password)
  mail.sendmail(mail_sender, mail_reciver, msg.as_string())
  mail.quit()
  print('Mail wysłany')


def wyslij_maila_bez_h(mail_opiekuna, msg_body):
  password = r"@Asg%rej189#%"
  mail_sender = 'rejestracja@asgard.gifts'
  mail_reciver = 'j.pawlewski@asgard.gifts'
  #mail_opiekuna = [f'{mail_opiekuna}','j.pawlewski@asgard.gifts']
  #mail_opiekuna = ['j.pawlewski@asgard.gifts']
  msg = MIMEMultipart()
  msg['Subject'] = "Aktywacja konta - www.asgard.gifts"
  msg['From'] = mail_sender
  msg['To'] = mail_reciver
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
  part = MIMEText(html, 'html')
  msg.attach(part)  
  # Send the message via local SMTP server.
  mail = smtplib.SMTP_SSL('smtp.asgard.gifts', 465)
  mail.login(mail_sender, password)
  odbiorcy = [mail_reciver]+mail_opiekuna
  mail.sendmail(mail_sender, odbiorcy, msg.as_string())
  mail.quit()
  print('Mail wysłany z proźbą o wpisanie do ')


def zapytanie_do_V(zmienna,dane):
  zapytanie_klient_vendo = ''
  if zmienna == 'NIP':
    zapytanie_klient_vendo = vendoApi.getJson(
      '/json/reply/CRM_Klienci_KlientRozszerzony',
      {"Token":vendoApi.USER_TOKEN,"Model":
      {"Nip":dane}})
  if zmienna == 'login':
    zapytanie_klient_vendo = vendoApi.getJson(
      '/json/reply/CRM_Klienci_KlientRozszerzony',
      {"Token":vendoApi.USER_TOKEN,"Model":
      {"Kod":dane}})
  return zapytanie_klient_vendo['Wynik']['Rekordy']


def wpisanie_do_bazy_evolve(vendo_id,login,imie,nazwisko):
  h = 'Wzwert1x'
  evolve_url = 'https://asgard.gifts/admin/'
  dodaj_klienta_url = r'https://asgard.gifts/admin/userEdit/0'
  chrome = webdriver.Chrome(r'C:\Users\asgard_48\Documents\chromedriver_win32\chromedriver.exe')
  chrome.get(evolve_url)
  chrome.maximize_window()
  chrome.find_element_by_name('uLogin').send_keys(l)
  chrome.find_element_by_name('uPasswd').send_keys(h)
  chrome.find_element_by_xpath('/html/body/div[2]/div[1]/form/button').click()
  chrome.get(dodaj_klienta_url)
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
  chrome.find_element_by_name('SaveClose').click()
  chrome.close()


def znajdz_opiekuna(kraj,kod_pocztowy):
  opiekun = []
  opiekPL = {('łódzkie','świętokrzyskie','opolskie','śląskie','małopolskie','podkarpackie'):['AgnieszkaSitek','TomaszPiszczola'],
  ('zachodniopomorskie','wielkopolskie','lubuskie','dolnośląskie'):['MagdalenaMikolajczyk','MarlenaKluszczynska'],
  ('pomorskie','warmińsko-mazurskie','podlaskie','mazowieckie','kujawsko-pomorskie'):['IwonaRosinska','CezaryIdziak']}
  
  opiekunowieWorld = {'JakubNieglos':['łotwa','litwa','rumunia','węgry','grecja','bułgaria'],
  'AleksandraMuszynska':['portugalia','hiszpania','francja','włochy','belgia','luksemburg','holandia','słowenia','chorwacja'],
  'MonikaBujakowska':['anglia','irlandia','austria','niemcy','luksemburg','dania','szwecja','finlandia','austria'],
  'MariannaPrange':['portugalia','hiszpania','francja','włochy','szwajcaria','belgia','holandia','norwegia','czechy','słowenia','chorwacja'],
  'PawelStrzelecki':['łotwa','litwa','ukraina','rumunia','węgry','bułgaria'],
  'LukaszUrbanczyk':['szwajcaria','austria','niemcy','dania','szwecja','finlandia','czechy','austria','słowacja',],
  'MalgorzataSobaszek':[],
  'AdriannaTrafna':[],}
  with open ('kody.csv','rb') as f:
    df_kody = pd.read_csv(f, sep=';',usecols=['KOD POCZTOWY','WOJEWÓDZTWO'])
  if kraj.lower() == 'polska':
    for index,row in df_kody.iterrows():
      if kod_pocztowy == str(row['KOD POCZTOWY']):
        wojewodztwo = row['WOJEWÓDZTWO'].lower().split()[1]
                #print(wojewodztwo)
        for key,value in opiekPL.items():
          if wojewodztwo in key:
            if len(opiekun) == 0:           
              opiekun.append(value[random.randint(0,1)])
    
  else:
    translator = Translator()
    tlumaczenie_kraj = translator.translate(f'{kraj}',dest='pl').text
    for key,value in opiekunowieWorld.items():
      if tlumaczenie_kraj in value:
        opiekun.append(key)
      else:
        print('Nie znalazlem')
         #print(key)
  
  
  imie = re.findall(r'(.*)[A-Z][a-z]*$',opiekun[0])[0]
  nazwisko = re.findall(r'([A-Z][a-z]*$)',opiekun[0])[0]
  mail_wysylki = f'{imie[0].lower()}.{nazwisko.lower()}@asgard.gifts'
  print(mail_wysylki)
  return mail_wysylki

      
# create an IMAP4 class with SSL 
imap = imaplib.IMAP4_SSL("imap.asgard.gifts")
# authenticate
imap.login(username, password)

status, messages = imap.select("INBOX")
# number of top emails to fetch
N = 50
# total number of emails
messages = int(messages[0])
print(messages)

for i in range(messages, messages-N, -1):
  msg_body = ''
  # fetch the email message by ID
  try:
      res, msg = imap.fetch(str(i), "(RFC822)")
  except:
      print('niemoglem odczytać maila')
      break
  czy_sprawdzac_mail = 1
  for response in msg:        
      if isinstance(response, tuple):
          print('*')
          # parse a bytes email into a message object
          msg = email.message_from_bytes(response[1])
          # decode the email subject
          subject, encoding = decode_header(msg["Subject"])[0]
          if isinstance(subject, bytes):
              # if it's a bytes, decode to str
              subject = subject.decode(encoding).strip()
          # decode email sender
          From, encoding = decode_header(msg.get("From"))[0]
          if isinstance(From, bytes):
              From = From.decode(encoding).strip()

          if From != 'Asgard <rejestracja@asgard.gifts>' and subject !='Zgłoszono nowe konto / Admin':
              print('Niezgadza się temat i nadawca')
              czy_sprawdzac_mail = 0
              break            
          print("="*100)
          print("Subject:", subject)
          print("From:", From)
          # if the email message is multipart
          if msg.is_multipart():
              # iterate over email parts
              for part in msg.walk():
                  # extract content type of email
                  content_type = part.get_content_type()
                  content_disposition = str(part.get("Content-Disposition"))
                  try:
                      # get the email body
                      msg_body = part.get_payload(decode=True).decode()
                  except:
                      pass
          else:
              # extract content type of email
              content_type = msg.get_content_type()
              # get the email body
              msg_body = msg.get_payload(decode=True).decode()
          #print(body)
       
  if czy_sprawdzac_mail:
      translator = Translator()
      msg_body = translator.translate(f'{msg_body}',dest='pl').text 
      #print(type(msg_body))
      login = re.findall(r"Login:(.*)<br />",msg_body)[0].strip()
      kraj = re.findall(r"Kraj:(.*)<br />",msg_body)[0].strip()
      NIP = re.findall(r"NIP:\s*\D*(\d*)",msg_body)[0].strip()
      imie = re.findall(r"Imię: (.*)<br />",msg_body)[0].strip()
      nazwisko = re.findall(r"Nazwisko: (.*)<br />",msg_body)[0].strip()
      kod_pocztowy = re.findall(r"Kod: (.*)<br />",msg_body)[0].strip()
      print(login)
      print(kraj)
      print(NIP)
      czy_klient_w_vendo = zapytanie_do_V('NIP',NIP)
      #print(czy_klient_w_vendo)       
      if len(czy_klient_w_vendo) != 0:
        czy_kod_klienta_zajety = zapytanie_do_V('login',login)
        if len(czy_kod_klienta_zajety) == 0:
          """
          Klient jest założony w Vendo pod tym numerem NIP i możemy zakładać konto
          """
          vendo_id = czy_klient_w_vendo[0]['Klient']['ID']
          print(vendo_id)
          print('zakładam konto w Evolve')
          dodaj_klienta(login,vendo_id,imie,nazwisko)
          continue
        else:
          problem = ''
          wyslij_maila_obsluga(login,problem)
      czy_kod_klienta_zajety = zapytanie_do_V('login',login)
      if len(czy_kod_klienta_zajety) != 0:
        """
        Dany mail jest już wykożystany w bazie Vendo i nie można go użyć
        Wyślij maila do JP żeby sprawdzićdaną sytuację ręcznie.
        Dodaj sprawdzanie i wysyłanie do handlowca powiadomienia jak jego klient wykonuje takie rzeczy
        """
        print('Klient juz istnieje')
        vendo_id = czy_kod_klienta_zajety[0]['Klient']['ID']
        print(f'konto klienta w vendo ma id: {vendo_id}')
        problem = f'konto o takim samym KODzie w Vendo ma id: {vendo_id}'
        wyslij_maila_obsluga(login,problem)
      else:
        """
        Wyslij maila do wybranego opiekuna o proźbe dopisania danego klienta do Vendo
        Kraj: Polska, dzialający skrypt w oparciu o kod pocztowy.
        Reszta świata ręczne tłumaczenie na j polski ?
        """
        try:
          znajdz_opiekuna(kraj,kod_pocztowy)
          print('wyslij maila z proźbą o wpisanie do Vendo')
        except:
          problem = 'Nie mogę znaleść opiekuna dla tego klienta'
          wyslij_maila_obsluga(login,problem)   
        #Vendo_id = odp[0]['Klient']['ID']
          #print(Vendo_id)
          # Create the body of the message (a plain-text and an HTML version).

        print("="*100)
    
# close the connection and logout
imap.close()
imap.logout()