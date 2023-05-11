import requests
import os
import re
from bs4 import BeautifulSoup
from docx import Document
from openpyxl import Workbook
import datetime
from email.message import EmailMessage
import ssl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


#Ingreso de la pagina que se desea analizar y el correo al que se desea enviar. 
#Se ingresa la pagina que desea analizar el usuario
pagina_Web = input("Ingrese la pagina que desea analizar: ")
#Se ingresa a que correo se va a enviar el mensaje junto a los archivos
correo_Receptor = input("Ingrese el correo al cual desea enviar el analisis de la pagina: ")

#Aqui se descarga la informacion de la pagina
respuesta_Pagina = requests.get(pagina_Web)

# Crear un objeto BeautifulSoup con el contenido de la página para realizar el analisis
soup = BeautifulSoup(respuesta_Pagina.content, 'html.parser')

#Aqui se recopila unicamente el texto principal de la pagina asi evitando llevarse basura
textoContenidoPagina = ''
for objeto in soup.find_all('div', class_='mw-parser-output'):
    for parrafo in objeto.find_all('p'):
        textoContenidoPagina += parrafo.text

#Se obtienen las palabras del texto principal pero se excluyen las etiquetas HTML para tener un texto mas limpio
palabras = re.findall(r'\b\w+\b', textoContenidoPagina)

#
#Guardado texto en los archivos
#

#Aqui se guarda el texto recopilado de la pagina ya procesado en
#un archivo de tipo ".docx"
archivoDOCX = Document()
archivoDOCX.add_paragraph(textoContenidoPagina)
archivoDOCX.save('Texto_de_la_pagina_web.docx')

#Todas las se convierten en minusculas para que a la hora de realizar elconteo no se cuente diferente palabras 
#Que son lo mismo como puede ser "El" y "el".
minusculasPalabras = [palabra.lower() for palabra in palabras]
numeroPalabras = {}
for palabra in minusculasPalabras:
    if palabra in numeroPalabras:
        numeroPalabras[palabra] += 1
    else:
        numeroPalabras[palabra] = 1

#Aqui se guaran las palabras que contiene el texto y el numero de veces que se repiten
#en un archivo de tipo ".xlsx".
hojaExelPalabras = Workbook()
hoja = hojaExelPalabras.active
hoja.title = 'Conteo_de_palabras'
hoja['A1'] = 'Palabra'
hoja['B1'] = 'Veces'
fila = 2
for palabra, conteo in numeroPalabras.items():
    hoja[f'A{fila}'] = palabra
    hoja[f'B{fila}'] = conteo
    fila += 1
hojaExelPalabras.save('Conteo_de_palabras.xlsx')

#Se analiza el texto y se obtiene el numero de palabras que contiene
tot_Palabras = len(palabras)

#
#Descargar las imagenes de la pagina
#

#Esta funcion es necesaria para descargar las imagenes
def descargarImagenes(url, nombreArchivo):
    respuesta_Pagina = requests.get(url)
    if respuesta_Pagina.status_code == 200:
        with open(nombreArchivo, 'wb') as f:
            f.write(respuesta_Pagina.content)

#Se descargan las imagenes en una carpeta
if not os.path.exists('Imagenes_de_la_pagina'):
    os.mkdir('Imagenes_de_la_pagina')
for imagen in soup.find_all('img'):
    URL_de_la_Imagen = imagen.get('src')
    if URL_de_la_Imagen.startswith('//'):
        URL_de_la_Imagen = 'https:' + URL_de_la_Imagen
    if URL_de_la_Imagen.startswith('http'):
        nombreArchivo = os.path.basename(URL_de_la_Imagen)
        #Algunas imagenes no se pueden descargar y estas mismas se saltan para evitar errores
        if "start?type=1x1" not in URL_de_la_Imagen:
            descargarImagenes(URL_de_la_Imagen, f'Imagenes_de_la_pagina/{nombreArchivo}')

#Se obtiene la fecha y hora en la que se realizo el analisis
Fecha_Hora_Del_Correo = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

#
#Despues de realizar el analisis de datos y guardarlos en documentos se envian los archivos ".docx" y ".xlsx"
# al correo ingresado por el usuario
#

correo = f'Despues de analizar la pagina se han encontrado "{tot_Palabras}" ' \
         f'palabras en la pagina ingresada la cual es:{pagina_Web} y la hora y fecha ' \
         f'en la que se realizo el analisis es:{Fecha_Hora_Del_Correo}. En el archivo DOCX "docs" se ' \
         f'almaceno el texto total de la pagina y en el "xlsx" se registro cada palabra con el numero de veses que se repite.'

#Datos necesarios para el correo
correo_Emisor = 'email'
contrasena_Correo_Emisor = 'Password'
asunto = f'Analisis de la pagina {pagina_Web}'

#Se empieza a organizar la informacion para enviar el correo
informacion_del_correo = MIMEMultipart()
informacion_del_correo['From'] = correo_Emisor
informacion_del_correo['To'] = correo_Receptor
informacion_del_correo['Subject'] = asunto

#Se adjuntan los archivos en el correo
informacion_del_correo.attach(MIMEText(correo))

archivo = 'Conteo_de_palabras.xlsx'
with open(archivo, 'rb') as f:
    adjunto = MIMEApplication(f.read(), _subtype='xls')
    adjunto.add_header('content-disposition', 'attachment', filename=archivo)
informacion_del_correo.attach(adjunto)

archivo2 = 'Texto_de_la_pagina_web.docx'
with open(archivo2, 'rb') as f:
    adjunto2 = MIMEApplication(f.read(), _subtype='docx')
    adjunto2.add_header('content-disposition', 'attachment', filename=archivo2)
informacion_del_correo.attach(adjunto2)

# Conectar al servidor SMTP y enviar correo
servidor_smtp = smtplib.SMTP('smtp.office365.com', 587)
servidor_smtp.starttls()
servidor_smtp.login(correo_Emisor, contrasena_Correo_Emisor)
servidor_smtp.sendmail(correo_Emisor, correo_Receptor, informacion_del_correo.as_string())
servidor_smtp.quit()