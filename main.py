from string import Template
import smtplib
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import os
from os import listdir
from os.path import isfile, join

#Función para leer los contactos de alumnos.txt

def get_contacto(filename):
    '''

    La función coge un archivo como parámetro de entrada, lo abre, lee cada línea
    y divide por el espacio el nombre y el email.
    :param filename:
    :return: dos listas con los nombres y los emails
    '''

    nombres = []
    emails = []
    with open(filename, mode='r', encoding='utf-8') as archivoContactos:
        for contacto in archivoContactos:
            nombres.append(contacto.split()[0])
            emails.append(contacto.split()[1])

    return sorted(nombres), sorted(emails)
nombres, emails = get_contacto('alumnos.txt')
# Función para leer la plantilla de texto

def leer_plantilla(filename):
    with open(filename, mode='r', encoding='utf-8') as archivoPlantilla:
        contenidoPlantilla = archivoPlantilla.read()

    return Template(contenidoPlantilla)

# Función para obtener los archivos de las correcciones.

def leer_directorios():
    '''
    La función coge el working directory, y añade la carpeta correcciones.
    Luego lista los archivos y se queda sólo con aquellos cuya extensión sea .docx.
    Después une el cwd con los archivos y genera el path completo para cada archivo.
    :return: lista con los directorios de cada archivo.
    '''

    cwd = os.getcwd()
    path = cwd + '/correcciones'
    archivos_pre = [f for f in listdir(path) if isfile(join(path, f))]
    archivos_pre = [f for f in archivos_pre if f[-4:] == 'docx']
    archivos_pre = sorted(archivos_pre)
    archivos = []
    nombre_archivos = []

    for archivo in archivos_pre:
        nombre_archivos.append(archivo)
        nombre_archivos = sorted(nombre_archivos)
        archivos.append(path +'/'+ archivo)
    return archivos, nombre_archivos

#Montamos el servidor de SMTP
s = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
s.starttls() #STARTTLS es una extensión a los protocolos de comunicación de texto plano
miCorreo = 'nachoejemplo@hotmail.com'
miContraseña = '1234'
s.login(miCorreo, miContraseña)
enCopia = 'prueba0@bridge.es, prueba1@bridge.es' #string entero separado por comas
asunto = 'Corrección proyecto EDA'

# Juntamos la información de contacto con su respectivo mensaje y archivo a adjuntar (correccion)

nombres, emails = get_contacto('alumnos.txt')
plantillaMensaje = leer_plantilla('mensaje.txt')
listaArchivos, nombre_archivos  = leer_directorios()

for nombre, email, directorio, archivo in zip(nombres, emails, listaArchivos, nombre_archivos):
    msj = MIMEMultipart()  #Generamos el mensaje

    # Añadimos el nombre de cada persona a la plantilla del mensaje
    mensaje = plantillaMensaje.substitute(PERSON_NAME=nombre.title())

    # Añadimos los parámetros del mensaje
    msj['From'] = miCorreo
    msj['To'] = email
    msj['Subject'] = "Corrección proyecto EDA"
    msj['Cc'] = enCopia

    # Añadimos el cuerpo del correo
    msj.attach(MIMEText(mensaje, 'plain'))

    #Añadimos el archivo adjunto
    archivo_adjunto = open(directorio, 'rb')  # Open the file as binary mode
    payload = MIMEBase('application', 'octet-stream')
    '''
    Con esto especificamos el tipo de contenido, en este caso un archivo binario.
    El motivo por el que lo hacemos es porque queremos que el archivo lo abra la aplicación
    de email (gmail, outlook, etc)
    '''

    payload.set_payload(archivo_adjunto.read())
    encoders.encode_base64(payload)  # Base64 es una codificación de binario a texto.
    # add payload header with filename
    payload.add_header('Content-Disposition',  'attachment', filename=archivo)
    msj.attach(payload)

    # Enviamos el mensaje a través del servidor que montamos antes.
    s.send_message(msj)

    del(msj)