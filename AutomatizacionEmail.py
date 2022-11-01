# Instalamos e importamos la libreria win32com          #EXPLICAR
import win32com.client

# Movemos el archivo CSV al path (pythonPrueba) y lo importamos para usarlo
import csv

# EXPLICAR
outlook = win32com.client.Dispatch('outlook.application')

with open('csv practica.csv') as my_file:
    reader = csv.DictReader(my_file, delimiter=';')      #Para saber cual es el delimitador abrir el csv con wordpad
    for row in reader:                                   #Creamos un bucle
        name_in_csv = row['Nombre']                      #Definimos los nombres de las columnas
        email_in_csv = row['Email']
        print(name_in_csv+" "+email_in_csv)              #Imprimimos para ver los nombres y los mails

        mail = outlook.CreateItem(0)                     #Creamos objeto mail y mensaje
        mail.To = 'administracion@gaston.com'           #Destinatario
        mail.Cc = 'administracion@gaston.com'           #Destinatario Copia
        mail.Subject = 'Prueba auto-Email -'+name_in_csv #Asunto + nombre del valor en csv
        mail.Body = "Este es un cuerpo normal"           #Cuerpo del mail
        mail.Display()                                   #Ejecuta la función sin enviar el e-mail



