#LIBRERIAS

# Instalamos e importamos la libreria win32com (se usa para automatizar con Office)
import win32com.client

# Movemos el archivo CSV al path (pythonPrueba) y lo importamos para usarlo
import csv

#Creamos la instancia para usar en este entorno
outlook = win32com.client.Dispatch('outlook.application')

#LEER ARCHIVO CSV Y DEFINIR VARIABLES

with open('csv practica.csv') as my_file:
    reader = csv.DictReader(my_file, delimiter=';')      #Para saber cual es el delimitador abrir el csv con wordpad
    for row in reader:                                   #Creamos un bucle
        name_in_csv = row['Nombre']                      #Definimos los nombres de las columnas
        emailTo_in_csv = row['Email']
        emailCc_in_csv = row['EmailCopia']
        surname_in_csv = row['Apellido']
        print(name_in_csv+" "+emailTo_in_csv)              #Imprimimos para ver los nombres y los mails

#EMAIL

        mail = outlook.CreateItem(0)                     #Creamos objeto email
        mail.To = ''+emailTo_in_csv           #Destinatario
        mail.Cc = ''+emailCc_in_csv           #Destinatario Copia
        mail.Subject = 'Factura de Honorarios - '+name_in_csv     #Asunto + Nombre declarado en csv

    #Cuerpo Email
#Cree un archivo txt para usarlo como cuerpo del email

        template_text = open('CuerpoMensaje.txt', 'r').read()    #Abrimos y leemos el txt
        final_mail_body = template_text.format(name_in_csv,"OCTUBRE")    #Definimos las referencias ( {} )
        mail.Body = final_mail_body           #Cuerpo del mail

    #Adjuntar archivos

#Cree una carpeta en el path con los archivos en pdf cuyos nombres son los mismos apellidos de sus destinatarios
#Para que cada destinatario tenga su archivo adjunto correspondiente

        file_path = r"C:\\pythonPrueba\Archivos adjuntos\\"+surname_in_csv+".pdf"    #Ruta de acceso a carpeta
        mail.attachments.Add(file_path)


        mail.Display()                                   #Ejecuta la función sin enviar el e-mail
        #mail.Send()                                     #Funcion para enviar directamente sin revisar los emails

