import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import xlwings as xw
import toml

# Charger la configuration depuis le fichier TOML
config = toml.load("config.toml")

# Accéder aux valeurs de la configuration
sender = config["email"]["sender"]
password = config["email"]["password"]

def print_data(path, df, status):
    """
    Print the status of the dataFrame in the Excel sheet
    :param path: path of the wordbook
    :param df: dataFrame
    :param column_status: column status of the dataframe that need to be printed
    """
    wb = xw.Book(path)  # Open the workbook
    sheet = wb.sheets['bq']  # select sheet
    sheet['G1'].options(index=False).value = df[status]  # write status in the file


def mail(sender, password, mail_receiver, mail_subject, invoice_list):
    """
    send a reminder to client that didn't pay their invoice
    :param sender: sender email
    :param password: application password
    :param mail_receiver: receiver email
    :param mail_subject: subject of the email
    :param invoice_list: list of tuple with invoice and amount of each client
    :return: sent an email
    """
    for invoice, amount in invoice_list:
        # Corps de l'e-mail
        body = f"""\
        Bonjour,
        Je vous informe que votre facture {invoice}, pour un montant de {amount} euros reste à ce jour impayée.
        Je vous pris de bien vouloir vous acquitter du montant de la facture dans un délai de 8 jours, dans le cas \
        contraire
        votre dossier sera transmis à notre service de recouvrement.
        Merci pour votre compréhension.
        """
        # Créez un objet MIMEMultipart
        message = MIMEMultipart()
        message["From"] = sender
        message["To"] = mail_receiver
        message["Subject"] = mail_subject
        # Attachez le corps au message
        message.attach(MIMEText(body, "plain"))
        # Initialisez la connexion SMTP
        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(sender, password)
            print("successful connection")
        except Exception as e:
            print("Erreur lors de la connexion au serveur SMTP : ", e)
            exit()
        # Envoyer l'e-mail
        try:
            server.sendmail(sender, mail_receiver, message.as_string())
            print("E-mail envoyé")
        except Exception as e:
            print("Erreur lors de l'envoi de l'e-mail : ", e)
        # Fermez la connexion SMTP
        server.quit()



def graph(table, tab):
    """
    create plots, stack them and save into a pdf file
    :param table: first dataframe to appear in the pdf
    :param tab:  second dataframe to appear in the pdf
    :return: a pdf with the two dataframe and their plots
    """
    print("Début de la fonction graph()")  # Ajoutez cette ligne pour vérifier si la fonction est appelée

    with PdfPages('Report.pdf') as pdf:
        print("Ouverture du fichier PDF")  # Ajoutez cette ligne pour vérifier si le fichier PDF est ouvert correctement

        plt.figure(figsize=(8, 13))
        plt.tight_layout(pad=7.0)  # Adjust layout to prevent clipping of titles
        plt.subplot(4, 1, 1)
        plt.axis('tight')
        plt.axis('off')
        plt.title('Tableau récapitulatif des données')
        plt.table(cellText=table.values, colLabels=table.columns, loc='center')
        plt.subplot(4, 1, 2)
        plt.bar(table['Status'], table['count'], color=['tab:blue', 'tab:orange', 'tab:green'])
        plt.xlabel('Status')
        plt.ylabel('Nombre de facture')
        plt.title('Bar chart')
        plt.subplot(4, 1, 3)
        plt.axis('tight')
        plt.axis('off')
        plt.title('Tableau récapitulatif des données par mois')
        plt.table(cellText=tab.values, colLabels=tab.columns, loc='center')
        plt.subplot(4, 1, 4)
        plt.bar(tab['Date'], tab['Total'])
        plt.title("Représentation du CA")
        # Save the current figure to the PDF
        pdf.savefig()
        print("Fichier PDF sauvegardé")  # Ajoutez cette ligne pour vérifier si le fichier PDF est correctement sauvegardé
        plt.close()

    print("Fin de la fonction graph()")  # Ajoutez cette ligne pour vérifier si la fonction est terminée
