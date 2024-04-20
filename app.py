import streamlit as st
import pandas as pd
from datetime import datetime
import Repository
from View import print_data, mail, graph
import toml
import logging


logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')


# Charger la configuration depuis le fichier TOML
config = toml.load("config.toml")

# Accéder aux valeurs de la configuration
path = config["paths"]["excel_file"]
sender = config["email"]["sender"]
password = config["email"]["password"]
receiver = config["email"]["receiver"]
subject = config["email"]["subject"]
status_column = config["columns"]["status"]
amount_column = config["columns"]["amount"]

# Chargement des données depuis le fichier Excel
df_bq = pd.read_excel(path, sheet_name="bq")
df_erp = pd.read_excel(path, sheet_name="ERP")
excel = pd.read_excel(path, sheet_name="tableau")
excel['Date'] = excel['Date'].apply(lambda x: datetime.strftime(x, "%B"))

# Formatage et fusion des dataframes
df_bq = Repository.reformat_file(df_bq)
df_erp = Repository.reformat_file(df_erp)
inner_join = pd.merge(df_erp, df_bq, on='Numéro', how='inner')
Repository.get_status(df=inner_join, column_status=status_column)

# Ouverture du classeur Excel et impression du statut dans la feuille BQ
print_data(path=path, df=inner_join, status=status_column)



# Enregistrer des messages de journalisation à différents niveaux de gravité
logging.debug('Les données ont été chargées avec succès depuis le fichier Excel.')
logging.info('Les données ont été formatées et fusionnées avec succès.')
logging.warning('Attention : certaines données peuvent être manquantes ou inattendues.')
logging.error('Erreur lors du chargement ou du formatage des données.')

# Obtenir la liste des factures et des montants pour lesquels un e-mail de rappel doit être envoyé
final_list = Repository.get_list_reminder(df_erp)

# Envoi d'un e-mail de rappel
#mail(sender, password, receiver, subject, final_list)

# Création du dataframe nécessaire pour le graphique et la table
table1, table2 = Repository.create_tables(excel, amount_column, status_column)

# Génération du graphique et du PDF
graph(table1, table2)

# Interface utilisateur Streamlit
st.title("Application de gestion des factures")

# Affichage des données
st.write("## Données des factures")
st.write(inner_join)

# Affichage du graphique
st.write("## Graphique")
st.pyplot(graph(table1, table2))

# Affichage des tables
st.write("## Tableau récapitulatif des données")
st.write(table1)
st.write("## Tableau récapitulatif des données par mois")
st.write(table2)


