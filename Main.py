from datetime import datetime
import pandas as pd
import Repository
from View import print_data, mail, graph
import toml

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

# get data from excel_file
df_bq = pd.read_excel(path, sheet_name="bq")
df_erp = pd.read_excel(path, sheet_name="ERP")
excel = pd.read_excel(path, sheet_name="tableau")
excel['Date'] = excel['Date'].apply(lambda x: datetime.strftime(x, "%B"))

# format and merge dataframes
df_bq = Repository.reformat_file(df_bq)
df_erp = Repository.reformat_file(df_erp)
inner_join = pd.merge(df_erp, df_bq, on='Numéro', how='inner')
Repository.get_status(df=inner_join, column_status=status_column)

# Open excel workbook and print the status in sheet BQ
print_data(path=path, df=inner_join, status=status_column)

# get invoices and amounts for which a reminder e-mail must be sent
final_list = Repository.get_list_reminder(df_erp)

# send reminder email
#mail(sender, password, receiver, subject, final_list)

# creating the necessary dataframe for graphic and table
table1, table2 = Repository.create_tables(excel, amount_column, status_column)

# generate chart and pdf
graph(table1, table2)
