from datetime import datetime
import numpy as np
import pandas as pd
import toml

# Charger la configuration depuis le fichier TOML
config = toml.load("config.toml")

# Accéder aux valeurs de la configuration
status_column = config["columns"]["status"]
amount_column = config["columns"]["amount"]

def reformat_file(dataframe):
    """
    add currency column, Montant column and drop columns Montant USD and EUR
    return a reformated data frame
    :param dataframe: DataFrame
    :return: reformated dataFrame
    """
    dataframe['CCY'] = np.where(np.isnan(dataframe['Montant EUR']), 'USD', 'EUR')
    dataframe['Montant'] = np.where(dataframe['CCY'] == 'EUR', dataframe['Montant EUR'], dataframe['Montant USD'])
    dataframe = dataframe.drop(columns=['Montant EUR', 'Montant USD'])
    return dataframe

def get_status(df, column_status):
    """ give the status for each receipt
    :param df: DataFrame
    :return: fill column Status
    """
    df[column_status] = np.where(df['Montant_x'] == df['Montant_y'], "Payed",
                                  np.where(df['Montant_x'] <= df['Montant_y'], "Payed with change gain",
                                           "Payed with losses"))

def create_tables(df, column_amount, column_status):
    """
    Create the tables necessary for the pdf
    :param df: Dataframe necessary for the table
    :param column_amount: amount
    :param column_status: status
    :return: list of tables
    """
    table = pd.pivot_table(df, values=column_amount, index=column_status,
                           aggfunc={column_amount: ["count", "sum", "mean", "std", "median"]})
    table.loc['Total'] = [len(df[column_status]), df[column_amount].mean(), df[column_amount].median(),
                          df[column_amount].std(), df[column_amount].sum()]
    table = table.round(decimals=2)
    table = table.reset_index(drop=False)

    tab = pd.pivot_table(df, values=column_amount, index='Date', columns=column_status, aggfunc="sum", fill_value=0,
                         margins=True, margins_name="Total")
    tab = tab.drop('Total', axis=0).reset_index().round(2)

    return table, tab


def get_list_reminder(df):
    """
    Get the list of info for each client that need to be reminded of their none payment
    :param df: Dataframe
    :return: list of tuple with client info
    """
    today = datetime.today()
    df['reminder'] = (today - df["Date de facture"]).dt.days
    df = df[df['reminder'] >= 30]
    final_list = list(zip(df['Numéro'], df['Montant']))
    return final_list
