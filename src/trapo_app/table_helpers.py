import pandas as pd

def clean_compare_table_po(df):
    # TODO
    headers = df.iloc[0]
    new_df = pd.DataFrame(df.values[1:], columns=headers)
    new_df = new_df.iloc[:,[4,6,7,9]]
    new_df = new_df.rename(columns= {"Daten ES/PS/PSO": "Adoptant"}, errors= "ignore")
    new_df = new_df.sort_values('Name')
    return new_df

def clean_compare_table_chat(df):
    # TODO
    headers = df.iloc[0]
    new_df = pd.DataFrame(df.values[1:], columns=headers)
    new_df = new_df.iloc[:,[3,4,6,14,18]]
    new_df = new_df.rename(columns= {"": "Adoptant", "Microchip": "Chip"}, errors= "ignore")
    new_df = new_df.sort_values('Name')
    return new_df

def compare(df1, df2):
    # TODO
    df1 = clean_compare_table_chat(df1)
    print(df1)
    df2 = clean_compare_table_po(df2)
    print(df2)
    return False