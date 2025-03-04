import pandas as pd

def clean_compare_table(df):
    df = df.rename(columns= {"Daten ES/PS/PSO": "Kontakt", "Microchip": "Chip"}, errors= "ignore")
    wanted_cols = ['Name', 'Ort', 'Chip', 'Kontakt', 'DOB']
    tab_cols = list(df.columns)
    got_cols = []
    for col in tab_cols:
        if col in wanted_cols:
            got_cols.append(col)

    new_df = df[got_cols]
    new_df = new_df.sort_values('Name')
    return new_df

def compare(df1, df2):
    # TODO
    df1 = clean_compare_table(df1)
    print(df1)
    df2 = clean_compare_table(df2)
    print(df2)
    return False