import pandas as pd
import utils
from pathlib import Path
import os

def save_xlsx(folder, name, df):
    filename = f"{folder}/{name}.xlsx"
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    with pd.ExcelWriter(filename) as writer:
        utils.to_formatted_xlsx(df, writer, sheet_name=name, index=False)

def kpi_databases(folder):
    databases = {}
    for path in folder.iterdir():
        if path.suffix == ".csv":
            # Load database from csv file
            df = pd.read_csv(path, header=5, sep=";")
            # Remove unnamed columns
            df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
            databases[path.name.removesuffix(".csv")] = df
            df["Client ID ( vista Brand )"] = df["Client ID ( vista Brand )"].apply(clean_client_id)
    return databases

if __name__ == "__main__":

    database_folder = Path("c:\\Users\cadae\OneDrive - Valentino S.p.A\Desktop\python_data\kpi\databases")
    shared_folder = Path("C:\\Users\cadae\OneDrive - Valentino S.p.A\CRM Valentino NA")
