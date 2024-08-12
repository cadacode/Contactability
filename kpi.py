import pandas as pd
import utils
from pathlib import Path
import os

def clean_client_id(client_id):
    '''
    Objective: 
        "VALENTINO/1976137" -> "1976137"
    Implementation:
        list = "VALENTINO/1976137".split("/")
        list == ["VALENTINO", "1976137"]
        list[-1] == "1976137"
    '''
    return int(client_id.split('/')[-1])

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
            df = pd.read_csv(path, header=0)
            # Remove unnamed columns
            df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
            databases[path.name.removesuffix(".csv")] = df
            if 'Client ID ( vista Brand )' in df.columns:
                df["Client ID ( vista Brand )"] = df["Client ID ( vista Brand )"].apply(clean_client_id)
    return databases

def store_folder_name(databse_name, store):
    if (databse_name == "sales" or databse_name == "MDSales"):
        if store == "NEW YORK SOHO":
            return "SPRING STREET POPUP (SSV)"
    return store


if __name__ == "__main__":

    database_folder = Path("C:\\Users\\cadae\\OneDrive - Valentino S.p.A\\Desktop\\python_data\\kpi\\databases\\december\\todo")
    shared_folder = Path("C:\\Users\cadae\OneDrive - Valentino S.p.A\CRM Valentino NA")
    
    for name, db in kpi_databases(database_folder).items():
        # Select the appropriate column based on file
        column = "Signage" if name == "sales" else "Dedicated Store"
        
        # Iterate over the stores in that colum
        for store in set(db[column]):
            # Create a database for the store, sorted by dedicated associate
            store_db = db[db[column] == store]
            #store_db = store_db[store_db["Loyalty CY"].str.contains('2. Loyal|1. Super Loyal')]
            store_db = store_db.sort_values("Dedicated Associate")

            # Save the database in its store's folder with the original name
            store_folder = store_folder_name(name, store)
            report_folder = shared_folder / store_folder / "Monthly_KPIs" / f"December_KPIs_{store_folder}"
            save_xlsx(report_folder, name, store_db)
        #print(store_db)
