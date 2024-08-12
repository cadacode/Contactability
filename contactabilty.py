import mail
import pandas as pd
import io
import utils
import modules.database as database
from pathlib import Path


def escape_html(text):
    return text.replace("\n", "<br>")


def email_message(store):
    return escape_html(
        f"""Ciao team {store},

Please see attached a recap of the Contactability report of last week for your store.
Please be careful to collect both the Marketing and Profiling consent and at least one Contact channel (E-mail or Mobile)

Having a high Contactability rate is crucial for Valentino and the basis to start the relationship with our clients.


"""
    )


def email_signature():
    return escape_html(
        
        """Thank you,
Best,
_____
Elisa Cadamuro
CRM Trainee

VALENTINO USA
11 West 42nd Street, 26th Floor
New York, NY 10036
Cell 646-300-1014
Elisa.Cadamuro@valentino.com 
Team Email: usa-it@valentino.com
www.valentino.com
"""
    )


def email_content(store, store_db):
    return email_message(store) + store_db.to_html(index=False) + email_signature()


def Store_conditions(db):
    contactable = (db["Marketing Permission"] == "Allowed") and (
        db["Profiling Permission"] == "Allowed"
    )
    return 1 if contactable else 0

def contactability_marketing(db):
    contactable = (db["Marketing Permission"] == "Allowed")
    return 1 if contactable else 0

def contactability_profiling(db):
    contactable = (db["Profiling Permission"] == "Allowed")
    return 1 if contactable else 0

def email_conditions(db):
    email = (db["Goodness: E-mail"] == "X" and db["Contactable"] == 1)
    return 1 if email else 0

def mobile_conditions(db):
    mobile = (db["Goodness: Mobile"] == "X" and db["Contactable"] == 1)
    return 1 if mobile else 0

def channel_conditions(db):
    channel = (db["Contactable"] == 1 and (db["Goodness: Mobile"] == "X" or db["Goodness: E-mail"] == "X"))
    return 1 if channel else 0

def contactability_database():
    base = Path(
        "V:\\Retail\\CRM\\23. BI\\2022\\New\\Extract\\sales extract.csv"
    )
    return database.create_client_database(
        base,
        6,
        [
            "Calendar year / week",
            "Client ID",
            "Signage",
            "Associate",
            "Marketing Permission",
            "Profiling Permission",
            "Goodness: E-mail",
            "Goodness: Mobile"
        ],
    )


#def email_total(password, recipients, store, store_db):
    subject = f"Contactability | {store}"
    body = email_content(store, store_db)
    message = f"<html><body><h4>{subject}</h4>{body}</body></html>"
    cc = []
    bcc = []
    attachments = []

    mail.send_valentino_email(
        recipients, subject, message, password, cc, bcc, attachments
    )
    print("Sent email to store:", store)


def test_email_total(password, recipients, store, store_db):
    recipients = ["eli.cadamuro@gmail.com"]
    subject = f"Contactability | {store}"
    body = email_content(store, store_db)
    message = f"<html><body><h4>{subject}</h4>{body}</body></html>"
    attachments = []
    cc = []
    bcc = []

    mail.send_valentino_email(
        recipients, subject, message, password, cc, bcc, attachments
    )
    print("Sent test email to store:", store)


if __name__ == "__main__":

    password = input("Type your email password: ")

    # Generate database
    db = contactability_database()


    # Remove duplicated clients for each associate
    #db = db.drop_duplicates(subset=["Client ID", "Associate"])


    for store in set(db["Signage"]):
    #for store in ["ASPEN"]:
        recipients = utils.get_store_emails("data/mailing_list.xlsx", store)
        if recipients:
    
            store_db = db[db["Signage"] == store].copy()

             # Create marketing and email permissions columns
            store_db["Date"] = store_db.apply(store_db["Date"], axis=1)
            store_db["Marketing Contactable"] = store_db.apply(contactability_marketing, axis=1)
            store_db["Profiling Contactable"] = store_db.apply(contactability_profiling, axis=1)
            store_db["Email Contactable"] = store_db.apply(email_conditions, axis=1)
            store_db["Mobile Contactable"] = store_db.apply(mobile_conditions, axis=1)
            #store_db["Email/Mobile Contactable"] = store_db.apply(mobile_conditions, axis=1)

            # Count clients using their contactability status
            store_db["Total Clients"] = store_db.groupby(["Associate"])["Client ID"].transform("count")
            store_db["Contactable Clients"] = store_db.groupby(["Associate"])["Contactable"].transform("sum")
            store_db["Marketing Permission"] = store_db.groupby(["Associate"])["Marketing Contactable"].transform("sum")
            store_db["Profiling Permission"] = store_db.groupby(["Associate"])["Profiling Contactable"].transform("sum")
            store_db["Contactable by Email"] = store_db.groupby(["Associate"])["Email Contactable"].transform("sum")
            store_db["Contactable by Mobile"] = store_db.groupby(["Associate"])["Mobile Contactable"].transform("sum")
            #store_db["Email/Mobile Clients"] = store_db.groupby(["Associate"])["Email/Mobile Contactable"].transform("sum")

            # Calculate contactability perrcentages
            store_db["Contactability %"] = ((store_db["Contactable Clients"] / store_db["Total Clients"]) * 100)
            #store_db["Email/Mobile %"] = ((store_db["Email/Mobile Clients"] / store_db["Total Clients"]) * 100)

            # Select only relevant columns and drop duplicates 
            store_db = store_db[
                [
                    "Associate",
                    "Total Clients",
                    "Contactable Clients",
                    "Marketing Permission",
                    "Profiling Permission",
                    "Contactability %",
                    "Contactable by Email",
                    "Contactable by Mobile",
                ]
            ].drop_duplicates(subset=["Associate"])
            
            store_db = store_db.sort_values("Associate")
            #store_db = store_db.reset_index(drop=True)

            # Calculate total number of contactable clients for the store
            store_db.loc["Total"] = store_db.sum(numeric_only=True)

            total_clients = store_db.at["Total", "Total Clients"]
            contact_ratio = (store_db.at["Total", "Contactable Clients"] / total_clients)
            #channel_ratio = (store_db.at["Total", "Email/Mobile Clients"] / total_clients)
            store_db.at["Total", "Contactability %"] = round(contact_ratio * 100)
            #store_db.at["Total", "Email/Mobile %"] = round(channel_ratio * 100)
            
            # Format the columns appropriately
            for col in store_db.columns:
                if col == "Associate": 
                    store_db.at["Total", col] = "TOTAL"
                elif '%' in col:
                    store_db[col] = store_db[col].apply(lambda x: str(int(x)) + "%")
                else:
                    store_db[col] = store_db[col].astype(int)
            
    
            print(store_db)

            test_email_total(password, recipients, store, store_db)
        else:
            print("Unknow email for store", store)
