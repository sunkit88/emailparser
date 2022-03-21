import os, re, html
import time
import win32com.client
from bs4 import BeautifulSoup
import pandas as pd



pd.set_option("display.max_rows", None, "display.max_columns", None)


# Set up database
# db = setup()

# Create an folder input dialog with tkinter
# folder_path = os.path.normpath(askdirectory(title='Select Folder'))
folder_path = r'D:/TEMP/email'

# Initialise & populate list of emails
email_list = [file for file in os.listdir(folder_path) if file.endswith(".msg")]

# Connect to Outlook with MAPI
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

dvd_list = []
# text_list = []
date_list = []

noneed = ["vr", "dmlk", "jdxa", "glod", "_", "sweets"]

# Iterate through every email
for i, _ in enumerate(email_list):

    # Create variable storing info from current email being parsed
    msg = outlook.OpenSharedItem(os.path.join(folder_path, email_list[i]))

    print(i+1)
    print(str(msg.SentOn))

    soup = BeautifulSoup(msg.HTMLBody, 'lxml')

    imgs = soup.find_all('img', src=True)
    for img in imgs:
        if ('src' in img.attrs):
            if img['src'].endswith('.jpg'):
                print(img['src'])
                dvd_list.append(img['src'])
                # print(img['alt'])
                # text_list.append(img['alt'])
                date_list.append(str(msg.SentOn))

df_import = pd.read_csv(folder_path+"/result.csv")
df_raw = pd.DataFrame(list(zip(dvd_list,date_list)), columns=['dvd', 'date'])
df = df_raw.copy()



df['dvd'] = df['dvd'].astype(str).str.split("/").str.get(-1)
print(df)
df1 =pd.DataFrame()
df1['name'] = df['dvd'].replace(r'\D+$', r'', regex=True)
df1['name'] = df1['name'].replace(r'^\d+', r'', regex=True)
df1 = df1[~df1["name"].str.contains('|'.join(noneed))]
df1 = df1[df1['name'].str.strip().astype(bool)]
# df1[0] = df1[0].apply(lambda x: str(x)[:-1] if str.endswith(r'[A-Za-z]?'))
df1 = df1['name'].str.split(r'(\d+$)', expand=True)
print(df1)



df1['num']=df1[1].apply(lambda x: x[-3:] if len(x)>3 else x)
df['dvd'] = df1[0].astype(str) + " " + df1['num'].astype(str)
df = df.dropna()
# df = df[~df["dvd"].str.contains("vr")]
df = df.sort_values('date', ascending=True)
df_new = df.drop_duplicates(subset = "dvd")
df_new = pd.concat([df_import,df_new],ignore_index=True)
df_new = df_new.drop_duplicates(subset = "dvd")


outpath = r'D:/TEMP/email'
filename1 = "result_"+(time.strftime("%Y%m%d")+".csv")
filename2 = "result.csv"
df_new.to_csv(outpath + "\\" + filename1, index = False)
df_new.to_csv(outpath + "\\" + filename2, index = False)
# df.to_csv(r'D:/TEMP/email/result.csv', index = False)

