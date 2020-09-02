import win32com.client
import openpyxl
import os
import pandas as pd
CapexUpload = ""
month="Jan20"

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
folder = inbox.Folders('Capex Test')
subfolder = folder.Folders(month)
subfoldermessages = subfolder.Items

email_list = []
for email in subfoldermessages:
    email_list.append(email)

attachments_list=[]
for email in email_list:
    for attachment in email.Attachments:
        if "Capex" in attachment.FileName:
            attachments_list.append(attachment)

for attachment in attachments_list:
    attachment.SaveAsFile(CapexUpload + "/"+ attachment.FileName)

uploadfile = "To be uploaded"
dataframe_all = pd.DataFrame()
for filename in os.listdir(CapexUpload):
    if uploadfile in filename:
        continue
    if filename.endswith(".xlsx"):
        dataframe = pd.read_excel(f"{CapexUpload}/{filename}")
    if filename.endswith(".csv"):
        dataframe = pd.read_csv(f"{CapexUpload}/{filename}")
    dataframe = dataframe.iloc[:, 0:7]
    dataframe_all = dataframe_all.append(dataframe,sort=False)
dataframe_all.to_csv(f"{CapexUpload}/{uploadfile}.csv", index= False)


