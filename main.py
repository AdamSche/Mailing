import win32com.client as win32
import pandas as pd

# Read the Excel file containing the email addresses
df = pd.read_excel(r'Base.XLSX')

# Connect to Outlook
outlook = win32.Dispatch('outlook.application')

# Get the email template
template = outlook.CreateItemFromTemplate(r'Template.oft')
template.HTMLBody += "<Logo>"

Str_mail = ""
with open(r'Base.Send.TXT', mode="r") as file:
    for line in file:
        Str_mail += line

    list_mail = Str_mail.split("; ")

new_mail = ""
lista_wew = []
# Loop through the email addresses
for index, row in df.iterrows():
    # Create a new email

    if row['email_address'] in list_mail:
        continue

    if row['email_address'] in lista_wew:
        continue

    lista_wew.append(row['email_address'])
    mail = outlook.CreateItem(0)
    mail.To = row['email_address']
    mail.Subject = f'Subject dla {row["company_name"]}'
    mail.htmlBody = template.HTMLBody
    mail.Attachments.Add(r'Attachment')
    mail.Display()

    mail.send
    new_mail = row['email_address'] + "; "
    with open(r'Base.Send.TXT', mode="a") as file2:
        file2.write(new_mail)

