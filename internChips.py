# import required libraries
import requests
from bs4 import BeautifulSoup
import pandas as pd
import datetime
import smtplib
from email.message import EmailMessage

# fetch mailing list
mailing_list_sheet = 'link to your google sheet containing emails collected from google form responses'
response = requests.get(mailing_list_sheet)
soup = BeautifulSoup(response.content, "html.parser")
table = soup.find("table")

receivers = []
for row in table.find_all("tr")[1:]:
    cols = row.find_all("td")[1].text.strip()
    if cols:
        receivers.append(cols)

receivers.pop(0)
mailing_list = list(set(receivers))
print(f'{len(mailing_list)} emails fetched...')
print('')


# check if any chips due today
internChips_sheet = "link to your google sheet containing internship details"
response = requests.get(internChips_sheet)
soup = BeautifulSoup(response.content, "html.parser")
table = soup.find("table")

data = []
for row in table.find_all("tr"):
    cols = [col.text.strip() for col in row.find_all("td")]
    data.append(cols)
    
df = pd.DataFrame(data)
mask = df[0] == str(datetime.date.today())
expiring = df[mask].iloc[:, 1:3]
joined = expiring[1] + ': ' + expiring[2]
listed = '\n'.join([" - " + x for x in joined])
print(f'{len(joined)} chips expiring today...')
print('')


# if so, mail details
due_message_1 = f"""Hello,

Today is the last day the following chip will be valid, so make sure to apply before midnight:
{listed}

Apart from the one listed above, if you know of any that expire today, please reach out to any one of the Chips' Admins on WhatsApp with the details at the earliest, so that we may update everyone.

That being said, good luck applying!"""

due_message_2 = f"""Hello,

Today is the last day the following chips will be valid, so make sure to apply before midnight:

{listed}

Apart from the ones listed above, if you know of any that expire today, please reach out to any one of the Chips' Admins on WhatsApp with the details at the earliest, so that we may update everyone.

That being said, good luck applying!"""

if len(listed) > 0:
    print('sending reminder...')
    
    sender = 'insert bot account email here' 
    password = 'insert bot account password here'
    reminder = EmailMessage()
    if len(joined) == 1:
        reminder['Subject'] = '1 chip expiring TODAY!🍟'
    else:
        reminder['Subject'] = f'{len(joined)} chips expiring TODAY!🍟'
    reminder['From'] = sender
    reminder['Bcc'] = mailing_list
    if len(joined) == 1:
        reminder.set_content(due_message_1)
    else:
        reminder.set_content(due_message_2)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(sender, password)
        server.send_message(reminder)
        server.quit()


# if not, mail accordingly
not_due_message = f"""Hello,
    
No chips are expiring today (to our knowledge).

If you think we've made a mistake, and know of any that DO expire today, please reach out to any one of the Chips' Admins on WhatsApp with the details at the earliest, so that we may inform everyone.
    
Have a good day!"""

if len(listed) == 0:
    print('mailing accordingly...')
    
    sender = 'insert bot account email here' 
    password = 'insert bot account password here'
    reminder = EmailMessage()
    reminder['Subject'] = 'No chips expiring today🍟'
    reminder['From'] = sender
    reminder['Bcc'] = mailing_list
    reminder.set_content(not_due_message)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(sender, password)
        server.send_message(reminder)
        server.quit()