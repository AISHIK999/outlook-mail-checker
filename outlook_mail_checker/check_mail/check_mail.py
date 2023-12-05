import win32com.client
import pandas as pd
import os
import re


# Constants
REP = "re: "
URGENT = "urgent"
HTML_TABLE_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<style>
 table {
     width: 100%;
     border-collapse: collapse;
 }
 th, td {
     border: 1px solid black;
     padding: 8px;
     text-align: left;
 }
 th {
     background-color: #f2f2f2;
 }
 .unread {
     color: red;
 }
</style>
</head>
<body>
<p>You have <span class="unread">
"""


# Functions
def count_mails(inbox):
    """
    count_mails counts the number of mails in the folder specified that contains the word "urgent" in the subject or body unless it is a trailing message.

    :param inbox: Folder in Outlook where we are going to search for mails.

    :return: Number of mails in the folder specified that contains the word "urgent" in the subject or body.
    """
    urgent_count = 0
    for message in inbox.Items:
        if not message.Subject.lower().startswith(REP):
            if re.findall(r'\b' + URGENT + r'\b', message.Subject.lower()) or re.findall(r'\b' + URGENT + r'\b', message.Body.lower()):
                urgent_count += 1
    return urgent_count


def get_sender_count(inbox):
    """
    get_sender_count gets the number of mails from each sender in the folder specified that contains the word "urgent" in the subject or body unless it is a trailing message.

    :param inbox: Folder in Outlook where we are going to search for mails.
    """
    sender_count = {}
    for message in inbox.Items:
        if not message.Subject.lower().startswith(REP):
            if re.findall(r'\b' + URGENT + r'\b', message.Subject.lower()) or re.findall(r'\b' + URGENT + r'\b', message.Body.lower()):
                sender = message.Sender.Name
                if sender in sender_count:
                    sender_count[sender] += 1
                else:
                    sender_count[sender] = 1
    return sender_count


def get_last_email_date(inbox, sender_count):
    """
    get_last_email_date gets the last email date from each sender in the folder specified that contains the word "urgent" in the subject or body unless it is a trailing message.

    :param inbox: Folder in Outlook where we are going to search for mails.
    """
    last_email_date = {}
    for message in inbox.Items:
        if not message.Subject.lower().startswith(REP):
            if re.findall(r'\b' + URGENT + r'\b', message.Subject.lower()) or re.findall(r'\b' + URGENT + r'\b', message.Body.lower()):
                sender = message.Sender.Name
                email_date = message.ReceivedTime.date()
                if sender in last_email_date:
                    if email_date > last_email_date[sender]:
                        last_email_date[sender] = email_date
                else:
                    last_email_date[sender] = email_date
    return last_email_date


def create_html_table(urgent_count, sender_count, last_email_date):
    """
    create_html_table creates the HTML table that will be sent in the email.
    It contains the summary of the mails we analyzed.

    :param urgent_count: Number of mails in the folder specified that contains the word "urgent" in the subject or body.
    :param sender_count: Number of mails from each sender in the folder specified that contains the word "urgent" in the subject or body.
    :param last_email_date: Last email date from each sender in the folder specified that contains the word "urgent" in the subject or body.

    :return: HTML table that will be sent in the email, containing urgent mail information.
    """
    html_table = (
        HTML_TABLE_TEMPLATE
        + str(urgent_count)
        + """
</span> important emails in your inbox.</p>
<table>
<tr>
<th>Sender</th>
<th>Count</th>
<th>Last Email Date</th>
</tr>
"""
    )
    for sender, count in sender_count.items():
        html_table += f"""
<tr>
<td>{sender}</td>
<td>{count}</td>
<td>{last_email_date[sender].strftime('%d-%m-%Y')}</td>
</tr>
"""
    html_table += """
</table>
</body>
</html>
"""
    return html_table


def checker():
    """
    checker is the main function of the module.
    It calls all the other functions and returns the HTML table that will be sent in the email alongwith storing the data in a csv file.

    :return: HTML table that will be sent in the email, containing urgent mail information.
    """
    ol = win32com.client.Dispatch("Outlook.Application")
    inbox = ol.GetNamespace("MAPI").GetDefaultFolder(6)
    # sub_folder = inbox.Folders.Item("sub_folder")
    # sub_sub_folder = sub_folder.Folders.Item("sub_sub_folder")

    # Replace below inbox with subfolders if necessary
    urgent_count = count_mails(inbox)
    sender_count = get_sender_count(inbox)
    last_email_date = get_last_email_date(inbox, sender_count)

    df = pd.DataFrame(list(sender_count.items()), columns=["Sender", "Count"])
    df["Last Email Date"] = df["Sender"].map(last_email_date)
    if not os.path.exists(".\\outlook_mail_checker\\temp"):
        os.makedirs(".\\outlook_mail_checker\\temp")
    df.to_csv(".\\outlook_mail_checker\\temp\\email_data.csv", index=False)

    html_table = create_html_table(urgent_count, sender_count, last_email_date)

    return html_table
