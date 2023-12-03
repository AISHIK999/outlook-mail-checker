import win32com.client
from check_mail.check_mail import checker
import os

# Constants
# Change the below constants as per your requirement
# Separate multiple email addresses with a semicolon (;)
RECIPIENT_EMAIL = "username@domain.com"
SUBJECT = "Important mail report"


# Functions
def send_email(
    outlook,
    recipient_email,
    subject,
    body,
    attachment_path,
):
    mail = outlook.CreateItem(0)
    mail.To = recipient_email
    mail.Subject = subject
    mail.HTMLBody = body
    mail.Attachments.Add(attachment_path)
    mail.Send()
    os.remove(attachment_path)


def main():
    outlook = win32com.client.Dispatch("Outlook.Application")
    body = checker()
    current_dir = os.path.dirname(os.path.realpath(__file__))
    attachment_path = os.path.join(current_dir, "temp\\email_data.csv")
    send_email(
        outlook,
        RECIPIENT_EMAIL,
        SUBJECT,
        body,
        attachment_path,
    )


# Driver code
if __name__ == "__main__":
    main()
