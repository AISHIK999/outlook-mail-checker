import os
import asyncio
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import aiosmtplib

from checker import checker


async def main():
    load_dotenv()

    email_address = "your.email@gmail.com"
    email_password = os.getenv("OFF_PASS")
    email_server = "smtp.gmail.com"
    email_port = 587

    msg = MIMEMultipart()
    msg["From"] = email_address
    msg["To"] = "to.email@mail.com"
    msg["Subject"] = "Email status report"
    message = checker()
    msg.attach(MIMEText(message, "html"))

    with open("email_data.csv", "rb") as file:
        attachment = MIMEBase("application", "octet-stream")
        attachment.set_payload(file.read())
    encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", "attachment", filename="your_file.csv")
    msg.attach(attachment)

    smtp = aiosmtplib.SMTP(hostname=email_server, port=email_port, start_tls=True)
    await smtp.connect()
    await smtp.login(email_address, email_password)
    await smtp.send_message(msg)
    await smtp.quit()


if __name__ == "__main__":
    asyncio.run(main())
    os.remove("email_data.csv")
