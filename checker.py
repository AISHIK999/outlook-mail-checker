import win32com.client
import pandas as pd


def checker():
    ol = win32com.client.Dispatch("Outlook.Application")

    inbox = ol.GetNamespace("MAPI").GetDefaultFolder(6)
    mdm_folder = inbox.Folders.Item("MDM")

    unread_count = 0
    sender_count = {}
    last_email_date = {}
    for message in mdm_folder.Items:
        if ("urgent" in message.Subject.lower()) or ("urgent" in message.Body.lower()):
            unread_count += 1
            sender = message.Sender.Name
            if sender in sender_count:
                sender_count[sender] += 1
            else:
                sender_count[sender] = 1

            email_date = message.ReceivedTime.date()  # get the date of the email
            if sender in last_email_date:
                if email_date > last_email_date[sender]:
                    last_email_date[sender] = email_date
            else:
                last_email_date[sender] = email_date

    df = pd.DataFrame(list(sender_count.items()), columns=["Sender", "Count"])
    df["Last Email Date"] = df["Sender"].map(last_email_date)
    df.to_csv("email_data.csv", index=False)

    html_table = (
        """
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
        + str(unread_count)
        + """
</span> urgent emails in MDM folder.</p>
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
<td>{last_email_date[sender].strftime('%m-%d-%Y')}</td>
</tr>
"""

    html_table += """
</table>
</body>
</html>
"""

    return html_table
