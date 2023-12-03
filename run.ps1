if (Test-Path venv) {
    Write-Output "Running..."
    .\venv\Scripts\activate
    # Change the path below to the path of the script if needed
    python .\outlook_mail_checker\send_mail.py
    Write-Output "Execution successuful!"
} else {
    Write-Output "Initial setup..."
    python -m venv venv
    .\venv\Scripts\activate
    pip install -r requirements.txt
    Write-Output "Setup successuful!"
}