# DATA

This repository contains a simple Python script for sending email via SMTP.

## mail_sender.py

`mail_sender.py` provides a small function and an example usage for sending an
email through an SMTP server.

### Usage

Edit the configuration variables at the bottom of `mail_sender.py` to match
your SMTP server settings and run:

```bash
python mail_sender.py
```

The script will connect to the server and send a basic test email.

## mail_scheduler.py

`mail_scheduler.py` provides a graphical interface to prepare and schedule
bulk emails. The tool lets you select an Excel file with recipient data,
choose an HTML template, optionally add a signature and schedule multiple
send times using an integrated calendar widget.

Run the script with Python 3:

```bash
python mail_scheduler.py
```

The application will open a window where you can configure your message and
export the settings to `config.json`.
