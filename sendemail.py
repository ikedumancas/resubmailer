import csv
import os
from smtplib import SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def get_env_or_die(name):
    try:
        return os.environ[name]
    except KeyError:
        raise KeyError('Set the environment variable %s' % name)


def generate_email_content(fname):
    text_email = """\
    Hi %s!
    This is a test email text version.
    """ % (fname)
    html_email = """\
    <html>
      <head></head>
      <body>
        <p>Hi %s!</p>
        <p>This is a test emai HTML version.</p>
      </body>
    </html>
    """ % (fname)
    return html_email, text_email


# SETTINGS
SERVER_EMAIL = "BlindsOnSale.com <design@blindsonsale.com>"
CSV_FILE = 'BOS-1999-2013-Email-List-test.csv'
EMAIL_HOST = 'smtp.webfaction.com'
EMAIL_PORT = 587
EMAIL_USER = 'bos_sendmail'
EMAIL_PASSWORD = get_env_or_die('EMAIL_HOST_PASSWORD')

with open(CSV_FILE) as csvfile:
    emaillistreader = csv.reader(csvfile)
    x = 0
    for row in emaillistreader:
        x += 1
        if x > 1:
            to_email = row[0]
            fname = row[1]
            html_email, text_email = generate_email_content(fname)
            try:
                # Create message container
                msg = MIMEMultipart('alternative')
                msg['Subject'] = "Resub"
                msg['From'] = SERVER_EMAIL

                # Record the MIME types of both parts - text/plain and text/html.
                part1 = MIMEText(text_email, 'plain')
                part2 = MIMEText(html_email, 'html')

                # Attach parts into message container.
                # The last part of a multipart message, is best and preferred.
                msg.attach(part1)
                msg.attach(part2)

                # Connect to SMTP server
                smtp = SMTP()
                smtp.connect(EMAIL_HOST, EMAIL_PORT)
                smtp.set_debuglevel(True)
                smtp.login(EMAIL_USER, EMAIL_PASSWORD)

                # Send Email
                try:
                    smtp.sendmail(SERVER_EMAIL, to_email, msg.as_string())
                except Exception as exc:
                    print('------')
                    print('Error Sending', exc)
                finally:
                    smtp.quit()
            except Exception as exc:
                print('------')
                print('Error Connecting', exc)
