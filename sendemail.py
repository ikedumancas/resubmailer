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

BODY_TEXT = """\
I know it's been a while since we've connected.

You asked for our newsletter a while back. We put it on hold while we were making some changes.

We're going to send out 1 free tip a week on interior design. We'll also have giveaways and exclusive discounts for you.

Are you still interested? Let us know by resubscribing. http://bit.ly/2dHWYwz

If not, don't do anything. You'll be unsubscribed automatically.

This week's tip is about How to Make a Room Look Taller. A little extra height in a room is always a good thing. Here are 6 quick ways to make your room look taller. http://bit.ly/2dmdTDG

Don't forget to resubscribe if you want to keep getting these. http://bit.ly/2dHWYwz

Thanks!
BlindsOnSale.com
"""

BODY_HTML = """
<p>I know it&#8217;s been a while since we&#8217;ve connected.</p>

<p>You asked for our newsletter a while back. We put it on hold while we were making some changes.</p>

<p>We&#8217;re going to send out 1 free tip a week on interior design. We&#8217;ll also have giveaways and exclusive discounts for you.</p>

<p>Are you still interested? Let us know by resubscribing. <a href="%s">http://bit.ly/2dHWYwz</a></p>

<p>If not, don&#8217;t do anything. You&#8217;ll be unsubscribed automatically.</p>

<p>This week&#8217;s tip is about <a href="http://bit.ly/2dmdTDG">How to Make a Room Look Taller</a>. A little extra height in a room is always a good thing. Here are 6 quick ways to make your room look taller.</p>

<p>Don&#8217;t forget to resubscribe if you want to keep getting these. <a href="%s">http://bit.ly/2dHWYwz</a></p></p>

<p>Thanks!<br>
BlindsOnSale.com</p>"""

def generate_email_content(to_email, fname):
    text_email = BODY_TEXT
    html_email = BODY_HTML

    # Append a greeting if we have the person's name
    if fname:
        text_email = "Hey " + fname + "!\n\n" + text_email
        html_email = "<p>Hey " + fname + "!</p>" + html_email

    # Wrap in HTML boilerplate
    html_email = "<html><head></head><body>" + html_email + "<body>"

    # Populate email and name in the URL
    url = 'https://www.blindsonsale.com/subscribe?email=%s&name=%s' % (to_email, fname)
    html_email = html_email % (url, url)

    return html_email, text_email


# SETTINGS
SERVER_EMAIL = "Design at BOS <design@blindsonsale.com>"
CSV_FILE = 'mike.csv'
EMAIL_HOST = 'smtp.webfaction.com'
EMAIL_PORT = 587
EMAIL_USER = 'bos_sendmail'
EMAIL_PASSWORD = get_env_or_die('EMAIL_HOST_PASSWORD')

# Connect to SMTP server
try:
    smtp = SMTP()
    smtp.connect(EMAIL_HOST, EMAIL_PORT)
    smtp.login(EMAIL_USER, EMAIL_PASSWORD)
except Exception as exc:
    print('Error Connecting', exc)
    exit()

# Process CSV file
with open(CSV_FILE) as csvfile:
    emaillistreader = csv.reader(csvfile)

    # Skip the header
    next(emaillistreader, None)

    for row in emaillistreader:
        to_email = row[0]
        fname = row[1]
        html_email, text_email = generate_email_content(to_email, fname)

        # Create message container
        msg = MIMEMultipart('alternative')
        msg['Subject'] = "Our Last Email :( Unless..."
        msg['From'] = SERVER_EMAIL

        # Record the MIME types of both parts - text/plain and text/html.
        part1 = MIMEText(text_email, 'plain')
        part2 = MIMEText(html_email, 'html')

        # Attach parts into message container.
        # The last part of a multipart message, is best and preferred.
        msg.attach(part1)
        msg.attach(part2)

        # Send Email
        print("Sending email to %s <%s>... " % (fname, to_email), end='')
        try:
            smtp.sendmail(SERVER_EMAIL, to_email, msg.as_string())
            print("OK")
        except Exception as exc:
            print('ERR', exc)

# Close connection
smtp.quit()
