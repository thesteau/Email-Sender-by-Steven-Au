# Writes and send the emails
import smtplib
import time
import mimetypes
from email.message import Message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.audio import MIMEAudio
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders

import pandas as pd
from program import parameters as params


class EmailWriter:
    """ Writes and sends the emails based on data values."""

    def __init__(self, sender, recipients):
        self._the_word = params.Parameters().get_word()
        self._sender = sender
        self._recipients = recipients




if __name__ == "__main__":
    pass
