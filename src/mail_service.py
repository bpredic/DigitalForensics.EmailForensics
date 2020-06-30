import datetime
import imaplib
import email
import re
from email.header import decode_header
import config

class MailService:
    def __init__(self):
        self.__imap = None

    def connect(self):
        self.__imap = imaplib.IMAP4_SSL(config.mail_server, config.mail_port)

    def authenticate(self, username, password):
        self.__imap.login(username, password)

    def list_folders(self):
        folder_names = [name.decode('utf8') for name in self.__imap.list()[1]]
        return folder_names

    def select_mailbox(self, mailbox):
        a = self.__imap.select(mailbox, readonly=True)
        print(a)

    def close_mailbox(self):
        self.__imap.close()

    def logout(self):
        self.__imap.logout()

    def get_message_info_for_period(self, period_start, period_end):
        start = period_start.strftime('%d-%b-%Y')
        end = period_end.strftime('%d-%b-%Y')
        status, messages = self.__imap.search(None, f'(SINCE "{start}" BEFORE "{end}")')
        message_info = []
        i = 1
        num_messages = len(messages[0].split())
        for message_id in messages[0].split():
            print(f"Fetching {i}/{num_messages}")
            i += 1
            status, data = self.__imap.fetch(message_id, '(RFC822)')
            try:
                for response_part in data:
                    if isinstance(response_part, tuple):
                        message = email.message_from_bytes(response_part[1])
                        subject = self.__parse_subject(message['Subject'])
                        sender = self.__parse_email(message['From'])[0]
                        recievers = self.__parse_email(message['To'])
                        cc = None
                        if 'CC' in message:
                            cc = self.__parse_email(message['CC'])
                        bcc = None
                        if 'BCC' in message:
                            bcc = self.__parse_email(message['BCC'])
                        date = self.__parse_email_datetime(message['Date'])
                        body = self._parse_message_body(message)
                        dict = {'Sender': sender, 'Recievers': recievers, 'CC': cc, 'BCC': bcc, 'Date': date,
                                'Subject': subject, 'Text-Body': body}
                        message_info.append(dict)
            except:
                print("Error")
        return message_info

    def __parse_subject(self, header):
        subject = self.__header_to_decoded_string(header)
        return subject

    def __parse_email(self, header):
        string_header = self.__header_to_decoded_string(header)
        email = re.findall('[\w\-\.]+@[\w\-\.]+\.[\w\-\.]+', string_header)
        return email

    def __parse_email_datetime(self, header):
        date_tuple = email.utils.parsedate_tz(header)
        local_date = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
        return local_date

    def _parse_message_body(self, message):
        body = None
        if message.is_multipart():
            for part in message.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    body = part.get_payload(decode=True).decode()
        else:
            content_type = message.get_content_type()
            if content_type == "text/plain":
                body = message.get_payload(decode=True).decode()
        return body

    def __header_to_decoded_string(self, header):
        parts = []
        header = decode_header(header)
        for content, encoding in header:
            try:
                if type(content) is str:
                    parts.append(content)
                else:
                    parts.append(content.decode(encoding or "utf-8"))
            except:
                parts.append(content.decode("utf-8"))
        return "".join(parts)