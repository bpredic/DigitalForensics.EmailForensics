from datetime import datetime
import re
import string
from collections import defaultdict
import config
from mail_service import MailService
from dateutil.relativedelta import *
from nltk import word_tokenize

class MailTool:
    def __init__(self):
        self.mail_service = MailService()
        self.email = None

    def connect(self):
        email, password = self.__get_credentials()
        self.mail_service.connect()
        self.mail_service.authenticate(email, password)

    def disconnect(self):
        self.mail_service.logout()

    def count_sent_messages_hourly(self, day):
        period_start = datetime(day.year, day.month, day.day)
        period_end = period_start + relativedelta(days=+1)
        messages = self.__get_sent_messages(period_start, period_end)

        time_dictionary = self.__generate_time_dictionary(period_start, period_end, "Hourly")
        for message in messages:
            key = message['Date'].strftime("%Y-%m-%d %H:00")
            time_dictionary[key] += 1
        return time_dictionary

    def count_sent_messages_daily(self, month):
        period_start = datetime(month.year, month.month, 1)
        period_end = period_start + relativedelta(months=+1)
        messages = self.__get_sent_messages(period_start, period_end)

        time_dictionary = self.__generate_time_dictionary(period_start, period_end, "Daily")
        for message in messages:
            key = message['Date'].strftime("%Y-%m-%d")
            time_dictionary[key] += 1
        return time_dictionary

    def count_sent_messages_monthly(self, year):
        period_start = datetime(year.year, 1, 1)
        period_end = datetime(year.year, 12, 31)
        messages = self.__get_sent_messages(period_start, period_end)

        time_dictionary = self.__generate_time_dictionary(period_start, period_end, "Monthly")
        for message in messages:
            key = message['Date'].strftime("%Y-%m")
            time_dictionary[key] += 1
        return time_dictionary

    def count_sent_messages_by_domain(self, period_start, period_end):
        messages = self.__get_sent_messages(period_start, period_end)
        domain_dict = defaultdict(int)
        for message in messages:
            recievers = message['Recievers']
            if "BCC" in message and message['BCC']:
                recievers.extend(message['BCC'])
            if "CC" in message and message['CC']:
                recievers.extend(message['CC'])
            domains = self.__parse_domains(recievers)
            for domain in domains:
                domain_dict[domain] += 1
        sorted_dict = self.__sort_dictionary_by_value(domain_dict)
        return sorted_dict

    def count_most_used_keywords(self, period_start, period_end):
        messages = self.__get_sent_messages(period_start, period_end)
        token_dict = defaultdict(int)
        for message in messages:
            text_body = message['Text-Body']
            subject = message['Subject']
            text_to_process = None
            if text_body and subject:
                text_to_process = text_body + " " + subject
            elif text_body:
                text_to_process = text_body
            else:
                text_to_process = subject

            if text_to_process:
                text = re.sub('(https|http){1}:\/\/[^\s]+\.[\w\d]+\/{0,1}', '', text_to_process, flags=re.MULTILINE)  # remove links
                text = re.sub('<.*>', '', text, flags=re.MULTILINE)
                tokens = word_tokenize(text)
                tokens = list(filter(lambda token: token not in string.punctuation and token not in ['``', '\'\'', '\"\"', '...', ], tokens))
                for token in tokens:
                    token_dict[token] += 1

        sorted_dict = self.__sort_dictionary_by_value(token_dict)
        return sorted_dict

    def get_contact_interaction_weights(self, period_start, period_end):
        sent_messages = self.__get_sent_messages(period_start, period_end)
        recieved_messages = self.__get_recieved_messages(period_start, period_end)

        info = self.__get_contact_relationship_info(sent_messages, recieved_messages)
        for c in info:
            print(c, info[c])
        params = self.__get_contact_params(info)
        for c in params:
            print(c, params[c])
        weights = self.__get_contact_weights(params)

        sorted_ = self.__sort_dictionary_by_value(weights)
        for c in sorted_:
            print(c, weights[c])
        return sorted_

    def __get_contact_relationship_info(self, sent_messages, recieved_messages):
        contacts = defaultdict(lambda: {'First-Contact': datetime.max, 'Last-Contact': datetime.min,
            'From-Me': 0, 'From-Me-Cc': 0, 'From-Me-Bcc': 0, 'To-Me': 0, 'To-Me-Cc': 0, 'To-Me-Bcc': 0, 'To-Me-Groups': 0})

        #messages the current contact has sent
        for i, message in enumerate(sent_messages):
            for recipient in message['Recievers']:
                contacts[recipient]['From-Me'] += 1
                if message['Date'] < contacts[recipient]['First-Contact']:
                    contacts[recipient]['First-Contact'] = message['Date']
                if message['Date'] > contacts[recipient]['Last-Contact']:
                    contacts[recipient]['Last-Contact'] = message['Date']
            if message['CC']:
                for cc_recipient in message['CC']:
                    contacts[cc_recipient]['From-Me-Cc'] += 1
                    if message['Date'] < contacts[cc_recipient]['First-Contact']:
                        contacts[cc_recipient]['First-Contact'] = message['Date']
                    if message['Date'] > contacts[cc_recipient]['Last-Contact']:
                        contacts[cc_recipient]['Last-Contact'] = message['Date']
            if message['BCC']:
                for bcc_recipient in message['BCC']:
                    contacts[bcc_recipient]['From-Me-Bcc'] += 1
                    if message['Date'] < contacts[bcc_recipient]['First-Contact']:
                        contacts[bcc_recipient]['First-Contact'] = message['Date']
                    if message['Date'] > contacts[bcc_recipient]['Last-Contact']:
                        contacts[bcc_recipient]['Last-Contact'] = message['Date']

        #messages the current person has recieved
        email = str.strip(self.__get_email())
        print(email)
        for i, message in enumerate(recieved_messages):
            sender = message['Sender']
            if email in message['Recievers']:
                contacts[sender]['To-Me'] += 1
            elif message['CC'] is not None and email in message['CC']:
                contacts[sender]['To-Me-Cc'] += 1
            elif message['BCC'] is not None and email in message['BCC']:
                contacts[sender]['To-Me-Bcc'] += 1
            else:
                contacts[sender]['To-Me-Groups'] += 1

            if message['Date'] < contacts[sender]['First-Contact']:
                contacts[sender]['First-Contact'] = message['Date']
            if message['Date'] > contacts[sender]['Last-Contact']:
                contacts[sender]['Last-Contact'] = message['Date']
        return contacts

    def __get_contact_params(self, contact_info):
        params = defaultdict(lambda: {})
        for contact in contact_info:
            entry = contact_info[contact]
            num_recieved = entry['To-Me'] + entry['To-Me-Cc'] + entry['To-Me-Bcc'] + entry['To-Me-Groups']
            num_sent = entry['From-Me'] + entry['From-Me-Cc'] + entry['From-Me-Bcc']
            num_to_me = entry['To-Me']
            num_from_me =  entry['From-Me']
            num_sec = entry['To-Me-Cc'] + entry['To-Me-Bcc'] + entry['To-Me-Groups'] \
                + entry['From-Me-Cc'] + entry['From-Me-Bcc']
            num_total = num_to_me + num_from_me + num_sec
            length = (entry['Last-Contact'] - entry['First-Contact']).days
            length = length if length > 0 else 1

            params[contact]['Recen'] = (datetime.now() - entry['Last-Contact']).days
            params[contact]['Len'] = length
            params[contact]['Sent-Freq'] = float(num_sent) / length
            params[contact]['Recv-Freq'] = float(num_recieved) / length
            params[contact]['To-Me'] = float(num_to_me) / num_total
            params[contact]['From-Me'] = float(num_from_me) / num_total
            params[contact]['Sec'] = float(num_sec) / num_total
            params[contact]['Recip'] = 1 - (abs(num_recieved - num_sent) / float(num_total))
        return params

    def __get_contact_weights(self, params):
        weights = defaultdict(int)
        min_sent_freq = min(params.values(), key = lambda x : x['Sent-Freq'])['Sent-Freq']
        max_sent_freq = max(params.values(), key = lambda x : x['Sent-Freq'])['Sent-Freq']
        min_recv_freq = min(params.values(), key = lambda x : x['Recv-Freq'])['Recv-Freq']
        max_recv_freq = max(params.values(), key = lambda x : x['Recv-Freq'])['Recv-Freq']
        min_recen = min(params.values(), key=lambda x: x['Recen'])['Recen']
        max_recen = max(params.values(), key=lambda x: x['Recen'])['Recen']

        print('aaa', min_sent_freq, max_sent_freq, min_recv_freq, max_recv_freq, min_recen, max_recen)
        for contact in params:
            entry = params[contact]
            norm_sent_freq = (entry['Sent-Freq'] - min_sent_freq) / (max_sent_freq - min_sent_freq)
            norm_recv_freq = (entry['Recv-Freq'] - min_recv_freq) / (max_recv_freq - min_recv_freq)
            norm_recen = (entry['Recen'] - min_recen) / (max_recen - min_recen)

            most_infl = entry['Recip'] + entry['From-Me'] + norm_sent_freq
            medium_infl = norm_recv_freq + entry['To-Me']
            less_infl = norm_recen + entry['Sec']
            weights[contact] = most_infl + 0.5 * medium_infl + 0.3 * less_infl
        return weights

    def __get_sent_messages(self, period_start, period_end):
        self.mail_service.select_mailbox(config.sent_folder)
        sent_messages = self.mail_service.get_message_info_for_period(period_start, period_end)
        self.mail_service.close_mailbox()
        return sent_messages

    def __get_recieved_messages(self, period_start, period_end):
        self.mail_service.select_mailbox(config.recieved_folder)
        recieved_messages = self.mail_service.get_message_info_for_period(period_start, period_end)
        self.mail_service.close_mailbox()
        return recieved_messages

    def __get_credentials(self):
        f = open("credentials.txt", "r")
        email = f.readline(400)
        password = f.readline(400)
        return email, password

    def __get_email(self):
        f = open("credentials.txt", "r")
        email = f.readline(400)
        return email

    def __generate_time_dictionary(self, period_start, period_end, mode):
        day = period_start
        time_dictionary = {}
        while day < period_end:
            if mode == "Hourly":
                key = day.strftime("%Y-%m-%d %H:00")
                day = day + relativedelta(hours=+1)
            elif mode == "Daily":
                key = day.strftime("%Y-%m-%d")
                day = day + relativedelta(days=+1)
            else:
                key = day.strftime("%Y-%m")
                day = day + relativedelta(months=+1)
            time_dictionary[key] = 0
        return time_dictionary

    def __parse_domains(self, emails):
        domains = []
        for email in emails:
            match = re.search('@[\w\-\.]+\.[\w\-\.]+', email)
            if match:
                domains.append(match.group(0))
        return domains

    def __sort_messages_by_date(self, messages):
        sorted_ = sorted(messages, key=lambda message: message['Date'], reverse=True)
        return sorted_

    def __sort_dictionary_by_value(self, dict):
        sorted_ = {key: value for key, value in sorted(dict.items(), key=lambda item: item[1], reverse=True)}
        return sorted_






