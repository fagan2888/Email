class Email_Helper:
    '''
    Retrieve and send emails.
    '''
    __email_add = ''
    __password = ''
    __smtpObj = None
    __mail_client = None
    __email_uid_list = []
    
    def __init__(self, email_add, password):
        self.__email_add = email_add
        self.__password = password
        
    def login(self, email_client='outlook.com'):
        self.__smtpObj = smtplib.SMTP(email_client, 587)
        self.__smtpObj.starttls()
        self.__smtpObj.login(self.__email_add, self.__password)
        
    def send_email(self, email_recepient, subject='', body=''):
        msg = MIMEText(body)
        msg['Subject'] = subject
        msg['From'] = self.__email_add
        msg['To'] = email_recepient
        
        self.__smtpObj.sendmail(self.__email_add, email_recepient, msg.as_string())
        
    def logout(self):
        self.__smtpObj.quit()
        
    def login_imap(self, email_add='', password='', email_client='outlook.com'):
        self.__mail_client = imaplib.IMAP4_SSL('imap.' + email_client)
        
        if email_add == '' or password == '':
            if self.__email_add != '' and self.__password != '':
                self.__mail_client.login(self.__email_add, self.__password)
            else:
                print('Missing email address or password.')
                return
        else:
            self.__mail_client.login(email_add, password)
    
    def get_mail_folders(self):
        if self.__mail_client == None:
            print('Call login_imap() first.')
            return
        else:
            return self.__mail_client.list()
    
    def select_mail_folder(self, folder):
        if self.__mail_client == None:
            print('Call login_imap() first.')
            return
        else:
            self.__mail_client.select(folder)
    
    def get_mail_UIDs(self, criteria1=None, criteria2='ALL'):
        '''
        UIDs unique to each email.
        
        Other option: mail.uid('search', "SUBJECT", "Test")
        '''
        try:
            result, all_mail_data = self.__mail_client.uid('search', criteria1, criteria2)
        except:
            print('Folder selected is not valid.')
            return
            
        ids = all_mail_data[0] # data is a list.
        self.__email_uid_list = ids.split() # ids is a space separated string
        
        return self.__email_uid_list

    def __parse_email_body(self, email_message):
        email_body=email_message.get_payload()[0].get_payload();
        if type(email_body) is list:
            email_body=','.join(str(v) for v in email_body)

        final_email_body = ''
        paragraphs = justext.justext(email_body, justext.get_stoplist("English"))
        for j, paragraph in enumerate(paragraphs):
#             print('{}. \n{}'.format(j, paragraph.text))
            if 'Microsoft' not in paragraph.text:
                final_email_body += paragraph.text + '\n'
            
        return final_email_body
    
    def parse_emails(self):
        '''
        Returns email_list: [Date, To, From, Subject, Body].
        '''
        
        if self.__mail_client == None:
            print('Call login_imap() first.')
            return
        
        email_list = []
        for i in range(len(self.__email_uid_list)):
            email_uid = self.__email_uid_list[i]
            result, email_data = self.__mail_client.uid('fetch', email_uid, '(RFC822)')
            raw_email = email_data[0][1]
            raw_email_string = raw_email.decode('utf-8')
            email_message = email.message_from_string(raw_email_string)

            email_date = datetime.datetime(*email.utils.parsedate_tz(email_message['Date'])[0:6])
            email_to = '\n'.join(email_message['To'].replace('\r\n\t', '').split(','))
            email_from = '\n'.join(email_message['From'].replace('\r\n\t', '').split(','))
            email_subject = email_message['Subject'].strip()

            email_body = ''
            try:
                email_body = self.__parse_email_body(email_message)
            except:
                pass

#             print('{}.\nDATE: {}\n\nTO:\n{}\n\nFROM:\n{}\n\nSubject:\n{}\n\nBody: \n{}\n\n\n'.format(
#                     i+1,
#                     email_date,
#                     email_to, 
#                     email_from,
#                     email_subject,
#                     email_body))
            
            email_list.append([email_date, email_to, email_from, email_subject, email_body])

    def download_email_attachments(self, dwld_dir='C:/Users/Alex/Downloads'):
        '''
        Just use email ids, instead of UIDs. Problem with decoding.
        '''
        if self.__mail_client is None:
            print('Call login_imap() first.')
            return

        typ, msgs = self.__mail_client.search(None, 'ALL')
        msgs = msgs[0].split()
        
        for emailid in msgs:
            resp, data = mail.fetch(emailid, "(RFC822)")
            email_body = data[0][1] 
            m = email.message_from_string(email_body.decode('utf-8'))

            if m.get_content_maintype() != 'multipart':
                continue

            for part in m.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue

                filename=part.get_filename()
                if filename is not None:
                    sv_path = os.path.join(dwld_dir, filename)
                    if not os.path.isfile(sv_path):
        #                 print (sv_path)      
                        try:
                            fp = open(sv_path, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()
                        except:
                            try:
                                # filename is too long
                                sv_path = sv_path.split('.')[0][:50] + '.' + \
                                          sv_path.split('.')[1]
                                fp = open(sv_path, 'wb')
                                fp.write(part.get_payload(decode=True))
                                fp.close()
                            except:
                                # bad characters of filename
                                sv_path = sv_path.replace('?', '').replace('=', '')
                                fp = open(sv_path, 'wb')
                                fp.write(part.get_payload(decode=True))
                                fp.close()