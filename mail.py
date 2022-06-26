from pandas import DataFrame
from pandas import read_csv
from smtplib import SMTP
from smtplib import SMTPConnectError
from smtplib import SMTPServerDisconnected
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pretty_html_table import build_table
from socket import gaierror
from io import StringIO
from sys import exc_info
from log import Log

class SendMail():
    def __init__(self):
        pass

    def set_config(self, cfg_mail, cfg_smtp):
        
        # return [style, sender_email, receiver_email]
        self.style = cfg_mail[0]
        self.sender_email = cfg_mail[1]
        self.receiver_email1 = cfg_mail[2]

        # return [smtp_server, port]
        self.smtp_server = cfg_smtp[0]
        self.port = cfg_smtp[1]
        
    def set_content(self, content):

        style = self.style
        content_out_html_final = ""

        for sheet in content:
            sheet = StringIO(sheet)
            content_out = read_csv(sheet, sep=",")
            content_out = DataFrame(content_out)

            if(content_out.empty != True):
                content_out_text = content_out.to_string()
                content_out_text = content_out_text.replace("NaN", "Brak wpisu")

                content_out_html = build_table(content_out, style)
                content_out_html_final = content_out_html_final + content_out_html + "\n<br />\n"

        self.text = f"""\
                Zajrzyj do pliku excel z dostępami. Zaraz jakieś wygasną !

                {content_out_text}

                ----------------------------
                Powered by Majster
                """

        self.html = f"""\
                <div style="font-weight: bold; font-size: 20; font-famili: Comic Sans MS;">
                    Zajrzyj do pliku excel z dostępami.
                    <div style="color: red;">
                        Zaraz jakieś wygasną !
                    </div>
                </div> 
                <br /><br />

                {content_out_html_final}

                <br />----------------------------<br />
                Powered by <a href="mailto://rachuna.mikolaj@gmail.com" style="color: red; text-decoration: none; font-weight: bold">Majster</a>
                """ 

    def send_mail(self, temat):

        log_error = Log().log_error

        sender_email = self.sender_email
        receiver_email1 = self.receiver_email1
        text = self.text
        html = self.html
        smtp_server = self.smtp_server
        port = self.port

        message = MIMEMultipart("alternative")
        message["Subject"] = temat
        message["From"] = sender_email

        part1 = MIMEText(text, "plain")
        part2 = MIMEText(html, "html")

        message.attach(part1)
        message.attach(part2)

        try:
            server = SMTP(smtp_server, port)
            server.ehlo_or_helo_if_needed()

            for receiver_email in receiver_email1:
                message["To"] = receiver_email
                server.sendmail(sender_email, receiver_email, message.as_string())
                receiver_email = None
                
            server.close()
            return True
            
        except SMTPConnectError as err:
            msg = "Błąd wysyłania maila. SMTP Connection Error."
            print(msg)
            log_error(err)
            return False
        
        except SMTPServerDisconnected as err:
            msg = "Błąd wysyłania maila. SMTP Auth Error."
            print(msg)
            log_error(f"SMTP Error: {err}\nSprawdź dane wysyłającego")
            return False

        except TimeoutError as err:
            msg = "Błąd wysyłania maila. Timeout Error."
            print(msg)
            log_error(f"SMTP Error: {err}\nSprawdź port. Czy ruch jest otwarty?\n")
            return False

        except gaierror as err:
            msg = "Błąd wysyłania maila. SMTP Address Error."
            print(msg)
            log_error(f"SMTP Error: {err}\nSprawdź nazwę serwera SMTP w configu")
            return False

        except:
            print(exc_info()[0])
            log_error(exc_info()[0])
            return False

    def __del__(self):
        pass