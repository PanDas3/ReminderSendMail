from configparser import ConfigParser
from configparser import MissingSectionHeaderError
from sys import exc_info
from sys import exit
from log import Log

class Configuration():
    def __init__(self):
        self.log_error = Log().log_error
        self.log_info = Log().log_info

    def read_config(self, file_config):
        try:
            config = ConfigParser()
            config.read(file_config)

            # Read Excel Section
            self.expired_days = config["Excel"]["days"]
            self.expired_days_again = config["Excel"]["days_again"]
            self.excel_path = config["Excel"]["excel_path"]
            self.excel_column_date = config["Excel"]["column_date"]
            self.excel_sheets = config["Excel"]["sheets"]

            # Read SendMail Section
            self.style = config["SendMail"]["style"]
            self.sender_email = config["SendMail"]["sender"]
            self.receiver_email = config["SendMail"]["receivers"]

            # Read SMTP Section
            self.smtp_server = config["SMTP"]["server"]
            self.port = config["SMTP"]["port"]

        except FileNotFoundError as err:
            print(err)
            self.log_error(f"Config Error: {err}")
            exit()

        except MissingSectionHeaderError as err:
            print(err)
            self.log_error(f"Config Error: {err}")
            exit()
        
        except:
            print(exc_info()[0])
            self.log_error(exc_info()[0])
            exit()

    def get_excel_params(self):
        expired_days = self.expired_days
        expired_days_again = self.expired_days_again
        excel_path = self.excel_path
        excel_column_date = self.excel_column_date
        excel_sheets = self.excel_sheets
       
        expired_days = int(expired_days)
        expired_days_again = int(expired_days_again)

        excel_sheets = excel_sheets.replace(",", ", ").replace("  ", " ")
        excel_sheets = list(excel_sheets.split(", "))
     
        return [expired_days, expired_days_again, excel_path, excel_column_date, excel_sheets]

    def get_sendmail_params(self):
        style = self.style
        sender_email = self.sender_email
        receiver_email = self.receiver_email

        receiver_email = receiver_email.replace(",", ", ").replace("  ", " ")
        receiver_email = list(receiver_email.split(", "))

        return [style, sender_email, receiver_email]

    def get_smtp_params(self):
        smtp_server = self.smtp_server
        port = int(self.port)

        return [smtp_server, port]

    def create_config(self, file_config):

        default_cfg = """[Excel]
# Nazwy arkuszy w pliku Excel - Program będzie szukał kolumny "Data" w niżej wymienionych arkuszach/zakładkach - !! Wielkość znaków ma znaczenie !!
sheets = Arkusz1, Arkusz2
# Nazwa kolumny z datą w pliku Excelu - !! Wielkość znaków ma znaczenie - kolumna musi się tak samo nazywać we wszystkich arkuszach/zakładkach !!
column_date = Data_obowiazywania
# Liczba dni poprzedzająca wygaśnięcie wniosku (mail zostanie wysłany tyle dni przed)
days = 7
# Ponowne przypomnienie
days_again = 3
# Lokalizacja pliku Excel - Lokalizacja pisana ciągiem
excel_path = c:\\x.xlsx

[SendMail]
# Adres E-Mail, z którego będzie wysyłany mail
sender = przypominajka@gmail.com
# Lista odbiorców
receivers = test@gmail.com, test@yahoo.com
# Jaki styl tabelki? (grey_dark) Do wyboru: blue_light/dark, grey_light/dark, orange_light/dark, yellow_light/dark, green_light/dark, red_light/dark
style = red_light

[SMTP]
# Nazwa serwera SMTP - np. smtp.google.com
server = bpsmtpqa.qa.bpsa.pl
# Port serwera SMTP - np. 465 lub 587
port = 465"""

        try:
            open(file_config)

        except IOError:
            with open("config.ini", mode="w", encoding="UTF-8") as default_config:
                    default_config.write(default_cfg)
                    default_config.close()

            err = "Config Error: Nie znaleziono config.ini lub jest on uszkodzony !"
            print(err)
            self.log_error(f"{err}\nUsuń plik, skrypt wygeneruje domyślną konfigurację.")
            exit()

        except:
            print(exc_info()[0])
            self.log_error(exc_info()[0])
            exit()        

    def __del__(self):
        pass