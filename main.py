# pip install pandas, pyinstaller, configparser, smtplib, configupdater, xlrd, openxl, Jinja2, pretty_html_table

from time import sleep
from sys import exit
from mail import SendMail
from log import Log
from config import Configuration
from excel import Excel

if __name__ == "__main__":

    file_config = "config.ini"
    print("##########################")
    print("### Powered by Majster ###")
    print("##########################\n")

    # New istance class
    cfg = Configuration()
    excel = Excel()
    sendmail = SendMail()
    log_error = Log().log_error
    log_info = Log().log_info

    # Read information from config.ini
    cfg.create_config(file_config)                  # If not exist -> Create default and exit()
    cfg.read_config(file_config)                    # Read all configuration parameters
    cfg_excel = cfg.get_excel_params()              # Sort excel parameters
    cfg_mail_config = cfg.get_sendmail_params()     # Sort mail parameters
    cfg_smtp_config = cfg.get_smtp_params()         # Sort smtp parameters

    # Execute results
    excel.set_config_params(cfg_excel)               # Move excel parameters
    excel_out = excel.read_sheets()                  # Build and return excel content
    if(excel_out != False):
        temat = excel_out[0]
        content = excel_out[1]
    else:
        del cfg
        del excel
        del sendmail
        exit()

    # Exectue send mail
    sendmail.set_config(cfg_mail_config, cfg_smtp_config)   # Move frame mail parameters
    sendmail.set_content(content)                           # Build content

    if sendmail.send_mail(temat) == True:                   # Send mail
        log_info("Znaleziono pasujące kryteria")
        log_info("Send Mail - SUCCESS")
        sleep(2.0)
        
    else:
        print("Błąd wysyłki Maila")
        log_error("Send Mail - FAILED")
        sleep(2.0)

    del cfg
    del excel
    del sendmail
    exit()