from pandas import read_excel
from pandas import DataFrame
from pandas import to_datetime
from pandas import core
# from pandas import core.frame.DataFrame
from datetime import date
from datetime import timedelta
from sys import exit
from log import Log
from mail import SendMail
from sys import exc_info

class Excel():
    def __init__(self):
        self.log_error = Log().log_error
        self.log_info = Log().log_info

    def set_config_params(self, param):
    
        self.expired_days = param[0]
        self.expired_days_again = param[1]
        self.excel_path = param[2]
        self.excel_column_date = param[3]
        self.excel_sheets = param[4]

    def read_sheets(self):

        expired_days = self.expired_days
        expired_days_again = self.expired_days_again
        excel_path = self.excel_path
        excel_column_date = self.excel_column_date
        excel_sheets = self.excel_sheets

        today = date.today()
        day = today + timedelta(expired_days)           # Pierwotne przypomnienie
        day_again = today + timedelta(expired_days_again)    # Powtorne przypomnienie

        day = day.strftime("%Y-%m-%d")
        day_again = day_again.strftime("%Y-%m-%d")

        print("Dzisiaj jest:", today)
        print("Zakres od", day_again, "do", day,"\n")

        dataframe = core.frame.DataFrame(None)
        table = []

        try:

            for sheet in excel_sheets:
                excel = read_excel(excel_path, sheet_name=sheet)
                df = DataFrame(excel)
                df[excel_column_date] = to_datetime(df[excel_column_date])

                mask = (df[excel_column_date] >= day_again) & (df[excel_column_date] <= day)
                df = df.loc[mask]
                dataframe = dataframe.append(df)
                csv = df.to_csv(index=False)
                table.append(f"{csv}")

            excel = str(dataframe[excel_column_date])

            if((day_again != day) and ((day_again in excel) and (day in excel))):
                print("Znaleziono - Zbiorcze !")
                temat = "!!! Przypominajka - Zbiorcza !!!"
                return [temat, table]

            elif((day in excel) or (day_again == day)):
                print("Znaleziono !")
                temat = "!!! Przypominajka !!!"
                return [temat, table]

            elif(day_again in excel):
                print("Znaleziono - Przypomnienie !")
                temat = "!!! Przypominajka - Powtorne powiadomienie !!!" # Jeżeli jest znak specjalny np. ó to mail się nie wyślę
                return [temat, table]

            else:
                print("Brak znalezionych kryetriów do wysyłki przypomnienia")
                self.log_info("Brak znalezionych kryetriów do wysyłki przypomnienia")
                return False

        except FileNotFoundError as err:
            print(err)
            self.log_error(f"Config/Excel Error: {err}")
            exit()

        except ValueError as err:
            print("Nie znaleziono kolumny wskazanej w config.ini")
            self.log_error(f"Config/Excel Error: {err} \nCzy plik config.ini jest zapisany jako ASCII ?")
            quit()

        except KeyError as err:
            print("Błąd pliku config.ini")
            self.log_error(f"Config Error: {err}")
            exit()
        
        except:
            print(exc_info()[0])
            self.log_error(exc_info()[0])
            exit()

    def __del__(self):
        pass