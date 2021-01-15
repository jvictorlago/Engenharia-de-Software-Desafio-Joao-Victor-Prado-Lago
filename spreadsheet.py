import gspread
from gspread.models import Spreadsheet
import time

#Connects to the spreadsheet
gc = gspread.service_account('client_secret.json')
sh = gc.open(" Engenharia de Software – Desafio João Victor Prado Lago")
worksheet = sh.get_worksheet(0)

def grades_and_absence():
    row = 4
    column = 4
    x = 0
    while x < 25:
        #calculate the grades
        val = float(worksheet.cell(row, column).value)
        column += 1
        val2 = float(worksheet.cell(row, column).value)
        column += 1
        val3 = float(worksheet.cell(row, column).value)
        column = 4
        m = (val+val2+val3)/3

        #calculate the abesence percentage
        absence = float(worksheet.cell(row, 3).value)
        if absence > 15:
            worksheet.update_cell(row, 7, 'Reprovado por Falta')
            worksheet.update_cell(row, 8, '0')
        else:
            if m < 50.00:
                worksheet.update_cell(row, 7, 'Reprovado por Nota')
                worksheet.update_cell(row, 8, '0')
            elif 50.00 <= m < 70.00:
                worksheet.update_cell(row, 7, 'Exame Final')
                #Calculates the final exam grade
                naf = 100 - m
                worksheet.update_cell(row, 8, naf)
            elif m>= 70.00:
                worksheet.update_cell(row, 7, 'Aprovado')
                worksheet.update_cell(row, 8, '0')
        print ("{0:.2f}".format(m))
        row += 1
        x += 1
        #I'm using "sleep" for two seconds 'cause the Google Sheets API has a limit of requests per user per minute
        time.sleep(2)

grades_and_absence()
