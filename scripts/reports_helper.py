
import sys
import json
import requests
import os
from os import listdir
import xlsxwriter
import datetime




def getSpecReport(sheet_id, PID):
    report_path = os.environ['APP_HOME']+'logs'
    date = str(datetime.datetime.now())
    # Get the previuos report name and delete it
    lis = listdir(report_path)
    for i in range(len(lis)):
        if PID in lis[i]:
            try:
                print lis[i]
                os.remove(report_path+'/'+lis[i])
            except:
                continue


    # Get that sheet out of there
    with open(os.environ['APP_HOME']+"logs/"+sheet_id) as f:
        sheet = f.read()


    # Convert the string of the sheet to a sheet in JSON
    sheet = json.loads(sheet)


    # Get entries
    # Variables
    entries = []
    keys = ['ACT_ID','Descripcion', 'Nombre_Ventana', 'Fecha_Ventana', 'Horario_Ventana', 'NCE_Ventana', 'Entrega_Doc_Ventana', 'Requerimientos_Ventana','Estado_Ventana','Emails_Ventana']

    # Get the list of columnId
    columnsId = {}

    for column in sheet["columns"]:
        columnsId[column["title"]]= column["id"]

        # Get Line
    for row in sheet["rows"]:
        # Check if the Cell does have the VM code else continue with the next row
        try:
            if PID+"-VM" in row["cells"][0]["value"]:
                pass
            else:
                continue
        except:
            continue
        line = []
        # If the cell does have the VM code will execute this
        for cell in row["cells"]:
            for key in keys:
                if cell["columnId"] == columnsId[key]:
                    try:
                        line.append(cell["value"])
                    except:
                        line.append(" ")
        dic={}
        for i in range(len(keys)):
            dic[keys[i]] = line[i]
        entries.append(dic)

    # Set the new report name
    filename = 'Reporte_Ventanas_'+PID+'_'+date+'.xlsx'
    workbook = xlsxwriter.Workbook(report_path+'/'+filename)
    worksheet = workbook.add_worksheet()
    worksheet.set_column(0, 0, 15)
    worksheet.set_column(1, 1, 25)
    worksheet.set_column(2, 2, 40)
    worksheet.set_column(3, 4, 15)
    worksheet.set_column(5, 5, 20)
    worksheet.set_column(6, 6, 15)
    worksheet.set_column(7, 7, 40)
    worksheet.set_column(8, 8, 15)
    worksheet.set_column(9, 9, 30)

    format1=workbook.add_format({'bold': True, 'font_color': 'white', 'border':1, 'bg_color':'gray', 'align':'center'})
    format2=workbook.add_format({'border':1, 'valign':'vcenter', 'align':'center'})
    format2.set_text_wrap()

    row = 0
    col = 0
    # Re order the Keys for XSLS
    headers = ['ACT_ID','Nombre_Ventana','Descripcion', 'Fecha_Ventana', 'Horario_Ventana', 'NCE_Ventana', 'Entrega_Doc_Ventana', 'Requerimientos_Ventana','Estado_Ventana','Emails_Ventana']
    # Write the Headers
    for col in range(len(headers)):
        worksheet.write(0, col, headers[col], format1)
    # Write the data

    for entry in entries:
        row = row+1
        for col in range(len(keys)):
            worksheet.write(row, col, entry[headers[col]],format2)

    workbook.close()



    return filename



def getEntries(sheet_id, PID, token):
    #Create a text file wich contains all the content of the selected sheet_id= '6595213704619908', token=os.environ['SMARTSHEET_ACCESS_TOKEN'],
    #Get the whole sheet of a selected sheet
    url='https://api.smartsheet.com/2.0/sheets/'+sheet_id
    headers = {'Authorization': 'Bearer '+token,'Content-Type': 'application/json'}
    response = requests.get(url,headers=headers)
    sheet = response.text
    sheet = sheet.encode('utf-8')
    with open(os.environ['APP_HOME']+"logs/"+sheet_id,"w") as f:
        f.write(sheet)
    # Convert the string of the sheet to a sheet in JSON
    sheet = json.loads(sheet)

    # Get entries
    # Variables
    entries = []
    keys = ['ACT_ID','Descripcion', 'Nombre_Ventana', 'Fecha_Ventana', 'Horario_Ventana', 'NCE_Ventana', 'Entrega_Doc_Ventana', 'Requerimientos_Ventana','Estado_Ventana','Emails_Ventana']

    # Get the list of columnId
    columnsId = {}

    for column in sheet["columns"]:
        columnsId[column["title"]]= column["id"]

        # Get Line
    for row in sheet["rows"]:
        # Check if the Cell does have the VM code else continue with the next row
        try:
            if PID+"-VM" in row["cells"][0]["value"]:
                pass
            else:
                continue
        except:
            continue
        line = []
        # If the cell does have the VM code will execute this
        for cell in row["cells"]:
            for key in keys:
                if cell["columnId"] == columnsId[key]:
                    try:
                        line.append(cell["value"])
                    except:
                        line.append(" ")
        dic={}
        for i in range(len(keys)):
            dic[keys[i]] = line[i]
        entries.append(dic)

    return entries


def getGeneralReport(projectsList, token):
    report_path = os.environ['APP_HOME']+'logs'
    date = str(datetime.datetime.now())
    # Get the previuos report name and delete it
    lis = listdir(report_path)
    for i in range(len(lis)):
        if 'Reporte_General_Ventanas' in lis[i]:
            try:
                os.remove(report_path+'/'+lis[i])
            except:
                continue

    # Build a list with all PIDs
    report = []
    for l in projectsList:
        sheet_id = str(l[1])
        PID= l[0]
        entry = getEntries(sheet_id, PID, token)
        for ent in entry:
            report.append(ent)

    # Set the new report name
    filename = 'Reporte_General_Ventanas_'+date+'.xlsx'
    workbook = xlsxwriter.Workbook(report_path+'/'+filename)
    worksheet = workbook.add_worksheet()
    worksheet.set_column(0, 0, 15)
    worksheet.set_column(1, 1, 25)
    worksheet.set_column(2, 2, 40)
    worksheet.set_column(3, 4, 15)
    worksheet.set_column(5, 5, 20)
    worksheet.set_column(6, 6, 15)
    worksheet.set_column(7, 7, 40)
    worksheet.set_column(8, 8, 15)
    worksheet.set_column(9, 9, 30)

    format1=workbook.add_format({'bold': True, 'font_color': 'white', 'border':1, 'bg_color':'gray', 'align':'center'})
    format2=workbook.add_format({'border':1, 'valign':'vcenter', 'align':'center'})
    format2.set_text_wrap()

    row = 0
    col = 0
    # Re order the Keys for XSLS
    headers = ['ACT_ID','Nombre_Ventana','Descripcion', 'Fecha_Ventana', 'Horario_Ventana', 'NCE_Ventana', 'Entrega_Doc_Ventana', 'Requerimientos_Ventana','Estado_Ventana','Emails_Ventana']
    # Write the Headers
    for col in range(len(headers)):
        worksheet.write(0, col, headers[col], format1)
    # Write the data
    for entry in report:
        row = row+1
        for col in range(len(headers)):
            worksheet.write(row, col, entry[headers[col]],format2)

    workbook.close()

    return filename
