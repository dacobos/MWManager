################################  MODULE  INFO  ################################
# Author: David  Cobos
# Cisco Systems Solutions Integrations Architect
# Mail: cdcobos1999@gmail.com  / dacobos@cisco.com
##################################  IMPORTS   ##################################

# Following task shall be done
# TODO: Install python libraries: pyOpenSSL, requests, json, smtplib, flask, threading, subprocess, shutil
# TODO: Define environment variable SMARTSHEET_ACCESS_TOKEN as the credentiasl to access smartsheet
# TODO: Define environment variable GMAIL_ACCOUNT as the source of the Emails_Ventana
# TODO: Define environment variable GMAIL_PWD as the password of the GMAIL_ACCOUNT

import sys
import json
import requests
import os


# Version1
def getSheet(sheet_id, token):
    #Get the whole sheet of a selected sheet
    url='https://api.smartsheet.com/2.0/sheets/'+sheet_id
    headers = {'Authorization': 'Bearer '+token,'Content-Type': 'application/json'}
    r = requests.get(url,headers=headers)
    sheet = json.loads(r.content)
    return sheet

def getRow(rows, act_id):
    #Get the row that matches the act_id value
    for row in rows:
        try:
            if row["cells"][0]["value"] == act_id:
                return row["cells"]
        except KeyError:
            continue

def getActi(sheet_id, token):
    sheet = getSheet(sheet_id, token)
    acti = []

# sheet = getSheet(sheet_id = "85520752633732", token=os.environ['SMARTSHEET_ACCESS_TOKEN'])
    for row in sheet["columns"]:
        if row["index"] == 0:
            ActId_columnId=row["id"]
        elif row["index"] == 1:
            Name_columnId=row["id"]

    for row in sheet["rows"]:
        for column in row["cells"]:
            if column["columnId"] == ActId_columnId:
                ActID =  column["value"]
            elif column["columnId"] == Name_columnId:
                Name =  column["value"]
        acti.append({"actId":ActID,"Name":Name})

    # Ex: acti = [{"actId":"77777-VM04", "Name":"Software Upgrade"},{"actId":"0001-VM01", "Name":"Integracion"}]
    return acti











#Version2.0
def getSheet2(sheet_id, token):
    #Get the whole sheet of a selected sheet
    url='https://api.smartsheet.com/2.0/sheets/'+sheet_id
    headers = {'Authorization': 'Bearer '+token,'Content-Type': 'application/json'}
    response = requests.get(url,headers=headers)
    return response.text

def updateVal(sheet_id, rowId, columnId, newStatus, token):
    url='https://api.smartsheet.com/2.0/sheets/'+sheet_id+'/rows'
    payload = "{{\"id\": {}, \"cells\": [{{\"columnId\": {},\"value\": \"{}\"}}]}}"
    payload = payload.format(rowId, columnId, newStatus)

    # headers omitted from here for privacy
    headers = {'Authorization': 'Bearer '+token,'Content-Type': 'application/json'}
    response = requests.request("PUT", url, data=payload, headers=headers)
    return response

def getPosition(sheet_id, ActId, Val_Header, token):
    sheet = getSheet(sheet_id, token)
    result = {}
    try:
        for row in sheet["columns"]:
            if row["title"] == Val_Header:
                Val_index = row["index"]
            #
        for row in sheet["rows"]:
            try:
                if row["cells"][0]["value"] == ActId:
                    result["rowId"] = row["id"]
                    result["columnId"] = row["cells"][Val_index]["columnId"]
            except:
                continue
    except:
        result = {}

    return result
# Call of function Ex: recipients = getVal(sheet_id = '6595213704619908', ActId = '840358-VM01', Val_Header="Emails_Ventana", token=os.environ['SMARTSHEET_ACCESS_TOKEN']).split(", ")
def getVal(sheet_id, ActId, Val_Header, token):
    with open("/Users/dacobos/Development/MWManager/logs/"+sheet_id) as f:
        sheet = f.read()
    # Convert the string of the sheet to a sheet in JSON
    sheet = json.loads(sheet)

    for row in sheet["columns"]:
        if row["title"] == Val_Header:
            Val_index = row["index"]

    for row in sheet["rows"]:
        try:
            if row["cells"][0]["value"] == ActId:
                result = row["cells"][Val_index]["value"]
        except:
            continue


    return result

def getFolders(workspace, token):
    #Get all the sheets_id of a selected workspace and return each name and sheet_id
    url='https://api.smartsheet.com/2.0/workspaces/'+workspace
    headers = {'Authorization': 'Bearer '+token,'Content-Type': 'application/json'}
    r = requests.get(url,headers=headers)
    folders = json.loads(r.content)
    return folders


def getSheets(folderId, token):
    #Get all the sheets_id of a selected workspace and return each name and sheet_id
    url='https://api.smartsheet.com/2.0/folders/'+folderId
    headers = {'Authorization': 'Bearer '+token,'Content-Type': 'application/json'}
    r = requests.get(url,headers=headers)
    sheets = json.loads(r.content)
    return sheets

# Call of function Ex:
# columnId = ['2540241916585860']
# value = ['Ventana de Integracion PE ASR9006']
# response = updateRow(sheet_id='6595213704619908', row_id='3475399882631044', columnId=columnId, value = value, token = os.environ['SMARTSHEET_ACCESS_TOKEN'])

def updateRow(sheet_id, row_id, columnId, value, token):
    url='https://api.smartsheet.com/2.0/sheets/'+sheet_id+'/rows'
    payload = "{{\"id\": {}, \"cells\":[\
    {{\"columnId\": {},\"value\": \"{}\"}},\
    {{\"columnId\": {},\"value\": \"{}\"}},\
    {{\"columnId\": {},\"value\": \"{}\"}},\
    {{\"columnId\": {},\"value\": \"{}\"}},\
    {{\"columnId\": {},\"value\": \"{}\"}},\
    {{\"columnId\": {},\"value\": \"{}\"}},\
    {{\"columnId\": {},\"value\": \"{}\"}},\
    {{\"columnId\": {},\"value\": \"{}\"}},\
    {{\"columnId\": {},\"value\": \"{}\"}}]}}"

    # Remove any "\n or \r"
    value = [w.replace('\n',' ') for w in value]

    # Fill all the args from columnId and value
    payload = payload.format(row_id, columnId[0], \
    value[0], columnId[1], value[1], columnId[2], \
    value[2], columnId[3], value[3], columnId[4], \
    value[4], columnId[5], value[5], columnId[6], \
    value[6], columnId[7], value[7], columnId[8], \
    value[8])

    headers = {'Authorization': 'Bearer '+token,'Content-Type': 'application/json'}
    response = requests.request("PUT", url, data=payload, headers=headers)
    return response
