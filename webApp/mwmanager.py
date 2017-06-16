################################  MODULE  INFO  ################################
# Author: David  Cobos
# Cisco Systems Solutions Integrations Architect
# Mail: cdcobos1999@gmail.com  / dacobos@cisco.com
##################################  IMPORTS   ##################################

# Following task shall be done
# TODO: Install python libraries: pyOpenSSL, requests, json, smtplib, flask, threading, subprocess, shutil
# TODO: Define environment variable SMARTSHEET_ACCESS_TOKEN as the credentiasl to access smartsheet
# TODO: Define environment variable SS_WORKSPACE as the workspace where all the projects are located
# TODO: Define environment variable GMAIL_ACCOUNT as the source of the Emails_Ventana
# TODO: Define environment variable GMAIL_PWD as the password of the GMAIL_ACCOUNT

# Modules Required for python web application
from flask import Flask, request, session, g, redirect, url_for, abort, \
     render_template, flash, Response, send_from_directory
import sqlite3
from os import rename
import os
import threading
import subprocess
import sys
from shutil import copyfile


# Support Scripts for web application
sys.path.insert(0, '/Users/dacobos/Development/MWManager/scripts')
from rest_helper import *

sys.path.insert(0, '/Users/dacobos/Development/MWManager/scripts')
from email_helper import *


app = Flask(__name__) # create the application instance :)
app.config.from_object(__name__) # load config from this file

# Load default config and override config from an environment variable
app.config.update(dict(
    DATABASE=os.path.join(app.root_path, 'mwmanager.db'),
    SECRET_KEY='development key',
    USERNAME='admin',
    PASSWORD='default',
    LOG_FILE = '/Users/dacobos/Development/MWManager/logs/logfile.log',
    APP_HOME = '/Users/dacobos/Development/MWManager/webApp/',
    ERROR = None,
    TOKEN=os.environ['SMARTSHEET_ACCESS_TOKEN'],
    WORKSPACE = os.environ['SS_WORKSPACE']


))

# app.config.from_envvar('FLASKR_SETTINGS', silent=True)


# Functions to CRUD on SQL Database
def connect_db():
    """Connects to the specific database."""
    rv = sqlite3.connect(app.config['DATABASE'])
    rv.row_factory = sqlite3.Row
    return rv

def get_db():
    """Opens a new database connection if there is none yet for the
    current application context.
    """
    if not hasattr(g, 'sqlite_db'):
        g.sqlite_db = connect_db()
    return g.sqlite_db

@app.teardown_appcontext
def close_db(error):
    """Closes the database again at the end of the request."""
    if hasattr(g, 'sqlite_db'):
        g.sqlite_db.close()


def init_db():
    db = get_db()
    with app.open_resource('schema.sql', mode='r') as f:
        db.cursor().executescript(f.read())
    db.commit()
    update_db()


def update_db():
    db = get_db()
    folders = getFolders(workspace = app.config['WORKSPACE'], token=app.config['TOKEN'])
    xxx = []
    for folder in folders["folders"]:
        sheets = getSheets(folderId = str(folder["id"]), token=app.config['TOKEN'])
        for sheet in sheets["sheets"]:
            sheetName = sheet["name"].split("-")
            sheetId = sheet["id"]
            try:
                projectId = sheetName[0]
                country = sheetName[1]
                capexCicle = sheetName[2]
                projectManager = sheetName[3]
                projectName = sheetName[4]
            except IndexError:
                continue
            db.execute('insert into projects (sheetId, projectId, country, capexCicle, projectManager, projectName) values (?,?,?,?,?,?)',[sheetId, projectId, country, capexCicle, projectManager, projectName])
    db.commit()


@app.cli.command('initdb')
def initdb_command():
    """Initializes the database."""
    init_db()
    """Clear Logfile."""
    print('Initialized the database.')

# Start of Web App
@app.route('/')
def index():
    return render_template('index.html')

# Version1
@app.route('/mwindows')
def mwindows():
    try:
        db = get_db()
        sheets = db.execute('select * from projects').fetchall()
        return render_template('mwindows.html', sheets = sheets)
    except sqlite3.Error as e:
        error = "Could't complete the Query: "+e.args[0]
        return render_template('mwindows.html', error = error)


def getQuery(request):
    sheet = getSheet(sheet_id = request.form["sheetId"], token=app.config['TOKEN'])
    row = getRow(rows = sheet["rows"], act_id = request.form["actId"])
    if row:
        pidName = getPidName(request.form["sheetId"])
        entry = {}
        entries = []
        keys = ['Nombre', 'Descripcion', 'Fecha', 'Horario', 'NCE', 'Entrega_Doc', 'Requerimientos','Estado']
        for i in range(len(keys)):
            entry[keys[i]] = row[i+1]["value"]
        entry["ActID"] = request.form["actId"]
        entry["PID"] = pidName[0][0]
        entry["Project"] = pidName[0][1]
        entries.append(entry)
        return entries
    else:
        error = "Couldn't find the Act ID: " + request.form["actId"]
        return render_template('mwindows.html', error = error)




@app.route('/query_mw', methods = ['GET','POST'])
def query_mw():
    if request.method == 'GET':
        return render_template('query_mw.html')
    else:

        if request.form["button"] == 'Schedule_Home':
            return redirect('mwindows')
        elif request.form["button"] == 'Email_Request':
            # Check Smartsheet status is Planned
            status = getVal(sheet_id = request.form["sheetId"], ActId = request.form['ActID'], Val_Header="Estado", token=app.config['TOKEN'])
            # Check Smartsheet status is Planned

            if status == "Planned":
                recipients = getVal(sheet_id = request.form["sheetId"], ActId = request.form['ActID'], Val_Header="Email_Recipients", token=app.config['TOKEN']).split(", ")

                if recipients == []:
                    error = "Email list is empty"
                    return render_template('mwindows.html', error = error)

                inp = [request.form['Project'],request.form['PID'],request.form['ActID'],request.form['Nombre'], request.form['Descripcion'], request.form['Fecha'], request.form['Horario'], request.form['NCE'], request.form['Entrega_Doc'], request.form['Requerimientos'],"Requested"]
                if mwMIME(inp, recipients, subject = "SDVM:Solicitud de ventana de mantenimiento") == None:
                    # Update Smartsheet status to Requested
                    position = getPosition(sheet_id = request.form["sheetId"], ActId = request.form['ActID'], Val_Header = "Estado", token=os.environ['SMARTSHEET_ACCESS_TOKEN'])
                    response = updateVal(sheet_id = request.form["sheetId"], rowId = position["rowId"], columnId = position["columnId"],  newStatus="Requested", token=os.environ['SMARTSHEET_ACCESS_TOKEN'])
                    # Update Smartsheet status to Requested
                    flash('Notification: Email requesting MW sent successfully')
                    return redirect(url_for('query_mw'))

            else:
                error = "MW Cannot be requested"
                return render_template('mwindows.html', error = error)


        elif request.form["button"] == 'Email_Program':
            # Check Smartsheet status is Requested
            status = getVal(sheet_id = request.form["sheetId"], ActId = request.form['ActID'], Val_Header="Estado", token=app.config['TOKEN'])
            # Check Smartsheet status is Requested

            if status == "Requested":
                recipients = getVal(sheet_id = request.form["sheetId"], ActId = request.form['ActID'], Val_Header="Email_Recipients", token=app.config['TOKEN']).split(", ")
                if recipients == []:
                    error = "Email list is empty"
                    return render_template('mwindows.html', error = error)

                inp = [request.form['Project'],request.form['PID'],request.form['ActID'],request.form['Nombre'], request.form['Descripcion'], request.form['Fecha'], request.form['Horario'], request.form['NCE'], request.form['Entrega_Doc'], request.form['Requerimientos'],"Programed"]
                if mwMIME(inp, recipients, subject = "PDVM:Programacion de ventana de mantenimiento") == None:
                    # Update Smartsheet status to Programed
                    position = getPosition(sheet_id = request.form["sheetId"], ActId = request.form['ActID'], Val_Header = "Estado", token=os.environ['SMARTSHEET_ACCESS_TOKEN'])
                    response = updateVal(sheet_id = request.form["sheetId"], rowId = position["rowId"], columnId = position["columnId"],  newStatus="Programed", token=os.environ['SMARTSHEET_ACCESS_TOKEN'])
                    # Update Smartsheet status to Programed
                    flash('Notification: Email programing MW sent successfully')
                    return redirect(url_for('query_mw'))


            else:
                error = "MW Cannot be programed"
                return render_template('mwindows.html', error = error)

        elif request.form["button"] == 'Next':
            if request.form["sheetId"] == '--':
                error = "You have to select a project in order to Query"
                return render_template('mwindows.html')
            else:
                param = [{"sheetId":request.form["sheetId"]}]
                # Provide an activities List "acti" of the selected sheet
                acti = getActi(sheet_id = request.form["sheetId"], token=app.config['TOKEN'])
                return render_template('act_id.html' , param = param, acti = acti)
        elif request.form["button"] == 'Query_MW':
            param = [{"sheetId":request.form["sheetId"]}]
            # Call rest helper to get the sheet and activity details
            entries = getQuery(request)
            return render_template('query_mw.html' , param = param, entries = entries)


# Version 2.0
def getPN(sheet_id):
    try:
        db = get_db()
        pidName = db.execute('select projectId, projectName from projects where sheetId = ?',[sheet_id]).fetchall()
        return pidName
    except sqlite3.Error as e:
        error = "Could't complete the Query: "+e.args[0]
        return render_template('schedule_home.html', error = error)


@app.route('/schedule_home', methods = ['GET','POST'])
def schedule_home():
    # Read the DB and enable select of work sheet
    try:
        db = get_db()
        sheets = db.execute('select * from projects').fetchall()
        return render_template('schedule_home.html', sheets = sheets)
    except sqlite3.Error as e:
        error = "Could't complete the Query: "+e.args[0]
        return render_template('schedule_home.html')



@app.route('/open_query', methods = ['GET','POST'])
def open_query():

    # # Get a list of ActID
    # with open("/Users/dacobos/Development/MWManager/logs/"+request.form["sheetId"]) as f:
    #     sheet = f.read()
    # # Convert the string of the sheet to a sheet in JSON
    # sheet = json.loads(sheet)
    try:
        sheet_id = request.form["sheetId"]
    except:
        sheet_id = request.args.get('sheetId')


    #Create a text file wich contains all the content of the selected sheet_id= '6595213704619908', token=os.environ['SMARTSHEET_ACCESS_TOKEN'],
    sheet = getSheet2(sheet_id = sheet_id, token=os.environ['SMARTSHEET_ACCESS_TOKEN'])
    sheet = sheet.encode('utf-8')
    with open("/Users/dacobos/Development/MWManager/logs/"+sheet_id,"w") as f:
        f.write(sheet)
    # Convert the string of the sheet to a sheet in JSON
    sheet = json.loads(sheet)

    param = []
    val = {}

    pidName = getPN(sheet_id)
    val["PID"] = pidName[0][0]
    val["Project"] = pidName[0][1]
    param.append(val)
    PID = str(pidName[0][0])

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

        # line = [u'881828-VM03', u'Migracion CMTS Sonsonate', u'Se migrara el CMTS de Sonsonate', u'27/11/2017', u'00:00-04:00', u'Lord Lizarazo', u'20/11/2017', u'Permisos de configuracion', u'Programed', u'dacobos@cisco.com, cdcobos1999@icloud.com']
        # Evaluate the value of the status and defined the value of the key Action in order to create each button

        if line[8] == "Propuesta":
            accion = "Enviar Solicitud"
        elif line[8] == "Solicitada":
            accion = "Enviar Programacion"
        elif line[8] == "Programada":
            accion = "Enviar Cancelacion"
        elif line[8] == "Cancelada":
            accion = "Enviar Solicitud"
        else:
            accion = "Enviar Solicitud"
        # line = [u'881828-VM03', u'Migracion CMTS Sonsonate', u'Se migrara el CMTS de Sonsonate', u'27/11/2017', u'00:00-04:00', u'Lord Lizarazo', u'20/11/2017', u'Permisos de configuracion', u'Programed', u'dacobos@cisco.com, cdcobos1999@icloud.com','Accion']

        dic={}

        for i in range(len(keys)):
            dic[keys[i]] = line[i]

        dic["Accion"] = accion
        entries.append(dic)
# Exit Get Entries

    for i in range(len(entries)):
        entries[i]["sheetId"] = sheet_id
    return render_template('open_query.html', entries = entries, param = param)

@app.route('/action', methods = ['GET','POST'])
def action():
    # Read the DB to get the SheetId based on PID
    try:
        db = get_db()
        data = db.execute('select projectName, projectId, sheetId from projects where projectId = ?',[request.form["PID"]]).fetchall()
    except sqlite3.Error as e:
        error = "Could't complete the Query: "+e.args[0]
        return render_template('schedule_home.html')
    sheet_id = str(data[0][2])
    PID = str(data[0][1])

    # Get a list of ActID
    with open("/Users/dacobos/Development/MWManager/logs/"+sheet_id) as f:
        sheet = f.read()
    # Convert the string of the sheet to a sheet in JSON
    sheet = json.loads(sheet)
    l = []
    for row in sheet["rows"]:
        try:
            if PID+"-VM" in row["cells"][0]["value"]:
                l.append(row["cells"][0]["value"])
                pass
            else:
                continue
        except:
            continue


    # Iterate the list looking who press the button
    for elem in l:
        try:
            accion = request.form[elem]
            activity = elem
        except:
            pass




    # Get recipients list from Smartsheet
    recipients = getVal(sheet_id = sheet_id, ActId = activity, Val_Header="Emails_Ventana", token=app.config['TOKEN']).split(", ")

    # Validate that there is at leas on email in the recipients list
    if recipients == []:
        error = "Email list is empty"
        return render_template('schedule_home.html', error = error)

    # Get the entries of selected mwindows
    with open("/Users/dacobos/Development/MWManager/logs/"+sheet_id) as f:
        sheet = f.read()
    # Convert the string of the sheet to a sheet in JSON
    sheet = json.loads(sheet)


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

        # line = [u'881828-VM03', u'Migracion CMTS Sonsonate', u'Se migrara el CMTS de Sonsonate', u'27/11/2017', u'00:00-04:00', u'Lord Lizarazo', u'20/11/2017', u'Permisos de configuracion', u'Programed', u'dacobos@cisco.com, cdcobos1999@icloud.com']
        # Exit Get Entries
        # Evaluate the value of the status and defined the value of the key Action in order to create each button
        if line[0] == activity:
            entries = line
        else:
            pass

    if request.form[activity] == 'Enviar Solicitud':
        newStatus = "Solicitada"
        subject = "SDVM:Solicitud de ventana de mantenimiento"
        notification = "Notification: Email requesting MW sent successfully"
    elif request.form[activity] == 'Enviar Programacion':
        newStatus = "Programada"
        subject = "PDVM:Programacion de ventana de mantenimiento"
        notification = "Notification: Email programing MW sent successfully"
    elif request.form[activity] == 'Enviar Cancelacion':
        newStatus = "Cancelada"
        subject = "CDVM:Cancelacion de ventana de mantenimiento"
        notification = "Notification: Email canceling MW sent successfully"

    # Define data to fill HTML for Multi Part Mail
    inp = [str(data[0][0]),str(data[0][1]),entries[0],entries[1],entries[2],entries[3],entries[4],entries[5],entries[6],entries[7],newStatus]
    # Send and Evaluate if the Mail return an error
    if mwMIME(inp, recipients, subject) == None:
        # Get Position at Smartsheet of State in order to update
        position = getPosition(sheet_id = sheet_id, ActId = activity, Val_Header = "Estado_Ventana", token=os.environ['SMARTSHEET_ACCESS_TOKEN'])
        # Update the Smartsheet of state Value using the position

        response = updateVal(sheet_id = sheet_id, rowId = position["rowId"], columnId = position["columnId"],  newStatus=newStatus, token=os.environ['SMARTSHEET_ACCESS_TOKEN'])
        flash(notification)
        return redirect(url_for('open_query', sheetId=sheet_id))
    else:
        error = "Could't sent the email, did not changed anything"
        return render_template('open_query',sheetId=sheet_id, error=error)

@app.route('/edit', methods = ['GET','POST'])
def edit():
    if request.method == ['GET']:
        print "Using method GET"
        return render_template('edit.html')

    elif request.method == ['POST']:
        print "Using method POST"
        return render_template('edit.html')

    # Split argument
    PID = request.args.get('ActID').split("-")[0]
    ActID = request.args.get('ActID')
    # Read the DB to get the SheetId based on PID
    try:
        db = get_db()
        data = db.execute('select projectName, projectId, sheetId from projects where projectId = ?',[PID]).fetchall()
    except sqlite3.Error as e:
        error = "Could't complete the Query: "+e.args[0]
        return render_template('schedule_home.html')

    sheet_id = str(data[0][2])

    # Get the entries of selected mwindow
    with open("/Users/dacobos/Development/MWManager/logs/"+sheet_id) as f:
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

        # line = [u'881828-VM03', u'Migracion CMTS Sonsonate', u'Se migrara el CMTS de Sonsonate', u'27/11/2017', u'00:00-04:00', u'Lord Lizarazo', u'20/11/2017', u'Permisos de configuracion', u'Programed', u'dacobos@cisco.com, cdcobos1999@icloud.com']
        # Evaluate the value of the status and defined the value of the key Action in order to create each button

        if line[0] == ActID:
            l = line
        else:
            pass

    dic={}

    for i in range(len(keys)):
        dic[keys[i]] = l[i]

    entries.append(dic)

# Exit Get Entries

    for i in range(len(entries)):
        entries[i]["sheetId"] = sheet_id

    return render_template('edit.html', entries = entries)


@app.route('/update', methods = ['GET','POST'])
def update():
    # Split PID argument
    PID = request.form["ACT_ID"].split("-")[0]
    ActID = request.form["ACT_ID"]
    keys = ['ACT_ID','Descripcion', 'Nombre_Ventana', 'Fecha_Ventana', 'Horario_Ventana', 'NCE_Ventana', 'Entrega_Doc_Ventana', 'Requerimientos_Ventana','Estado_Ventana','Emails_Ventana']

    # Read the DB to get the SheetId based on PID
    try:
        db = get_db()
        data = db.execute('select projectName, projectId, sheetId from projects where projectId = ?',[PID]).fetchall()
    except sqlite3.Error as e:
        error = "Could't complete the Query: "+e.args[0]
        return render_template('schedule_home.html')

    # Get the sheet_id from the db query result
    sheet_id = str(data[0][2])

    # Get a list of ActID
    with open("/Users/dacobos/Development/MWManager/logs/"+sheet_id) as f:
        sheet = f.read()
    # Convert the string of the sheet to a sheet in JSON
    sheet = json.loads(sheet)

    # Get the list of columnId
    columnsId = {}

    for column in sheet["columns"]:
        columnsId[column["title"]]= column["id"]


    for row in sheet["rows"]:
        # Check if the Cell does have the VM code else continue with the next row
        try:
            if PID+"-VM" in row["cells"][0]["value"]:
                pass
            else:
                continue
        except:
            continue
        columnId = []
        # If the cell does have the VM code will execute this
        for cell in row["cells"]:
            for key in keys:
                if cell["columnId"] == columnsId[key]:
            		columnId.append(cell["columnId"])


    # Trim the list to remove the ActID because this value should not change
    columnId= columnId[1:10]

    # Get the row_id
    for row in sheet["rows"]:
        try:
            if row["cells"][0]["value"] == ActID:
                row_id = row["id"]
        except:
            continue

    # Get the new values to update the entry
    value = [request.form["Descripcion"], request.form["Nombre_Ventana"], request.form["Fecha_Ventana"],request.form["Horario_Ventana"],request.form["NCE_Ventana"], request.form["Entrega_Doc_Ventana"], request.form["Requerimientos_Ventana"], request.form["Estado_Ventana"], request.form["Emails_Ventana"]]

    # sheet_id = '85520752633732'
    # row_id = '7381189202864004'
    # columnId = ['5617640476567428','3365840662882180','7869440290252676','7114333500008324','2239940756039556','6743540383410052','4491740569724804','8995340197095300','736437256644484']



    response = updateRow(sheet_id=sheet_id, row_id=row_id, columnId=columnId, value = value, token = os.environ['SMARTSHEET_ACCESS_TOKEN'])
    print response.text
    #Update the local file of query
    sheet = getSheet2(sheet_id = sheet_id, token=app.config['TOKEN'])
    sheet = sheet.encode('utf-8')
    with open("/Users/dacobos/Development/MWManager/logs/"+sheet_id,"w") as f:
        f.write(sheet)
    flash('Notification: Entry updated successfully')
    return redirect(url_for('edit', ActID = ActID, sheetId=sheet_id))
