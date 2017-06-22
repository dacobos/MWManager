################################  MODULE  INFO  ################################
# Author: David  Cobos
# Cisco Systems Solutions Integrations Architect
# Mail: cdcobos1999@gmail.com  / dacobos@cisco.com
##################################  IMPORTS   ##################################

# Following task shall be done
# TODO: Install python libraries: pyOpenSSL, requests, json, smtplib, flask, threading, subprocess, shutil, Pysocks
# TODO: Define environment variable SMARTSHEET_ACCESS_TOKEN as the credentiasl to access smartsheet
# TODO: Define environment variable GMAIL_ACCOUNT as the source of the Emails_Ventana
# TODO: Define environment variable GMAIL_PWD as the password of the GMAIL_ACCOUNT


import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import socks

# Uncomment this in order to use the APP from the Server in SJU
# socks.setdefaultproxy(socks.SOCKS5, 'proxy-rtp-1', 1080)
# socks.wrapmodule(smtplib)


def mwMIME(inp, recipients, subject):
    origin = os.environ['GMAIL_ACCOUNT']
    origin_pwd = os.environ['GMAIL_PWD']

    # Create message container - the correct MIME type is multipart/alternative.
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = origin
    msg['To'] = ", ".join(recipients)

    # Create the body of the message (a plain-text and an HTML version).
    # text = "Hi!\nHow are you?\nHere is the link you wanted:\nhttps://www.python.org"
    html_header = """\
    <html>
      <head>
    <style>
    table {
    	font-family:Arial, Helvetica, sans-serif;
    	color:#666;
    	font-size:12px;
    	text-shadow: 1px 1px 0px #fff;
    	background:#eaebec;
    	border:#ccc 1px solid;

    	-moz-border-radius:3px;
    	-webkit-border-radius:3px;
    	border-radius:3px;

    	-moz-box-shadow: 0 1px 2px #d1d1d1;
    	-webkit-box-shadow: 0 1px 2px #d1d1d1;
    	box-shadow: 0 1px 2px #d1d1d1;
    }
    table th {
    	padding:21px 25px 22px 25px;
    	border-top:1px solid #fafafa;
    	border-bottom:1px solid #e0e0e0;

    	background: #ededed;
    	background: -webkit-gradient(linear, left top, left bottom, from(#ededed), to(#ebebeb));
    	background: -moz-linear-gradient(top,  #ededed,  #ebebeb);
    }
    table th:first-child {
    	text-align: left;
    	padding-left:20px;
    }
    table tr:first-child th:first-child {
    	-moz-border-radius-topleft:3px;
    	-webkit-border-top-left-radius:3px;
    	border-top-left-radius:3px;
    }
    table tr:first-child th:last-child {
    	-moz-border-radius-topright:3px;
    	-webkit-border-top-right-radius:3px;
    	border-top-right-radius:3px;
    }
    table tr {
    	text-align: center;
    	padding-left:20px;
    }
    table td:first-child {
    	text-align: left;
    	padding-left:20px;
    	border-left: 0;
    }
    table td {
    	padding:18px;
    	border-top: 1px solid #ffffff;
    	border-bottom:1px solid #e0e0e0;
    	border-left: 1px solid #e0e0e0;

    	background: #fafafa;
    	background: -webkit-gradient(linear, left top, left bottom, from(#fbfbfb), to(#fafafa));
    	background: -moz-linear-gradient(top,  #fbfbfb,  #fafafa);
    }
    table tr.even td {
    	background: #f6f6f6;
    	background: -webkit-gradient(linear, left top, left bottom, from(#f8f8f8), to(#f6f6f6));
    	background: -moz-linear-gradient(top,  #f8f8f8,  #f6f6f6);
    }
    table tr:last-child td {
    	border-bottom:0;
    }
    table tr:last-child td:first-child {
    	-moz-border-radius-bottomleft:3px;
    	-webkit-border-bottom-left-radius:3px;
    	border-bottom-left-radius:3px;
    }
    table tr:last-child td:last-child {
    	-moz-border-radius-bottomright:3px;
    	-webkit-border-bottom-right-radius:3px;
    	border-bottom-right-radius:3px;
    }

    </style>
      </head>
    """
    html_body = """\
      <body>
    <div id="title">
    <h1>{}</h1>
    </div>
    <br>
    <table style="display: table;" id="query">
      <tbody><tr>
        <td>Proyecto</td>
        <td>{}</td>
      </tr>
      <tr>
        <td>PID</td>
        <td>{}</td>
      </tr>
      <tr>
        <td>ACT ID</td>
        <td>{}</td>
      </tr>

    </tbody></table>
    <br>
    <table style="display: table;" id="query">
      <tbody><tr>
        <th>Nombre de la Ventana
        </th>
        <th>Descripcion Actividades
        </th>
        <th>Fecha Actividad
        </th>
        <th>Horario
        </th>
        <th>NCE
        </th>
        <th>Entrega Documentos
        </th>
        <th>Requerimientos
        </th>
        <th>Estado
        </th>
      </tr>

      <tr>

        <td>{}
        </td>
        <td>{}
        </td>
        <td>{}
        </td>
        <td>{}
        </td>
        <td>{}
        </td>
        <td>{}
        </td>
        <td>{}
        </td>
        <td>{}
        </td>
      </tr>
    </tbody></table>

      </body>
    </html>

    """

    html_body_formated = html_body.format(subject,inp[0],inp[1],inp[2],inp[3],inp[4],inp[5],inp[6],inp[7],inp[8],inp[9],inp[10])
    html_message = html_header + html_body_formated
    part1 = MIMEText(html_message, 'html')

    # Attach parts into message container.
    # According to RFC 2046, the last part of a multipart message, in this case
    # the HTML message, is best and preferred.
    # msg.attach(part1)
    msg.attach(part1)

    # Send the message via local SMTP server.
    try:
        s = smtplib.SMTP('smtp.gmail.com:587')
        s.ehlo()
        s.starttls()
        s.login(origin, origin_pwd)
        # sendmail function takes 3 arguments: sender's address, recipient's address
        # and message to send - here it is sent as one string.
        s.sendmail(origin, recipients, msg.as_string())
        s.quit()
        error = None
    except:
        error = "There was a problem, couln't send email"
        print error

    return error
# 17:37
