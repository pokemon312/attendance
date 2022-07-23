import openpyxl
import os
import random
import smtplib
from datetime import datetime, timedelta
from random import randrange
from uuid import uuid4

import pymongo as mongo
from bson.objectid import ObjectId
from dotenv import load_dotenv
from flask import Flask, jsonify, request
from flask_cors import CORS
app = Flask(__name__)
CORS(app, resources={r"*": {"origins": "*"}})
app.config['CORS_HEADERS'] = 'Content-Type, Access-Control-Allow-Origin'
def op(year,specialization,sec,timing):
    xfile = openpyxl.load_workbook('C:/Users/ramuk/OneDrive/Desktop/attendance_name_update/attendance/test.xlsx')
    sheet = xfile.get_sheet_by_name('Sheet1')
    sheet['B1'] = 'Kannan Ramu'
    sheet['E1']='22.07.22'
    lis=[121212,233232,232323,232323,23232,232323,232323,233322,121212,233232,232323,232323,23232,232323,232323,233322]
    for i in range(3,len(lis)+3):
        sheet['A{}'.format(i)]=year
        sheet['B{}'.format(i)]=specialization
        sheet['C{}'.format(i)]=sec
        sheet['D{}'.format(i)]=timing
        sheet['E{}'.format(i)]=lis[i-3]
    xfile.save('attendance/text2.xlsx')
    return jsonify({
        'status':True,
        "done":"kjnsjna"
    })


@app.route("/", methods=["GET"])
def main():
	return "ok"
@app.route("/workpannu", methods=["POST"])
def workpannu():
	if request.method == "POST":
		year = request.form.get("year")
		sp = request.form.get("sp")
		otp_verification_status = op(year, sp,'C1','20pm')
		return otp_verification_status    

# year='III/V'
# specialization='CSE'
# sec='C3'
# timing="1pm-2pm"
if __name__ == "__main__":
	app.run(debug=True)