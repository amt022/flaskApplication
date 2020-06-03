from flask import Flask
from flask import request
import requests
import datetime
import os
from flask import send_from_directory
from pandas import ExcelWriter,ExcelFile
import xlsxwriter
app = Flask(__name__)

def get_segregated_values(r):
    dictionary=dict()
    for i in range(len(r)):
        date = r[i]['DateTime'][:10]
        length = r[i]['Length']
        weight = r[i]['Weight']
        quantity = r[i]['Quantity']
        if(date in dictionary):
            dictionary[date].append((r[i]['DateTime'],length,weight,quantity))
        else:
            dictionary[date]=[]
            dictionary[date].append((r[i]['DateTime'],length,weight,quantity))
    return dictionary


@app.route("/")
def home():
    return "Invalid Endpoint!!"

@app.route("/total",methods=['GET'])
def total():
    date = request.args.get('day')
    date_obj = datetime.datetime.strptime(date,'%d-%m-%Y')
    #print(date_obj.date())
    r=requests.get('https://assignment-machstatz.herokuapp.com/excel').json()
    total_weight=0
    total_length=0
    total_quantity=0
    for i in range(len(r)):
        date=r[i]['DateTime'][:10]
        json_date_obj = datetime.datetime.strptime(date,'%Y-%m-%d')
        #print(json_date_obj.date())
        if json_date_obj.day==date_obj.day and json_date_obj.month==date_obj.month and json_date_obj.year==date_obj.year:
            total_weight=total_weight+r[i]['Weight']
            total_length=total_length+r[i]['Length']
            total_quantity=total_quantity+r[i]['Quantity']
    retVal={"totalWeight":round(total_weight,2),"totalLength":round(total_length,2),"totalQuantity":round(total_quantity,2)}
    return retVal
    
@app.route("/excelreport",methods=['GET'])
def excelreport():
    print("Generating report")
    workbook = xlsxwriter.Workbook('Report.xlsx')
    row=1
    #col=0
    r=requests.get('https://assignment-machstatz.herokuapp.com/excel').json()
    dictionary = get_segregated_values(r)
    for key in dictionary:
        worksheet = workbook.add_worksheet()
        worksheet.write('A1','DateTime')
        worksheet.write('B1','Length')
        worksheet.write('C1','Weight')
        worksheet.write('D1','Quantity')
        values = dictionary[key]
        row=1
        for val in values:
            worksheet.write(row,0,val[0])
            worksheet.write(row,1,val[1])
            worksheet.write(row,2,val[2])
            worksheet.write(row,3,val[3])
            row=row+1

    workbook.close()
    print(app.instance_path)
    return send_from_directory(os.getcwd(),'Report.xlsx')

if __name__ == "__main__":
    app.run(debug=True)