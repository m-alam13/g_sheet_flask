from oauth2client.service_account import ServiceAccountCredentials
import gspread
import json
from flask import Flask, render_template,Response,url_for,request,redirect
  
scopes = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']

credcentials = ServiceAccountCredentials.from_json_keyfile_name("auto-pump.json", scopes) #access the json key you downloaded earlier 
client = gspread.authorize(credcentials)
sheet = client.open("alam").worksheet('Sheet1')


app = Flask(__name__)
@app.route('/home', methods=['GET', 'POST'])
def home():
    message = request.args.get('message')
    msg = request.args.get('msg')
    cell_n = request.args.get('cell_n')
    # if msg==None:
    #     msg=""
    if message!=None:
        return render_template("home.html",message=message)
    
    if cell_n!=None:
        return render_template("home.html",msg=msg,cell=cell_n)
    return render_template("home.html")

@app.route('/update', methods=['GET', 'POST'])
def update():
    try:
        if request.method=="POST":
            cell_n = request.form['cell'] #get data from form
            value_n = request.form['value']
            sheet.update_acell(cell_n, value_n)
            return redirect(url_for('home',message="Value update successfully"))
    except:
        return redirect(url_for('home',message="Somthing going wrong!!!"))
@app.route('/delete' ,methods=['GET','POST'])
def delete():
    try:
        if request.method=="POST":
            cell_n= request.form['cell']
            value_n=''
            sheet.update_acell(cell_n, value_n)
            return redirect(url_for('home',message="Value delete successfully"))
    except:
        return redirect(url_for('home',message="Somthing going wrong!!!"))

@app.route('/append',methods=["GET","POST"])
def append():
    try:
        if request.method=="POST":
            value = request.form['value']
            column_name = request.form['col']
            # print(column_name)
            # Find the last row with data in the specified column
            column_values = sheet.col_values(ord(column_name.lower()) - 96)  # Convert column letter to index
            last_row = len(column_values) + 1
            # print(last_row)
            # Append the value to the specific cell in the Google Sheet
            sheet.update_cell(last_row, ord(column_name.lower()) - 96, value)  # Convert column letter to index
            return redirect(url_for('home',message="Value append successfully"))
    except:
        return redirect(url_for('home',message="Somthing going wrong!!!"))
@app.route('/read',methods=['GET','POST'])
def read():
    try:
        if request.method=="POST":
            cell_n= request.form['cell']
            m=sheet.acell(cell_n).value
            print(m)
            return redirect(url_for('home',msg=str(m),cell_n=cell_n))
    except:
        return redirect(url_for('home',message="Somthing going wrong!!!"))


if __name__ == '__main__':
    app.run(host="0.0.0.0",debug=True)