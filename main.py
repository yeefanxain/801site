import flask
from flask import Flask
from flask import render_template
import pygsheets
import random


#google sheets
gc = pygsheets.authorize(service_file='credentials.json')
sht = gc.open_by_url(
'https://docs.google.com/spreadsheets/d/1RQTrnU8BFLfkeC-rMwHYO-XCv1SBxJ-xDh6L5sOXvZQ/'
)

wks = sht.worksheet_by_title("填表區")
wks2 = sht.worksheet_by_title("需留校名單")
allpeople = wks.cell('I29:J29')
print(allpeople.value)

#web server
app = Flask(__name__)

@app.route("/")
def index():
    allpeople = wks.cell('I29:J29')
    allpeople=allpeople.value
    allnumber= wks.cell('L29:M29')
    allnumber=allnumber.value
    avrnumber=int(allnumber)/25
    avrnumber=round(avrnumber, 1)
    totaltest= wks.cell('B31')
    totaltest=totaltest.value
    
    return render_template("site.html", allpeople=allpeople, allnumber=allnumber, avrnumber=avrnumber,totaltest=totaltest)

@app.route("/list")
def list():
    name="通過"
    d1= wks2.cell('B2')
    d3= wks2.cell('B3')
    d4= wks2.cell('B4')
    d5= wks2.cell('B5')
    d6= wks2.cell('B6')
    d7= wks2.cell('B7')
    d9= wks2.cell('B8')
    d10= wks2.cell('B9')
    d11= wks2.cell('B10')
    d12= wks2.cell('B11')
    d13= wks2.cell('B12')
    d14= wks2.cell('B13')
    d21= wks2.cell('B14')
    d22= wks2.cell('B15')
    d23= wks2.cell('B16')
    d24= wks2.cell('B17')
    d25= wks2.cell('B18')
    d26= wks2.cell('B19')
    d27= wks2.cell('B20')
    d28= wks2.cell('B21')
    d29= wks2.cell('B22')
    d30= wks2.cell('B23')
    d31= wks2.cell('B24')
    d32= wks2.cell('B25')
    d33= wks2.cell('B26')

    return render_template("list.html",name=name,d1=d1.value,d3=d3.value,d4=d4.value,d5=d5.value,d6=d6.value,d7=d7.value,d9=d9.value,d10=d10.value,d11=d11.value,d12=d12.value,d13=d13.value,d14=d14.value,d21=d21.value,d22=d22.value,d23=d23.value,d24=d24.value,d25=d25.value,d26=d26.value,d27=d27.value,d28=d28.value,d29=d29.value,d30=d30.value,d31=d31.value,d32=d32.value,d33=d33.value)



      

@app.route("/list2")
def list2():
    name="留校"
    ln1= wks2.cell('C2')
    ln3= wks2.cell('C3')
    ln4= wks2.cell('C4')
    ln5= wks2.cell('C5')
    ln6= wks2.cell('C6')
    ln7= wks2.cell('C7')
    ln9= wks2.cell('C8')
    ln10= wks2.cell('C9')
    ln11= wks2.cell('C10')
    ln12= wks2.cell('C11')
    ln13= wks2.cell('C12')
    ln14= wks2.cell('C13')
    ln21= wks2.cell('C14')
    ln22= wks2.cell('C15')
    ln23= wks2.cell('C16')
    ln24= wks2.cell('C17')
    ln25= wks2.cell('C18')
    ln26= wks2.cell('C19')
    ln27= wks2.cell('C20')
    ln28= wks2.cell('C21')
    ln29= wks2.cell('C22')
    ln30= wks2.cell('C23')
    ln31= wks2.cell('C24')
    ln32= wks2.cell('C25')      
    ln33= wks2.cell('C26')
    return render_template("list.html",name=name,d1=ln1.value,d3=ln3.value,d4=ln4.value,d5=ln5.value,d6=ln6.value,d7=ln7.value,d9=ln9.value,d10=ln10.value,d11=ln11.value,d12=ln12.value,d13=ln13.value,d14=ln14.value,d21=ln21.value,d22=ln22.value,d23=ln23.value,d24=ln24.value,d25=ln25.value,d26=ln26.value,d27=ln27.value,d28=ln28.value,d29=ln29.value,d30=ln30.value,d31=ln31.value,d32=ln32.value,d33=ln33.value)

@app.route("/apk")
def apk():
  return render_template("apk.html")


@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html'), 404

@app.errorhandler(500)
def server_error(e):
    return render_template('500.html'), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=random.randint(2000,9000))
    

