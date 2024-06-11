from flask import Flask,render_template,request,flash, make_response,redirect,url_for,Response
import re
import pypyodbc
import io
import xlwt

app=Flask(__name__)
app.secret_key="gfhdsfjfsd"
DRIVER_NAME="{SQL SERVER}"
SERVER_NAME="NIKESH"
DATABASE_NAME="schooldb"

connection_string=f"""
DRIVER={DRIVER_NAME};
SERVER={SERVER_NAME};
DATABASE={DATABASE_NAME};
Trust_Connection=yes;
"""

connection=pypyodbc.connect(connection_string)
cursor=connection.cursor()
@app.route('/')
def home():
    return render_template("home.html")

@app.route('/register')
def register():
    return render_template("register.html")

@app.route('/about')
def about():
    return render_template("about.html")

@app.route('/admission')
def admission():
    return render_template("admission.html")


# @app.route('/addemployee_fromdashbord')
# def adddash():
#     return render_template("add_student_dashbord.html")

@app.route('/dash')
def dashbord():
    return render_template("dashbord.html")

@app.route('/login')
def login():
    return render_template("login.html")

@app.route("/form_data_store" ,methods=['POST'])
def submitdata():
    if request.method == "POST":
        FIRST_NAME = request.form.get("FN")
        LAST_NAME = request.form.get("LN")
        EMAIL = request.form.get("ID")
        GENDER = request.form.get("GENDER")
        AGE = request.form.get("AGE")
        STANDARD = request.form.get("STD")
        TEACHER = request.form.get("CT")
        PASSWORD = request.form.get("SP")
        check = request.form.get("ck")
        ACTIVATE = "ACTIVE"
        NONACTIVE = "INACTIVE" 
        if check == 'on':
                    cursor.execute('execute spSaveData ?,?,?,?,?,?,?,?,?', (FIRST_NAME, LAST_NAME, EMAIL, GENDER, AGE, STANDARD, TEACHER, PASSWORD, ACTIVATE))
                    cursor.commit()
                    flash("SUCCESSFULLY ADDED!")
               
        else:
            cursor.execute('EXEC spSaveData ?,?,?,?,?, ?,?,?,?',
                           (FIRST_NAME, LAST_NAME, EMAIL, GENDER, AGE, STANDARD, TEACHER, PASSWORD, ACTIVATE))
            cursor.commit()
            flash("SUCCESSFULLY ADDED!")
            return render_template("dashbord.html")
    # q = "SELECT * FROM tblstudentdata"
    cursor.execute("execute spGetAll")
    display = cursor.fetchall()
    return render_template("dashbord.html", display=display)
  
@app.route('/loginvalidation',methods=['GET','POST'])
def validate():
    email=request.form.get("LOGIN_EMAILID")
    password=request.form.get("LOGIN_PASSWORD")
    # query='select COUNT(*) from tblstudentdata WHERE EMAIL=? and STUDENT_PASSWORD=? '
    cursor.execute('spValidate ?,?',(email,password))
    result=cursor.fetchone()[0]
    if result>0:
        cursor.execute('execute spGetAll')
        display=cursor.fetchall()
        flash("YOU ARE SUCCESSFULLY LOGGED IN")
        return render_template("dashbord.html",display=display)
    else:
        flash("PLEASE TRY AGAIN INVALID USERNAME OR PASSWORD")
        return render_template("login.html")
    

@app.route("/edit")
def add():
    return render_template("edit.html")

@app.route("/edituser/<int:id>", methods=['GET','POST'])
def update(id):
    if request.method=='POST':
         FIRST_NAME=request.form.get("field1")
         LAST_NAME=request.form.get("field2")
         EMAIL=request.form.get("field3")
         GENDER=request.form.get("field4")
         AGE=request.form.get("field5")
         STANDARD=request.form.get("field6")
         TEACHER=request.form.get("field7")
         PASSWORD=request.form.get("field8")
         SCENARIO=request.form.get("field9")
         cursor.execute("execute spUpdateuser ?,?,?,?,?,?,?,?,?,?",[FIRST_NAME,LAST_NAME,EMAIL,GENDER,AGE,STANDARD,TEACHER,PASSWORD,SCENARIO,id])
         cursor.commit()
        # cursor.
        #  q="SELECT * FROM tblstudentdata"
         cursor.execute("execute spGetAll")
         display=cursor.fetchall()
         return render_template("dashbord.html",display=display)
    result=cursor.execute('spGetById ?',[id])
    result=cursor.fetchone()
    return render_template("edit.html",data=result)
    
@app.route('/deleteuser/<int:id>',methods=["GET","POST"])
def drop(id):
    # query="delete from tblstudentdata WHERE ID=?"
    cursor.execute('execute spDel ?',[id])
    cursor.commit()
    # q="SELECT * FROM tblstudentdata"
    cursor.execute("execute spGetAll")
    display=cursor.fetchall()
    return render_template("dashbord.html",display=display)


#THIS WILL ADD THE DATA FROM Login page register button
@app.route('/addempo', methods=['POST'])
def adduser():
 if request.method == "POST":
    FIRST_NAME = request.form.get("FN")
    LAST_NAME = request.form.get("LN")
    EMAIL = request.form.get("ID")
    GENDER = request.form.get("GENDER")
    AGE = request.form.get("AGE")
    STANDARD = request.form.get("STD")
    TEACHER = request.form.get("CT")
    PASSWORD = request.form.get("SP")
    check = request.form.get("ck")
    ACTIVATE = "ACTIVE"
    NONACTIVE = "INACTIVE"
    
    if check == 'on':
        cursor.execute('EXEC spSaveData ?,?,?,?,?, ?,?,?,?',
                       (FIRST_NAME, LAST_NAME, EMAIL, GENDER, AGE, STANDARD, TEACHER, PASSWORD, ACTIVATE))
        cursor.commit()
        flash("SUCCESSFULLY ADDED!")
    else:
        cursor.execute('EXEC spSaveData ?,?,?,?,?, ?,?,?,?',
                       (FIRST_NAME, LAST_NAME, EMAIL, GENDER, AGE, STANDARD, TEACHER, PASSWORD, NONACTIVE))
        cursor.commit()
        flash("SUCCESSFULLY ADDED!")
        return render_template("login.html")
 return render_template("login.html")
 

      
    
@app.route('/vieweuser/<int:id>',methods=["GET","POST"])
def view(id):
    cursor.execute('execute spGetById ?',[id])
    user=cursor.fetchone()
    return render_template("view.html",user=user)


@app.route('/search',methods=['GET','POST'])
def searchs():
    if request.method=="POST":
        requests=request.form.get("search")
        cursor.execute('execute spGetData ?,?,?,?,?,?,?,?',[requests,requests,requests,requests,requests,requests,requests,requests])
        display=cursor.fetchall()
        cursor.commit()
        return render_template("search.html",display=display)
    else:
        flash("NO SUCH STUDENT EXIST")
        cursor.execute('execute spGetAll')
        display=cursor.fetchall()
        return render_template("dashbord.html", display=display)

@app.route("/back-dashbord")
def dash_bord():
    cursor.execute('execute spGetAll')
    display=cursor.fetchall()
    return render_template("dashbord.html", display=display)

@app.route('/export')
def export_excel():
    cursor.execute("execute spGetAll")
    result=cursor.fetchall()
    output=io.BytesIO()
    workbook=xlwt.Workbook()
    sh=workbook.add_sheet('STUDENTDATA')
    sh.write(0,0,'ID')
    sh.write(0,1,'FIRST_NAME')
    sh.write(0,2,'LAST_NAME')
    sh.write(0,3,'EMAIL')
    sh.write(0,4,'GENDER')
    sh.write(0,5,'AGE')
    sh.write(0,6,'STANDARD')
    sh.write(0,7,'CLASS_TEACHER')
    sh.write(0,8,'STUDENT_PASSWORD')
    sh.write(0,9,'SCENARIO')

    idx=0
    for row in result:
        sh.write(idx+1,0,str(row[0]))
        sh.write(idx+1,1,row[1])
        sh.write(idx+1,2,row[2])
        sh.write(idx+1,3,row[3])
        sh.write(idx+1,4,row[4])
        sh.write(idx+1,5,row[5])
        sh.write(idx+1,6,row[6])
        sh.write(idx+1,7,row[7])
        sh.write(idx+1,8,row[8]) 
        sh.write(idx+1,9,row[9]) 
        idx +=1
    workbook.save(output)
    output.seek(0)
    return Response(output,mimetype="application/ms-excel",headers={"content-Disposition":"attachment;filename=employee.xlm"}) 


#THIS WILL ADD THE DATA FROM DASHBORD
@app.route('/from_dash', methods=['POST'])
def from_dash():
    if request.method == "POST":
        FIRST_NAME = request.form.get("FN")
        LAST_NAME = request.form.get("LN")
        EMAIL = request.form.get("ID")
        GENDER = request.form.get("GENDER")
        AGE = request.form.get("AGE")
        STANDARD = request.form.get("STD")
        TEACHER = request.form.get("CT")
        PASSWORD = request.form.get("SP")
        check = request.form.get("ck")
        ACTIVATE = "ACTIVE"
        NONACTIVE = "INACTIVE"
        if check == 'on':
                    cursor.execute('execute spSaveData ?, ?, ?, ?, ?, ?, ?, ?, ?', (FIRST_NAME, LAST_NAME, EMAIL, GENDER, AGE, STANDARD, TEACHER, PASSWORD, ACTIVATE))
                    cursor.commit()
                    flash("SUCCESSFULLY ADDED!")
               
        else:
            cursor.execute('execute spSaveData ?,?,?,?,?,?,?,?,?',
                           (FIRST_NAME, LAST_NAME, EMAIL, GENDER, AGE, STANDARD, TEACHER, PASSWORD, NONACTIVE))
            cursor.commit()
            flash("SUCCESSFULLY ADDED!")
            cursor.execute('execute spGetAll')
            display=cursor.fetchall()
            return render_template("dashbord.html", display=display)

    cursor.execute('execute spGetAll')
    display=cursor.fetchall()
    return render_template("dashbord.html", display=display)

@app.route('/logout')
def logout():
    flash("YOU HAVE LOGED OUT")
    return render_template('login.html')

        
    

if __name__ == "__main__":
    app.run(debug=True)
