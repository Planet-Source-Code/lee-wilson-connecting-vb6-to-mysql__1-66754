VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Visual Basic 6 MySql tutorial By Lee Wilson 11/10/06
'Feel free to use this code in your own projects.
'This tutorial was created because it took me over 4 hours to figure this out
'Had to mess around when it should have been copy and paste,
'Hope this helps ,Lee Wilson

'Any questions comments or concerns, or any bugs found with this mail me.
'lee88eng@hotmail.com

Private Sub Form_Load()
'Ok READ THIS, if you want to get this to run. You need to do two things
' 1) download the obcd driver as found here.

'   http://dev.mysql.com/downloads/connector/odbc/3.51.html

' 2) So go to Project >> References...
'      Dialog box should open up
'      Find Microsoft ActiveX Data Objects 2.7 Library
'      Tick it and go ok.
 
'You should be able to copy and paste the code directly into your project
'And its all pretty much explained.





'These two are required connection and recordset allow
'you to connect to the db and manipulate it
Dim CNN As Connection
Dim rs As Recordset


'These are the variables that determine where your database is
'how to connect to it, and what to do once its connected
'you can hard code these in if you want


Dim ssql As String           'sql to be used to manipulate datebase
Dim serverip As String       'ip of server or "localhost" if its on your machine
Dim port As String           'port used my MySql 3306 by default
Dim user As String           'username
Dim datebasename As String   'database name
Dim password As String       'password
Dim table_name As String     'table to manipulate



'Setting variables
serverip = "localhost"
port = "3306"
datebasename = "test"
user = "root"
password = "password"

table_name = "test_table"

'Creating a new connection, and defineing the variables

Set CNN = New Connection
CNN.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                      & "SERVER=" & serverip & ";" _
                      & " DATABASE=" & datebasename & ";" _
                      & " PORT=" & port & ";" _
                      & "UID=" & user & ";PWD=" & password & "; OPTION=3"

CNN.Open       'Opening the database

Set rs = New Recordset  'Rs allows you to move through the records with simple commands such as rs.MoveNext
    
    'This example is inserting data, this data is hardcoded but most of the time variables would be used
    'In fact if you look at the code its been made for variables and not hard coding
    ssql = "INSERT INTO " & table_name & " VALUES ('" & "first_data" & "', '" & "second_data" & "')"
   ' Hoping you know sql, should make it all nice and simple. Only thing to mention is table_name variable
   'This ssql only uploads to bits of data into this database, data piece one, "first_date" and "second_data"
   'by altering the sql you can insert more less, or retrieve infomation, go to w3school for more help on sql
   
    
    rs.Open ssql, CNN   'activates the sql
Set rs = Nothing        'To reduce bugs rs and cnn are set to nothing.
    

    
CNN.Close
Set CNN = Nothing
MsgBox ("upload complete") 'Thats it, simple eh


End Sub
