Option Explicit
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset

Sub CreateDb()
Set con = New ADODB.Connection
con.Open "DRIVER={MySql ODBC 8.0 Unicode Driver};Server=localhost;Database=practice1;User=root;Password=tq3wt2117;Option=3;"
Set rs = New ADODB.Recordset
Dim rslt As String
Dim insertString As String
insertString = "INSERT INTO EMP (ID, fName, lName, Manager, HireDate, Position) VALUES"

Sheets("EmpData").Select

'Find data region
Dim DataEnd As Integer
Range("A2").End(xlDown).Select
DataEnd = Selection.Row

Dim I As Integer
Dim X As Integer


For I = 2 To DataEnd
    insertString = insertString & "("
    For X = 1 To 6
        Dim currentValue As Variant
        Dim strSnip As String
        currentValue = Cells(I, X).Value
        strSnip = Chr(34) & currentValue & Chr(34)
        If X <> 6 Then strSnip = strSnip & ","
        insertString = insertString & strSnip
    Next X
    insertString = insertString & ")"
    If I <> DataEnd Then insertString = insertString & ","
Next I

MsgBox insertString


con.Execute ("CREATE TABLE Emp (ID INT NOT NULL AUTO_INCREMENT,fName VARCHAR(100) NOT NULL,lName VARCHAR(100) NOT NULL,Manager INT NOT NULL,HireDate DATE NOT NULL, Position Varchar(100),Primary Key(ID));")
con.Execute (insertString)
Set rs = con.Execute("Select * From Emp")
MsgBox rs.GetString

con.Close
End Sub

