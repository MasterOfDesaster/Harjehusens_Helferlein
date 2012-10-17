Attribute VB_Name = "CreateComputerVersion"
Option Compare Database


Function NewComputerVersion(c_id As Integer) As Integer
    Dim db As Database
    Set db = CurrentDb()
    db.Execute ("INSERT INTO computerversion (computer_id, erfassdatum) VALUES (" & c_id & ",'" & Now & "')")
    NewComputerVersion = db.OpenRecordset("SELECT * FROM computerversion WHERE computer_id = " & c_id & " ORDER BY erfassdatum DESC")![id]
    db.Close
End Function
