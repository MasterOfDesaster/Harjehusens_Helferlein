Attribute VB_Name = "ReadComputer"
Option Compare Database

Function GetComputerId() As Integer
    Dim title As String
    Dim raum_id As String
    Dim os As String
    Dim ip As String
    Dim mac As String
    
    title = Environ("COMPUTERNAME")
    raum_id = "0"
    os = "Windows"
    ip = "127.0.0.1"
    mac = "0"
    
    Dim db As Database
    Dim Rec As DAO.Recordset
    Set db = CurrentDb()
    
    'Check for duplicate, insert new if necessary
    Set Rec = db.OpenRecordset("SELECT * FROM computer")
    If Rec.EOF <> False Then
        db.Execute ("INSERT INTO computer (titel, raum_id, os, ip, mac) VALUES ('" & title & "','" & raum_id & "','" & os & "','" & ip & "','" & mac & "')")
        Set Rec = db.OpenRecordset("SELECT * FROM computer")
        Debug.Print ("Inserted " & title)
    End If
    
    'Get programm id for further use
    GetComputerId = Rec![id]
    db.Close
    
End Function

Function NewDataset()
     CreateProgramList.NewProgramList (CreateComputerVersion.NewComputerVersion(GetComputerId()))
End Function
