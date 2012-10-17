Attribute VB_Name = "CreateProgramList"
Option Compare Database


Function NewProgramList(cv_id As Integer)
    Dim db As Database
    Set db = CurrentDb()
    Dim arrProgramIds() As Integer
        
    arrProgramIds = ReadRegistry.GetAllPrograms()
    For Each i In arrProgramIds
        If i <> 0 Then
            'If entry does not exist
            If db.OpenRecordset("SELECT * FROM programmliste WHERE computerversion_id = " & cv_id & " AND programm_id = " & i).EOF <> False Then
                db.Execute ("INSERT INTO programmliste (computerversion_id, programm_id) VALUES (" & cv_id & "," & i & ")")
            End If
        End If
    Next
        
    db.Close
    
End Function
