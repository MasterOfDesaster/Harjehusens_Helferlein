Attribute VB_Name = "ReadRegistry"
Option Compare Database

    Const HKLM = &H80000002
    Function GetAllPrograms() As Integer()
        Dim temp As Object
        Dim strComputer As String
        Dim rPath As String
        Dim arrSubKeys()
        Dim strAsk
        Dim arrProgramIds() As Integer
        Dim index As Integer
        index = 0
        
        'Set StdRegProv class
        strComputer = "."
        Set temp = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
            strComputer & "\root\default:StdRegProv")
            
        'Get Subkeys from rPath
        rPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
        temp.EnumKey HKLM, rPath, arrSubKeys
        ReDim arrProgramIds(0 To UBound(arrSubKeys))
        'Process results
        For Each strAsk In arrSubKeys
        
            'Get DisplayName
            Dim strTitle
            temp.GetStringValue HKLM, rPath & strAsk, "DisplayName", strTitle
            
            'Get DisplayVersion
            Dim strVersion
            temp.GetStringValue HKLM, rPath & strAsk, "DisplayVersion", strVersion
            
            'Filter out all entries with empty Name
            If (strTitle <> "") Then
                arrProgramIds(index) = GetProgramId(strTitle, strVersion)
                index = index + 1
            End If
        Next
        GetAllPrograms = arrProgramIds
    End Function
    
    Private Function GetProgramId(strTitle, strVersion) As Integer
    
        Dim db As Database
        Dim Rec As DAO.Recordset
        
        'Escape characters
        strTitle = Replace(strTitle, "'", "''")
        
        Set db = CurrentDb()
        'Check for duplicate, insert new if necessary
        Set Rec = db.OpenRecordset("SELECT * FROM programm WHERE titel = '" & strTitle & "' AND version = '" & strVersion & "'")
        If Rec.EOF <> False Then
            db.Execute "INSERT INTO programm (titel, version) VALUES ('" & strTitle & "', '" & strVersion & "')"
            Set Rec = db.OpenRecordset("SELECT * FROM programm WHERE titel = '" & strTitle & "' AND version = '" & strVersion & "'")
            Debug.Print ("Inserted " & strTitle)
        End If
        
        'Get programm id for further use
        GetProgramId = Rec![id]
        db.Close
    End Function
