Attribute VB_Name = "ReadRegistry"
Option Compare Database

    Const HKLM = &H80000002
    Function GetPrograms()
        Dim temp As Object
        Dim strComputer As String
        Dim rPath As String
        Dim arrSubKeys()
        Dim strAsk
        
        'Set StdRegProv class
        strComputer = "."
        Set temp = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
            strComputer & "\root\default:StdRegProv")
            
        'Get Subkeys from rPath
        rPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
        temp.EnumKey HKLM, rPath, arrSubKeys
        
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
                Debug.Print ("Title: " & strTitle)
                InsertIntoPrograms (strTitle)
            End If
        Next
    End Function
    
    Sub InsertIntoPrograms(strTitle As String)
        Dim db As Database
        Dim check As String
        Dim Sql As DAO.Recordset
        
        Set db = CurrentDb()
        Set Sql = db.OpenRecordset("SELECT * FROM programm WHERE titel = '" & strTitle & "'")
        If Sql.EOF <> False Then
            db.Execute "INSERT INTO programm (titel) VALUES ('" & strTitle & "')"
        End If
        Set cn = Nothing
    End Sub
