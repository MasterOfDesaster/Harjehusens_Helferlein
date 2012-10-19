Attribute VB_Name = "ReadComputer"
Option Compare Database

Function GetComputerId() As Integer
    Dim title As String
    Dim raum_id As String
    Dim os As String
    Dim ip As String
    Dim mac As String
    
    title = GetComputerName()
    raum_id = "0"
    os = GetRegistryKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
    ip = GetIPAddress()
    mac = GetMACAddress()
    
    
    Dim db As Database
    Dim Rec As DAO.Recordset
    Set db = CurrentDb()
    
    'Check for duplicate, insert new if necessary
    Set Rec = db.OpenRecordset("SELECT * FROM computer WHERE mac = '" & mac & "'")
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


Function GetComputerName() As String
    GetComputerName = Environ("COMPUTERNAME")
End Function

Function GetIPAddress() As String
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMIService.ExecQuery( _
        "SELECT * FROM Win32_NetworkAdapterConfiguration", , 48)
        
    For Each objItem In colItems
        If Not IsNull(objItem.IPAddress) Then
            GetIPAddress = objItem.IPAddress(0)
        End If
    Next
End Function

Function GetMACAddress() As String
    ' get a list of enabled adaptor names and MAC addresses
    ' from msdn.microsoft.com/en-us/library/windows/desktop/aa394217(v=vs.85).aspx
    
    Dim objVMI As Object
    Dim vAdptr As Variant
    Dim objAdptr As Object
    
    Set objVMI = GetObject("winmgmts:\\" & "." & "\root\cimv2")
    Set vAdptr = objVMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    
    For Each objAdptr In vAdptr
        GetMACAddress = objAdptr.MACAddress
    Next objAdptr

End Function

Function GetRegistryKey(rPath As String)
    Dim regApi As Object
    Set regApi = CreateObject("WScript.Shell")
    GetRegistryKey = regApi.RegRead(rPath)
End Function
