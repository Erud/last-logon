
Option Explicit

Dim objRootDSE, strConfig, objConnection, objCommand, strQuery
Dim objRecordSet, objDC
Dim strDNSDomain, objShell, lngBiasKey, lngBias, k, arrstrDCs() , inLoop
Dim strDN, dtmDate, objDate, lngDate, objList, strUser , lngDateCr , objListD , objListCr , objUser , objExcel
Dim strBase, strFilter, strAttributes, strSave , strAMAaccount
'

strAMAaccount = InputBox("Enter account name")

' Determine configuration context and DNS domain from RootDSE object.
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfig = objRootDSE.Get("configurationNamingContext")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

' Use ADO to search Active Directory for ObjectClass nTDSDSA.
' This will identify all Domain Controllers.
Set objCommand = CreateObject("ADODB.Command")
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection

strBase = "<LDAP://" & strConfig & ">"
strFilter = "(objectClass=nTDSDSA)"
strAttributes = "AdsPath"
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

objCommand.CommandText = strQuery
objCommand.Properties("Page Size") = 100
objCommand.Properties("Timeout") = 60
objCommand.Properties("Cache Results") = False

Set objRecordSet = objCommand.Execute

' Enumerate parent objects of class nTDSDSA. Save Domain Controller
' AdsPaths in dynamic array arrstrDCs.
k = 0
Do Until objRecordSet.EOF
  Set objDC = _
    GetObject(GetObject(objRecordSet.Fields("AdsPath")).Parent)
  ReDim Preserve arrstrDCs(k)
  arrstrDCs(k) = objDC.DNSHostName
  k = k + 1
  objRecordSet.MoveNext
Loop
' Excel Inint
Set objExcel = WScript.CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.ActiveSheet.Name = strAMAaccount
objExcel.ActiveSheet.Range("A4").Activate

objExcel.ActiveCell.Value = "Total of " & k & " servers."
objExcel.ActiveCell.Offset(1,0).Activate				'move 1 down
' Retrieve lastLogon attribute for each user on each Domain Controller.
objExcel.ActiveCell.Value = "Server"						'col header 1
objExcel.ActiveCell.Offset(0,1).Value = "whenChanged"
objExcel.ActiveCell.Offset(1,0).Activate				'move 1 down

strSave = "1/1/1601"

For k = 0 To Ubound(arrstrDCs)
' For k = 0 To 0
  strBase = "<LDAP://" & arrstrDCs(k) & "/" & strDNSDomain & ">"

  strFilter = "(&(objectClass=computer)(Name=" & strAMAaccount & "))"
  strAttributes = "whenChanged"
  strQuery = strBase & ";" & strFilter & ";" & strAttributes _
    & ";subtree"
  objCommand.CommandText = strQuery
  On Error Resume Next
  Set objRecordSet = objCommand.Execute

  If Err.Number <> 0 Then
    On Error GoTo 0
    Wscript.Echo "Domain Controller not available: " & arrstrDCs(k)
  Else
    On Error GoTo 0
      Do Until objRecordSet.EOF

       strDN = objRecordSet.Fields("whenChanged")
      
       'if strDN = "1/1/1601" then strDN = "N/A"
       objExcel.ActiveCell.Value = k + 1 & " " & strBase
       objExcel.ActiveCell.Offset(0,1).Value = strDN
       objExcel.ActiveCell.Offset(1,0).Activate				'move 1 down
       If DateDiff("d", strSave , strDN) > 0 Then strSave = strDN
      objRecordSet.MoveNext
    Loop
  End If
Next
' Wscript.Quit
' Output latest lastLogon date for each user.
objExcel.ActiveSheet.Range("A1").Activate
objExcel.ActiveCell.Value = "whenChanged"						'col header 1
objExcel.ActiveCell.Offset(1,0).Activate
objExcel.ActiveCell.Value = strSave
objExcel.ActiveCell.Offset(0,1).Value = DateDiff("d", strSave, now)
' Clean up.
objConnection.Close
Set objRootDSE = Nothing
Set objConnection = Nothing
Set objCommand = Nothing
Set objRecordSet = Nothing
Set objDC = Nothing
