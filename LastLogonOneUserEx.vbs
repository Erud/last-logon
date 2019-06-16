
Option Explicit

Dim objRootDSE, strConfig, objConnection, objCommand, strQuery
Dim objRecordSet, objDC
Dim strDNSDomain, objShell, lngBiasKey, lngBias, k, arrstrDCs() , inLoop
Dim strDN, dtmDate, objDate, lngDate, objList, strUser , lngDateCr , objListD , objListCr , objUser , objExcel
Dim strBase, strFilter, strAttributes, lngHigh, lngLow , bitSW , strLLDate , strCrDate , strAMAaccount
'
'strAMAaccount = "RRay"
strAMAaccount = InputBox("Enter account name")
' Use a dictionary object to track latest lastLogon for each user.
Set objList = CreateObject("Scripting.Dictionary")
objList.CompareMode = vbTextCompare
Set objListCr = CreateObject("Scripting.Dictionary")
objListCr.CompareMode = vbTextCompare
Set objListD = CreateObject("Scripting.Dictionary")
objListCr.CompareMode = vbTextCompare

' Obtain local Time Zone bias from machine registry.
Set objShell = CreateObject("Wscript.Shell")
lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\" _
  & "TimeZoneInformation\ActiveTimeBias")
If UCase(TypeName(lngBiasKey)) = "LONG" Then
  lngBias = lngBiasKey
ElseIf UCase(TypeName(lngBiasKey)) = "VARIANT()" Then
  lngBias = 0
  For k = 0 To UBound(lngBiasKey)
    lngBias = lngBias + (lngBiasKey(k) * 256^k)
  Next
End If

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
objExcel.ActiveCell.Offset(0,1).Value = "Last Logon"
objExcel.ActiveCell.Offset(0,2).Value = "Creation Date"
objExcel.ActiveCell.Offset(0,3).Value = "Last Bad Logon"
objExcel.ActiveCell.Offset(0,4).Value = "Bad Logons Count"
objExcel.ActiveCell.Offset(1,0).Activate				'move 1 down

For k = 0 To Ubound(arrstrDCs)
' For k = 0 To 0
  strBase = "<LDAP://" & arrstrDCs(k) & "/" & strDNSDomain & ">"

  strFilter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=" & strAMAaccount _
  & ")(!useraccountcontrol:1.2.840.113556.1.4.803:=2))"
  strAttributes = "distinguishedName,lastLogon,createTimeStamp,badPasswordTime,badPwdCount"
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
    inLoop = 0
      Do Until objRecordSet.EOF

      strDN = objRecordSet.Fields("distinguishedName")
      lngDate = objRecordSet.Fields("lastLogon")
      lngDateCr = objRecordSet.Fields("createTimeStamp")
        On Error Resume Next
        Set objDate = lngDate
        If Err.Number <> 0 Then
          On Error GoTo 0
          dtmDate = #01/01/1601 12:00 AM#
        Else
          On Error GoTo 0
          lngHigh = objDate.HighPart
          lngLow = objDate.LowPart
          If lngLow < 0 Then
            lngHigh = lngHigh + 1
          End If
          If (lngHigh = 0) And (lngLow = 0 ) Then
            dtmDate = #01/01/1601 12:00 AM#
          Else
            dtmDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
              + lngLow)/600000000 - lngBias)/1440
          End If
        End If
' /=/=/=/=/=/=/==/=/=/=/=/=/==/=/=/=/=/=/=
        If objList.Exists(strDN) Then
          If dtmDate > objList(strDN) Then
            objList(strDN) = dtmDate
          End If
        Else
          objList.Add strDN, dtmDate
        End If
        If Not objListCr.Exists(strDN) Then objListCr.Add strDN, lngDateCr
'-2-2-2-2-2-2-2-2-2
        On Error Resume Next
        lngDate = objRecordSet.Fields("badPasswordTime")
        Set objDate = lngDate
        If Err.Number <> 0 Then
          On Error GoTo 0
          dtmDate = #01/01/1601 12:00 AM#
        Else
          On Error GoTo 0
          lngHigh = objDate.HighPart
          lngLow = objDate.LowPart
          If lngLow < 0 Then
            lngHigh = lngHigh + 1
          End If
          If (lngHigh = 0) And (lngLow = 0 ) Then
            dtmDate = #01/01/1601 12:00 AM#
          Else
            dtmDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
              + lngLow)/600000000 - lngBias)/1440
          End If
        End If
'-2-2-2-2-2-2-2-2-2
       strLLDate = FormatDateTime(objList(strDN), 0)
       if strLLDate = "1/1/1601" then strLLDate = "N/A"
       strCrDate = FormatDateTime(dtmDate, 0)
       if strCrDate = "1/1/1601" then strCrDate = "N/A"
       objExcel.ActiveCell.Value = k + 1 & " " & strBase
       objExcel.ActiveCell.Offset(0,1).Value = strLLDate
       objExcel.ActiveCell.Offset(0,2).Value = FormatDateTime(lngDateCr, 0)
       objExcel.ActiveCell.Offset(0,3).Value = strCrDate
       objExcel.ActiveCell.Offset(0,4).Value = objRecordSet.Fields("badPwdCount")
       objExcel.ActiveCell.Offset(1,0).Activate				'move 1 down

      objRecordSet.MoveNext
      inLoop = 1
    Loop
    if inLoop = 0 then
       objExcel.ActiveCell.Value = k + 1 & " " & strBase
       objExcel.ActiveCell.Offset(0,1).Value = "N/A"
       objExcel.ActiveCell.Offset(0,2).Value = "N/A"
       objExcel.ActiveCell.Offset(0,3).Value = "N/A"
       objExcel.ActiveCell.Offset(0,4).Value = " "
       objExcel.ActiveCell.Offset(1,0).Activate
    End If
  End If
Next
' Wscript.Quit
' Output latest lastLogon date for each user.
objExcel.ActiveSheet.Range("A1").Activate
objExcel.ActiveCell.Value = "Last Logon"						'col header 1
objExcel.ActiveCell.Offset(0,1).Value = "Creation Date"
objExcel.ActiveCell.Offset(0,2).Value = "User Name"
objExcel.ActiveCell.Offset(0,3).Value = "Display Name"
objExcel.ActiveCell.Offset(0,4).Value = "Name"
objExcel.ActiveCell.Offset(0,5).Value = "Description"
objExcel.ActiveCell.Offset(0,6).Value = "Days Inactive"
objExcel.ActiveCell.Offset(0,7).Value = "Distinguished Name"
objExcel.ActiveCell.Offset(1,0).Activate

For Each strUser In objList

     Set objUser = GetObject("LDAP://" & strUser )
     strLLDate = FormatDateTime(objList(strUser), 2)
     if strLLDate = "1/1/1601" then strLLDate = "N/A"
     strCrDate = FormatDateTime(objListCr(strUser), 2)
     if strCrDate = "1/1/1601" then strCrDate = "N/A"
     objExcel.ActiveCell.Value = strLLDate						'col header 1
     objExcel.ActiveCell.Offset(0,1).Value = strCrDate
     objExcel.ActiveCell.Offset(0,2).Value = objUser.sAMAccountName
     objExcel.ActiveCell.Offset(0,3).Value = objUser.displayName
     objExcel.ActiveCell.Offset(0,4).Value = objUser.cn
     objExcel.ActiveCell.Offset(0,5).Value = objUser.description
     objExcel.ActiveCell.Offset(0,6).Value = DateDiff("d", objList(strUser), Now)
     objExcel.ActiveCell.Offset(0,7).Value = strUser
     objExcel.ActiveCell.Offset(1,0).Activate

     Set objUser = Nothing

Next

' Clean up.
objConnection.Close
Set objUser = Nothing
Set objRootDSE = Nothing
Set objConnection = Nothing
Set objCommand = Nothing
Set objRecordSet = Nothing
Set objDC = Nothing
Set objDate = Nothing
Set objList = Nothing
Set objListCr = Nothing
Set objShell = Nothing