' LastLogon.vbs
' VBScript program to determine when each user in the domain last logged
' on.
'

Option Explicit

Const ADS_UF_PASSWD_CANT_CHANGE = &H40
Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000
Const ADS_UF_PASSWD_NOTREQD = &H20
Const ADS_UF_PASSWORD_EXPIRED = &H800000

Dim objRootDSE, strConfig, objConnection, objCommand, strQuery
Dim objRecordSet, objDC
Dim strDNSDomain, objShell, lngBiasKey, lngBias, k, arrstrDCs(), lngFlag, strPswdOPt
Dim strDN, dtmDate, objDate, lngDate, objList, strUser , lngDateCr , objListD , objListCr , objUser , objExcel
Dim strBase, strFilter, strAttributes, lngHigh, lngLow , bitSW , strLLDate , strCrDate , strDays 
Dim objListBadD, objListBadCount, intBadCnt, strBdDate
' Just to be safe??????????
Set objList = Nothing
' Use a dictionary object to track latest lastLogon for each user.
Set objList = CreateObject("Scripting.Dictionary")
objList.CompareMode = vbTextCompare
Set objListCr = CreateObject("Scripting.Dictionary")
objListCr.CompareMode = vbTextCompare
Set objListBadD = CreateObject("Scripting.Dictionary")
objListBadD.CompareMode = vbTextCompare
Set objListBadCount = CreateObject("Scripting.Dictionary")
objListBadCount.CompareMode = vbTextCompare


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
Wscript.Echo "Total of " & k & " servers."
' Retrieve lastLogon attribute for each user on each Domain Controller.
For k = 0 To Ubound(arrstrDCs)
' For k = 0 To 1
  strBase = "<LDAP://" & arrstrDCs(k) & "/" & strDNSDomain & ">"
  Wscript.Echo   k + 1 & " " & strBase
  strFilter = "(&(objectCategory=person)(objectClass=user)(!useraccountcontrol:1.2.840.113556.1.4.803:=2))"
  strAttributes = "distinguishedName,lastLogon,createTimeStamp,badPwdCount,badPasswordTime"
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
      strDN = objRecordSet.Fields("distinguishedName")
      bitSW = 0
      bitSW = Instr(1, strDN, "OU=Service_Accounts", 1)
      bitSW = bitSW + Instr(1, strDN, "OU=NS_TEST", 1)
      bitSW = bitSW + Instr(1, strDN, "OU=SMS_Accounts", 1)
      bitSW = bitSW + Instr(1, strDN, "OU=Physicians", 1)
      bitSW = bitSW + Instr(1, strDN, "OU=Training Accounts", 1)
      if bitSW = 0 then
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
        If objList.Exists(strDN) Then
          If dtmDate > objList(strDN) Then
            objList(strDN) = dtmDate
          End If
        Else
          objList.Add strDN, dtmDate
        End If
        If Not objListCr.Exists(strDN) Then objListCr.Add strDN, lngDateCr
        lngDate = objRecordSet.Fields("badPasswordTime")
        intBadCnt = objRecordSet.Fields("badPwdCount")  
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
        If objListBadD.Exists(strDN) Then
          If dtmDate > objListBadD(strDN) Then
            objListBadD(strDN) = dtmDate
          End If
        Else
          objListBadD.Add strDN, dtmDate
        End If
        If (Not IsNumeric(intBadCnt)) Then intBadCnt = 0 
        If objListBadCount.Exists(strDN) Then
           objListBadCount(strDN) = objListBadCount(strDN) + intBadCnt 
        Else
          objListBadCount.Add strDN, intBadCnt
        End If
      End if
      '
      objRecordSet.MoveNext
    Loop
  End If
Next
Wscript.Echo "Total of " & objList.Count & " users."
' Excel Inint
Set objExcel = WScript.CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.ActiveSheet.Name = "Users"
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
objExcel.ActiveCell.Offset(0,8).Value = "Last Bad Logon"
objExcel.ActiveCell.Offset(0,9).Value = "Bad Password Count"
objExcel.ActiveCell.Offset(0,10).Value = "Password Options"

objExcel.ActiveCell.Offset(1,0).Activate

k = 0
For Each strUser In objList
  If (DateDiff("d", objList(strUser), Now) > 120) and (DateDiff("d", objListCr(strUser), Now) > 120) then
     On Error Resume Next
     Set objUser = GetObject("LDAP://" & strUser )
     lngFlag = objUser.Get("userAccountControl")
     strPswdOPt = ""
     If (lngFlag And ADS_UF_PASSWD_CANT_CHANGE) <> 0 Then strPswdOPt = "Cant Change"
     If (lngFlag And ADS_UF_DONT_EXPIRE_PASSWD) <> 0 Then strPswdOPt = strPswdOPt & " Dont Expire"
     If (lngFlag And ADS_UF_PASSWD_NOTREQD) <> 0 Then strPswdOPt = strPswdOPt & " Not Req"
     If (lngFlag And ADS_UF_PASSWORD_EXPIRED) <> 0 Then strPswdOPt = strPswdOPt & " Expired"
     If objUser.pwdLastSet = 0 Then strPswdOPt = strPswdOPt & " Must change"
     
     On Error GoTo 0
     If Err.Number <> 0 Then
         strLLDate = FormatDateTime(objList(strUser), 2)
         if strLLDate = "1/1/1601" then strLLDate = "N/A"
         strCrDate = FormatDateTime(objListCr(strUser), 2)
         if strCrDate = "1/1/1601" then strCrDate = "N/A"
         Wscript.Echo strLLDate & ";" & strCrDate & ";;;;;" _
         & ";" & DateDiff("d", objList(strUser), Now) & ";" & Replace(strUser ,",OU=", ";OU=",1,1,1)
     Else
         strLLDate = FormatDateTime(objList(strUser), 2)
         if strLLDate = "1/1/1601" then 
            strLLDate = "N/A"
            strDays   = "N/A"
         Else
            strDays = DateDiff("d", objList(strUser), Now)
         End If
         strCrDate = FormatDateTime(objListCr(strUser), 2)
         if strCrDate = "1/1/1601" then strCrDate = "N/A"
         strBdDate = FormatDateTime(objListBadD(strUser), 2)
         if strBdDate = "1/1/1601" then strBdDate = "N/A"
         objExcel.ActiveCell.Value = strLLDate						'col header 1
         objExcel.ActiveCell.Offset(0,1).Value = strCrDate
         objExcel.ActiveCell.Offset(0,2).Value = objUser.sAMAccountName
         objExcel.ActiveCell.Offset(0,3).Value = objUser.displayName
         objExcel.ActiveCell.Offset(0,4).Value = objUser.cn
         objExcel.ActiveCell.Offset(0,5).Value = objUser.description
         objExcel.ActiveCell.Offset(0,6).Value = strDays
         objExcel.ActiveCell.Offset(0,7).Value = strUser
         objExcel.ActiveCell.Offset(0,8).Value = strBdDate
         objExcel.ActiveCell.Offset(0,9).Value = objListBadCount(strUser)
         objExcel.ActiveCell.Offset(0,10).Value = strPswdOPt
         objExcel.ActiveCell.Offset(1,0).Activate
         k = k + 1
     End if
     Set objUser = Nothing
  End If
Next
Wscript.Echo "Total of " & k & " old users."
objExcel.ActiveSheet.Columns("A:K").AutoFit
objExcel.ActiveSheet.Range("A1:I1").Font.Bold = True
objExcel.ActiveSheet.Select
objExcel.ActiveSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True

' Save the spreadsheet and close the workbook.
objExcel.ActiveWorkbook.SaveAs "Last Logon times"
objExcel.ActiveWorkbook.Close

' Quit Excel.
objExcel.Application.Quit

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
Set objExcel = Nothing