' LastLogon.vbs
' VBScript program to determine when each user in the domain last logged
' on.
'E. Rudakov 01/10/06
'01/31/2006 EER added Excel dialog

Option Explicit

Const ADS_UF_PASSWD_CANT_CHANGE = &H40
Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000
Const ADS_UF_PASSWD_NOTREQD = &H20
Const ADS_UF_PASSWORD_EXPIRED = &H800000

Dim objFSO, objFile
Dim objRootDSE, strConfig, objConnection, objCommand, strQuery, objNet
Dim objRecordSet, objDC
Dim strDNSDomain, objShell, lngBiasKey, lngBias, k, arrstrDCs(), lngFlag, strPswdOPt
Dim strDN, dtmDate, objDate, lngDate, objList, strUser , lngDateCr , objListD , objListCr , objUser , objExcel
Dim strBase, strFilter, strAttributes, lngHigh, lngLow , bitSW , strLLDate , strCrDate , strDays, strExcFileName 
Dim objListBadD, objListBadCount, intBadCnt, strBdDate, intLen, intIndex, strCn, strOU, binTest, intErrNum
' Just to be safe??????????
Set objList = Nothing

' See where we are?
'call VerifyLoc("\\RLICFS01\SHARED\IS\AdminScripts\LastLogonPhUEx.vbs")

strExcFileName = "h:\last logon times " & Replace(FormatDateTime(Now, 2), "/", "_") & ".xls"
' Get Name of Input File and Check to see if its valid
strExcFileName = InputBox("Enter file name for output Excel spreadsheet." & vblf & "File will be deleted.",,strExcFileName)
If strExcFileName = "" Then
   MsgBox ("Operation Cancelled, no Excel file supplied")
   Wscript.Quit(1)
End if 
Set objNet = CreateObject("WScript.NetWork")
If objNet.UserName = "erudakov" Then 
   binTest = True
Else
   binTest = False
End If 
binTest = False

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
If binTest Then Wscript.Echo "Total of " & k & " servers."
' Retrieve lastLogon attribute for each user on each Domain Controller.
For k = 0 To Ubound(arrstrDCs)
' For k = 0 To 1
  strBase = "<LDAP://" & arrstrDCs(k) & "/" & strDNSDomain & ">"
  If binTest Then Wscript.Echo   k + 1 & " " & strBase
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
      bitSW = bitSW + Instr(1, strDN, "Shared_PC_Accounts", 1)
      bitSW = bitSW + Instr(1, strDN, "OU=Vendors", 1)
      bitSW = bitSW + Instr(1, strDN, "OU=Service_Accounts", 1)
      bitSW = bitSW + Instr(1, strDN, "OU=service account", 1)
      bitSW = bitSW + Instr(1, strDN, "OU=Training Accounts", 1)
      bitSW = bitSW + Instr(1, strDN, "OU=WSCN (College of Nursing)", 1)
      bitSW = bitSW + Instr(1, strDN, "OU=HMI_SERVICE_ACCTS", 1)
      bitSW = bitSW + Instr(1, strDN, "$,CN=USERS,DC=", 1)
      bitSW = bitSW + Instr(1, strDN, "CN=Toomey\, Joseph", 1)
      bitSW = bitSW + Instr(1, strDN, "CN=Rollins\, Skip", 1)
      bitSW = bitSW + Instr(1, strDN, "CN=Veres\, Ludovic", 1)
      bitSW = bitSW + Instr(1, strDN, "CN=tandberg", 1)
      bitSW = bitSW + Instr(1, UCase(strDN), "CN=CONSULTANTROLE", 1)
      bitSW = bitSW + Instr(1, UCase(strDN), "CN=EMPLOYEEROLE", 1)
      bitSW = bitSW + Instr(1, UCase(strDN), "CN=STUDENTROLE", 1)
      bitSW = bitSW + Instr(1, UCase(strDN), "CN=RESIDENTROLE", 1) 
      
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
If binTest Then Wscript.Echo "Total of " & objList.Count & " users."
'
' Delete file if it alread exists in the destination.
Set objFSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set objFile = objFSO.GetFile(strExcFileName)
If Err.Number = 0 Then objFile.Delete ()
On Error Goto 0
'
' Excel Inint
Set objExcel = WScript.CreateObject("Excel.Application")
objExcel.Visible = binTest
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
'objExcel.ActiveCell.Offset(0,7).Value = "Distinguished Name"
objExcel.ActiveCell.Offset(0,7).Value = "CN Name"
objExcel.ActiveCell.Offset(0,8).Value = "OU Name"
objExcel.ActiveCell.Offset(0,9).Value = "Last Bad Logon"
objExcel.ActiveCell.Offset(0,10).Value = "Bad Password Count"
objExcel.ActiveCell.Offset(0,11).Value = "Password Options"

objExcel.ActiveCell.Offset(1,0).Activate

k = 0

For Each strUser In objList
  'If k > 49 Then WScript.Echo strUser
  If (DateDiff("d", objList(strUser), Now) > 120) and (DateDiff("d", objListCr(strUser), Now) > 120) And _
     (InStr(1, strUser, "OU=Physicians", 1) = 0) And (CheckQS(strUser) > 0) Then
     On Error Resume Next
     intErrNum = 0
     Set objUser = GetObject("LDAP://" & strUser )
     If Err.Number <> 0 Then intErrNum = Err.Number
     If intErrNum <> 0 Then
        Wscript.Echo strUser & " erorr " & intErrNum
        On Error GoTo 0
     Else
        lngFlag = objUser.Get("userAccountControl")
        If Err.Number <> 0 Then intErrNum = Err.Number
        strPswdOPt = ""
        If (lngFlag And ADS_UF_PASSWD_CANT_CHANGE) <> 0 Then strPswdOPt = "Cant Change"
        If (lngFlag And ADS_UF_DONT_EXPIRE_PASSWD) <> 0 Then strPswdOPt = strPswdOPt & " Dont Expire"
        If (lngFlag And ADS_UF_PASSWD_NOTREQD) <> 0 Then strPswdOPt = strPswdOPt & " Not Req"
        If (lngFlag And ADS_UF_PASSWORD_EXPIRED) <> 0 Then strPswdOPt = strPswdOPt & " Expired"
        Set objDate = objUser.PwdLastSet
        If Err.Number <> 0 Then intErrNum = Err.Number
        lngHigh = objDate.HighPart
        lngLow = objdate.LowPart
        If (lngHigh = 0) And (lngLow = 0) Then strPswdOPt = strPswdOPt & " Must change"
     End If 
     On Error GoTo 0
     If intErrNum <> 0 Then
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
'         objExcel.ActiveCell.Offset(0,7).Value = strUser
         intIndex = 0
         intLen = Len(strUser)
         intIndex = Instr(1, strUser, ",OU=", 1)
         If intIndex > 0 Then    
            strCn = Mid(strUser, 1, intIndex - 1)
            strOU = Mid(strUser,intIndex + 1)
         Else
            strCn = strUser
            strOU = ""
         End If
         objExcel.ActiveCell.Offset(0,7).Value = strCn
         objExcel.ActiveCell.Offset(0,8).Value = strOU         
         objExcel.ActiveCell.Offset(0,9).Value = strBdDate
         objExcel.ActiveCell.Offset(0,10).Value = objListBadCount(strUser)
         objExcel.ActiveCell.Offset(0,11).Value = strPswdOPt
         objExcel.ActiveCell.Offset(1,0).Activate
         k = k + 1
     End if
     Set objUser = Nothing
  End If
Next
If binTest Then Wscript.Echo "Total of " & k & " old users."

objExcel.ActiveSheet.Range("A1:Z1").Font.Bold = True
objExcel.ActiveSheet.Columns("A:Z").AutoFit
objExcel.ActiveSheet.Select
objExcel.ActiveSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True
'Physicians
objExcel.Sheets(2).Activate
objExcel.ActiveSheet.Name = "Physicians"
' Output latest lastLogon date for each user.

objExcel.ActiveSheet.Range("A1").Activate
objExcel.ActiveCell.Value = "Last Logon"						'col header 1
objExcel.ActiveCell.Offset(0,1).Value = "Creation Date"
objExcel.ActiveCell.Offset(0,2).Value = "User Name"
objExcel.ActiveCell.Offset(0,3).Value = "Display Name"
objExcel.ActiveCell.Offset(0,4).Value = "Name"
objExcel.ActiveCell.Offset(0,5).Value = "Description"
objExcel.ActiveCell.Offset(0,6).Value = "Days Inactive"
objExcel.ActiveCell.Offset(0,7).Value = "CN Name"
objExcel.ActiveCell.Offset(0,8).Value = "OU Name"
objExcel.ActiveCell.Offset(0,9).Value = "Last Bad Logon"
objExcel.ActiveCell.Offset(0,10).Value = "Bad Password Count"
objExcel.ActiveCell.Offset(0,11).Value = "Password Options"

objExcel.ActiveCell.Offset(1,0).Activate

k = 0
For Each strUser In objList
  If (DateDiff("d", objList(strUser), Now) > 120) and (DateDiff("d", objListCr(strUser), Now) > 120) And _
     (InStr(1, strUser, "OU=Physicians", 1) > 0) Then
     On Error Resume Next
     Set objUser = GetObject("LDAP://" & strUser )
     intErrNum = 0
     If Err.Number <> 0 Then 
        Wscript.Echo strUser, Err.Number
        intErrNum = Err.Number
     End If
     lngFlag = objUser.Get("userAccountControl")
     strPswdOPt = ""
     If (lngFlag And ADS_UF_PASSWD_CANT_CHANGE) <> 0 Then strPswdOPt = "Cant Change"
     If (lngFlag And ADS_UF_DONT_EXPIRE_PASSWD) <> 0 Then strPswdOPt = strPswdOPt & " Dont Expire"
     If (lngFlag And ADS_UF_PASSWD_NOTREQD) <> 0 Then strPswdOPt = strPswdOPt & " Not Req"
     If (lngFlag And ADS_UF_PASSWORD_EXPIRED) <> 0 Then strPswdOPt = strPswdOPt & " Expired"
     Set objDate = objUser.PwdLastSet
     lngHigh = objDate.HighPart
     lngLow = objdate.LowPart
     If (lngHigh = 0) And (lngLow = 0) Then strPswdOPt = strPswdOPt & " Must change"
     
     On Error GoTo 0
     If (Err.Number <> 0) Or (intErrNum <> 0) Then
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
'         objExcel.ActiveCell.Offset(0,7).Value = strUser
         intIndex = 0
         intLen = Len(strUser)
         intIndex = Instr(1, strUser, ",OU=", 1)
         If intIndex > 0 Then    
            strCn = Mid(strUser, 1, intIndex - 1)
            strOU = Mid(strUser,intIndex + 1)
         Else
            strCn = strUser
            strOU = ""
         End If
         objExcel.ActiveCell.Offset(0,7).Value = strCn
         objExcel.ActiveCell.Offset(0,8).Value = strOU 
         objExcel.ActiveCell.Offset(0,9).Value = strBdDate
         objExcel.ActiveCell.Offset(0,10).Value = objListBadCount(strUser)
         objExcel.ActiveCell.Offset(0,11).Value = strPswdOPt
         objExcel.ActiveCell.Offset(1,0).Activate
         k = k + 1
     End if
     Set objUser = Nothing
  End If
Next
If binTest Then Wscript.Echo "Total of " & k & " old Physicians."


objExcel.ActiveSheet.Range("A1:Z1").Font.Bold = True
objExcel.ActiveSheet.Columns("A:Z").AutoFit
objExcel.ActiveSheet.Select
objExcel.ActiveSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True

' Save the spreadsheet and close the workbook.
objExcel.ActiveWorkbook.SaveAs strExcFileName
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

WScript.Echo "<***** script completed *****>"


WScript.Quit

Function CheckQS(strName)
'CheckQS = 99
'Exit Function

Dim struserWorkstations, arrayComp, comp, strUcomp, strUName, sUser, sDomain, strUUname
Dim objRootDSE, objItem, objWMIService, colProc, colItems, oProcess
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

strName = UCase(strName)
If (InStr(1, strName, "OU=QUICK START ACCOUNTS", 1) = 0) Then
   CheckQS = 99
   Exit Function
End If
CheckQS = 99
   Exit Function
arrayComp = Split(strName, ",", -1, 1)
comp = arrayComp(0)
arrayComp = Split(comp, "=", -1, 1)
strUName = UCase("RHCMASTER\" & arrayComp(1))
'***********************************************

On Error Resume Next
Set objItem = GetObject("LDAP://" & strName )
If Err.Number > 0 Then 
   Err.Clear ' Clear the error. 
   On Error GoTo 0
   CheckQS = 0
   WScript.Echo "LDAP Error " & strName
   Exit Function
End if   
'***********************************************

struserWorkstations = objItem.Get("userWorkstations")
If Err.Number > 0 Then 
   Err.Clear ' Clear the error. 
   On Error GoTo 0
   CheckQS = 0
   WScript.Echo "Get workstation error " & strName
   Exit Function
End if 
On Error GoTo 0
arrayComp = Split(struserWorkstations, ",", -1, 1)
For Each comp in arrayComp
   strUcomp = UCase(comp)
   If (Left(strUcomp,5)) <> "RHCTS" And (Right(strUcomp,4) <> "WS01") Then 
      'WScript.Echo strUcomp
      ' and who is at that workstation?
      If IsConnectable(strUcomp) Then		

      'Connect to machine and check owner of "explorer.exe"		
          On Error Resume Next
          Set objWMIService = GetObject("winmgmts:\\" & strUcomp & "\root\CIMV2")
          If Err.Number > 0 Then 
             Err.Clear ' Clear the error. 
             On Error GoTo 0	
             CheckQS = 0
             WScript.Echo "winmgmts Error " & strName & " " & strUcomp
             Exit Function
          End If
             
          Set colProc = objWmiService.ExecQuery("Select Name from Win32_Process"  & _
              " Where Name='explorer.exe' and  SessionID=0")
          If Err.Number > 0 Then 
             Err.Clear ' Clear the error. 
             On Error GoTo 0	
             CheckQS = 0
             WScript.Echo "Select Name from Win32_Process Error " & strName & " " & strUcomp
             Exit Function
          End If
          
          On Error GoTo 0
          If colProc.Count > 0 Then
              For Each oProcess In colProc
                  oProcess.GetOwner sUser, sDomain
                  strUUname = UCase(sDomain & "\" & sUser)
                  If strUUname = strUName Then 
                     CheckQS = 0 ' Ok
                     Exit Function
                  Else
                     CheckQS = 4 'some one on
                  End If
              Next
          Else
              CheckQS = 2 'no one logon
          End If
      
      Else
          CheckQS = 1 'unreacable
          Exit Function
      End If

      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
   End If   
Next 

End Function

Function IsConnectable(strCcomp)

Dim objPing, objStatus

'If machine is local machine, just exit the function.	
    If strCcomp = "." Then
        IsConnectable = True
        Exit Function
    End If

'Check if the remote machine is online. 
    
    Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
        ("select Replysize from Win32_PingStatus where address = '" & strCcomp & "'")

    For Each objStatus in objPing
        If  IsNull(objStatus.ReplySize) Then 
            IsConnectable=False
        Else 
            IsConnectable = True
        End If
    Next
    Set objPing=Nothing
    Set objStatus=Nothing

End Function

Sub VerifyLoc(strPatchL)
   Dim strScriptPath, numMSG, strMsg, strDrive, objFs, objDrv
   
   strScriptPath = WScript.ScriptFullName
   
   If Mid(strScriptPath,2,1) = ":" Then 
   
      strDrive = Left(strScriptPath,1)
      Set objFs = CreateObject("Scripting.FileSystemObject") 
      Set objDrv = objFs.GetDrive(strDrive) 
      strDrive = Mid(strScriptPath,3)
      strScriptPath = objDrv.ShareName  & strDrive
   
   End If
   
   ' WScript.Echo strScriptPath
   If LCase(strScriptPath) <> LCase(strPatchL) Then
      strMsg = "This script cannot be executed as" & vbCrLf & vbNewLine & strScriptPath 
      numMSG = MsgBox (strMsg, vbCritical, "Restricted script")
      WScript.Quit
   End If  
End Sub