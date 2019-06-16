' LastLogon.vbs
' VBScript program to determine when each user in the domain last logged
' on.
'E. Rudakov 01/10/06
'01/31/2006 EER added Excel dialog

Option Explicit


Dim objFSO, objFile
Dim objRootDSE, strConfig, objConnection, objCommand, strQuery, objNet, a
Dim objRecordSet, objDC
Dim strDNSDomain, objShell, lngBiasKey, lngBias, k, arrstrDCs(), lngFlag, strPswdOPt, strBat 
Dim strDN, dtmDate, objDate, lngDate, objList, strUser , lngDateCr , objListD , objListCr , objUser , objExcel
Dim strBase, strFilter, strAttributes, lngHigh, lngLow , bitSW , strLLDate , strCrDate , strDays, strExcFileName 
Dim objListBadD, objListBadCount, intBadCnt, strBdDate, intLen, intIndex, strCn, strOU, binTest
Set objList = CreateObject("Scripting.Dictionary")
objList.CompareMode = vbTextCompare
strExcFileName = "h:\Profile batch file " & Replace(FormatDateTime(Now, 2), "/", "_") & ".xls"
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
'binTest = False


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
' 

objExcel.ActiveSheet.Range("A1").Activate

objExcel.ActiveCell.Value = "Account name"
objExcel.ActiveCell.Offset(0,1).Value = "Name"
objExcel.ActiveCell.Offset(0,2).Value = "OU Name"
objExcel.ActiveCell.Offset(0,3).Value = "batch file"


objExcel.ActiveCell.Offset(1,0).Activate


  strBase = "<LDAP://rlicdc2.reshealthcare.org/" & strDNSDomain & ">"
  If binTest Then Wscript.Echo strBase
  objCommand.Properties("Page Size") = 1000
  strFilter = "(&(objectCategory=person)(objectClass=user)(!useraccountcontrol:1.2.840.113556.1.4.803:=2))"
  strAttributes = "distinguishedName,scriptPath,sAMAccountName"
  strQuery = strBase & ";" & strFilter & ";" & strAttributes _
    & ";subtree"
  objCommand.CommandText = strQuery
  On Error Resume Next
  Set objRecordSet = objCommand.Execute
  If Err.Number <> 0 Then
    On Error GoTo 0
    Wscript.Echo "Domain Controller not available: " & strBase
  Else
    'On Error GoTo 0
      k = 0 
      Do Until objRecordSet.EOF
         k = k + 1
         strDN = objRecordSet.Fields("distinguishedName")
         bitSW = 0
         bitSW = Instr(1, strDN, "OU=Service_Accounts", 1)
         bitSW = bitSW + Instr(1, strDN, "OU=NS_TEST", 1)
         bitSW = bitSW + Instr(1, strDN, "OU=SMS_Accounts", 1)
         'bitSW = bitSW + Instr(1, strDN, "OU=Physicians", 1)
         bitSW = bitSW + Instr(1, strDN, "OU=Training Accounts", 1)
         bitSW = bitSW + Instr(1, strDN, "OU=Quick Start Accounts", 1)
         strBat = Trim(LCase(objRecordSet.Fields("scriptPath")))
         If strBat = "runme.bat" Then bitSW = 1
         if bitSW = 0 Then
            intIndex = 0
            intLen = Len(strDN)
            intIndex = Instr(1, strDN, ",OU=", 1)
            If intIndex > 0 Then    
               strCn = Mid(strDN, 1, intIndex - 1)
               strOU = Mid(strDN,intIndex + 1)
            Else
               strCn = strDN
               strOU = ""
            End If   
            If Not IsNull(strBat) Then   
               If Not objList.Exists(strBat) Then objList.Add strBat, strBat
            End If
            objExcel.ActiveCell.Value = objRecordSet.Fields("sAMAccountName")
            objExcel.ActiveCell.Offset(0,1).Value = strCn
            objExcel.ActiveCell.Offset(0,2).Value = strOu
            objExcel.ActiveCell.Offset(0,3).Value = strBat 
            objExcel.ActiveCell.Offset(1,0).Activate
         End If
         objRecordSet.MoveNext
    Loop
  End If

Wscript.Echo "Total of " & k & " users"

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
a = objList.Items   
For k = 0 To objList.Count -1
   Wscript.Echo a(k)
Next

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