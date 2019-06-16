' LastLogon.vbs
' VBScript program to determine when each user in the domain last logged
' on.
'
' ----------------------------------------------------------------------
' Copyright (c) 2002 Richard L. Mueller
' Hilltop Lab web site - http://www.rlmueller.net
' Version 1.0 - December 7, 2002
' Version 1.1 - January 17, 2003 - Account for null value for lastLogon.
' Version 1.2 - January 23, 2003 - Account for DC not available.
' Version 1.3 - February 3, 2003 - Retrieve users but not contacts.
' Version 1.4 - February 19, 2003 - Standardize Hungarian notation.
' Version 1.5 - March 11, 2003 - Remove SearchScope property.
' Version 1.6 - May 9, 2003 - Account for error in IADsLargeInteger
'                             property methods HighPart and LowPart.
' Version 1.7 - January 25, 2004 - Modify error trapping.
'
' Because the lastLogon attribute is not replicated, every Domain
' Controller in the domain must be queried to find the latest lastLogon
' date for each user. The lastest date found is kept in a dictionary
' object. The program first uses ADO to search the domain for all Domain
' Controllers. The AdsPath of each Domain Controller is saved in an
' array. Then, for each Domain Controller, ADO is used to search the
' copy of Active Directory on that Domain Controller for all user
' objects and return the lastLogon attribute. The lastLogon attribute is
' a 64-bit number representing the number of 100 nanosecond intervals
' since 12:00 am January 1, 1601. This value is converted to a date. The
' last logon date is in UTC (Coordinated Univeral Time). It must be
' adjusted by the Time Zone bias in the machine registry to convert to
' local time.
'
' You have a royalty-free right to use, modify, reproduce, and
' distribute this script file in any way you find useful, provided that
' you agree that the copyright owner above has no warranty, obligations,
' or liability for such use.

Option Explicit

Dim objRootDSE, strConfig, objConnection, objCommand, strQuery
Dim objRecordSet, objDC
Dim strDNSDomain, objShell, lngBiasKey, lngBias, k, arrstrDCs()
Dim strDN, dtmDate, objDate, lngDate, objList, strUser , dtmDateCr , lngDateCr , objListCr
Dim strBase, strFilter, strAttributes, lngHigh, lngLow

' Use a dictionary object to track latest lastLogon for each user.
Set objList = CreateObject("Scripting.Dictionary")
objList.CompareMode = vbTextCompare
Set objListCr = CreateObject("Scripting.Dictionary")
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
Wscript.Echo "Total of " & k & " servers."
' Retrieve lastLogon attribute for each user on each Domain Controller.
For k = 0 To Ubound(arrstrDCs)
' For k = 0 To 1
  strBase = "<LDAP://" & arrstrDCs(k) & "/" & strDNSDomain & ">"
  Wscript.Echo   k + 1 & " " & strBase
  strFilter = "(&(objectCategory=person)(objectClass=user)(!useraccountcontrol:1.2.840.113556.1.4.803:=2))"
  strAttributes = "distinguishedName,lastLogon,createTimeStamp"
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
      lngDate = objRecordSet.Fields("lastLogon")
      lngDateCr = objRecordSet.Fields("createTimeStamp")
      On Error Resume Next
      Set objDate = lngDate

      If VarType (lngDate) <> 7 Then lngDate = #01/01/1601 12:00 AM#

      If objList.Exists(strDN) Then
        If lngDate > objList(strDN) Then
          objList(strDN) = lngDate
        End If
      Else
        objList.Add strDN, lngDate
      End If
      If Not objListCr.Exists(strDN) Then objListCr.Add strDN, lngDateCr
      objRecordSet.MoveNext
      'Exit do
    Loop
  End If
Next

' Output latest lastLogon date for each user.
For Each strUser In objList
  Wscript.Echo strUser & " ; " & objList(strUser) & " ; " & objListCr(strUser)
Next

' Clean up.
objConnection.Close
Set objRootDSE = Nothing
Set objConnection = Nothing
Set objCommand = Nothing
Set objRecordSet = Nothing
Set objDC = Nothing
Set objDate = Nothing
Set objList = Nothing
Set objListCr = Nothing
Set objShell = Nothing