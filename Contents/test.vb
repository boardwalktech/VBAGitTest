'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module  :    DBMail
'* Description :
'* Created :    03-16-2018 11:25
'* Modified:    03-16-2018 11:25
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub/Fn  :    DataBaseMailCall
'* Description :
'* Created :    03-16-2018 11:35
'* Modified:    03-16-2018 11:35
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub DataBaseMailCall()
    On Error GoTo Err_DataBaseMailCall
    Dim strTolist As String
    Dim strCClist As String
    Dim strSubject As String
    Dim strBody As String
    If checkUserLogin = False Then
        Exit Sub
    End If
    Dim objWSF As New clsBWWorksheetFunction
    objWSF.setApplicationStatusBar "DataBase mail sending started.."
    strTolist = "sujith.yetipathi@boardwalktech.com;lakshman.sai@boardwalktech.com"
    strCClist = ""
    strSubject = "Test Subject"
    strBody = "Test mail body"
    Dim strResult As String
    strResult = DataBaseMail(strTolist, strCClist, strSubject, strBody)
    If strResult = "sent" Then
        MsgBox "Mail sent succesfully.", vbOKOnly + vbInformation, "DataBase Mail Status"
        objWSF.setApplicationStatusBar "Mail sent succesfully."
    Else
        MsgBox "There is an error in sending the Database Mail. Please check the database mail logs.", vbOKOnly + vbCritical, "DataBase Mail Status"
    End If
    objWSF.setApplicationStatusBar "Ready."
    Exit Sub
Err_DataBaseMailCall:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error occurred in DataBaseMailCall"

End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub/Fn  :    DataBaseMail
'* Description :
'* Created :    03-16-2018 11:29
'* Modified:    03-16-2018 11:29
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function DataBaseMail(ByVal strTolist As String, ByVal strCClist As String, ByVal strSubject As String, ByVal strBody As String) As String
    On Error GoTo Err_DataBaseMail
    
    If checkUserLogin = False Then
        DataBaseMail = "Please login"
        Exit Function
    End If
    If strTolist = "" Or strSubject = "" Or strBody = "" Then
        DataBaseMail = "strTolist, strSubject, strBody should not be blank"
        Exit Function
    End If
    If strCClist = "" Then
        strCClist = " "
    End If
    
    Dim StrInput As String
    StrInput = strTolist & "|" & strCClist & "|" & strSubject & "|" & strBody
    Dim strReportName As String: strReportName = "DBMailSheet"
    Dim rangeReturn As Range: Set rangeReturn = getDefinedExternalData(strReportName, EXT_QUERIES.DATABASE_MAIL, StrInput)
    
    DataBaseMail = rangeReturn.Cells(2, 2).Value
    Application.DisplayAlerts = False
    rangeReturn.Parent.Delete
    
    
    Exit Function
Err_DataBaseMail:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error occurred in DataBaseMail"

End Function

