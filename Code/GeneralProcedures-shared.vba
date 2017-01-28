Option Compare Database
Option Explicit
Function StartUp()
On Error GoTo err_startup
DoCmd.OpenForm "Login", acNormal, , , acFormEdit, acDialog
SetCurrentVersion
DoCmd.OpenForm "Front End", acNormal, , , acFormReadOnly 'open main menu
DoCmd.Maximize
Forms![Front End].Refresh
Exit Function
err_startup:
    Call General_Error_Trap
End Function
Sub General_Error_Trap()
    MsgBox "The system has encountered an error. The message is as follows:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Error Code: " & Err.Number, vbOKOnly, "System Error"
End Sub
Function GetCurrentVersion()
On Error GoTo err_GetCurrentVersion
    GetCurrentVersion = VersionNumber
Exit Function
err_GetCurrentVersion:
    Call General_Error_Trap
End Function
Function SetCurrentVersion()
On Error GoTo err_SetCurrentVersion
Dim retVal
retVal = "v"
If DBName <> "" Then
    Dim mydb As Database, myrs As DAO.Recordset
    Dim sql
    Set mydb = CurrentDb()
    sql = "SELECT [Version_Num] FROM [Database_Interface_Version_History] WHERE [MDB_Name] = '" & DBName & "' AND not isnull([DATE_RELEASED]) ORDER BY [Version_Num] DESC;"
    Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
    If Not (myrs.BOF And myrs.EOF) Then
        myrs.MoveFirst
        retVal = retVal & myrs![Version_num]
    End If
    myrs.Close
    Set myrs = Nothing
    mydb.Close
    Set mydb = Nothing
Else
    retVal = retVal & "X"
End If
VersionNumber = retVal
SetCurrentVersion = retVal
Exit Function
err_SetCurrentVersion:
    Call General_Error_Trap
End Function
Sub SetGeneralPermissions(username, pwd, connStr)
On Error GoTo err_SetGeneralPermissions
Dim tempVal, msg, usr
Dim mydb As DAO.Database
Dim myq1 As QueryDef
    Set mydb = CurrentDb
    Set myq1 = mydb.CreateQueryDef("")
    myq1.Connect = connStr & ";UID=" & username & ";PWD=" & pwd
    myq1.ReturnsRecords = True
    myq1.sql = "sp_table_privilege_overview_for_user '%', 'dbo', null, '" & username & "'"
    Dim myrs As DAO.Recordset
    Set myrs = myq1.OpenRecordset
    If myrs.Fields(0).Value = "" Then
        tempVal = "RO"
        msg = "Your permissions on the database cannot be defined, you have been assigned READ ONLY permissions from now on." & Chr(13) & Chr(13) & "If this is incorrect please re-start the application and then if problems persist contact the Database Administrator."
    Else
        usr = UCase(myrs.Fields(0).Value)
        If InStr(usr, "RO") <> 0 Then
            tempVal = "RO"
        ElseIf InStr(usr, "ADMIN") <> 0 Then
            tempVal = "ADMIN"
        ElseIf InStr(usr, "RW") <> 0 Then
            tempVal = "RW"
        Else
            tempVal = "RO"
            msg = "The system is unsure of the rights of your login name so you have been assigned " & _
                "READ ONLY permissions on this occassion." & Chr(13) & Chr(13) & "Please contact" & _
                " the Database Administrator with the following message:" & Chr(13) & Chr(13) & "The login '" & _
                username & "' does not fall into any of the known types, please update the " & _
                "SetGeneralPermissions code"
        End If
    End If
myrs.Close
Set myrs = Nothing
myq1.Close
Set myq1 = Nothing
mydb.Close
Set mydb = Nothing
If msg <> "" Then
    MsgBox msg, vbInformation, "Permissions setup"
End If
GeneralPermissions = tempVal
Exit Sub
err_SetGeneralPermissions:
    GeneralPermissions = "RO"
    msg = "An error has occurred in the procedure: SetGeneralPermissions " & Chr(13) & Chr(13)
    msg = msg & "The system is unsure of the rights of your login name so you have been assigned " & _
                "READ ONLY permissions on this occassion." & Chr(13) & Chr(13) & "Please contact" & _
                " the Database Administrator with the following message:" & Chr(13) & Chr(13) & "The login '" & _
                username & "' does not fall into any of the known types"
    MsgBox msg, vbInformation, "Permissions setup"
    Exit Sub
End Sub
Function GetGeneralPermissions()
On Error GoTo err_GetCurrentVersion
    If GeneralPermissions = "" Then
        SetGeneralPermissions "", "", ""
    End If
    GetGeneralPermissions = GeneralPermissions
Exit Function
err_GetCurrentVersion:
    Call General_Error_Trap
End Function
