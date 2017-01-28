Option Compare Database
Option Explicit
Function LogUserIn(username As String, pwd As String)
On Error GoTo err_LogUserIn
Dim retVal
If username <> "" And pwd <> "" Then
    Dim mydb As DAO.Database, I, errmsg, connStr
    Dim tmptable As TableDef
    Set mydb = CurrentDb
    Dim myq As QueryDef
    Set myq = mydb.CreateQueryDef("")
    connStr = ""
    For I = 0 To mydb.TableDefs.Count - 1 'loop the tables collection
         Set tmptable = mydb.TableDefs(I)
        If tmptable.Connect <> "" Then
            If connStr = "" Then connStr = tmptable.Connect
            On Error Resume Next
                myq.Connect = tmptable.Connect & ";UID=" & username & ";PWD=" & pwd
                myq.ReturnsRecords = False 'don't waste resources bringing back records
                myq.sql = "select [ID] from [Diary_Table] WHERE [ID] = 1" 'this is a shared and core table so should always be avail, the record doesn't have to exist
                myq.Execute
            If Err <> 0 Then 'the login deails are incorrect
                GoTo err_LogUserIn
            Else
                On Error GoTo err_LogUserIn:
                tmptable.Connect = tmptable.Connect & ";UID=" & username & ";PWD=" & pwd
                tmptable.RefreshLink
            End If
            Exit For 'only necessary for one table for Access to set up the correct link to SQL Server
        End If
    Next I
Else
    MsgBox "Both a username and password are required to operate the system correctly. Please quit and restart the application.", vbCritical, "Login problem encountered"
End If
LogUserIn = True
cleanup:
    myq.Close
    Set myq = Nothing
    Set tmptable = Nothing
    mydb.Close
    Set mydb = Nothing
Exit Function
err_LogUserIn:
    If Err.Number = 3059 Or Err.Number = 3151 Then
        errmsg = "Sorry but the system cannot log you into the database. There are three reasons this may have occurred:" & Chr(13) & Chr(13)
        errmsg = errmsg & "1. Your login details have been entered incorrectly" & Chr(13) & Chr(13)
        errmsg = errmsg & "2. There is no ODBC connection to the database setup on this computer." & Chr(13) & "    See http://www.catalhoyuk.com/database/odbc.html for details." & Chr(13) & Chr(13)
        errmsg = errmsg & "3. Your computer is not connected to the Internet at this time." & Chr(13) & Chr(13)
        errmsg = errmsg & "Do you wish to try logging in again?"
        retVal = MsgBox(errmsg, vbCritical + vbYesNo, "Login Failure")
        If retVal = vbYes Then
            GoTo cleanup 'used to be resume before querydef intro, now just cleanup and leave so user can try again
        Else
            retVal = MsgBox("Are you really sure you want to quit and close the system?", vbCritical + vbYesNo, "Confirm System Closure")
            If retVal = vbNo Then
                GoTo cleanup 'on 2nd thoughts the user doesn't want to quit so now just cleanup and leave so user can try again
            Else
                MsgBox "The system will now quit" & Chr(13) & Chr(13) & "The error reported was: " & Err.Description, vbCritical, "Login Failure"
            End If
        End If
    Else
        MsgBox Err.Description & Chr(13) & Chr(13) & "The system will now quit", vbCritical, "Login Failure"
    End If
    LogUserIn = False
    DoCmd.Quit acQuitSaveAll
End Function
