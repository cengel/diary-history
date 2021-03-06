Option Compare Database
Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_Click
Dim retVal
retVal = MsgBox("The system cannot continue without a login name and password." & Chr(13) & Chr(13) & "Are you sure you want to quit the system?", vbCritical + vbYesNo, "Confirm System Closure")
    If retVal = vbYes Then
        MsgBox "The system will now quit", vbCritical + vbOKOnly, "Invalid Login"
        DoCmd.Quit acQuitSaveAll
    End If
    DoCmd.GoToControl "txtLogin"
Exit Sub
cmdCancel_Click:
    Call General_Error_Trap
End Sub
Private Sub cmdOK_Click()
On Error GoTo cmdOK_Click
Dim retVal
If IsNull(Me![txtLogin]) Or IsNull(Me![txtPwd]) Then
    retVal = MsgBox("Sorry but the system cannot continue without both a login name and a password. Do you want to try again?", vbCritical + vbYesNo, "Login required")
    If retVal = vbYes Then 'try again
        DoCmd.GoToControl "txtLogin"
        Exit Sub
    Else 'no, don't try again so quit system
        retVal = MsgBox("The system cannot continue without a login name and password." & Chr(13) & Chr(13) & "Are you sure you want to quit the system?", vbCritical + vbYesNo, "Confirm System Closure")
        If retVal = vbYes Then
            MsgBox "The system will now quit", vbCritical + vbOKOnly, "Invalid Login"
            DoCmd.Quit acQuitSaveAll
        Else 'no I don't want to quit system, ie: try again
            DoCmd.GoToControl "txtLogin"
            Exit Sub
        End If
    End If
Else
    If LogUserIn(Me![txtLogin], Me![txtPwd]) = True Then
        DoCmd.Close acForm, "Login" 'shut form as modal
    Else
    End If
End If
Exit Sub
cmdOK_Click:
    Call General_Error_Trap
    DoCmd.Close acForm, "Excavation_Login" 'this may be better as a simply quit the system, will see, however must shut form as modal
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdOK_Click
End If
End Sub
