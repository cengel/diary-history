Option Compare Database
Private Sub cmdCloseSingleEntryPopup_Click()
On Error GoTo Err_Close_Form_Click
    DoCmd.Close
Exit_Close_Form_Click:
    Exit Sub
Err_Close_Form_Click:
    MsgBox Err.Description
    Resume Exit_Close_Form_Click
End Sub
Private Sub Post_a_reply_Click()
On Error GoTo Err_Post_a_reply_Click
   DoCmd.OpenForm "Diary Reply Form", acNormal, , , , , ID
Exit_Post_a_reply_Click:
    Exit Sub
Err_Post_a_reply_Click:
    MsgBox Err.Description
    Resume Exit_Post_a_reply_Click
End Sub
