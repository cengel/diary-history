Option Compare Database   'Use database order for string comparisons
Option Explicit
Sub EnableLists(Action)
On Error GoTo err_EnableLists
    If Action = "disable" Then
        Me![lblMsg].Visible = True
        Me![Diary_Units subform].Enabled = False
        Me![Diary_Features subform].Enabled = False
        Me![Diary_Spaces subform].Enabled = False
        Me![Diary_Buildings subform].Enabled = False
        Me![Diary_Tags_Subform].Enabled = False
        Me![Diary_Tags_Pulldown_subform1].Enabled = False
        Me![Diary_Tags_Pulldown_subform2].Enabled = False
        Me![Diary_Tags_Pulldown_subform3].Enabled = False
        Me![Diary_Tags_Pulldown_subform4].Enabled = False
        Me![Diary_Tags_Pulldown_subform5].Enabled = False
        Me![Diary_Tags_Pulldown_subform6].Enabled = False
    Else
        Me![lblMsg].Visible = False
        Me![Diary_Units subform].Enabled = True
        Me![Diary_Features subform].Enabled = True
        Me![Diary_Spaces subform].Enabled = True
        Me![Diary_Buildings subform].Enabled = True
        Me![Diary_Tags_Subform].Enabled = True
        Me![Diary_Tags_Pulldown_subform1].Enabled = True
        Me![Diary_Tags_Pulldown_subform2].Enabled = True
        Me![Diary_Tags_Pulldown_subform3].Enabled = True
        Me![Diary_Tags_Pulldown_subform4].Enabled = True
        Me![Diary_Tags_Pulldown_subform5].Enabled = True
        Me![Diary_Tags_Pulldown_subform6].Enabled = True
    End If
Exit Sub
err_EnableLists:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Command26_Click()
End Sub
Private Sub Excavation_Click()
On Error GoTo Err_Excavation_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Excavation"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Diary Form"
Exit_Excavation_Click:
    Exit Sub
Err_Excavation_Click:
    MsgBox Err.Description
    Resume Exit_Excavation_Click
End Sub
Private Sub Master_Control_Click()
On Error GoTo Err_Master_Control_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Catal Data Entry"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Diary Form"
Exit_Master_Control_Click:
    Exit Sub
Err_Master_Control_Click:
    MsgBox Err.Description
    Resume Exit_Master_Control_Click
End Sub
Sub New_Diary_Entry_Click()
On Error GoTo Err_New_Diary_Entry_Click
    DoCmd.GoToRecord , , acNewRec
Exit_New_Diary_Entry_Click:
    Exit Sub
Err_New_Diary_Entry_Click:
    MsgBox Err.Description
    Resume Exit_New_Diary_Entry_Click
End Sub
Sub Diary_Go_to_New_Click()
On Error GoTo Err_Diary_Go_to_New_Click
    DoCmd.GoToRecord , , acNewRec
Exit_Diary_Go_to_New_Click:
    Exit Sub
Err_Diary_Go_to_New_Click:
    MsgBox Err.Description
    Resume Exit_Diary_Go_to_New_Click
End Sub
Sub New_Diary_Entry2_Click()
On Error GoTo Err_New_Diary_Entry2_Click
    New_Diary_Entry_Click
Exit_New_Diary_Entry2_Click:
    Exit Sub
Err_New_Diary_Entry2_Click:
    MsgBox Err.Description
    Resume Exit_New_Diary_Entry2_Click
End Sub
Private Sub close_Click()
On Error GoTo Err_Close_Form_Click
    DoCmd.Close
Exit_Close_Form_Click:
    Exit Sub
Err_Close_Form_Click:
    MsgBox Err.Description
    Resume Exit_Close_Form_Click
End Sub
Private Sub cmdOpenSketch_Click()
On Error GoTo err_opensketch
    DoCmd.OpenForm "frm_pop_graphic", acNormal, , , acFormReadOnly, , Me![txtSketch_Name]
Exit Sub
err_opensketch:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Diary_AfterUpdate()
On Error GoTo err_diary
    EnableLists "enable"
Exit Sub
err_diary:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Field20_AfterUpdate()
On Error GoTo err_field20
    EnableLists "enable"
Exit Sub
err_field20:
    Call General_Error_Trap
    Exit Sub
End Sub
Sub find_Click()
On Error GoTo Err_find_Click
    Screen.PreviousControl.SetFocus
    Me![Diary].SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70
Exit_find_Click:
    Exit Sub
Err_find_Click:
    MsgBox Err.Description
    Resume Exit_find_Click
End Sub
Private Sub Close_Form_Click()
On Error GoTo Err_Close_Form_Click
    DoCmd.Close
Exit_Close_Form_Click:
    Exit Sub
Err_Close_Form_Click:
    MsgBox Err.Description
    Resume Exit_Close_Form_Click
End Sub
Private Sub Form_Current()
On Error GoTo err_Current
If IsNull(Me![ID]) Then
    EnableLists "disable"
Else
    EnableLists "enable"
End If
If Me![txtSketch_Name] <> "" Then
    Me![cmdOpenSketch].Enabled = True
Else
    Me![cmdOpenSketch].Enabled = False
End If
Exit Sub
err_Current:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
If DataErr = 3146 Then
    MsgBox "You must fill out all relevant information - please ensure you have filled out your name"
    Response = 2
End If
End Sub
Private Sub Form_Load()
    Me![txtReplyTo].DefaultValue = Me.OpenArgs
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_frm
    DoCmd.RunCommand acCmdRecordsGoToNew
Exit Sub
err_frm:
    If Err.Number = 2046 Then
        Resume Next
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub
Sub next_Click()
On Error GoTo Err_next_Click
    DoCmd.GoToRecord , , acNext
Exit_next_Click:
    Exit Sub
Err_next_Click:
    MsgBox Err.Description
    Resume Exit_next_Click
End Sub
Sub last_Click()
On Error GoTo Err_last_Click
    DoCmd.GoToRecord , , acLast
Exit_last_Click:
    Exit Sub
Err_last_Click:
    MsgBox Err.Description
    Resume Exit_last_Click
End Sub
Sub prev_Click()
On Error GoTo Err_prev_Click
    DoCmd.GoToRecord , , acPrevious
Exit_prev_Click:
    Exit Sub
Err_prev_Click:
    MsgBox Err.Description
    Resume Exit_prev_Click
End Sub
Sub first_Click()
On Error GoTo Err_first_Click
    DoCmd.GoToRecord , , acFirst
Exit_first_Click:
    Exit Sub
Err_first_Click:
    MsgBox Err.Description
    Resume Exit_first_Click
End Sub
Private Sub cmdSave_Click()
On Error GoTo Err_cmdSave_Click
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
Exit_cmdSave_Click:
    Exit Sub
Err_cmdSave_Click:
    MsgBox Err.Description
    Resume Exit_cmdSave_Click
End Sub
Private Sub txtSketch_Name_AfterUpdate()
On Error GoTo err_Name
If Me![txtSketch_Name] <> "" Then
    Me![cmdOpenSketch].Enabled = True
Else
    Me![cmdOpenSketch].Enabled = False
End If
Exit Sub
err_Name:
    Call General_Error_Trap
    Exit Sub
End Sub
