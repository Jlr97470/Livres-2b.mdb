Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdAddNew_Click()
    If Me.Dirty Then
        ' Check for minimum data here...
        If IsNothing(Me!Gorp) Then
            MsgBox "Must save current record with a valid ... before adding another.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
    End If
    ' Execute a Data Entry command
    DoCmd.RunCommand acCmdDataEntry
    ' Move the focus off this button
    Me![Gorp].SetFocus
    Me!cmdAddNew.Visible = False
    Me!cmdShowAll.Visible = True
    Me!cmdSearch.Enabled = False

End Sub

Private Sub cmdCancel_Click()
    ' If Me.Dirty Then Me.Undo  (Dirty valid only on bound form)
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdClose_Click()
'    If Me.Dirty Then  (Dirty valid only on bound form)
        If IsNothing(Me!Gorp) Then
            MsgBox "Must have an .... to save a record.  Click Cancel to exit without saving.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
'    End If

    DoCmd.Close acForm, Me.Name

End Sub
Private Sub cmdSave_Click()
On Error GoTo SAExit
Dim strKey As String
'    If Me.Dirty Then  (Dirty valid only on bound form)
        If IsNothing(Me!Gorp) Then
            MsgBox "Record must have a valid .... to save it.", vbCritical, "Microsoft Press"
            Me!Gorp.SetFocus
            Exit Sub
        End If
        ' Save the current "dirty" record
        DoCmd.RunCommand acCmdSaveRecord
      '  Else
        Exit Sub
'    End If

    If Me!cmdShowAll.Visible = True And Me!cmdSearch.Enabled = False Then
        Application.Echo False
        strKey = Me!Gorp
        cmdShowAll_Click
        Me!Gorp.SetFocus
'        DoCmd.FindRecord strKey
        Application.Echo True
    End If
SAExit:
    If Err <> 0 Then
        MsgBox Error$, vbCritical
    End If
    Exit Sub


End Sub

Private Sub cmdSearch_Click()
'    If Me.Dirty Then  (Dirty valid only on bound form)
        If IsNothing(Me!Gorp) Then
            MsgBox "Must save current record with a valid ... before performing a search.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
'    End If
'    DoCmd.OpenForm "frmXXXSearch"

End Sub

Private Sub cmdShowAll_Click()
On Error GoTo SHacExit
'    If Me.Dirty Then  (Dirty valid only on bound form)
        If IsNothing(Me!Gorp) Then
            MsgBox "Must save current record with a valid ... before viewing all records.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
'    End If
    DoCmd.ShowAllRecords
'    gstrWhereXXX = ""
    Me![Gorp].SetFocus
    Me!cmdAddNew.Visible = True
    Me!cmdShowAll.Visible = False
    Me!cmdSearch.Enabled = True
SHacExit:
    If Err <> 0 Then MsgBox Error$, vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
    Me.SetFocus
    DoCmd.MoveSize 0, 0
End Sub
