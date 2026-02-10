Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Form to add a single new category
Dim intClosing As Integer

Private Sub cmdCancel_Click()
    ' Clear any edits and close
    If Me.Dirty Then Me.Undo
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdClose_Click()
    ' Save any changes before closing
    If Me.Dirty Then
        If IsNothing(Me!ClassName) Or IsNothing(Me!Category) Then
            MsgBox "Must have a class and category to save a record.  Click Cancel to exit without saving.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
    End If

    ' Tell AfterUpdate we're closing (don't close twice)
    intClosing = True
    DoCmd.Close acForm, Me.Name

End Sub
Private Sub cmdSave_Click()
On Error GoTo SAExit
    ' Nothing to do if not "dirty"
    If Me.Dirty Then
        If IsNothing(Me!ClassName) Or IsNothing(Me!Category) Then
            MsgBox "Record must have a class and category to save it.", vbCritical, "Microsoft Press"
            Me!FirstName.SetFocus
            Exit Sub
        End If
        ' Save the current "dirty" record
        DoCmd.RunCommand acCmdSaveRecord
        Else
        Exit Sub
    End If

SAExit:
    If Err <> 0 Then
        MsgBox Error$, vbCritical
    End If
    Exit Sub


End Sub


Private Sub Form_AfterUpdate()
' Close after saving a new record - but not if close already issued
    If Not intClosing Then
        DoCmd.Close acForm, Me.Name
    End If
End Sub

Private Sub Form_Load()
    ' Check for an open argument - exit if none
    If IsNothing(Me.OpenArgs) Then Exit Sub
    ' Set Category = Openargs
    Me!Category = Me.OpenArgs

End Sub
