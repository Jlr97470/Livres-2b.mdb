Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim intClosing As Integer

Private Sub cmdCancel_Click()
    ' Cancel requested - undo edits and close
    If Me.Dirty Then Me.Undo
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdClose_Click()
    ' Close requested.  If changes, then save them first
    If Me.Dirty Then
        If IsNothing(Me!LastName) Then
            MsgBox "Must have an Author Name to save a record.  Click Cancel to exit without saving.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
    End If

    ' Tell AfterUpdate we're closing (don't close twice)
    intClosing = True
    DoCmd.Close acForm, Me.Name

End Sub
Private Sub cmdSave_Click()
On Error GoTo SAExit
Dim lngAuthID As Long
    ' Nothing to do if form is not "dirty"
    If Me.Dirty Then
        If IsNothing(Me!LastName) Then
            MsgBox "Record must have a valid Author Name to save it.", vbCritical, "Microsoft Press"
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
Dim strAuth As String, intComma As Integer
    ' Check for an open argument - exit if none
    If IsNothing(Me.OpenArgs) Then Exit Sub
    ' Parse out first and last names - set default values for a new row
    strAuth = Me.OpenArgs
    intComma = InStr(strAuth, ",")
    If intComma = 0 Then
        Me![LastName].DefaultValue = """" & strAuth & """"
        Else
        Me![LastName].DefaultValue = """" & Left(strAuth, intComma - 1) & """"
        Me![FirstName].DefaultValue = """" & Mid(strAuth, intComma + 2) & """"
    End If

End Sub
Public Sub FormHelp()
' Custom assistant help method for this form
Dim balCurrent As Balloon

    ' Create a new assistant baloon
    Set balCurrent = Assistant.NewBalloon

    With balCurrent
        ' Set balloon type to bullets
        .BalloonType = msoBalloonTypeBullets
        ' Set mode to auto-close as soon as the user clicks anywhere else
        .Mode = msoModeAutoDown
        ' Add an OK button just for looks
        .Button = msoButtonSetOK
        ' Define the balloon title
        .Heading = "Sample Balloon Help for Author Add"
        ' This is a tip, so use the right icon
        .Icon = msoIconTip
        ' Set up labels to display helpful hints
        .Labels(1).Text = "You must enter at least an author Last Name."
        .Labels(2).Text = "Right-click on the photo box and choose Insert Object to add the author's picture."
        .Labels(3).Text = "Click Save or Close to save the new row and exit.  Click Cancel to exit without saving."
        ' Set up title text inside the balloon
        .Text = "Use this form to add a new author!"
        ' Activate the balloon
        .Show
        ' Code will stop here until user dismisses the baloon
    End With
End Sub
