Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub AuthorID_Enter()
' The following GoToPage realigns the page when you tab forward
'  from the previous record.  Has no effect otherwise.
    Me.GoToPage 1
End Sub

Private Sub Bio_Enter()
' The following GoToPage realigns the page when you tab backward
'  from the next record.  It is also activated when you key PageDn
'  in the subform control on the previous page.  The KeyDown event
'  puts the focus here, and this aligns the page correctly.
    Me.GoToPage 2
End Sub

Private Sub cmdAddNew_Click()
    ' This button not enabled if not an admin user
    ' First, put focus back where it was
    On Error Resume Next  ' Avoid errors
    Screen.PreviousControl.SetFocus
    On Error GoTo 0
    ' Request to add a new record - check for changes first
    If Me.Dirty Then
        If IsNothing(Me!LastName) Then
            MsgBox "Must save current record with a valid author name before adding another.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        ' Save the record if there are changes
        DoCmd.RunCommand acCmdSaveRecord
    End If
    ' Execute a Form, Records, Data Entry command
    DoCmd.RunCommand acCmdDataEntry
    ' Put focus on FirstName and set some options
    Me![FirstName].SetFocus
    Me!cmdAddNew.Visible = False
    Me!cmdShowAll.Visible = True
    Me!cmdSearch.Enabled = False
    Me!cmdSave.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    ' Request to bail - clear any edits and close
    If Me.Dirty Then Me.Undo
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdClose_Click()
    ' Request to close - save any changes first
    If Me.Dirty Then
        If IsNothing(Me!LastName) Then
            MsgBox "Must have an Author Name to save a record.  Click Cancel to exit without saving.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
    End If

    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdSave_Click()
On Error GoTo SAExit
Dim lngAuthID As Long
    ' First, put focus back where it was
    On Error Resume Next  ' Avoid errors
    Screen.PreviousControl.SetFocus
    On Error GoTo 0
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
    ' If we were in "add new" mode, then
    If Me!cmdShowAll.Visible = True And Me!cmdSearch.Enabled = False Then
        Application.Echo False
        ' Save the newly added author
        lngAuthID = Me!AuthorID
        cmdShowAll_Click
        Me!AuthorID.SetFocus
        ' Put focus back on new row after doing a "show all"
        DoCmd.FindRecord lngAuthID
        Application.Echo True
    End If
SAExit:
    If Err <> 0 Then
        MsgBox Error$, vbCritical
    End If
    Exit Sub


End Sub

Private Sub cmdSearch_Click()
    ' First, put focus back where it was
    On Error Resume Next  ' Avoid errors
    Screen.PreviousControl.SetFocus
    On Error GoTo 0
    ' Request to search - save any changes first
    If Me.Dirty Then
        If IsNothing(Me!LastName) Then
            MsgBox "Must save current record with a valid author name before performing a search.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        cmdSave_Click
    End If
    DoCmd.OpenForm "frmAuthorSearch"

End Sub

Private Sub cmdShowAll_Click()
On Error GoTo SHacExit
    ' Going to remove filters - save changes first
    If Me.Dirty Then
        If IsNothing(Me!LastName) Then
            MsgBox "Must save current record with a valid author name before viewing all authors.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        cmdSave_Click
    End If
    ' Move focus to avoid an error setting my property
    Me![FirstName].SetFocus
    ' Remove any filter and reset the global filter string
    DoCmd.ShowAllRecords
    gstrWhereAuthor = ""
    ' Show the Add button only if Admin user
    If gintIsAdmin Then Me!cmdAddNew.Visible = True
    Me!cmdShowAll.Visible = False
    Me!cmdSearch.Enabled = True
    Me!cmdSave.Enabled = True
SHacExit:
    If Err <> 0 Then MsgBox Error$, vbCritical
    Exit Sub
End Sub

Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    ' Filter has been applied - save it in the global filter string
    gstrWhereAuthor = Me.Filter
    If IsNothing(Me.Filter) Or ApplyType = acShowAllRecords Then
    ' Nothing in the filter, so set up like they clicked Show All
        gstrWhereAuthor = ""
        ' Show the Add button only if Admin user
        If gintIsAdmin Then Me!cmdAddNew.Visible = True
        Me!cmdShowAll.Visible = False
        Me!cmdSearch.Enabled = True
        Me!cmdSave.Enabled = True
        Else
        Me!cmdAddNew.Visible = False
        Me!cmdShowAll.Visible = True
    End If
End Sub

Private Sub Form_Filter(Cancel As Integer, FilterType As Integer)
    ' User about to build a filter, so clear previous
    Me.Filter = ""
    gstrWhereAuthor = ""
End Sub

Private Sub Form_Load()
    ' Turn off stuff if this is not an admin user
    If Not gintIsAdmin Then
        Me.AllowEdits = False
        Me.AllowDeletions = False
        Me.AllowAdditions = False
        Me!cmdAddNew.Visible = False
        Me!cmdAddNew.Enabled = False
        Me!cmdSave.Visible = False
        Me!fsubAuthorBooks.Locked = True
    End If
End Sub

Private Sub fsubAuthorBooks_Enter()
On Error Resume Next
    ' If tabbing into fsubAuthorBooks from control "Bio"
    If Screen.PreviousControl.Name = "Bio" Then
        ' Do some funny stuff to align the "page" correctly
        Application.Echo False
        Me!AuthorID.SetFocus
        Me.GoToPage 1
        Me!fsubAuthorBooks.SetFocus
        Application.Echo True
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
        .Heading = "Sample Balloon Help for Authors"
        ' This is a tip, so use the right icon
        .Icon = msoIconTip
        ' Set up labels to display helpful hints
        .Labels(1).Text = "Click Search to look for a different author."
        .Labels(2).Text = "Click Details next to a book to see info about that book."
        .Labels(3).Text = "Select a book and click " & vbCrLf & "Order > Selected Book to add the book to your latest open order."
        .Labels(4).Text = "Press Page Down to see author bio information."
        ' Set up title text inside the balloon
        .Text = "You can do lots of things with the Authors form!"
        ' Activate the balloon
        .Show
        ' Code will stop here until user dismisses the baloon
    End With
End Sub
