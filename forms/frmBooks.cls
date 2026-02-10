Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdAddNew_Click()
    ' If changes on this row, then save first
    If Me.Dirty Then
        If IsNothing(Me!ISBNNumber) Or IsNothing(Me!Title) Then
            MsgBox "Must save current record with a valid ISBN and title before adding another.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
    End If
    ' Execute a Data Entry command
    DoCmd.RunCommand acCmdDataEntry
    Me![Title].SetFocus
    Me!cmdAddNew.Visible = False
    Me!cmdShowAll.Visible = True
    Me!cmdSearch.Enabled = False

End Sub

Private Sub cmdCancel_Click()
    ' Request to cancel - undo any edits first
    If Me.Dirty Then Me.Undo
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdClose_Click()
    ' Request to close - save any changes first
    If Me.Dirty Then
        If IsNothing(Me!ISBNNumber) Or IsNothing(Me!Title) Then
            MsgBox "Must have an ISBN and title to save a record.  Click Cancel to exit without saving.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
    End If

    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdOrder_Click()
    On Error Resume Next
    Screen.PreviousControl.SetFocus
    ' Call the common "add to open order" function
    If OrderThis(Me!ISBNNumber) Then
        MsgBox "Book was successfully added to your most current open order.", vbInformation, "Microsoft Press"
        If IsLoaded("frmOrders") Then
            Forms!frmOrders!fsubOrderBooks.Form.Requery
        End If
        Else
        MsgBox "Book was not ordered.", vbInformation, "Microsoft Press"
    End If
End Sub

Private Sub cmdReports_Click()
    On Error Resume Next
    Screen.PreviousControl.SetFocus
    ' Save any changes first
    If Me.Dirty Then
        If IsNothing(Me!ISBNNumber) Or IsNothing(Me!Title) Then
            MsgBox "Record must have a valid ISBN and Title to save it.", vbCritical, "Microsoft Press"
            Me!ISBNNumber.SetFocus
            Exit Sub
        End If
        ' Save the current "dirty" record
        DoCmd.RunCommand acCmdSaveRecord
    End If
    ' Open the report options form...
    DoCmd.OpenForm "frmBookPrintParms"
End Sub

Private Sub cmdSave_Click()
On Error GoTo SAExit
Dim strISBN As String
    ' If not "dirty" then nothing to do
    If Me.Dirty Then
        If IsNothing(Me!ISBNNumber) Or IsNothing(Me!Title) Then
            MsgBox "Record must have a valid ISBN and Title to save it.", vbCritical, "Microsoft Press"
            Me!ISBNNumber.SetFocus
            Exit Sub
        End If
        ' Save the current "dirty" record
        DoCmd.RunCommand acCmdSaveRecord
        Else
        Exit Sub
    End If

    ' If we were in "add new" mode...
    If Me!cmdShowAll.Visible = True And Me!cmdSearch.Enabled = False Then
        Application.Echo False
        ' Save ISBN of new row
        strISBN = Me!ISBNNumber
        ' .. reveal all rows
        cmdShowAll_Click
        Me!ISBNNumber.SetFocus
        ' .. and then move to newly added row
        DoCmd.FindRecord strISBN
        Application.Echo True
    End If
SAExit:
    If Err <> 0 Then
        MsgBox Error$, vbCritical
    End If
    Exit Sub


End Sub

Private Sub cmdSearch_Click()
    ' Save any changes first
    If Me.Dirty Then
        If IsNothing(Me!ISBNNumber) Or IsNothing(Me!Title) Then
            MsgBox "Must save current record with a valid ISBN and title before performing a search.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
    End If
    ' Open the search form
    DoCmd.OpenForm "frmBookSearch"

End Sub

Private Sub cmdShowAll_Click()
On Error GoTo SHacExit
    ' Save any changes first
    If Me.Dirty Then
        If IsNothing(Me!ISBNNumber) Or IsNothing(Me!Title) Then
            MsgBox "Must save current record with a valid ISBN and Title before viewing all books.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
    End If
    Me!ISBNNumber.SetFocus
    ' Reveal all rows
    DoCmd.ShowAllRecords
    ' Clear the global filter string
    gstrWhereBook = ""
    Me![ISBNNumber].SetFocus
    If gintIsAdmin Then
        Me!cmdAddNew.Visible = True
        Else
        Me!cmdAddNew.Visible = False
    End If
    Me!cmdShowAll.Visible = False
    Me!cmdSearch.Enabled = True
SHacExit:
    If Err <> 0 Then MsgBox Error$, vbCritical
    Exit Sub
End Sub

Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    ' Filter has been applied - save it in the global filter string
    gstrWhereBook = Me.Filter
    If IsNothing(Me.Filter) Or ApplyType = acShowAllRecords Then
    ' Nothing in the filter, so set up like they clicked Show All
        gstrWhereBook = ""
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
    gstrWhereBook = ""
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
        .Heading = "Sample Balloon Help for Books"
        ' This is a tip, so use the right icon
        .Icon = msoIconTip
        ' Set up labels to display helpful hints
        .Labels(1).Text = "Click Search to look for a different book."
        .Labels(2).Text = "Select a tab to see details about book categories, authors, or liner notes."
        .Labels(3).Text = "When viewing authors, click Details next to an author to see info about that author."
        .Labels(4).Text = "Click the Reports button to see book report options."
        .Labels(5).Text = "Click Order This to add this book to your latest open order."
        ' Set up title text inside the balloon
        .Text = "You can do lots of things with the Books form!"
        ' Activate the balloon
        .Show
        ' Code will stop here until user dismisses the baloon
    End With
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
        Me!fsubBookCategories.Locked = True
        Me!fsubBookAuthors.Locked = True
    End If
End Sub
