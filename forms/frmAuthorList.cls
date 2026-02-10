Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub Close_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdAll_Click()
    ' Edit all - clear the filter string
    gstrWhereAuthor = ""
    ' .. and open authors unfiltered
    DoCmd.OpenForm "frmAuthors"
    DoCmd.Close acForm, Me.Name
    Forms!frmAuthors.SetFocus
End Sub

Private Sub cmdSome_Click()
Dim strWhere As String, varItem As Variant
    ' Request to edit items selected in the list box
    ' If no items selected, then nothing to do
    If Me!lstAName.ItemsSelected.Count = 0 Then Exit Sub
    ' Loop through the items selected collection
    For Each varItem In Me!lstAName.ItemsSelected
        ' Grab the AuthorID column for each selected item
        strWhere = strWhere & Me!lstAName.Column(0, varItem) & ","
    Next varItem
    ' Throw away the extra comma on the "IN" string
    strWhere = Left$(strWhere, Len(strWhere) - 1)
    ' Open the authors form filtered on the selected authors
    gstrWhereAuthor = "[AuthorID] IN (" & strWhere & ")"
    DoCmd.OpenForm FormName:="frmAuthors", WhereCondition:=gstrWhereAuthor
    ' Hide the New button, but show the Show All button
    Forms!frmAuthors!cmdAddNew.Visible = False
    Forms!frmAuthors!cmdShowAll.Visible = True
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Load()
    ' auto-select the first item in the list box
    Me!lstAName.Selected(0) = True
    ' Check for admin status
    If gintIsAdmin Then
        Me!cmdNew.Visible = True
        Me!cmdSome.Caption = "&Edit"
        Me!cmdSome.ControlTipText = "Edit selected author(s)"
        Me!cmdAll.Caption = "Edit &All"
        Me!cmdAll.ControlTipText = "Edit all authors"
    End If
End Sub

Private Sub lstAName_DblClick(Cancel As Integer)
    ' Double-clicking a name is the same as choosing the name
    ' and then clicking the edit button.
    cmdSome_Click
End Sub

Private Sub cmdNew_Click()
Dim frm As Form
    ' This button not enabled if not an admin user
    Me.Visible = False
    ' Open Authors in data entry and set some options
    DoCmd.OpenForm FormName:="frmAuthors", DataMode:=acAdd
    Set frm = Forms!frmAuthors
    frm.SetFocus
    frm![FirstName].SetFocus
    frm!cmdAddNew.Visible = False
    frm!cmdShowAll.Visible = True
    frm!cmdSearch.Enabled = False
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
    ' Put focus on main form because we're opening
    ' a modal form
    Forms!frmMain.SetFocus
    ' Close me
    DoCmd.Close acForm, Me.Name
    ' .. and open the modal form
    DoCmd.OpenForm "frmAuthorSearch"
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
        .Heading = "Sample Balloon Help for Author List"
        ' This is a tip, so use the right icon
        .Icon = msoIconTip
        ' Show help based on Admin status
        If gintIsAdmin Then
            ' Use labels to display helpful hints
            .Labels(1).Text = "Click Edit to edit authors you have chosen in the list."
            .Labels(2).Text = "Click Edit All to edit all author records."
            .Labels(3).Text = "Click Search... to open the author search criteria form."
            .Labels(4).Text = "Click New to add a new author."
        Else
            .Labels(1).Text = "Click View to view authors you have chosen in the list."
            .Labels(2).Text = "Click View All to view all author records."
            .Labels(3).Text = "Click Search... to open the author search criteria form."
        End If
        ' Set up title text inside the balloon
        .Text = "Choose one or more authors, then select the option you want."
        ' Activate the balloon
        .Show
        ' Code will stop here until user dismisses the baloon
    End With
End Sub
