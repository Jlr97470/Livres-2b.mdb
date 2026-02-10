Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdAll_Click()
    ' Asked to review all - clear the global filter string
    gstrWhereBook = ""
    ' .. and open form on all books
    DoCmd.OpenForm "frmBooks"
    DoCmd.Close acForm, Me.Name
    Forms!frmBooks.SetFocus
End Sub

Private Sub cmdSome_Click()
Dim strWhere As String, varItem As Variant
    ' If no items selected, then nothing to do
    If Me!lstBName.ItemsSelected.Count = 0 Then Exit Sub
    ' Loop through the items selected collection
    For Each varItem In Me!lstBName.ItemsSelected
        ' Grab the ISBN column for each selected item
        strWhere = strWhere & Chr$(39) & Me!lstBName.Column(0, varItem) & Chr$(39) & ","
    Next varItem
    ' Throw away the extra comma on the "IN" string
    strWhere = Left$(strWhere, Len(strWhere) - 1)
    ' Open the books form filtered on the selected books
    gstrWhereBook = "[ISBNNumber] IN (" & strWhere & ")"
    DoCmd.OpenForm FormName:="frmBooks", WhereCondition:=gstrWhereBook
    ' Hide the New button, but show the Show All button
    Forms!frmBooks!cmdAddNew.Visible = False
    Forms!frmBooks!cmdShowAll.Visible = True
    DoCmd.Close acForm, Me.Name
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
        .Heading = "Sample Balloon Help for Book List"
        ' This is a tip, so use the right icon
        .Icon = msoIconTip
        ' Show help based on Admin status
        If gintIsAdmin Then
            ' Use labels to display helpful hints
            .Labels(1).Text = "Click Edit to edit books you have chosen in the list."
            .Labels(2).Text = "Click Edit All to edit all book records."
            .Labels(3).Text = "Click Search... to open the book search criteria form."
            .Labels(4).Text = "Click New to add a new book."
        Else
            .Labels(1).Text = "Click View to view books you have chosen in the list."
            .Labels(2).Text = "Click View All to view all book records."
            .Labels(3).Text = "Click Search... to open the book search criteria form."
        End If
        ' Set up title text inside the balloon
        .Text = "Choose one or more books, then select the option you want."
        ' Activate the balloon
        .Show
        ' Code will stop here until user dismisses the baloon
    End With
End Sub

Private Sub Form_Load()
    ' auto-select the first item in the list box
    ' Note that this list box has column headings, so
    '  the "first" item is 1, not 0.
    Me!lstBName.Selected(1) = True
    ' Check for admin status
    ' Check for admin status
    If gintIsAdmin Then
        Me!cmdNew.Visible = True
        Me!cmdSome.Caption = "&Edit"
        Me!cmdSome.ControlTipText = "Edit selected book(s)"
        Me!cmdAll.Caption = "Edit &All"
        Me!cmdAll.ControlTipText = "Edit all books"
    End If
End Sub

Private Sub lstBName_DblClick(Cancel As Integer)
    ' Double-click on an item is the same as selecting it
    ' and clicking Edit
    cmdSome_Click
End Sub

Private Sub cmdNew_Click()
Dim frm As Form
    ' This button disabled if not an admin user
    Me.Visible = False
    ' Open books in data entry
    DoCmd.OpenForm FormName:="frmBooks", DataMode:=acAdd
    Set frm = Forms!frmBooks
    frm.SetFocus
    frm![ISBNNumber].SetFocus
    ' .. and set some options
    frm!cmdAddNew.Visible = False
    frm!cmdShowAll.Visible = True
    frm!cmdSearch.Enabled = False
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    ' Put focus on the main switchboard - about to open a modal form
    Forms!frmMain.SetFocus
    ' .. close me
    DoCmd.Close acForm, Me.Name
    ' .. open the search form
    DoCmd.OpenForm "frmBookSearch"
End Sub
