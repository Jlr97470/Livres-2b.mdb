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
    gstrWhereStore = ""
    DoCmd.OpenForm "frmStores"
    DoCmd.Close acForm, Me.Name
    Forms!frmStores.SetFocus
End Sub

Private Sub cmdView_Click()
Dim strWhere As String, varItem As Variant
    ' If no items selected, then nothing to do
    If Me!lstSName.ItemsSelected.Count = 0 Then Exit Sub
    ' Loop through the items selected collection
    For Each varItem In Me!lstSName.ItemsSelected
        ' Grab the StoreID column for each selected item
        strWhere = strWhere & Me!lstSName.Column(0, varItem) & ","
    Next varItem
    ' Throw away the extra comma on the "IN" string
    strWhere = Left$(strWhere, Len(strWhere) - 1)
    ' Open the stores form filtered on the selected stores
    gstrWhereStore = "[StoreID] IN (" & strWhere & ")"
    DoCmd.OpenForm FormName:="frmStores", WhereCondition:=gstrWhereStore
    ' Enable the Show All button
    Forms!frmStores!cmdShowAll.Enabled = True
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
        .Heading = "Sample Balloon Help for Store List"
        ' This is a tip, so use the right icon
        .Icon = msoIconTip
        ' Show help based on Admin status
        .Labels(1).Text = "Click View Selected to view stores you have chosen in the list."
        .Labels(2).Text = "Click View All to view all stores."
        .Labels(3).Text = "Click Search... to open the store search criteria form."
        ' Set up title text inside the balloon
        .Text = "Choose one or more stores, then select the option you want."
        ' Activate the balloon
        .Show
        ' Code will stop here until user dismisses the baloon
    End With
End Sub
Private Sub Form_Load()
    ' auto-select the first item in the list box
    Me!lstSName.Selected(0) = True
End Sub

Private Sub lstSName_DblClick(Cancel As Integer)
    cmdView_Click
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
    Forms!frmMain.SetFocus
    DoCmd.Close acForm, Me.Name
    DoCmd.OpenForm "frmStoreSearch"
End Sub
