Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdReports_Click()
    On Error Resume Next
    Screen.PreviousControl.SetFocus
    DoCmd.OpenForm "frmStorePrintParms"
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
    Screen.PreviousControl.SetFocus
    DoCmd.OpenForm "frmStoreSearch"
End Sub

Private Sub cmdShowAll_Click()
On Error GoTo SHacExit
    Me![StoreID].SetFocus
    DoCmd.ShowAllRecords
    gstrWhereStore = ""
    Me!cmdShowAll.Enabled = False
    Me!cmdSearch.Enabled = True
SHacExit:
    If Err <> 0 Then MsgBox Error$, vbCritical
    Exit Sub
End Sub

Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Select Case ApplyType
        Case acApplyFilter
            Me!cmdShowAll.Enabled = True
            gstrWhereStore = Me.Filter
        Case acShowAllRecords
            Me!cmdShowAll.Enabled = False
            gstrWhereStore = ""
    End Select
End Sub

Private Sub Form_Filter(Cancel As Integer, FilterType As Integer)
    ' User about to build a filter, so clear previous
    Me.Filter = ""
    gstrWhereStore = ""
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
        .Heading = "Sample Balloon Help for Stores"
        ' This is a tip, so use the right icon
        .Icon = msoIconTip
        ' Use labels to display helpful hints
        .Labels(1).Text = "Click Search... to open the store search criteria form."
        .Labels(2).Text = "Click Reports to see available store report options."
        .Labels(3).Text = "When you have searched for or selected only some stores, click Show All to see all stores."
        ' Set up title text inside the balloon
        .Text = "Here's the things you can do on this form:"
        ' Activate the balloon
        .Show
        ' Code will stop here until user dismisses the baloon
    End With
End Sub
