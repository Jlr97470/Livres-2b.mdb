Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub Close_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdAuthors_Click()
    DoCmd.OpenReport ReportName:="rptAuthors", _
        View:=gintPrintPreview
    DoCmd.Close acForm, Me.Name
    ' If this user has print preview option set,
    If gintPrintPreview = acPreview Then
        ' Then make sure focus goes to the report window
        DoCmd.SelectObject acReport, "rptAuthors"
    End If
End Sub

Private Sub cmdBooks_Click()
    DoCmd.OpenReport ReportName:="rptBooks", _
        View:=gintPrintPreview
    DoCmd.Close acForm, Me.Name
    ' If this user has print preview option set,
    If gintPrintPreview = acPreview Then
        ' Then make sure focus goes to the report window
        DoCmd.SelectObject acReport, "rptBooks"
    End If
End Sub

Private Sub cmdBookSales_Click()
On Error Resume Next
    Forms!frmMain.SetFocus
    DoCmd.OpenForm "frmBookSalesParm"
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdCustomers_Click()
    DoCmd.OpenReport ReportName:="rptCustomers", _
        View:=gintPrintPreview
    DoCmd.Close acForm, Me.Name
    ' If this user has print preview option set,
    If gintPrintPreview = acPreview Then
        ' Then make sure focus goes to the report window
        DoCmd.SelectObject acReport, "rptCustomers"
    End If
End Sub


Private Sub cmdCustSales_Click()
On Error Resume Next
    Forms!frmMain.SetFocus
    DoCmd.OpenForm "frmCustSalesParm"
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdStores_Click()
    DoCmd.OpenReport ReportName:="rptStores", _
        View:=gintPrintPreview
    DoCmd.Close acForm, Me.Name
    ' If this user has print preview option set,
    If gintPrintPreview = acPreview Then
        ' Then make sure focus goes to the report window
        DoCmd.SelectObject acReport, "rptStores"
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
        .Heading = "Sample Balloon Help for Reports"
        ' This is a tip, so use the right icon
        .Icon = msoIconTip
        ' Set up title text inside the balloon
        .Text = "Click on a button to see reports available for that category."
        ' Activate the balloon
        .Show
        ' Code will stop here until user dismisses the baloon
    End With
End Sub
