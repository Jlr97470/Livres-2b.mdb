Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub Close_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub StoreName_DblClick(Cancel As Integer)
    Details_Click
End Sub

Private Sub StoreID_DblClick(Cancel As Integer)
    Details_Click
End Sub

Private Sub Details_Click()
    gstrWhereStore = "[StoreID] = " & Me!StoreID
    DoCmd.OpenForm FormName:="frmStores", WhereCondition:=gstrWhereStore
    DoCmd.Close acForm, Me.Name
    Forms!frmStores.SetFocus
    Forms!frmStores!cmdShowAll.Enabled = True
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
        .Heading = "Sample Balloon Help for Store Summary Search list."
        ' This is a tip, so use the right icon
        .Icon = msoIconTip
        ' Set up title text inside the balloon
        .Text = "Click on View Details or double-click on a store row to see details about that store."
        ' Activate the balloon
        .Show
        ' Code will stop here until user dismisses the baloon
    End With
End Sub
