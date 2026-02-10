Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Close_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Details_Click()
    ' Picked one from the list, so open books filtered ...
    ' (Also activated if they double-click anywhere on the row)
    gstrWhereBook = "[ISBNNumber] = '" & Me!ISBNNumber & "'"
    DoCmd.OpenForm FormName:="frmBooks", WhereCondition:=gstrWhereBook
    DoCmd.Close acForm, Me.Name
    Forms!frmBooks.SetFocus
    Forms!frmBooks!cmdAddNew.Visible = False
    Forms!frmBooks!cmdShowAll.Visible = True

End Sub

Private Sub Disk_DblClick(Cancel As Integer)
    Details_Click
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
        .Heading = "Sample Balloon Help for Books Summary Search list."
        ' This is a tip, so use the right icon
        .Icon = msoIconTip
        ' Set up title text inside the balloon
        .Text = "Click on View Details or double-click on a book row to see details about that book."
        ' Activate the balloon
        .Show
        ' Code will stop here until user dismisses the baloon
    End With
End Sub

Private Sub EditionNumber_DblClick(Cancel As Integer)
    Details_Click
End Sub

Private Sub ISBNNumber_DblClick(Cancel As Integer)
    Details_Click
End Sub

Private Sub SuggPrice_DblClick(Cancel As Integer)
    Details_Click
End Sub

Private Sub Title_DblClick(Cancel As Integer)
    Details_Click
End Sub
