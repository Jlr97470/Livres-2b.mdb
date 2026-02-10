Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    If Me!OutOfPrint Then
        Me!OutOfPrint.Visible = True
        Else
        Me!OutOfPrint.Visible = False
    End If
End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)
    ' Don't produce group header at the top of a page (it will be a duplicate)
    If Me.Top < 2900 Then Cancel = True
End Sub
