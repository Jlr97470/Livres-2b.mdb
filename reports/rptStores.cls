Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim intLineNo As Integer

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    intLineNo = intLineNo + 1
    If intLineNo Mod 5 = 0 Then
        Me!linSep.Visible = True
        Else
        Me!linSep.Visible = False
    End If
End Sub

Private Sub PageHeader_Format(Cancel As Integer, FormatCount As Integer)
    intLineNo = 0
End Sub
