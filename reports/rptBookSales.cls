Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim intLinCount As Integer

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' Draw a line every 5 lines
    intLinCount = intLinCount + 1
    If intLinCount Mod 5 = 0 Then
        Me!linUp.Visible = True
        Else
        Me!linUp.Visible = False
    End If

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)
' Reset the line counter at the top of this section
    intLinCount = 0
' Don't produce this section when near the top of a page
' This code works as long as top margin is 1" and Page Header is 1.5" (3600 twips)
    If Me.Top < 3650 Then Cancel = True

End Sub

Private Sub PageHeader_Format(Cancel As Integer, FormatCount As Integer)
    ' Reset the line counter at the top of each page
    intLinCount = 0
End Sub
