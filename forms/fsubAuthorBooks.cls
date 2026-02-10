Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdDetails_Click()
Dim frm As Form
    If Not IsNothing(Me!ISBNNumber) Then
        ' Open the books form filtered on this book
        gstrWhereBook = "[ISBNNumber] = '" & Me!ISBNNumber & "'"
        DoCmd.OpenForm FormName:="frmBooks", WhereCondition:=gstrWhereBook
        ' Set up an object variable for efficiency
        Set frm = Forms!frmBooks
        ' Hide the New button, but show the Show All button
        frm!cmdAddNew.Visible = False
        frm!cmdShowAll.Visible = True
        ' Turn on the author display option
        frm!tabDisplay = 1
        ' and set to show the author info...
    End If

End Sub

Private Sub cmdOrder_Click()
    On Error Resume Next
    Me!ISBNNumber.SetFocus
    ' Call the common "add to open order" function
    If OrderThis(Me!ISBNNumber) Then
        MsgBox "Book " & Me!Title & " was successfully added to your most current open order.", vbInformation, "Microsoft Press"
        Else
        MsgBox "Book was not ordered.", vbInformation, "Microsoft Press"
    End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Dim varSeq As Variant
    If IsNothing(Me!SeqNo) Then
        varSeq = DMax("[SeqNo]", "tblBookAuthor", "[ISBNNumber] = '" & Me![ISBNNumber] & "'")
        If IsNull(varSeq) Then
            Me!SeqNo = 1
            Else
            Me!SeqNo = varSeq + 1
        End If
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' Trap all page down keystrokes.  Normally, a page down would do
' nothing while on the subform.  We're cancelling the keystroke
' and putting the focus on the first control on the second page
' of the outer form.  The Enter event of that control does a GoToPage
' to align the page correctly (you can't do GoToPage for the outer
' form from here).
    If KeyCode = 34 Then
        KeyCode = 0
        Me.Parent.SetFocus
        Me.Parent!Bio.SetFocus
    End If
End Sub
