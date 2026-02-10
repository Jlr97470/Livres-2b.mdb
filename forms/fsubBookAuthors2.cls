Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub AuthorID_NotInList(NewData As String, Response As Integer)
Dim strAuth As String, strLast As String, strFirst As String, intComma As Integer
Dim intReturn As Integer, varName As Variant
    strAuth = NewData
    intComma = InStr(strAuth, ",")
    If intComma = 0 Then
        strLast = strAuth
        Else
        strLast = Left(strAuth, intComma - 1)
        strFirst = Mid(strAuth, intComma + 2)
    End If
    intReturn = MsgBox("Author " & strAuth & _
        " is not in the system.  Do you want to add this author?", _
        vbQuestion + vbYesNo, "Microsoft Press")
    If intReturn = vbYes Then
        DoCmd.OpenForm FormName:="frmAuthorAdd", _
            DataMode:=acAdd, WindowMode:=acDialog, OpenArgs:=strAuth
        If IsNull(DLookup("AuthorID", "tblAuthors", "[LastName] = '" & strLast & "'")) Then
            Response = acDataErrContinue
            Else
            Response = acDataErrAdded
        End If
        Exit Sub
    End If
    Response = acDataErrDisplay
End Sub

Private Sub cmdDetails_Click()
    If Not IsNothing(Me!AuthorID) Then
        ' Open the authors form filtered on this author
        gstrWhereAuthor = "[AuthorID] = " & Me!AuthorID
        DoCmd.OpenForm FormName:="frmAuthors", WhereCondition:=gstrWhereAuthor
        ' Hide the New button, but show the Show All button
        Forms!frmAuthors!cmdAddNew.Visible = False
        Forms!frmAuthors!cmdShowAll.Visible = True
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
