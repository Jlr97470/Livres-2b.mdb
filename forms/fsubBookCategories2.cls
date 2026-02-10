Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CategoryID_NotInList(NewData As String, Response As Integer)
Dim strCat As String
Dim intReturn As Integer, varName As Variant
    strCat = NewData
    intReturn = MsgBox("Category " & strCat & _
        " is not in the system.  Do you want to add this category?", _
        vbQuestion + vbYesNo, "Microsoft Press")
    If intReturn = vbYes Then
        DoCmd.OpenForm FormName:="frmCategoryAdd", _
            DataMode:=acAdd, WindowMode:=acDialog, OpenArgs:=strCat
        If IsNull(DLookup("CategoryID", "tblCategories", "[Category] = '" & strCat & "'")) Then
            Response = acDataErrContinue
            Else
            Response = acDataErrAdded
        End If
        Exit Sub
    End If
    Response = acDataErrDisplay
End Sub
