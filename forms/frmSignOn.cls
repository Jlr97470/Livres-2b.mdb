Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmbUserID_NotInList(NewData As String, Response As Integer)
Dim strUser As String, strLast As String, strFirst As String, intComma As Integer
Dim intReturn As Integer, varName As Variant
    ' The following code handles a new name --
    ' It "parses" the name entered and offers to let the user add themselves as a new customer
    strUser = NewData
    intComma = InStr(strUser, ",")
    If intComma = 0 Then
        strLast = strUser
        Else
        strLast = Left(strUser, intComma - 1)
        strFirst = Mid(strUser, intComma + 2)
    End If
    intReturn = MsgBox("Customer " & strUser & _
        " is not in the system.  Do you want to add this customer?", _
        vbQuestion + vbYesNo, "Microsoft Press")
    If intReturn = vbYes Then
        DoCmd.OpenForm FormName:="frmCustAdd", _
            DataMode:=acAdd, WindowMode:=acDialog, OpenArgs:=strUser
        If IsNull(DLookup("CustomerID", "tblCustomers", "[LastName] = '" & strLast & "'")) Then
            Response = acDataErrContinue
            Else
            Response = acDataErrAdded
        End If
        Exit Sub
    End If
    Response = acDataErrDisplay

End Sub

Private Sub Form_Load()
    ' Make sure current customer is reset when this form loads
    glngCurrentCustID = 0
    gintPrintPreview = acPreview
    gintIsAdmin = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Form_Unload_Err

    ' Final check when form is closing...
    If glngCurrentCustID = 0 Then
        ' If userid not set, then quit
        DoCmd.SelectObject acTable, "tblAuthors", True
        gintPrintPreview = acPreview
    End If


Form_Unload_Exit:
    Exit Sub

Form_Unload_Err:
    MsgBox Error$
    Resume Form_Unload_Exit

End Sub

Private Sub cmdSignOn_Click()
On Error GoTo cmdSignOn_Click_Err
Static intPswdCount As Integer
    If Me!txtPassword = Me![cmbUserID].Column(2) Or Me![cmbUserID].Column(2) = "" Then
        Else
        MsgBox "Password you entered is incorrect.", vbCritical, "Microsoft Press"
        intPswdCount = intPswdCount + 1
        If intPswdCount < 3 Then Exit Sub
        Me.Undo
        DoCmd.Close acForm, Me.Name
        DoCmd.SelectObject acTable, "tblAuthors", True
        Exit Sub
    End If

    Me.Visible = False
    ' Copy selected customer ID to Public variable
    glngCurrentCustID = Me!cmbUserID.Column(1)
    ' Copy admin status to Public variable
    gintIsAdmin = Me!cmbUserID.Column(3)
    ' Set the global print option
    If Me!PrintPreview Then
        gintPrintPreview = acPreview
        Else
        gintPrintPreview = acNormal
    End If
    ' If not called with an open parameter
    If IsNothing(Me.OpenArgs) Then
        ' .. then open the main switchboard form
        DoCmd.OpenForm "frmMain"
    End If
    ' .. close this form
    DoCmd.Close acForm, Me.Name

cmdSignOn_Click_Exit:
    Exit Sub

cmdSignOn_Click_Err:
    MsgBox "Error during Sign On: " & Err & ", " & Error
    Resume cmdSignOn_Click_Exit

End Sub


Private Sub cmdExit_Click()
On Error GoTo cmdExit_Click_Err

    ' Exit button - Close the form
    DoCmd.Close acForm, Me.Name
    DoCmd.SelectObject acTable, "tblAuthors", True

cmdExit_Click_Exit:
    Exit Sub

cmdExit_Click_Err:
    MsgBox Error$
    Resume cmdExit_Click_Exit

End Sub
