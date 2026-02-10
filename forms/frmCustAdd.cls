Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim intClosing As Integer

Private Sub cmdCancel_Click()
    ' Cancel request - clear all edits before closing
    If Me.Dirty Then Me.Undo
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdClose_Click()
    ' Close - save any edits first
    If Me.Dirty Then
        If IsNothing(Me!FirstName) Or IsNothing(Me!LastName) Then
            MsgBox "Must have a first and last name to save a record.  Click Cancel to exit without saving.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
    End If

    ' Tell AfterUpdate we're closing (don't close twice)
    intClosing = True
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdSave_Click()
On Error GoTo SAExit
    ' If not "dirty" then nothing to do
    If Me.Dirty Then
        If IsNothing(Me!FirstName) Or IsNothing(Me!LastName) Then
            MsgBox "Record must have a first and last name to save it.", vbCritical, "Microsoft Press"
            Me!FirstName.SetFocus
            Exit Sub
        End If
        ' Save the current "dirty" record
        DoCmd.RunCommand acCmdSaveRecord
        Else
        Exit Sub
    End If

SAExit:
    ' Skip actions canceled by Before Update
    If Err = 2501 Then Exit Sub
    If Err <> 0 Then
        MsgBox Error$, vbCritical
    End If
    Exit Sub


End Sub


Private Sub Form_AfterUpdate()
' Close after saving a new record - but not if close already issued
    If Not intClosing Then
        DoCmd.Close acForm, Me.Name
    End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Dim intPLen As Integer
    ' Customer records with a credit card number or admin status
    '  must also have a password
    If Not IsNothing(Me!CCNumber) Or Me!Admin Then
        intPLen = Len(Nz(Me!Password))
        If intPLen < 5 Then
            MsgBox "A customer with a credit card number or admin status must also " & _
                " have a password." & vbCrLf & vbCrLf & _
                "Enter a password at least 5 characters long or clear the credit " & _
                "card number and admin status.", _
                vbExclamation, "Microsoft Press"
            Cancel = True
        End If
    End If

End Sub

Private Sub Form_Load()
Dim strUser As String, intComma As Integer, strFirst As String, strMid As String
    ' Check for an open argument - exit if none
    If IsNothing(Me.OpenArgs) Then Exit Sub
    ' Parse out first and last names - set default values for a new row
    strUser = Me.OpenArgs
    intComma = InStr(strUser, ",")
    If intComma = 0 Then
        Me![LastName].DefaultValue = """" & strUser & """"
        Else
        Me![LastName].DefaultValue = """" & Left(strUser, intComma - 1) & """"
        strFirst = Mid(strUser, intComma + 2)
        ' Check for a middle initial
        If Right(strFirst, 1) = "." And Mid(strFirst, Len(strFirst) - 2, 1) = " " Then
            strMid = Mid(strFirst, Len(strFirst) - 1, 1)
            strFirst = Left(strFirst, Len(strFirst) - 3)
            Me![FirstName].DefaultValue = """" & strFirst & """"
            Me![MiddleInit].DefaultValue = """" & strMid & """"
            Else
            Me![FirstName].DefaultValue = """" & strFirst & """"
        End If
    End If

End Sub

Private Sub PayBy_AfterUpdate()
    ' After choosing a payment method,
    ' .. reveal the CCNumber if necessary and set the input mask
    ' based on type of credit card
    Select Case Me!PayBy
        Case 1, 2
            Me!CCNumber = Null
            Me!CCNumber.Visible = False
        Case 3
            Me!CCNumber.Visible = True
            Me!CCNumber.InputMask = "0000\-000000\-00000;0;_"
        Case Else
            Me!CCNumber.Visible = True
            Me!CCNumber.InputMask = "0000\-0000\-0000\-0000;0;_"
    End Select
End Sub
