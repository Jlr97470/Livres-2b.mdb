Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub cmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Print_Err
    ' Print what they asked for...
    Dim frmO As Form, cn As ADODB.Connection, strSQL As String
    Set frmO = Forms!frmOrders
    Select Case Me!optPrtOptions
        Case 1
            ' Print the current order --
            Me.Visible = False
            DoCmd.OpenReport ReportName:="rptOrders", View:=gintPrintPreview, _
                WhereCondition:="[OrderID] = " & frmO!OrderID
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptOrders"
            End If
            If Not (frmO!Printed) Then
                ' If not marked printed, then do so
                strSQL = "UPDATE tblOrders SET [Printed] = True WHERE [OrderID] = " & frmO!OrderID
                Set cn = CurrentProject.Connection
                cn.Execute strSQL
                ' and call the orders form sub to lock all controls
                ' NOTE: The PUBLIC sub in frmOrders is like an object METHOD!
                frmO.SetLocked True
            End If
            DoCmd.Close acForm, Me.Name
        Case 2
            ' Customer Label
            DoCmd.OpenReport ReportName:="rptCustLabels", View:=gintPrintPreview, _
                WhereCondition:="[CustomerID] = " & frmO!CustomerID
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptCustLabels"
            End If
        Case 3
            ' Store label --
            DoCmd.OpenReport ReportName:="rptStoreLabels", View:=gintPrintPreview, _
                WhereCondition:="[StoreID] = " & frmO!StoreID
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptStoreLabels"
            End If
    End Select
    Exit Sub

Print_Err:
    MsgBox "Error: " & Err & ", " & Error
    Exit Sub
End Sub


Private Sub Form_Load()
    ' Check current order - if no Store, then disable that option
    If IsNothing(Forms!frmOrders!StoreID) Then
        Me!optStore.Enabled = False
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    If Not IsLoaded("frmOrders") Then
        MsgBox "The Orders form must be open to use this form.", vbCritical, "Microsoft Press"
        Cancel = True
    End If
End Sub
