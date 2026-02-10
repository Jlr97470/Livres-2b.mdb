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
    Dim strWhere As String

    Select Case Me!optPrtOptions
        Case 1
            ' Displayed (filtered) stores --
            If IsNothing(gstrWhereStore) Then
                ' If no filter string saved, then open normally
                DoCmd.OpenReport ReportName:="rptStores", View:=gintPrintPreview
                Else
                ' Open using the current filter
                DoCmd.OpenReport ReportName:="rptStores", _
                    View:=gintPrintPreview, WhereCondition:=gstrWhereStore
            End If
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptStores"
            End If
        Case 2
            ' All Stores Report
            DoCmd.OpenReport ReportName:="rptStores", View:=gintPrintPreview
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptStores"
            End If
        Case 3
            ' Mailing label for the current store --
            ' Set up the filter string
            strWhere = "[StoreID] = " & Forms!frmStores!StoreID
            DoCmd.OpenReport ReportName:="rptStoreLabels", _
                WhereCondition:=strWhere, _
                View:=gintPrintPreview
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptStoreLabels"
            End If
        Case 4
            ' Labels for selected (filtered) stores --
            Me.Visible = False
            '
            If IsNothing(gstrWhereStore) Then
                ' If no filter string saved, then open normally
                DoCmd.OpenReport ReportName:="rptStoreLabels", _
                    View:=gintPrintPreview
                Else
                ' Open using the current filter
                DoCmd.OpenReport ReportName:="rptStoreLabels", _
                    View:=gintPrintPreview, WhereCondition:=gstrWhereStore
            End If
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptStoreLabels"
            End If
        Case 5
            ' Labels for all stores --
            Me.Visible = False
            DoCmd.OpenReport ReportName:="rptStoreLabels", _
                View:=gintPrintPreview
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
