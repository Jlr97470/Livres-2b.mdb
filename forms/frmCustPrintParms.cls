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
    ' Run the report based on option selected...
    Select Case Me!optPrtOptions
        Case 1
            ' Customer Report
            If IsNothing(gstrWhereCust) Then
                ' If no filter string saved, then open normally
                DoCmd.OpenReport ReportName:="rptCustomers", View:=gintPrintPreview
                Else
                ' Open using the current filter
                DoCmd.OpenReport ReportName:="rptCustomers", _
                    View:=gintPrintPreview, WhereCondition:=gstrWhereCust
            End If
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptCustomers"
            End If
        Case 2
            ' Mailing labels --
            If IsNothing(gstrWhereCust) Then
                ' If no filter string saved, then open normally
                DoCmd.OpenReport ReportName:="rptCustLabels", View:=gintPrintPreview
                Else
                ' Open using the current filter
                DoCmd.OpenReport ReportName:="rptCustLabels", _
                    View:=gintPrintPreview, WhereCondition:=gstrWhereCust
            End If
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptCustLabels"
            End If
        Case 3
            ' #10 envelope for the currently displayed customer
            DoCmd.OpenReport ReportName:="rptCustEnvelope", _
                WhereCondition:="CustomerID = " & Forms!frmCustomers!CustomerID, _
                View:=gintPrintPreview
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptCustEnvelope"
            End If
        Case 4
            ' Sales for the current customer --
            If Me!FromDate <= Me!ToDate Then
                Else
                MsgBox "'From' date must be less than or equal to 'To' date.", vbExclamation, "Microsoft Press"
                Exit Sub
            End If
            Me.Visible = False
            ' Set up the filter string
            strWhere = "(CustomerID = " & Forms!frmCustomers!CustomerID & ")" & _
                " AND ([OrderDate] Between #" & Format(Me!FromDate, "Short Date") & _
                "# AND #" & Format(Me!ToDate, "Short Date") & "#)"
            DoCmd.OpenReport ReportName:="rptCustSales", _
                View:=gintPrintPreview, WhereCondition:=strWhere
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptCustSales"
            End If
        Case 5
            ' Sales for all selected customers --
            If Me!FromDate <= Me!ToDate Then
                Else
                MsgBox "'From' date must be less than or equal to 'To' date.", vbExclamation, "Microsoft Press"
                Exit Sub
            End If
            Me.Visible = False
            ' Set up the filter string
            strWhere = gstrWhereCust
            If IsNothing(strWhere) Then
                strWhere = "([OrderDate] Between #" & Format(Me!FromDate, "Short Date") & _
                    "# AND #" & Format(Me!ToDate, "Short Date") & "#)"
                Else
                strWhere = strWhere & _
                    " AND ([OrderDate] Between #" & Format(Me!FromDate, "Short Date") & _
                    "# AND #" & Format(Me!ToDate, "Short Date") & "#)"
            End If
            DoCmd.OpenReport ReportName:="rptCustSales", _
                View:=gintPrintPreview, WhereCondition:=strWhere
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptCustSales"
            End If
    End Select
    Exit Sub

Print_Err:
    MsgBox "Error: " & Err & ", " & Error
    Exit Sub
End Sub


Private Sub optPrtOptions_AfterUpdate()
    ' If asking for a sales report...
    Select Case Me!optPrtOptions
        Case 1, 2, 3
            DoCmd.MoveSize , , , 3090
        Case Else
            ' .. then reveal the sales option dates
            DoCmd.MoveSize , , , 3990
            Me!FromDate.SetFocus
    End Select
End Sub
