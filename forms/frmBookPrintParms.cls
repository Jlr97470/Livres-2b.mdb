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
            ' Displayed (filtered) books --
            If IsNothing(gstrWhereBook) Then
                ' If no filter string saved, then open normally
                DoCmd.OpenReport ReportName:="rptBooks", View:=gintPrintPreview
                Else
                ' Open using the current filter
                DoCmd.OpenReport ReportName:="rptBooks", _
                    View:=gintPrintPreview, WhereCondition:=gstrWhereBook
            End If
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptBooks"
            End If
        Case 2
            ' All Books Report
            DoCmd.OpenReport ReportName:="rptBooks", View:=gintPrintPreview
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptBooks"
            End If
        Case 3
            ' Sales for the current book --
            If Me!FromDate <= Me!ToDate Then
                Else
                MsgBox "'From' date must be less than or equal to 'To' date.", vbExclamation, "Microsoft Press"
                Exit Sub
            End If
            ' Set up the filter string
            strWhere = "([ISBNNumber] = '" & Forms!frmBooks!ISBNNumber & "')" & _
                " AND ([OrderDate] Between #" & Format(Me!FromDate, "Short Date") & _
                "# AND #" & Format(Me!ToDate, "Short Date") & "#)"
            DoCmd.OpenReport ReportName:="rptBookSales", _
                WhereCondition:=strWhere, _
                View:=gintPrintPreview
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptBookSales"
            End If
        Case 4
            ' Sales for selected (filtered) books --
            If Me!FromDate <= Me!ToDate Then
                Else
                MsgBox "'From' date must be less than or equal to 'To' date.", vbExclamation, "Microsoft Press"
                Exit Sub
            End If
            Me.Visible = False
            ' Set up the filter string
            strWhere = gstrWhereBook
            If IsNothing(strWhere) Then
                strWhere = "([OrderDate] Between #" & Format(Me!FromDate, "Short Date") & _
                    "# AND #" & Format(Me!ToDate, "Short Date") & "#)"
                Else
                strWhere = strWhere & _
                    " AND ([OrderDate] Between #" & Format(Me!FromDate, "Short Date") & _
                    "# AND #" & Format(Me!ToDate, "Short Date") & "#)"
            End If
            DoCmd.OpenReport ReportName:="rptBookSales", _
                View:=gintPrintPreview, WhereCondition:=strWhere
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptBookSales"
            End If
        Case 5
            ' Sales for all books --
            If Me!FromDate <= Me!ToDate Then
                Else
                MsgBox "'From' date must be less than or equal to 'To' date.", vbExclamation, "Microsoft Press"
                Exit Sub
            End If
            Me.Visible = False
            ' Set up the filter string
            strWhere = "([OrderDate] Between #" & Format(Me!FromDate, "Short Date") & _
                "# AND #" & Format(Me!ToDate, "Short Date") & "#)"
            DoCmd.OpenReport ReportName:="rptBookSales", _
                View:=gintPrintPreview, WhereCondition:=strWhere
            DoCmd.Close acForm, Me.Name
            ' If this user has print preview option set,
            If gintPrintPreview = acPreview Then
                ' Then make sure focus goes to the report window
                DoCmd.SelectObject acReport, "rptBookSales"
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
        Case 1, 2
            DoCmd.MoveSize , , , 3090
        Case Else
        ' .. then flick open the sales date options
            DoCmd.MoveSize , , , 3990
            Me!FromDate.SetFocus
    End Select
End Sub
