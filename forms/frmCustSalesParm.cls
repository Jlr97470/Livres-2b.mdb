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
Dim strWhere As String

    ' Build a filter string for criteria specified...
    If Not IsNothing(Me!LastName) Then
        strWhere = "[LastName] Like " & Chr$(39) & Me!LastName
        If Right$(Me!LastName, 1) = "*" Then
            strWhere = strWhere & Chr$(39)
        Else
            strWhere = strWhere & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!FirstName) Then
        If IsNothing(strWhere) Then
            strWhere = "[FirstName] Like " & Chr$(39) & Me!FirstName
        Else
            strWhere = strWhere & " AND [FirstName] Like " & Chr$(39) & Me!FirstName
        End If
        If Right$(Me!FirstName, 1) = "*" Then
            strWhere = strWhere & Chr$(39)
        Else
            strWhere = strWhere & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!City) Then
        If IsNothing(strWhere) Then
            strWhere = "[City] Like " & Chr$(39) & Me!City
        Else
            strWhere = strWhere & " AND [City] Like " & Chr$(39) & Me!City
        End If
        If Right$(Me!City, 1) = "*" Then
            strWhere = strWhere & Chr$(39)
        Else
            strWhere = strWhere & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!State) Then
        If IsNothing(strWhere) Then
            strWhere = "[StateOrProvince] Like " & Chr$(39) & Me!State
        Else
            strWhere = strWhere & " AND [StateOrProvince] Like " & Chr$(39) & Me!State
        End If
        If Right$(Me!State, 1) = "*" Then
            strWhere = strWhere & Chr$(39)
        Else
            strWhere = strWhere & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!Country) Then
        If IsNothing(strWhere) Then
            strWhere = "[Country] Like " & Chr$(39) & Me!Country
        Else
            strWhere = strWhere & " AND [Country] Like " & Chr$(39) & Me!Country
        End If
        If Right$(Me!Country, 1) = "*" Then
            strWhere = strWhere & Chr$(39)
        Else
            strWhere = strWhere & "*" & Chr$(39)
        End If
    End If

    ' Check from/to dates --
    If Me!FromDate <= Me!ToDate Then
        Else
        MsgBox "'From' date must be less than or equal to 'To' date.", vbExclamation, "Microsoft Press"
        Exit Sub
    End If
    Me.Visible = False
    ' Add dates to the filter string
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


End Sub
