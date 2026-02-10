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
Dim strANameComp As String
Dim intI As Integer, strISBN As String, strWhere As String

    ' Set up a filter string based on criteria entered
    If Not IsNothing(Me!ISBNNumber) Then
        strISBN = Me!ISBNNumber
        ' The ISBN has an input mask to help the user out, but they could
        ' enter all sorts of dashes and numbers.  The following code does an
        ' additional "parse" to clean it up!

        ' First, replace all blanks in ISBN with question marks
        intI = 1
        Do Until intI > Len(strISBN)
            If Mid$(strISBN, intI, 1) = " " Then
                If intI = 1 Then
                    strISBN = "?" & Mid$(strISBN, intI + 1)
                    Else
                    strISBN = Left$(strISBN, intI - 1) & "?" & Mid$(strISBN, intI + 1)
                End If
            End If
            ' Watch out for dashes in the right/wrong places
            If intI = 2 Or intI = 8 Or intI = 12 Then
                ' If position 2, 8, or 12, then a dash is OK
                If Mid$(strISBN, intI, 1) = "-" Then
                Else
                ' Otherwise, move it down
                    strISBN = Left$(strISBN, intI - 1) & "?" & Mid$(strISBN, intI)
                End If
            Else
                ' Not position 2, 8, or 12, so should be no dash
                If Mid$(strISBN, intI, 1) = "-" Then
                    If intI = 1 Then
                        strISBN = "?" & strISBN
                        Else
                        ' Watch out for too many dashes!
                        If intI > 12 Then
                            ' Replace the dash if beyond position 12
                            strISBN = Left$(strISBN, intI - 1) & "?" & Mid$(strISBN, intI + 1)
                        Else
                            ' Shove the dash down if position <=12
                            strISBN = Left$(strISBN, intI - 1) & "?" & Mid$(strISBN, intI)
                        End If
                    End If
                End If
            End If
            intI = intI + 1
        Loop
        Me!ISBNNumber = Left$(strISBN, 13)
        ' Result at this point should be a nice ISBN in the form:
        ' ?-?????-???-?
        ' Where the user may have specified digits for any one or more of the
        ' question marks.
        strWhere = "[ISBNNumber] Like " & Chr$(39) & Me!ISBNNumber
        ' Add "*" on the end just in case...
        If Right$(Me!ISBNNumber, 1) = "*" Then
            strWhere = strWhere & Chr$(39)
        Else
            strWhere = strWhere & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!Title) Then
        If IsNothing(strWhere) Then
            strWhere = "[Title] Like " & Chr$(39) & Me!Title
            Else
            strWhere = strWhere & " AND [Title] Like " & Chr$(39) & Me!Title
        End If
        If Right$(Me!Title, 1) = "*" Then
            strWhere = strWhere & Chr$(39)
        Else
            strWhere = strWhere & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!LastName) Then
        strANameComp = "[LastName] Like " & Chr$(39) & Me!LastName
        If Right$(Me!LastName, 1) = "*" Then
            strANameComp = strANameComp & Chr$(39)
        Else
            strANameComp = strANameComp & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!FirstName) Then
        If IsNothing(strANameComp) Then
            strANameComp = "[FirstName] Like " & Chr$(39) & Me!FirstName
        Else
            strANameComp = strANameComp & " AND [FirstName] Like " & Chr$(39) & Me!FirstName
        End If
        If Right$(Me!FirstName, 1) = "*" Then
            strANameComp = strANameComp & Chr$(39)
        Else
            strANameComp = strANameComp & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(strANameComp) Then  ' Did we build any author compares?
        If IsNothing(strWhere) Then
            strWhere = "[ISBNNumber] IN (Select ISBNNumber From qryBookAuthors Where " & strANameComp & ")"
            Else
            strWhere = strWhere & " AND [ISBNNumber] IN (Select ISBNNumber From qryBookAuthors Where " & strANameComp & ")"
        End If
    End If

    If Not IsNothing(Me!LowPrice) Then
        If IsNothing(strWhere) Then
            strWhere = "[SuggPrice] >= " & Me!LowPrice
            Else
            strWhere = strWhere & " AND [SuggPrice] >= " & Me!LowPrice
        End If
    End If

    If Not IsNothing(Me!HighPrice) Then
        If IsNothing(strWhere) Then
            strWhere = "[SuggPrice] <= " & Me!HighPrice
            Else
            strWhere = strWhere & " AND [SuggPrice] <= " & Me!HighPrice
        End If
    End If

    If Me!Disk Then
        If IsNothing(strWhere) Then
            strWhere = "([Disk] Is Not Null)"
            Else
            strWhere = strWhere & " AND ([Disk] Is Not Null)"
        End If
    End If

    If Not IsNothing(Me!TypeCodes) Then
        If IsNothing(strWhere) Then
            strWhere = "[ISBNNumber] IN (Select ISBNNumber From qryBookCategories Where CategoryID IN (" & Me!TypeCodes & "))"
            Else
            strWhere = strWhere & " AND [ISBNNumber] IN (Select ISBNNumber From qryBookCategories Where CategoryID IN (" & Me!TypeCodes & "))"
        End If
    End If

    ' Sales for all selected books --
    If Me!FromDate <= Me!ToDate Then
        Else
        MsgBox "'From' date must be less than or equal to 'To' date.", vbExclamation, "Microsoft Press"
        Exit Sub
    End If
    Me.Visible = False
    ' Finish the filter string
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


End Sub

Private Sub cmdTypeClear_Click()
    ' Clears the type code list when clicked
    Me!Types = Null
    Me!TypeCodes = Null
    Me!TypeList = Null

End Sub


Private Sub TypeList_AfterUpdate()
    ' User selected another type code, so add it to the list
    If IsNothing(Me!TypeCodes) Then
        Me!TypeCodes = Me!TypeList
        Me!Types = Me!TypeList.Column(0)
        Else
        Me!TypeCodes = Me!TypeCodes & ", " & Me!TypeList
        Me!Types = Me!Types & " OR " & Me!TypeList.Column(0)
    End If

End Sub
