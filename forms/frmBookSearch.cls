Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub cmdCancel_Click()
    ' Decided not to search after all...
    Me.Visible = False
    If IsLoaded("frmBooks") Then ' If books is open
        ' Then unfilter it and put the focus there
        Forms!frmBooks.SetFocus
        DoCmd.ShowAllRecords
        If gintIsAdmin Then
            Forms!frmBooks!cmdAddNew.Visible = True
            Else
            Forms!frmBooks!cmdAddNew.Visible = False
        End If
        Forms!frmBooks!cmdShowAll.Visible = False
        gstrWhereBook = ""
    End If

    DoCmd.Close acForm, Me.Name

End Sub


Private Sub cmdSearch_Click()
Dim rcd As New ADODB.Recordset, cn As ADODB.Connection
Dim lngCount As Long, intRtn As Integer, strANameComp As String
Dim intI As Integer, strISBN As String

    ' Search parameters entered (we hope)...
    ' Clear the global book filter string
    gstrWhereBook = ""

    ' .. and parse out a new one:
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
        gstrWhereBook = "[ISBNNumber] Like " & Chr$(39) & Me!ISBNNumber
        ' Add "*" on the end for good measure...
        If Right$(Me!ISBNNumber, 1) = "*" Then
            gstrWhereBook = gstrWhereBook & Chr$(39)
        Else
            gstrWhereBook = gstrWhereBook & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!Title) Then
        If IsNothing(gstrWhereBook) Then
            gstrWhereBook = "[Title] Like " & Chr$(39) & Me!Title
            Else
            gstrWhereBook = gstrWhereBook & " AND [Title] Like " & Chr$(39) & Me!Title
        End If
        If Right$(Me!Title, 1) = "*" Then
            gstrWhereBook = gstrWhereBook & Chr$(39)
        Else
            gstrWhereBook = gstrWhereBook & "*" & Chr$(39)
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
        If IsNothing(gstrWhereBook) Then
            gstrWhereBook = "[ISBNNumber] IN (Select ISBNNumber From [qryBookAuthors]" & _
                " Where " & strANameComp & ")"
            Else
            gstrWhereBook = gstrWhereBook & " AND [ISBNNumber] IN (Select ISBNNumber" & _
                " From [qryBookAuthors] Where " & strANameComp & ")"
        End If
    End If

    If Not IsNothing(Me!LowPrice) Then
        If IsNothing(gstrWhereBook) Then
            gstrWhereBook = "[SuggPrice] >= " & Me!LowPrice
            Else
            gstrWhereBook = gstrWhereBook & " AND [SuggPrice] >= " & Me!LowPrice
        End If
    End If

    If Not IsNothing(Me!HighPrice) Then
        If IsNothing(gstrWhereBook) Then
            gstrWhereBook = "[SuggPrice] <= " & Me!HighPrice
            Else
            gstrWhereBook = gstrWhereBook & " AND [SuggPrice] <= " & Me!HighPrice
        End If
    End If

    If Me!Disk Then
        If IsNothing(gstrWhereBook) Then
            gstrWhereBook = "([Disk] Is Not Null)"
            Else
            gstrWhereBook = gstrWhereBook & " AND ([Disk] Is Not Null)"
        End If
    End If

    If Me!OutOfPrint Then
        If IsNothing(gstrWhereBook) Then
            gstrWhereBook = "([OutOfPrint] = False)"
            Else
            gstrWhereBook = gstrWhereBook & " AND ([OutOfPrint] = False)"
        End If
    End If

    If Not IsNothing(Me!TypeCodes) Then
        If IsNothing(gstrWhereBook) Then
            gstrWhereBook = "[ISBNNumber] IN (Select ISBNNumber From [qryBookCategories]" & _
                " Where CategoryID IN (" & Me!TypeCodes & "))"
            Else
            gstrWhereBook = gstrWhereBook & " AND [ISBNNumber] IN (Select ISBNNumber" & _
                " From [qryBookCategories] Where CategoryID IN (" & Me!TypeCodes & "))"
        End If
    End If

    ' If no criteria, then nothing to do!
    If IsNothing(gstrWhereBook) Then
        MsgBox "No criteria specified.", vbExclamation, "Microsoft Press"
        Exit Sub
    End If

    ' Hide myself and turn on Hourglass
    Me.Visible = False
    DoCmd.Hourglass True
    If IsLoaded("frmBooks") Then  ' If books form already open,
        ' Then just filter it
        Forms!frmBooks.SetFocus
        DoCmd.ApplyFilter , gstrWhereBook
        If Forms!frmBooks.Recordset.RecordCount = 0 Then
            DoCmd.Hourglass False
            MsgBox "No Books meet your criteria", vbExclamation, "Microsoft Press"
            DoCmd.ShowAllRecords
            If gintIsAdmin Then
                Forms!frmBooks!cmdAddNew.Visible = True
                Else
                Forms!frmBooks!cmdAddNew.Visible = False
            End If
            Forms!frmBooks!cmdShowAll.Visible = False
            Me.Visible = True
            Exit Sub
        End If

        If IsNothing(Me!TypeCodes) And (Not IsNothing(strANameComp)) Then
        ' If not book categories, but did do a name search,
            ' Turn on the author display option
            Forms!frmBooks!tabDisplay = 1
        End If

        Forms!frmBooks!cmdAddNew.Visible = False
        Forms!frmBooks!cmdShowAll.Visible = True
        DoCmd.Hourglass False
    Else
        ' Find out if any books satisfy the Where clause
        Set cn = CurrentProject.Connection
        rcd.Open "Select * From tblBooks", cn, adOpenKeyset, adLockOptimistic
        rcd.Filter = gstrWhereBook
        ' If none found, then tell them and make me visible to try again
        If rcd.RecordCount = 0 Then
            DoCmd.Hourglass False
            MsgBox "No Books meet your criteria", vbExclamation, "Microsoft Press"
            gstrWhereBook = ""
            Me.Visible = True
            rcd.Close
            Exit Sub
        End If
        ' Move to last row to get an accurate record count
        rcd.MoveLast
        lngCount = rcd.RecordCount
        DoCmd.Hourglass False
        ' If more than 10, then ask if they want to only see a summary
        If lngCount > 10 Then
            intRtn = MsgBox("More than 10 books meet your criteria. " & vbCrLf & vbCrLf & _
                "Click Yes to see a summary list for all " & lngCount & _
                " books found, " & _
                "No to see complete data on all that match, " & _
                "or Cancel to try again.", vbInformation + vbYesNoCancel, "Microsoft Press")
            Select Case intRtn
                Case vbCancel   ' Cancel - Try again
                    Me.Visible = True
                    Exit Sub
                Case vbYes      ' Yes - show summary form
                    DoCmd.OpenForm FormName:="frmBookSummary", WhereCondition:=gstrWhereBook
                    DoCmd.Close acForm, Me.Name
                    Forms!frmBookSummary.SetFocus
                    Set Forms!frmBookSummary.Recordset = rcd
                    Exit Sub
            End Select
        End If
        ' Replied NO or not more than 10, show full details
        DoCmd.OpenForm FormName:="frmBooks", WhereCondition:=gstrWhereBook
        Forms!frmBooks!cmdAddNew.Visible = False
        Forms!frmBooks!cmdShowAll.Visible = True
        Set Forms!frmBooks.Recordset = rcd
        If IsNothing(Me!TypeCodes) And (Not IsNothing(strANameComp)) Then
        ' If not book categories, but did do a name search,
            ' Turn on the author display option
            Forms!frmBooks!tabDisplay = 1
        End If
    End If
    ' Close me, and we're done
    DoCmd.Close acForm, Me.Name


End Sub

Private Sub cmdTypeClear_Click()
    ' This clears the type code list
    Me!Types = Null
    Me!TypeCodes = Null
    Me!TypeList = Null

End Sub


Private Sub TypeList_AfterUpdate()
    ' Selected a new type code, so add it to the search list
    If IsNothing(Me!TypeCodes) Then
        Me!TypeCodes = Me!TypeList
        Me!Types = Me!TypeList.Column(0)
        Else
        Me!TypeCodes = Me!TypeCodes & ", " & Me!TypeList
        Me!Types = Me!Types & " OR " & Me!TypeList.Column(0)
    End If

End Sub
