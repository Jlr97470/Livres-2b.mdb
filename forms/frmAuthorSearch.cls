Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub cmdCancel_Click()
    Me.Visible = False
    If IsLoaded("frmAuthors") Then ' If authors is open
        ' Then unfilter it and put the focus there
        Forms!frmAuthors.SetFocus
        DoCmd.ShowAllRecords
        If gintIsAdmin Then
            Forms!frmAuthors!cmdAddNew.Visible = True
            Else
            Forms!frmAuthors!cmdAddNew.Visible = False
        End If
        Forms!frmAuthors!cmdShowAll.Visible = False
        gstrWhereAuthor = ""
    End If

    DoCmd.Close acForm, Me.Name

End Sub


Private Sub cmdSearch_Click()
Dim cn As ADODB.Connection, rcd As New ADODB.Recordset
Dim lngCount As Long, intRtn As Integer
    ' Clear out the global author filter string
    gstrWhereAuthor = ""

    ' Parse out the request into a new filter string
    If Not IsNothing(Me!LastName) Then
        gstrWhereAuthor = "[LastName] Like " & Chr$(39) & Me!LastName
        If Right$(Me!LastName, 1) = "*" Then
            gstrWhereAuthor = gstrWhereAuthor & Chr$(39)
        Else
            gstrWhereAuthor = gstrWhereAuthor & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!FirstName) Then
        If IsNothing(gstrWhereAuthor) Then
            gstrWhereAuthor = "[FirstName] Like " & Chr$(39) & Me!FirstName
        Else
            gstrWhereAuthor = gstrWhereAuthor & "AND [FirstName] Like " & Chr$(39) & Me!FirstName
        End If
        If Right$(Me!FirstName, 1) = "*" Then
            gstrWhereAuthor = gstrWhereAuthor & Chr$(39)
        Else
            gstrWhereAuthor = gstrWhereAuthor & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!TypeCodes) Then
        If IsNothing(gstrWhereAuthor) Then
            gstrWhereAuthor = "[AuthorID] IN (Select AuthorID From qryAuthorCategories Where CategoryID IN (" & Me!TypeCodes & "))"
            Else
            gstrWhereAuthor = gstrWhereAuthor & " AND [AuthorID] IN (Select AuthorID From qryAuthorCategories Where CategoryID IN (" & Me!TypeCodes & "))"
        End If
    End If

    ' If no filter criteria, then bail
    If IsNothing(gstrWhereAuthor) Then
        MsgBox "No criteria specified.", vbExclamation, "Microsoft Press"
        Exit Sub
    End If

    ' Hide myself and turn on Hourglass
    Me.Visible = False
    DoCmd.Hourglass True
    If IsLoaded("frmAuthors") Then  ' If authors form already open,
        ' Then just filter it
        Forms!frmAuthors.SetFocus
        DoCmd.ApplyFilter , gstrWhereAuthor
        If Forms!frmAuthors.Recordset.RecordCount = 0 Then
            DoCmd.Hourglass False
            MsgBox "No Authors meet your criteria", vbInformation, "Microsoft Press"
            DoCmd.ShowAllRecords
            If gintIsAdmin Then
                Forms!frmAuthors!cmdAddNew.Visible = True
                Else
                Forms!frmAuthors!cmdAddNew.Visible = False
            End If
            Forms!frmAuthors!cmdShowAll.Visible = False
            Me.Visible = True
            Exit Sub
        End If

        Forms!frmAuthors!cmdAddNew.Visible = False
        Forms!frmAuthors!cmdShowAll.Visible = True
        DoCmd.Hourglass False
    Else
        ' Find out if any authors satisfy the Where clause
        Set cn = CurrentProject.Connection
        rcd.Open "Select * From tblAuthors", cn, adOpenKeyset, adLockOptimistic
        rcd.Filter = gstrWhereAuthor
        ' If none found, then tell them and make me visible to try again
        If rcd.RecordCount = 0 Then
            DoCmd.Hourglass False
            MsgBox "No Authors meet your criteria", vbInformation, "Microsoft Press"
            gstrWhereAuthor = ""
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
            intRtn = MsgBox("More than 10 authors meet your criteria. " & vbCrLf & vbCrLf & _
                "Click Yes to see a summary list for all " & lngCount & _
                " authors found, " & _
                "No to see complete data on all that match, " & _
                "or Cancel to try again.", vbInformation + vbYesNoCancel, "Microsoft Press")
            Select Case intRtn
                Case vbCancel   ' Cancel - Try again
                    Me.Visible = True
                    Exit Sub
                Case vbYes      ' Yes - show summary form
                    DoCmd.OpenForm FormName:="frmAuthorSummary", WhereCondition:=gstrWhereAuthor
                    DoCmd.Close acForm, Me.Name
                    Forms!frmAuthorSummary.SetFocus
                    Exit Sub
            End Select
        End If
        ' Replied NO or not more than 10, show full details
        DoCmd.OpenForm FormName:="frmAuthors", WhereCondition:=gstrWhereAuthor
        Forms!frmAuthors!cmdAddNew.Visible = False
        Forms!frmAuthors!cmdShowAll.Visible = True
    End If
    ' Close me, and we're done
    DoCmd.Close acForm, Me.Name


End Sub

Private Sub cmdTypeClear_Click()
    ' Clear the type code list
    Me!Types = Null
    Me!TypeCodes = Null
    Me!TypeList = Null

End Sub


Private Sub TypeList_AfterUpdate()
    ' When they pick a new type code, then add it to the type code string
    If IsNothing(Me!TypeCodes) Then
        Me!TypeCodes = Me!TypeList
        Me!Types = Me!TypeList.Column(0)
        Else
        Me!TypeCodes = Me!TypeCodes & ", " & Me!TypeList
        Me!Types = Me!Types & " OR " & Me!TypeList.Column(0)
    End If

End Sub
