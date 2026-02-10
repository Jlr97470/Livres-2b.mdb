Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub cmdCancel_Click()

    DoCmd.Close acForm, Me.Name

End Sub


Private Sub cmdSearch_Click()
Dim cn As ADODB.Connection, rcd As New ADODB.Recordset
Dim lngCount As Long, intRtn As Integer
Dim strWhere As String

    If Not IsNothing(Me!StoreName) Then
        strWhere = "[StoreName] Like " & Chr$(39) & Me!StoreName
        If Right$(Me!LastName, 1) = "*" Then
            strWhere = strWhere & Chr$(39)
        Else
            strWhere = strWhere & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!City) Then
        If IsNothing(strWhere) Then
            strWhere = "[City] Like " & Chr$(39) & Me!City
        Else
            strWhere = strWhere & "AND [City] Like " & Chr$(39) & Me!City
        End If
        If Right$(Me!City, 1) = "*" Then
            strWhere = strWhere & Chr$(39)
        Else
            strWhere = strWhere & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!StateOrProvince) Then
        If IsNothing(strWhere) Then
            strWhere = "[StateOrProvince] Like " & Chr$(39) & Me!StateOrProvince
        Else
            strWhere = strWhere & "AND [StateOrProvince] Like " & Chr$(39) & Me!StateOrProvince
        End If
        If Right$(Me!StateOrProvince, 1) = "*" Then
            strWhere = strWhere & Chr$(39)
        Else
            strWhere = strWhere & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!PostalCode) Then
        If IsNothing(strWhere) Then
            strWhere = "[PostalCode] Like " & Chr$(39) & Me!PostalCode
        Else
            strWhere = strWhere & "AND [PostalCode] Like " & Chr$(39) & Me!PostalCode
        End If
        If Right$(Me!PostalCode, 1) = "*" Then
            strWhere = strWhere & Chr$(39)
        Else
            strWhere = strWhere & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!Country) Then
        If IsNothing(strWhere) Then
            strWhere = "[Country] Like " & Chr$(39) & Me!Country
        Else
            strWhere = strWhere & "AND [Country] Like " & Chr$(39) & Me!Country
        End If
        If Right$(Me!Country, 1) = "*" Then
            strWhere = strWhere & Chr$(39)
        Else
            strWhere = strWhere & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!PhoneCode) Then
        If IsNothing(strWhere) Then
            strWhere = "[PhoneCode] = " & Chr$(39) & Me!PhoneCode & Chr$(39)
        Else
            strWhere = strWhere & "AND [PhoneCode] = " & Chr$(39) & Me!PhoneCode & Chr$(39)
        End If
    End If

    If IsNothing(strWhere) Then
        MsgBox "No criteria specified.", vbExclamation, "Microsoft Press"
        Exit Sub
    End If

    ' Turn on Hourglass
    DoCmd.Hourglass True
    ' If stores form not open or called with an open argument
    If (Not IsLoaded("frmStores")) Or Me.OpenArgs = "Orders" Then
        ' Find out if any stores satisfy the Where clause
        ' Get a copy of the current connection
        Set cn = CurrentProject.Connection
        rcd.Open "Select *, Left([PhoneNumber], 5) As PhoneCode From tblStores", cn, adOpenKeyset, adLockOptimistic
        rcd.Filter = strWhere
        If rcd.RecordCount = 0 Then
            DoCmd.Hourglass False
            MsgBox "No Stores meet your criteria.", vbInformation, "Microsoft Press"
            strWhere = ""
            rcd.Close
            Exit Sub
        End If
        ' Found at least one..  If called from Orders form
        If Me.OpenArgs = "Orders" Then
            ' Set StoreID in orders
            Forms!frmOrders!StoreID = rcd!StoreID
            rcd.Close
            DoCmd.Hourglass False
            DoCmd.Close acForm, Me.Name
            Exit Sub
        End If
        ' Doing a search from frmStoreList...
        ' Set the global store filter
        gstrWhereStore = strWhere
        ' and hide me..
        Me.Visible = False
        ' Move to last row to get an accurate record count
        rcd.MoveLast
        lngCount = rcd.RecordCount
        DoCmd.Hourglass False
        ' If more than 10, then ask if they want to only see a summary
        If lngCount > 10 Then
            intRtn = MsgBox("More than 10 stores meet your criteria. " & vbCrLf & vbCrLf & _
                "Click Yes to see a summary list for all " & lngCount & _
                " stores found, " & _
                "No to see complete data on all that match, " & _
                "or Cancel to try again.", vbInformation + vbYesNoCancel, "Microsoft Press")
            Select Case intRtn
                Case vbCancel   ' Cancel - Try again
                    Me.Visible = True
                    Exit Sub
                Case vbYes      ' Yes - show summary form
                    DoCmd.OpenForm FormName:="frmStoreSummary", WhereCondition:=strWhere
                    DoCmd.Close acForm, Me.Name
                    Forms!frmStoreSummary.SetFocus
                    Exit Sub
            End Select
        End If
        ' Replied NO or not more than 10, show full details
        DoCmd.OpenForm FormName:="frmStores", WhereCondition:=strWhere
        Forms!frmStores!cmdShowAll.Enabled = True
    Else
        ' Stores form already loaded - filter it
        Me.Visible = False
        Forms!frmStores.SetFocus
        DoCmd.ApplyFilter , strWhere
        If Forms!frmStores.Recordset.RecordCount = 0 Then
            DoCmd.Hourglass False
            MsgBox "No Stores meet your criteria.", vbInformation, "Microsoft Press"
            If IsNothing(gstrWhereStore) Then ' If no global filter
                ' Do a show all records
                DoCmd.ShowAllRecords
                Forms!frmStores!cmdShowAll.Enabled = False
                Else
                ' Otherwise, put the filter back
                DoCmd.ApplyFilter , gstrWhereStore
            End If
            Me.Visible = True
            Exit Sub
        End If
        ' Got a good search -- set the global filter
        gstrWhereStore = strWhere
        Forms!frmStores!cmdShowAll.Enabled = True
        DoCmd.Hourglass False
    End If
    ' Close me, and we're done
    DoCmd.Close acForm, Me.Name


End Sub
