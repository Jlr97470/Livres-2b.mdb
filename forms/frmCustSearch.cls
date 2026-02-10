Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub cmdCancel_Click()
    ' Decided not to search after all...
    Me.Visible = False
    If IsLoaded("frmCustomers") Then ' If customers is open
        ' Then unfilter it and put the focus there
        Forms!frmCustomers.SetFocus
        DoCmd.ShowAllRecords
        If gintIsAdmin Then
            Forms!frmCustomers!cmdAddNew.Visible = True
            Else
            Forms!frmCustomers!cmdAddNew.Visible = False
        End If
        Forms!frmCustomers!cmdShowAll.Visible = False
        gstrWhereCust = ""
    End If

    DoCmd.Close acForm, Me.Name

End Sub


Private Sub cmdSearch_Click()
    ' Search - build a new filter string
    gstrWhereCust = ""

    If Not IsNothing(Me!LastName) Then
        gstrWhereCust = "[LastName] Like " & Chr$(39) & Me!LastName
        If Right$(Me!LastName, 1) = "*" Then
            gstrWhereCust = gstrWhereCust & Chr$(39)
        Else
            gstrWhereCust = gstrWhereCust & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!FirstName) Then
        If IsNothing(gstrWhereCust) Then
            gstrWhereCust = "[FirstName] Like " & Chr$(39) & Me!FirstName
        Else
            gstrWhereCust = gstrWhereCust & " AND [FirstName] Like " & Chr$(39) & Me!FirstName
        End If
        If Right$(Me!FirstName, 1) = "*" Then
            gstrWhereCust = gstrWhereCust & Chr$(39)
        Else
            gstrWhereCust = gstrWhereCust & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!City) Then
        If IsNothing(gstrWhereCust) Then
            gstrWhereCust = "[City] Like " & Chr$(39) & Me!City
        Else
            gstrWhereCust = gstrWhereCust & " AND [City] Like " & Chr$(39) & Me!City
        End If
        If Right$(Me!City, 1) = "*" Then
            gstrWhereCust = gstrWhereCust & Chr$(39)
        Else
            gstrWhereCust = gstrWhereCust & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!State) Then
        If IsNothing(gstrWhereCust) Then
            gstrWhereCust = "[StateOrProvince] Like " & Chr$(39) & Me!State
        Else
            gstrWhereCust = gstrWhereCust & " AND [StateOrProvince] Like " & Chr$(39) & Me!State
        End If
        If Right$(Me!State, 1) = "*" Then
            gstrWhereCust = gstrWhereCust & Chr$(39)
        Else
            gstrWhereCust = gstrWhereCust & "*" & Chr$(39)
        End If
    End If

    If Not IsNothing(Me!Country) Then
        If IsNothing(gstrWhereCust) Then
            gstrWhereCust = "[Country] Like " & Chr$(39) & Me!Country
        Else
            gstrWhereCust = gstrWhereCust & " AND [Country] Like " & Chr$(39) & Me!Country
        End If
        If Right$(Me!Country, 1) = "*" Then
            gstrWhereCust = gstrWhereCust & Chr$(39)
        Else
            gstrWhereCust = gstrWhereCust & "*" & Chr$(39)
        End If
    End If

    ' If no criteria, then nothing to do...
    If IsNothing(gstrWhereCust) Then
        MsgBox "No criteria specified.", vbExclamation, "Microsoft Press"
        Exit Sub
    End If

    ' Hide myself and turn on Hourglass
    Me.Visible = False
    DoCmd.Hourglass True
    If IsLoaded("frmCustomers") Then  ' If customers form already open (it should be),
        ' Then just filter it
        Forms!frmCustomers.SetFocus
        DoCmd.ApplyFilter , gstrWhereCust
        If Forms!frmCustomers.Recordset.RecordCount = 0 Then
            DoCmd.Hourglass False
            MsgBox "No Customers meet your criteria", vbExclamation, "Microsoft Press"
            DoCmd.ShowAllRecords
            If gintIsAdmin Then
                Forms!frmCustomers!cmdAddNew.Visible = True
                Else
                Forms!frmCustomers!cmdAddNew.Visible = False
            End If
            Forms!frmCustomers!cmdShowAll.Visible = False
            Me.Visible = True
            Exit Sub
        End If

        Forms!frmCustomers!cmdAddNew.Visible = False
        Forms!frmCustomers!cmdShowAll.Visible = True
        DoCmd.Hourglass False
    Else
        ' Find out if any customers satisfy the Where clause
        DoCmd.Hourglass True
        DoCmd.OpenForm FormName:="frmCustomers", WhereCondition:=gstrWhereCust, _
            WindowMode:=acHidden
        If Forms!frmCustomers.Recordset.RecordCount = 0 Then
            DoCmd.Hourglass False
            MsgBox "No Customers meet your criteria", vbExclamation, "Microsoft Press"
            DoCmd.Close acForm, "frmCustomers"
            Me.Visible = True
            Exit Sub
        End If
        DoCmd.Hourglass False
        Forms!frmCustomers!cmdAddNew.Visible = False
        Forms!frmCustomers!cmdShowAll.Visible = True
        Forms!frmCustomer.Visible = True
    End If
    ' Close me, and we're done
    DoCmd.Close acForm, Me.Name


End Sub
