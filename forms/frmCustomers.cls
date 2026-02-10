Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdAddNew_Click()
    ' This command disabled if not an admin user...
    ' .. but you can establish yourself as a new admin user at
    ' signon -- what can I say? -- a security "hole"
    ' Save any changes first..
    If Me.Dirty Then
        If IsNothing(Me!FirstName) Or IsNothing(Me!LastName) Then
            MsgBox "Must save current record with a valid first and last name before adding another.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
    End If
    ' Execute a Form, Records, Data Entry command
    DoCmd.RunCommand acCmdDataEntry
    Me![FirstName].SetFocus
    Me!cmdAddNew.Visible = False
    Me!cmdShowAll.Visible = True
    Me!cmdSearch.Enabled = False

End Sub

Private Sub cmdCancel_Click()
    ' Cancel requested - clear edits and close
    If Me.Dirty Then Me.Undo
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdClose_Click()
    ' Close requested - save any changes first
    If Me.Dirty Then
        If IsNothing(Me!FirstName) Or IsNothing(Me!LastName) Then
            MsgBox "Must have a first and last name to save a record.  Click Cancel to exit without saving.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
    End If

    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdNewOrder_Click()
    ' This button disabled if not on the signed-on customer row
    ' Save any changes
    If Me.Dirty Then
        If IsNothing(Me!FirstName) Or IsNothing(Me!LastName) Then
            MsgBox "Must save current record with a first name and last name before working with orders.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
    End If

    gstrWhereOrder = "[CustomerID] = " & glngCurrentCustID

    DoCmd.OpenForm FormName:="frmOrders", WhereCondition:=gstrWhereOrder, _
        OpenArgs:="New"
End Sub

Private Sub cmdReports_Click()
    ' Want some reports?  Ok, open the options form...
    DoCmd.OpenForm FormName:="frmCustPrintParms"
End Sub

Private Sub cmdSave_Click()
On Error GoTo SAExit
Dim lngCustID As Long
    ' If not "dirty" then nothing to do.
    If Me.Dirty Then
        If IsNothing(Me!FirstName) Or IsNothing(Me!LastName) Then
            MsgBox "Record must have a first and last name to save it.", vbCritical, "Microsoft Press"
            Me!FirstName.SetFocus
            Exit Sub
        End If
        ' Save the current "dirty" record
        DoCmd.RunCommand acCmdSaveRecord
        Else
        Exit Sub
    End If

    ' If we were in "add new" mode, then...
    If Me!cmdShowAll.Visible = True And Me!cmdSearch.Enabled = False Then
        Application.Echo False
        ' Save ID of new customer
        lngCustID = Me!CustomerID
        ' Display all customers
        cmdShowAll_Click
        Me!CustomerID.SetFocus
        ' .. but move to row of added customer
        DoCmd.FindRecord lngCustID
        Me!FirstName.SetFocus
        Application.Echo True
    End If
SAExit:
    ' Skip actions canceled by Before Update
    If Err = 2501 Then Exit Sub
    If Err <> 0 Then
        MsgBox Error$, vbCritical
    End If
    Exit Sub


End Sub

Private Sub cmdSearch_Click()
    ' Save any changes before searching...
    If Me.Dirty Then
        If IsNothing(Me!FirstName) Or IsNothing(Me!LastName) Then
            MsgBox "Must save current record with a first and last name before performing a search.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
    End If
    DoCmd.OpenForm "frmCustSearch"

End Sub

Private Sub cmdShowAll_Click()
On Error GoTo SHacExit
    ' Save any changes
    If Me.Dirty Then
        If IsNothing(Me!FirstName) Or IsNothing(Me!LastName) Then
            MsgBox "Must save current record with a first and last name before viewing all customers.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
    End If
    Me![FirstName].SetFocus
    ' Unfilter the form
    DoCmd.ShowAllRecords
    ' .. and clear the global filter string
    gstrWhereCust = ""
    If gintIsAdmin Then
        Me!cmdAddNew.Visible = True
        Else
        Me!cmdAddNew.Visible = False
    End If
    Me!cmdShowAll.Visible = False
    Me!cmdSearch.Enabled = True
SHacExit:
    If Err <> 0 Then MsgBox Error$, vbCritical
    Exit Sub
End Sub

Private Sub cmdEditOrders_Click()
    ' Save any changes
    If Me.Dirty Then
        If IsNothing(Me!FirstName) Or IsNothing(Me!LastName) Then
            MsgBox "Must save current record with a first name and last name " & _
            "before working with orders.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
    End If
    ' See if any orders for this customer
    gstrWhereOrder = "[CustomerID] = " & Me!CustomerID
    If 0 = DCount("*", "tblOrders", gstrWhereOrder) Then
        MsgBox "No orders exist for this customer.", _
            vbInformation, "Microsoft Press"
        Exit Sub
    End If
    ' Yes - open orders filtered
    DoCmd.OpenForm FormName:="frmOrders", WhereCondition:=gstrWhereOrder
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Dim intPLen As Integer
    ' Customer records with a credit card number or admin status
    '  must also have a password
    If Not IsNothing(Me!CCNumber) Or Me!Admin Then
        intPLen = Len(Nz(Me!Password))
        If intPLen < 5 Then
            MsgBox "A customer with a credit card number or admin status must also " & _
                "have a password." & vbCrLf & vbCrLf & _
                "Enter a password at least 5 characters long or clear the credit " & _
                "card number and admin status.", _
                vbExclamation, "Microsoft Press"
            Cancel = True
        End If
    End If

End Sub

Private Sub Form_Current()
    ' If this is the signed-on customer
    If (Me!CustomerID = glngCurrentCustID) Then
        ' .. then allow new orders and unlock all controls
        Me!cmdNewOrder.Enabled = True
        SetLocked False
        Else
        ' Else disable orders and lock
        Me!cmdNewOrder.Enabled = False
        SetLocked True
    End If
    ' If this is an admin user...
    If gintIsAdmin Then
        ' Unlock all controls
        SetLocked False
    End If
    ' Show CCNumber and set input mask based on PayBy value
    Select Case Me!PayBy
        Case 1, 2
            Me!CCNumber.Visible = False
        Case 3
            Me!CCNumber.Visible = True
            Me!CCNumber.InputMask = "0000\-000000\-00000;0;_"
        Case Else
            Me!CCNumber.Visible = True
            Me!CCNumber.InputMask = "0000\-0000\-0000\-0000;0;_"
    End Select

End Sub

Private Sub Form_Delete(Cancel As Integer)
    ' Don't allow delete of the currently signed on customer
    If Me!CustomerID = glngCurrentCustID Then
        MsgBox "You can't delete the customer record for the currently signed on customer." & _
            vbCrLf & vbCrLf & "Exit the application and sign on as a different customer " & _
            "with admin status to delete this record.", vbCritical, "Microsoft Press"
        Cancel = True
    End If
    ' Don't allow customer delete unless Admin
    If Not gintIsAdmin Then
        MsgBox "You are not authorized to delete customers.", vbExclamation, "Microsoft Press"
        Cancel = True
    End If

End Sub

Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    ' Filter has been applied - save it in the global filter string
    gstrWhereCust = Me.Filter
    If IsNothing(Me.Filter) Or ApplyType = acShowAllRecords Then
    ' Nothing in the filter, so set up like they clicked Show All
        gstrWhereCust = ""
        ' Show the Add button only if Admin user
        If gintIsAdmin Then Me!cmdAddNew.Visible = True
        Me!cmdShowAll.Visible = False
        Me!cmdSearch.Enabled = True
        Me!cmdSave.Enabled = True
        Else
        Me!cmdAddNew.Visible = False
        Me!cmdShowAll.Visible = True
    End If
End Sub

Private Sub Form_Filter(Cancel As Integer, FilterType As Integer)
    ' User about to build a filter, so clear previous
    Me.Filter = ""
    gstrWhereCust = ""
End Sub

Public Sub FormHelp()
' Custom assistant help method for this form
Dim balCurrent As Balloon

    ' Create a new assistant baloon
    Set balCurrent = Assistant.NewBalloon

    With balCurrent
        ' Set balloon type to bullets
        .BalloonType = msoBalloonTypeBullets
        ' Set mode to auto-close as soon as the user clicks anywhere else
        .Mode = msoModeAutoDown
        ' Add an OK button just for looks
        .Button = msoButtonSetOK
        ' Define the balloon title
        .Heading = "Sample Balloon Help for Customers"
        ' This is a tip, so use the right icon
        .Icon = msoIconTip
        ' Set up labels to display helpful hints
        .Labels(1).Text = "Click Search to look for a different customer."
        .Labels(2).Text = "When viewing your own customer record, you can change the data and create new orders."
        .Labels(3).Text = "Click Reports to see available customer report options."
        .Labels(4).Text = "Click Edit Orders to see all orders for the current customer."
        ' Set up title text inside the balloon
        .Text = "You can do lots of things with the Customers form!"
        ' Activate the balloon
        .Show
        ' Code will stop here until user dismisses the baloon
    End With
End Sub

Private Sub Form_Load()
    ' Check for admin status
    If Not gintIsAdmin Then
        ' Not an admin user, so disable a few things...
        Me!cmdAddNew.Enabled = False
        Me!cmdAddNew.Visible = False
        Me!cmdSave.Visible = False
    End If
End Sub

Private Sub PayBy_AfterUpdate()
    ' Reveal CCNumber and set Input Mask based on payment method
    Select Case Me!PayBy
        Case 1, 2
            Me!CCNumber = Null
            Me!CCNumber.Visible = False
        Case 3
            Me!CCNumber.Visible = True
            Me!CCNumber.InputMask = "0000\-000000\-00000;0;_"
        Case Else
            Me!CCNumber.Visible = True
            Me!CCNumber.InputMask = "0000\-0000\-0000\-0000;0;_"
    End Select
End Sub

Private Sub SetLocked(intLocked As Integer)
' Generic routine to lock and unlock all controls
' based on current user or admin status
    Me!FirstName.Locked = intLocked
    Me!MiddleInit.Locked = intLocked
    Me!LastName.Locked = intLocked
    Me!Password.Locked = intLocked
    Me!Address.Locked = intLocked
    Me!Admin.Locked = intLocked
    Me!City.Locked = intLocked
    Me!StateOrProvince.Locked = intLocked
    Me!PostalCode.Locked = intLocked
    Me!Country.Locked = intLocked
    Me!PhoneNumber.Locked = intLocked
    Me!EmailAddress.Locked = intLocked
    Me!PayBy.Locked = intLocked
    Me!CCNumber.Locked = intLocked

End Sub
