Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdAddNew_Click()
    ' This button disabled unless customer is signed on
    ' Save changes before creating a new one...
    If Me.Dirty Then
        ' Check for minimum data here...
        If IsNothing(Me!CustomerID) Or IsNothing(Me!PayBy) Then
            MsgBox "Must save current record with a valid customer ID and payment method before adding another.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
    End If
    ' Execute a Data Entry command
    DoCmd.RunCommand acCmdDataEntry
    ' Move the focus off this button
    Me![CustomerID].SetFocus
    Me!cmdAddNew.Visible = False
    Me!cmdShowAll.Visible = True
    Me!cmdSearch.Enabled = False

End Sub

Private Sub cmdCancel_Click()
    ' Cancel requested - clear edits and exit
    If Me.Dirty Then Me.Undo
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdClose_Click()
    ' Save any changes before closing
    If Me.Dirty Then
        If IsNothing(Me!CustomerID) Or IsNothing(Me!PayBy) Then
            MsgBox "Must have a customer ID and payment method to save a record.  Click Cancel to exit without saving.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
    End If

    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdPrint_Click()
    On Error Resume Next
    Screen.PreviousControl.SetFocus
    ' Save any changes
    If Me.Dirty Then
        ' Check for minimum data here...
        If IsNothing(Me!CustomerID) Or IsNothing(Me!PayBy) Then
            MsgBox "Must save current record with a valid customer and payment method before adding another.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
    End If
    ' Open the orders print options form...
    DoCmd.OpenForm "frmOrderPrintParms"
End Sub

Private Sub cmdSave_Click()
On Error GoTo SAExit
Dim lngKey As String
    ' If not "dirty", then nothing to do.
    If Me.Dirty Then
        If IsNothing(Me!CustomerID) Or IsNothing(Me!PayBy) Then
            MsgBox "Record must have a valid customer and payment method to save it.", vbCritical, "Microsoft Press"
            Me!CustomerID.SetFocus
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
        ' Save the new order ID
        lngKey = Me!OrderID
        ' Unfilter the form
        cmdShowAll_Click
        Me!OrderID.SetFocus
        ' Put form on the new order
        DoCmd.FindRecord lngKey
        Me!CustomerID.SetFocus
        Application.Echo True
    End If
SAExit:
    If Err <> 0 Then
        MsgBox Error$, vbCritical
    End If
    Exit Sub


End Sub

Private Sub cmdSearch_Click()
    ' Save any changes before doing a search...
    If Me.Dirty Then
        If IsNothing(Me!CustomerID) Or IsNothing(Me!PayBy) Then
            MsgBox "Must save current record with a valid customer and payment method before performing a search.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
    End If
    DoCmd.OpenForm "frmOrderOptions"

End Sub

Private Sub cmdShowAll_Click()
On Error GoTo SHacExit
    ' Save changes before unfiltering the form
    If Me.Dirty Then
        If IsNothing(Me!CustomerID) Or IsNothing(Me!PayBy) Then
            MsgBox "Must save current record with a valid customer and payment method before viewing all records.", vbCritical, "Microsoft Press"
            Exit Sub
        End If
        DoCmd.RunCommand acCmdSaveRecord
    End If
    Me![CustomerID].SetFocus
    DoCmd.ShowAllRecords
    gstrWhereOrder = ""
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

Private Sub cmdStoreSearch_Click()
    ' Open store search with a parameter to let it know we want any result
    ' returned here...
    DoCmd.OpenForm FormName:="frmStoreSearch", WindowMode:=acDialog, _
        OpenArgs:="Orders"
End Sub

Private Sub CustomerID_AfterUpdate()
    ' When customer number changes, ...
    '  copy over default shipping info.
    Me!ShipName = Me![FirstName] & IIf(IsNull([MiddleInit]), "", " " & [MiddleInit] & ".") & " " & Me!LastName
    Me!ShipAddress = Me!Address
    Me!ShipCity = Me!City
    Me!ShipStateOrProvince = Me!StateOrProvince
    Me!ShipPostalCode = Me!PostalCode
    Me!ShipCountry = Me!Country
    Me!ShipPhoneNumber = Me!PhoneNumber
    Me!PayBy = Me!CustPayBy
    Me!CCNumber = Me!CustCCNumber
    If Me!PayBy < 3 Then Me!CCNumber = Null
End Sub

Private Sub CustomerID_BeforeUpdate(Cancel As Integer)
Dim strName As String
    ' Customer ID must = current signed-on customer
    If Me!CustomerID <> glngCurrentCustID Then
        If glngCurrentCustID <> 0 Then strName = DLookup("[LastUserName]", "tblLocalValues")
        MsgBox "You selected an invalid customer for this order." & vbCrLf & vbCrLf & _
            "You can create orders only for the currently signed-on customer - " & _
            strName & vbCrLf & vbCrLf & _
            "Exit the application and sign on as the customer for whom you wish to create an order.", _
            vbExclamation, "Microsoft Press"
        Cancel = True
    End If
End Sub

Private Sub CustomerID_NotInList(NewData As String, Response As Integer)
Dim strUser As String, strLast As String, strFirst As String, intComma As Integer
Dim intReturn As Integer, varName As Variant
' This is an interesting example of how to handle NotInList, however...
' Since the BeforeUpdate event won't allow any ID other than the signed on
' customer, and you have to be in the database to be signed on, this will never
' complete successfully.  You can add a new customer, but then you'll have to exit
' and sign on again to actually build a new order.  Oh well.
    strUser = NewData
    intComma = InStr(strUser, ",")
    If intComma = 0 Then
        strLast = strUser
        Else
        strLast = Left(strUser, intComma - 1)
        strFirst = Mid(strUser, intComma + 2)
    End If
    intReturn = MsgBox("Customer " & strUser & _
        " is not in the system.  Do you want to add this customer?", _
        vbQuestion + vbYesNo, "Microsoft Press")
    If intReturn = vbYes Then
        DoCmd.OpenForm FormName:="frmCustAdd", _
            DataMode:=acAdd, WindowMode:=acDialog, OpenArgs:=strUser
        If IsNull(DLookup("CustomerID", "tblCustomers", "[LastName] = '" & strLast & "'")) Then
            Response = acDataErrContinue
            Else
            Response = acDataErrAdded
        End If
        Exit Sub
    End If
    Response = acDataErrDisplay

End Sub

Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    ' Filter has been applied - save it in the global filter string
    gstrWhereOrder = Me.Filter
    If IsNothing(Me.Filter) Or ApplyType = acShowAllRecords Then
    ' Nothing in the filter, so set up like they clicked Show All
        gstrWhereOrder = ""
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

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If IsNothing(Me!ShipName) Then
        CustomerID_AfterUpdate
    End If
End Sub

Private Sub Form_Current()
    ' If order already printed or not for current signed on customer, ...
    If (Me!Printed) Or (Me!CustomerID <> glngCurrentCustID) Then
        ' Lock all controls and disable some buttons
        SetLocked True
        Me!cmdSave.Enabled = False
        Me!cmdAddNew.Enabled = False
        Me!cmdAddNew.Visible = False
        Me!cmdStoreSearch.Enabled = False
        Else
        SetLocked False
        Me!cmdSave.Enabled = True
        Me!cmdAddNew.Enabled = True
        Me!cmdStoreSearch.Enabled = True
    End If
    ' If Payment method is cash or check
    If Me!PayBy < 3 Then
        ' Then don't display CCNumber
        Me!CCNumber.Visible = False
        Else
        Me!CCNumber.Visible = True
    End If
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
        .Heading = "Sample Balloon Help for Orders"
        ' This is a tip, so use the right icon
        .Icon = msoIconTip
        ' Set up labels to display helpful hints
        .Labels(1).Text = "Click the Search button in the top row to look for a different order."
        .Labels(2).Text = "When viewing your own orders, you can change the data as long as you haven't printed the order.  Once printed, an order is ""final""."
        .Labels(3).Text = "Click Search> next to Store to look for a store near you."
        .Labels(4).Text = "When you have finished selecting books, click Print to print out your order and mark it finalized."
        ' Set up title text inside the balloon
        .Text = "Here's what you can do on the Orders form:"
        ' Activate the balloon
        .Show
        ' Code will stop here until user dismisses the baloon
    End With
End Sub

Private Sub Form_Filter(Cancel As Integer, FilterType As Integer)
    ' User about to build a filter, so clear previous
    Me.Filter = ""
    gstrWhereOrder = ""
End Sub

Private Sub Form_Load()
    ' If no customer signed on, then disable some stuff
    If IsNothing(glngCurrentCustID) Then
        Me!cmdAddNew.Enabled = False
        Me.AllowAdditions = False
        Exit Sub
        Else
        ' .. otherwise set the default value for customer id (not the quotes!)
        Me!CustomerID.DefaultValue = CStr(glngCurrentCustID)
    End If
    ' If called from frmCustomer, then go right into add new mode...
    If Me.OpenArgs = "New" Then cmdAddNew_Click
End Sub

Sub SetLocked(intTF As Integer)
    ' This is a generic routine to lock/unlock all controls
    ' NOTE:  This is also PUBLIC for this form so it can be called
    ' to lock all controls after the order successfully prints!
    ' See "frmOrderPrintParms" for an example.
    Me!CustomerID.Locked = intTF
    Me!ShipName.Locked = intTF
    Me!ShipAddress.Locked = intTF
    Me!ShipCity.Locked = intTF
    Me!ShipStateOrProvince.Locked = intTF
    Me!ShipPostalCode.Locked = intTF
    Me!ShipCountry.Locked = intTF
    Me!ShipPhoneNumber.Locked = intTF
    Me!cmdStoreSearch.Enabled = Not (intTF)
    Me!StoreID.Locked = intTF
    Me!PayBy.Locked = intTF
    Me!CCNumber.Locked = intTF
    Me!fsubOrderBooks.Locked = intTF
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

Private Sub ShipName_Enter()
    ' Nothing to do if there's already something in ShipName
    If IsNothing(Me!ShipName) Then
        ' If form isn't "dirty" Then
        If Not (Me.Dirty) Then
            ' Check for a customer ID default
            If IsNothing(Me!CustomerID.DefaultValue) Then
                ' No customer ID, so bail out
                Exit Sub
            End If
            ' Make the form "dirty" by force-setting CustomerID
            ' This makes the autolookup fire so we can get
            ' a good Ship To Name
            Me!CustomerID = Me!CustomerID.DefaultValue
        End If
        CustomerID_AfterUpdate
    End If
End Sub
