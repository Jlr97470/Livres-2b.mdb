Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit

Sub cmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdSell_Click()
Dim cn As ADODB.Connection, rcdO As New ADODB.Recordset, rcdOD As New ADODB.Recordset
Dim rcdC As New ADODB.Recordset, rcdB As New ADODB.Recordset
Dim varBeginDate As Variant, intNumDays As Integer, varCurrentDate As Variant
Dim intNumOrders As Integer, intNumBooks As Integer, intNumCopies As Integer
Dim intCountCustomers As Integer, intCountBooks As Integer
Dim intOrdersToday As Integer, intBooksThisOrder As Integer
Dim intO As Integer, intOD As Integer, intCust As Integer
Dim intBook As Integer, intBookIncr As Integer
Dim lngOrderID As Long, varRtn As Variant

    ' Capture values from the form
    varBeginDate = Me!BeginDate
    intNumDays = Me!NumDays
    intNumOrders = Me!NumOrders
    intNumBooks = Me!NumBooks
    intNumCopies = Me!NumCopies

    ' Initialize the randomizer on system clock
    Randomize

    ' Open all recordsets and set the current connection
    Set cn = CurrentProject.Connection
    rcdO.Open "tblOrders", cn, adOpenKeyset, adLockOptimistic
    rcdOD.Open "tblOrderDetails", cn, adOpenKeyset, adLockOptimistic
    rcdC.Open "tblCustomers", cn, adOpenKeyset, adLockOptimistic
    rcdB.Open "Select * From tblBooks Where Not [OutOfPrint];", cn, adOpenKeyset, adLockOptimistic

    ' Find out how many customers and books we have to work with
    rcdC.MoveLast
    intCountCustomers = rcdC.RecordCount
    If intCountCustomers = 0 Then
        MsgBox "Can't generate any orders -- no customers in this database!", vbCritical, "Microsoft Press"
        Exit Sub
    End If

    rcdB.MoveLast
    intCountBooks = rcdB.RecordCount
    If intCountBooks = 0 Then
        MsgBox "Can't generate any orders -- no books in this database!", vbCritical, "Microsoft Press"
        Exit Sub
    End If
    If intCountBooks < intNumBooks Then
        MsgBox "You have asked for more books per order than there are books in the database." & _
            " Books per order must not exceed " & intCountBooks & ".", vbExclamation, "Microsoft Press"
        Exit Sub
    End If

    ' Set up to bail if something funny happens (it shouldn't)
    On Error GoTo BailOut
    ' Protect everything with a transaction
    cn.BeginTrans
    ' Turn on the hourglass
    DoCmd.Hourglass True

    ' Initialize the status bar
    varRtn = SysCmd(acSysCmdInitMeter, "Creating orders...", intNumDays)

    ' Set up the outside loop on dates
    For varCurrentDate = varBeginDate To (varBeginDate + intNumDays)
        ' Figure out max orders for today
        intOrdersToday = Int((intNumOrders * Rnd) + 1)
        ' Loop to create orders
        For intO = 1 To intOrdersToday
            ' Start a new order
            rcdO.AddNew
            ' Randomly grab a customer record
            intCust = Int(intCountCustomers * Rnd)
            rcdC.MoveFirst
            If intCust > 0 Then rcdC.Move intCust
            ' Move customer info to the order record
            rcdO!CustomerID = rcdC!CustomerID
            rcdO!OrderDate = varCurrentDate
            rcdO!ShipName = rcdC!FirstName & " " & rcdC!LastName
            rcdO!ShipAddress = rcdC!Address
            rcdO!ShipCity = rcdC!City
            rcdO!ShipStateOrProvince = rcdC!StateOrProvince
            rcdO!ShipPostalCode = rcdC!PostalCode
            rcdO!ShipCountry = rcdC!Country
            rcdO!ShipPhoneNumber = rcdC!PhoneNumber
            rcdO!PayBy = rcdC!PayBy
            rcdO!CCNumber = rcdC!CCNumber
            rcdO!Printed = True
            ' Save the generated order number for detail lines
            lngOrderID = rcdO!OrderID
            ' Add the master order row
            rcdO.Update

            ' Calculate number of books for this order
            intBooksThisOrder = Int((intNumBooks * Rnd) + 1)
             ' Randomly grab a book record
            intBook = Int(intCountBooks * Rnd)
            rcdB.MoveFirst
            If intBook > 0 Then rcdB.Move intBook
            ' Calculate the "jump" interval to the next book
            ' Using integer divide here to always get a truncated result
            intBookIncr = intCountBooks \ intBooksThisOrder
            ' Make sure result is at least 1
            If intBookIncr = 0 Then intBookIncr = 1
            ' Loop to add books
            For intOD = 1 To intBooksThisOrder
                ' Add a detail row
                rcdOD.AddNew
                rcdOD!OrderID = lngOrderID
TryAgain:
                rcdOD!ISBNNumber = rcdB!ISBNNumber
                rcdOD!Quantity = Int((intNumCopies * Rnd) + 1)
                ' Set max possible discount based on quantity
                Select Case rcdOD!Quantity
                    Case Is < 5
                        rcdOD!Discount = 0
                    Case Is < 11
                        rcdOD!Discount = 0.05
                    Case Is < 21
                        rcdOD!Discount = 0.1
                    Case Else
                        rcdOD!Discount = 0.15
                End Select
                ' Set a specific trap for duplicate record
                On Error Resume Next
                rcdOD.Update
                If Err <> 0 Then
                    If Err <> -2147217887 Then GoTo BailOut
                    ' Ooops -- we must have looped around --
                    ' ... move one more to avoid the duplicate and go try again
                    rcdB.MoveNext
                    If rcdB.EOF Then rcdB.MoveFirst
                    On Error GoTo BailOut
                    GoTo TryAgain
                End If
                ' Reset the generic error trap
                On Error GoTo BailOut
                ' Skip forward to "next" book
                rcdB.Move intBookIncr
                ' If bounced past the end, the start over
                '  with first book
                If rcdB.EOF Then rcdB.MoveFirst
            Next intOD
        Next intO
        ' Update the status bar
        varRtn = SysCmd(acSysCmdUpdateMeter, Int(varCurrentDate - varBeginDate))
    ' Loop to the next date value
    Next varCurrentDate
    ' Done - commit all updates
    cn.CommitTrans
    ' Clear the status bar
    varRtn = SysCmd(acSysCmdClearStatus)
    ' Done with error trapping, too
    On Error GoTo 0
    ' Be nice and close everything up
    rcdO.Close
    rcdOD.Close
    rcdB.Close
    rcdC.Close
    ' Turn off the hourglass
    DoCmd.Hourglass False
    DoCmd.Close acForm, Me.Name
    Exit Sub

BailOut:
    MsgBox "Unexpected error: " & Err & ", " & Error
    cn.RollbackTrans
    ' Turn off the hourglass
    DoCmd.Hourglass False
    varRtn = SysCmd(acSysCmdClearStatus)
    Exit Sub

End Sub
