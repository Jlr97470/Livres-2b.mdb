Attribute VB_Name = "modBooks"
Option Compare Database
Option Explicit

Public Function OrderThis(strISBN As String) As Integer
Dim cn As ADODB.Connection, rcdC As New ADODB.Recordset
Dim rcdO As New ADODB.Recordset, rcdOD As New ADODB.Recordset
Dim strSQL As String, intTrans As Integer, lngOrderID As Long
On Error GoTo OrderThis_Err
    ' Make sure they're signed on
    If glngCurrentCustID = 0 Then
        If vbYes = MsgBox("You must be signed on to order a book." & _
            "  Would you like to sign on now?", vbQuestion + vbYesNo, "Microsoft Press") Then
            ' Try to sign them on - tell signon form to not go to frmMain
            DoCmd.OpenForm FormName:="frmSignon", WindowMode:=acDialog, OpenArgs:="NotMain"
            ' If still not signed on, then bail
            If glngCurrentCustID = 0 Then Exit Function
            Else
            Exit Function
        End If
    End If
    ' Get a copy of the current connection
    Set cn = CurrentProject.Connection
    cn.BeginTrans
    intTrans = True
    ' Set up to find the latest open order for this customer
    strSQL = "SELECT * FROM tblOrders" & _
        " WHERE [CustomerID] = " & glngCurrentCustID & _
        " AND [Printed] = False" & _
        " AND [OrderDate] = " & _
        "(Select Max([OrderDate]) From tblOrders As tblO2" & _
        " WHERE tblO2.[CustomerID] = " & glngCurrentCustID & _
        " AND tblO2.[Printed] = False);"
    rcdO.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    ' If no open order found, then...
    If rcdO.RecordCount = 0 Then
        ' Go find the customer record and add a new order
        strSQL = "SELECT * FROM tblCustomers WHERE [CustomerID] = " & glngCurrentCustID
        rcdC.Open strSQL, cn, adOpenKeyset, adLockOptimistic
        If rcdC.RecordCount = 0 Then
            Err.Raise 9001, , "Unexpected error.  Customer record for customer number " & glngCurrentCustID & " not found."
        End If
        rcdO.AddNew
        rcdO!CustomerID = glngCurrentCustID
        rcdO!OrderDate = Date
        rcdO!ShipName = rcdC!FirstName & IIf(IsNull(rcdC![MiddleInit]), "", " " & rcdC![MiddleInit] & ".") & " " & rcdC!LastName
        rcdO!ShipAddress = rcdC!Address
        rcdO!ShipCity = rcdC!City
        rcdO!ShipStateOrProvince = rcdC!StateOrProvince
        rcdO!ShipPostalCode = rcdC!PostalCode
        rcdO!ShipCountry = rcdC!Country
        rcdO!ShipPhoneNumber = rcdC!PhoneNumber
        rcdO!PayBy = rcdC!PayBy
        rcdO!CCNumber = rcdC!CCNumber
        rcdO!Printed = False
        ' Save the generated order number for detail lines
        lngOrderID = rcdO!OrderID
        ' Add the master order row
        rcdO.Update
        rcdC.Close
    Else
        lngOrderID = rcdO!OrderID
    End If
    rcdO.Close
    ' See if this book already in this order
    strSQL = "SELECT * FROM tblOrderDetails" & _
        " WHERE [OrderID] = " & lngOrderID & _
        " AND [ISBNNumber] = '" & strISBN & "';"
    rcdOD.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    If rcdOD.RecordCount = 0 Then
        ' Nope, add a new row
        rcdOD.AddNew
        rcdOD![OrderID] = lngOrderID
        rcdOD![ISBNNumber] = strISBN
    End If
    rcdOD![Quantity] = rcdOD![Quantity] + 1
    rcdOD.Update
    rcdOD.Close
    cn.CommitTrans
    intTrans = False
    OrderThis = True
    Exit Function

OrderThis_Err:
    MsgBox "Unexpected error: " & Err & ", " & Error
    If intTrans Then cn.RollbackTrans
    OrderThis = False
    Exit Function

End Function
