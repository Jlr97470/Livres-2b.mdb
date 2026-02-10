Operation =1
Option =0
Begin InputTables
    Name ="tblCustomers"
    Name ="tblOrders"
    Name ="tblBooks"
    Name ="tblOrderDetails"
End
Begin OutputColumns
    Expression ="tblCustomers.*"
    Alias ="CustName"
    Expression ="[LastName] & \", \" & [FirstName] & \" \" & [MiddleInit] & \".\""
    Expression ="tblOrders.OrderID"
    Expression ="tblOrders.OrderDate"
    Expression ="tblOrderDetails.ISBNNumber"
    Expression ="tblOrderDetails.Quantity"
    Expression ="tblOrderDetails.Discount"
    Expression ="tblBooks.Title"
    Expression ="tblBooks.SuggPrice"
End
Begin Joins
    LeftTable ="tblCustomers"
    RightTable ="tblOrders"
    Expression ="[tblCustomers].[CustomerID]=[tblOrders].[CustomerID]"
    Flag =1
    LeftTable ="tblBooks"
    RightTable ="tblOrderDetails"
    Expression ="[tblBooks].[ISBNNumber]=[tblOrderDetails].[ISBNNumber]"
    Flag =1
    LeftTable ="tblOrders"
    RightTable ="tblOrderDetails"
    Expression ="[tblOrders].[OrderID]=[tblOrderDetails].[OrderID]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
Begin
End
Begin
    State =0
    Left =24
    Top =24
    Right =744
    Bottom =408
    Left =-1
    Top =-1
    Right =713
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =149
        Bottom =113
        Top =0
        Name ="tblCustomers"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =113
        Top =0
        Name ="tblOrders"
        Name =""
    End
    Begin
        Left =439
        Top =5
        Right =564
        Bottom =112
        Top =0
        Name ="tblBooks"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =113
        Top =0
        Name ="tblOrderDetails"
        Name =""
    End
End
