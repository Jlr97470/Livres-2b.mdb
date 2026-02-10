Operation =1
Option =0
Begin InputTables
    Name ="tblStores"
    Name ="tblCustomers"
    Name ="tblOrders"
End
Begin OutputColumns
    Expression ="tblCustomers.LastName"
    Expression ="tblCustomers.FirstName"
    Expression ="tblCustomers.MiddleInit"
    Expression ="tblCustomers.Address"
    Expression ="tblCustomers.City"
    Expression ="tblCustomers.StateOrProvince"
    Expression ="tblCustomers.PostalCode"
    Expression ="tblCustomers.Country"
    Expression ="tblCustomers.PhoneNumber"
    Alias ="CustPayBy"
    Expression ="[tblCustomers].[PayBy]"
    Alias ="CustCCNumber"
    Expression ="[tblCustomers].[CCNumber]"
    Expression ="tblOrders.OrderID"
    Expression ="tblOrders.CustomerID"
    Expression ="tblOrders.StoreID"
    Expression ="tblOrders.OrderDate"
    Expression ="tblOrders.ShipName"
    Expression ="tblOrders.ShipAddress"
    Expression ="tblOrders.ShipCity"
    Expression ="tblOrders.ShipStateOrProvince"
    Expression ="tblOrders.ShipPostalCode"
    Expression ="tblOrders.ShipCountry"
    Expression ="tblOrders.ShipPhoneNumber"
    Expression ="tblOrders.PayBy"
    Expression ="tblOrders.CCNumber"
    Expression ="tblOrders.Printed"
    Expression ="tblStores.StoreName"
    Alias ="StAddress"
    Expression ="[tblStores].[Address]"
    Alias ="StCity"
    Expression ="[tblStores].[City]"
    Alias ="StPostalCode"
    Expression ="[tblStores].[PostalCode]"
    Alias ="StStateOrProvince"
    Expression ="[tblStores].[StateOrProvince]"
    Alias ="StCountry"
    Expression ="[tblStores].[Country]"
End
Begin Joins
    LeftTable ="tblCustomers"
    RightTable ="tblOrders"
    Expression ="[tblCustomers].[CustomerID]=[tblOrders].[CustomerID]"
    Flag =1
    LeftTable ="tblStores"
    RightTable ="tblOrders"
    Expression ="[tblStores].[StoreID]=[tblOrders].[StoreID]"
    Flag =3
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
        Left =317
        Top =6
        Right =413
        Bottom =113
        Top =0
        Name ="tblStores"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =113
        Top =0
        Name ="tblCustomers"
        Name =""
    End
    Begin
        Left =175
        Top =6
        Right =271
        Bottom =113
        Top =0
        Name ="tblOrders"
        Name =""
    End
End
