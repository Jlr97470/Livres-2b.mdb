Operation =1
Option =0
Begin InputTables
    Name ="tblCustomers"
End
Begin OutputColumns
    Expression ="tblCustomers.CustomerID"
    Expression ="tblCustomers.LastName"
    Expression ="tblCustomers.FirstName"
    Expression ="tblCustomers.MiddleInit"
    Expression ="tblCustomers.Address"
    Expression ="tblCustomers.City"
    Expression ="tblCustomers.StateOrProvince"
    Expression ="tblCustomers.PostalCode"
    Expression ="tblCustomers.Country"
    Expression ="tblCustomers.PhoneNumber"
    Expression ="tblCustomers.EmailAddress"
    Expression ="tblCustomers.Password"
    Expression ="tblCustomers.PayBy"
    Expression ="tblCustomers.CCNumber"
    Expression ="tblCustomers.Admin"
End
Begin OrderBy
    Expression ="[tblCustomers].[LastName]"
    Flag =0
    Expression ="[tblCustomers].[FirstName]"
    Flag =0
    Expression ="[tblCustomers].[MiddleInit]"
    Flag =0
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
        Right =134
        Bottom =113
        Top =11
        Name ="tblCustomers"
        Name =""
    End
End
