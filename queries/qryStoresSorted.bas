Operation =1
Option =0
Begin InputTables
    Name ="tblStores"
End
Begin OutputColumns
    Expression ="tblStores.StoreID"
    Expression ="tblStores.PostalCode"
    Expression ="tblStores.StoreName"
    Expression ="tblStores.Address"
    Expression ="tblStores.City"
    Expression ="tblStores.StateOrProvince"
    Expression ="tblStores.Country"
    Expression ="tblStores.PhoneNumber"
    Expression ="tblStores.FaxNumber"
    Expression ="tblStores.Email"
    Expression ="tblStores.WebAddress"
    Alias ="PhoneCode"
    Expression ="Left([PhoneNumber],5)"
End
Begin OrderBy
    Expression ="[tblStores].[PostalCode]"
    Flag =0
    Expression ="[tblStores].[StoreName]"
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
        Top =7
        Name ="tblStores"
        Name =""
    End
End
