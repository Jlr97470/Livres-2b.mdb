Operation =1
Option =0
Begin InputTables
    Name ="tblCustomers"
End
Begin OutputColumns
    Alias ="CustName"
    Expression ="[LastName] & \", \" & [FirstName] & \" \" & [MiddleInit] & \".\""
    Expression ="tblCustomers.CustomerID"
    Expression ="tblCustomers.Password"
    Expression ="tblCustomers.Admin"
End
Begin OrderBy
    Expression ="[LastName] & \", \" & [FirstName] & \" \" & [MiddleInit] & \".\""
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="CustName"
        dbInteger "ColumnWidth" ="2370"
        dbBoolean "ColumnHidden" ="0"
    End
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
