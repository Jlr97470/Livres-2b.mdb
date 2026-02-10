Operation =1
Option =0
Begin InputTables
    Name ="tblStores"
End
Begin OutputColumns
    Expression ="tblStores.StoreID"
    Expression ="tblStores.StoreName"
    Alias ="CityState"
    Expression ="[City] & \", \" & [StateOrProvince]"
End
Begin OrderBy
    Expression ="[tblStores].[StateOrProvince]"
    Flag =0
    Expression ="[tblStores].[City]"
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
    Begin
        dbText "Name" ="CityState"
        dbMemo "Caption" ="City, State"
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
        Top =2
        Name ="tblStores"
        Name =""
    End
End
