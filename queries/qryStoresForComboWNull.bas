dbMemo "SQL" ="SELECT NullLongID As StoreID, \"<No specific store>\" As StoreName, Null As City"
    "State, Null As State From ztblNullEntry\015\012UNION SELECT StoreID, StoreName, "
    "[City] & \", \" & [StateOrProvince] AS CityState, StateOrProvince As State\015\012"
    "FROM tblStores\015\012ORDER BY State, CityState, StoreName;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="StoreID"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="StoreName"
        dbInteger "ColumnWidth" ="2820"
        dbBoolean "ColumnHidden" ="0"
    End
End
