Operation =1
Option =0
Begin InputTables
    Name ="tblClasses"
    Name ="tblCategories"
End
Begin OutputColumns
    Expression ="tblCategories.Category"
    Expression ="tblClasses.ClassName"
    Expression ="tblCategories.CategoryID"
End
Begin Joins
    LeftTable ="tblClasses"
    RightTable ="tblCategories"
    Expression ="[tblClasses].[ClassID]=[tblCategories].[ClassID]"
    Flag =1
End
Begin OrderBy
    Expression ="[tblCategories].[Category]"
    Flag =0
    Expression ="[tblClasses].[ClassName]"
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
        Left =172
        Top =6
        Right =268
        Bottom =83
        Top =0
        Name ="tblClasses"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =98
        Top =0
        Name ="tblCategories"
        Name =""
    End
End
