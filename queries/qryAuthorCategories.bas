Operation =1
Option =0
Begin InputTables
    Name ="tblCategories"
    Name ="tblBooks"
    Name ="tblAuthors"
    Name ="tblBookAuthor"
    Name ="tblBookCategories"
End
Begin OutputColumns
    Expression ="tblAuthors.AuthorID"
    Expression ="tblCategories.CategoryID"
End
Begin Joins
    LeftTable ="tblAuthors"
    RightTable ="tblBookAuthor"
    Expression ="[tblAuthors].[AuthorID]=[tblBookAuthor].[AuthorID]"
    Flag =1
    LeftTable ="tblBooks"
    RightTable ="tblBookAuthor"
    Expression ="[tblBooks].[ISBNNumber]=[tblBookAuthor].[ISBNNumber]"
    Flag =1
    LeftTable ="tblBooks"
    RightTable ="tblBookCategories"
    Expression ="[tblBooks].[ISBNNumber]=[tblBookCategories].[ISBNNumber]"
    Flag =1
    LeftTable ="tblCategories"
    RightTable ="tblBookCategories"
    Expression ="[tblCategories].[CategoryID]=[tblBookCategories].[CategoryID]"
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
        Left =542
        Top =2
        Right =638
        Bottom =94
        Top =0
        Name ="tblCategories"
        Name =""
    End
    Begin
        Left =282
        Top =5
        Right =378
        Bottom =112
        Top =0
        Name ="tblBooks"
        Name =""
    End
    Begin
        Left =21
        Top =5
        Right =117
        Bottom =112
        Top =0
        Name ="tblAuthors"
        Name =""
    End
    Begin
        Left =143
        Top =6
        Right =239
        Bottom =98
        Top =0
        Name ="tblBookAuthor"
        Name =""
    End
    Begin
        Left =414
        Top =9
        Right =510
        Bottom =86
        Top =0
        Name ="tblBookCategories"
        Name =""
    End
End
