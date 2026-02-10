Operation =1
Option =0
Begin InputTables
    Name ="tblAuthors"
    Name ="tblBookAuthor"
End
Begin OutputColumns
    Expression ="tblBookAuthor.ISBNNumber"
    Expression ="tblBookAuthor.AuthorID"
    Expression ="tblBookAuthor.SeqNo"
    Expression ="tblAuthors.FirstName"
    Expression ="tblAuthors.LastName"
    Expression ="tblAuthors.MiddleInit"
End
Begin Joins
    LeftTable ="tblAuthors"
    RightTable ="tblBookAuthor"
    Expression ="[tblAuthors].[AuthorID]=[tblBookAuthor].[AuthorID]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="tblBookAuthor.AuthorID"
        dbInteger "ColumnWidth" ="2895"
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
        Left =172
        Top =6
        Right =268
        Bottom =113
        Top =0
        Name ="tblAuthors"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =98
        Top =0
        Name ="tblBookAuthor"
        Name =""
    End
End
