Operation =1
Option =0
Begin InputTables
    Name ="tblBooks"
End
Begin OutputColumns
    Expression ="tblBooks.ISBNNumber"
    Expression ="tblBooks.Title"
    Expression ="tblBooks.CoverPix"
    Expression ="tblBooks.CopyrightYear"
    Expression ="tblBooks.SuggPrice"
    Expression ="tblBooks.EditionNumber"
    Expression ="tblBooks.OutOfPrint"
    Expression ="tblBooks.Pages"
    Expression ="tblBooks.Disk"
    Expression ="tblBooks.LinerNotes"
    Expression ="tblBooks.TOCDoc"
End
Begin OrderBy
    Expression ="[tblBooks].[Title]"
    Flag =0
    Expression ="[tblBooks].[EditionNumber]"
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
        Top =3
        Name ="tblBooks"
        Name =""
    End
End
