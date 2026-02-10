Version =19
VersionRequired =19
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3600
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =180
    Top =90
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe7c88fbe7e3ce140
    End
    RecordSource ="qryBookAuthors"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin BreakLevel
            ControlSource ="ISBNNumber"
        End
        Begin BreakLevel
            ControlSource ="SeqNo"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =299
            Name ="Detail"
            Begin
                Begin TextBox
                    Width =3600
                    Name ="Text0"
                    ControlSource ="=[LastName] & \", \" & [FirstName] & IIf(IsNull([MiddleInit]),Null,\" \" & [Midd"
                        "leInit] & \".\")"
                End
            End
        End
    End
End
