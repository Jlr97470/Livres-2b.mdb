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
    Width =5820
    DatasheetFontHeight =10
    ItemSuffix =3
    Left =180
    Top =90
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xfa4e7ebe7e3ce140
    End
    RecordSource ="qryAuthorBooks"
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
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin BreakLevel
            ControlSource ="Title"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =390
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =1500
                    Top =90
                    Width =3600
                    ColumnWidth =4710
                    TabIndex =1
                    Name ="Title"
                    ControlSource ="Title"
                    StatusBarText ="Book title"
                End
                Begin TextBox
                    Left =5160
                    Top =90
                    Width =585
                    TabIndex =2
                    Name ="EditionNumber"
                    ControlSource ="EditionNumber"
                    StatusBarText ="Edition number"
                End
                Begin TextBox
                    Top =90
                    Name ="ISBNNumber"
                    ControlSource ="ISBNNumber"
                    StatusBarText ="International Standard Book Number"
                End
            End
        End
    End
End
