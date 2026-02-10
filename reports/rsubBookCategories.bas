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
    Width =5130
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =180
    Top =90
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb004a1be7e3ce140
    End
    RecordSource ="qryBookCategories"
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
            ControlSource ="ClassName"
        End
        Begin BreakLevel
            ControlSource ="Category"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    Width =2160
                    ColumnWidth =2100
                    Name ="txtClassName"
                    ControlSource ="=[ClassName] & \":\""
                    StatusBarText ="Operating System; Application; Audience; Book Type"
                End
                Begin TextBox
                    Left =2250
                    Width =2880
                    ColumnWidth =2055
                    TabIndex =1
                    Name ="Category"
                    ControlSource ="Category"
                    StatusBarText ="Category description"
                End
            End
        End
    End
End
