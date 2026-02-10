Version =19
VersionRequired =19
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =20
    GridY =24
    Width =4860
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =180
    Top =90
    DatasheetGridlinesColor =12632256
    Toolbar ="Book Print Tools"
    RecSrcDt = Begin
        0x4169322f1643e140
    End
    RecordSource ="qryCustomerSorted"
    Caption ="Customer Labels"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin TextBox
            AutoLabel = NotDefault
            OldBorderStyle =0
            Height =300
            FontSize =10
            FontWeight =600
            FontName ="Arial"
        End
        Begin BreakLevel
            ControlSource ="Country"
        End
        Begin BreakLevel
            ControlSource ="PostalCode"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =3360
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =576
                    Top =480
                    Width =3708
                    Name ="Text2"
                    ControlSource ="=Trim([FirstName] & \" \" & [MiddleInit] & \". \" & [LastName])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =576
                    Top =780
                    Width =3708
                    TabIndex =1
                    Name ="Text3"
                    ControlSource ="=Trim([Address])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =576
                    Top =1080
                    Width =3708
                    TabIndex =2
                    Name ="Text4"
                    ControlSource ="=Trim([City] & \", \" & [StateOrProvince] & \"  \" & [PostalCode])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =576
                    Top =1380
                    Width =3708
                    TabIndex =3
                    Name ="Text5"
                    ControlSource ="=Trim([Country])"
                End
            End
        End
    End
End
