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
    Width =5346
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =180
    Top =90
    DatasheetGridlinesColor =12632256
    Toolbar ="Book Print Tools"
    RecSrcDt = Begin
        0x8a14a23c1643e140
    End
    RecordSource ="qryStoresSorted"
    Caption ="Store Labels"
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
            TextFontFamily =18
            Height =360
            FontSize =12
            FontWeight =700
            FontName ="Times New Roman"
        End
        Begin BreakLevel
            ControlSource ="Country"
        End
        Begin BreakLevel
            ControlSource ="PostalCode"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =2160
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =144
                    Top =540
                    Width =5058
                    Name ="Text2"
                    ControlSource ="=Trim([StoreName])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =144
                    Top =900
                    Width =5058
                    TabIndex =1
                    Name ="Text3"
                    ControlSource ="=Trim([Address])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =144
                    Top =1260
                    Width =5058
                    TabIndex =2
                    Name ="Text4"
                    ControlSource ="=Trim([City] & \", \" & [StateOrProvince] & \"  \" & [PostalCode])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =144
                    Top =1620
                    Width =5058
                    TabIndex =3
                    Name ="Text5"
                    ControlSource ="=Trim([Country])"
                End
            End
        End
    End
End
