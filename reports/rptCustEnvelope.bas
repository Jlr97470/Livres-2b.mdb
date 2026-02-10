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
    Width =11520
    DatasheetFontHeight =10
    ItemSuffix =8
    Left =180
    Top =90
    DatasheetGridlinesColor =12632256
    Toolbar ="Book Print Tools"
    RecSrcDt = Begin
        0x03f193f81543e140
    End
    RecordSource ="qryCustomerSorted"
    Caption ="Customer #10 Envelope"
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
            Height =4845
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =5400
                    Top =2895
                    Width =6120
                    Name ="Text2"
                    ControlSource ="=Trim([FirstName] & \" \" & IIf(IsNull([MiddleInit]),Null,[MiddleInit] & \". \")"
                        " & [LastName])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =5400
                    Top =3255
                    Width =6120
                    TabIndex =1
                    Name ="Text3"
                    ControlSource ="=Trim([Address])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =5400
                    Top =3615
                    Width =6120
                    TabIndex =2
                    Name ="Text4"
                    ControlSource ="=Trim([City] & \", \" & [StateOrProvince] & \"  \" & [PostalCode])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =5400
                    Top =3975
                    Width =6120
                    TabIndex =3
                    Name ="Text5"
                    ControlSource ="=Trim(IIf([Country]<>\"USA\",[Country],Null))"
                End
                Begin Label
                    TextFontFamily =18
                    Width =2385
                    Height =810
                    FontSize =11
                    FontWeight =700
                    Name ="Label6"
                    Caption ="Microsoft Press\015\012One Microsoft Way\015\012Redmond, WA 98052"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Top =810
                    Height =300
                    FontSize =11
                    TabIndex =4
                    Name ="Text7"
                    ControlSource ="=Trim(IIf([Country]<>\"USA\",\"USA\",Null))"
                End
            End
        End
    End
End
