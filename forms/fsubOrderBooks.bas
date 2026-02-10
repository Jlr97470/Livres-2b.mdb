Version =19
VersionRequired =19
Begin Form
    NavigationButtons = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7935
    DatasheetFontHeight =10
    ItemSuffix =21
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1a77bf998aa7e140
    End
    RecordSource ="qryOrderBooks"
    Caption ="fsubOrderBooks"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextAlign =1
            FontWeight =700
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
        End
        Begin Line
            SpecialEffect =3
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
        End
        Begin TextBox
            AutoLabel = NotDefault
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ListBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            AutoLabel = NotDefault
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin FormHeader
            Height =495
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =120
                    Top =210
                    Width =1230
                    Height =285
                    FontSize =9
                    Name ="ISBNNumber Label"
                    Caption ="ISBN Number"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="ISBNNumber_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =0
                    TextFontFamily =18
                    Left =2010
                    Top =210
                    Width =495
                    Height =285
                    FontSize =9
                    Name ="Label6"
                    Caption ="Title"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =0
                    TextFontFamily =18
                    Left =4740
                    Top =210
                    Width =360
                    Height =285
                    FontSize =9
                    Name ="Label9"
                    Caption ="Ed."
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =0
                    TextFontFamily =18
                    Left =5340
                    Top =210
                    Width =450
                    Height =285
                    FontSize =9
                    Name ="Label12"
                    Caption ="SRP"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =0
                    TextFontFamily =18
                    Left =5931
                    Top =210
                    Width =450
                    Height =285
                    FontSize =9
                    Name ="Label15"
                    Caption ="Qty."
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =0
                    TextFontFamily =18
                    Left =7380
                    Top =210
                    Width =480
                    Height =285
                    FontSize =9
                    Name ="Label16"
                    Caption ="Price"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =0
                    TextFontFamily =18
                    Left =6420
                    Width =795
                    Height =495
                    FontSize =9
                    Name ="Label17"
                    Caption ="Potential\015\012Discount"
                    FontName ="Times New Roman"
                End
            End
        End
        Begin Section
            Height =330
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    ColumnCount =3
                    ListWidth =5760
                    Left =60
                    Top =43
                    Width =1380
                    Height =255
                    ColumnWidth =2310
                    Name ="ISBNNumber"
                    ControlSource ="ISBNNumber"
                    RowSourceType ="Table/Query"
                    RowSource ="qryBooksForCombo"
                    ColumnWidths ="1296;3744;360"
                    StatusBarText ="International Standard Book Number"
                    FontName ="Arial"
                    InputMask ="0\\-00000\\-000\\-A;0;_"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =5890
                    Top =43
                    Width =495
                    Height =255
                    ColumnWidth =600
                    TabIndex =4
                    Name ="Quantity"
                    ControlSource ="Quantity"
                    StatusBarText ="Quantity in this order"
                    FontName ="Arial"
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1495
                    Top =43
                    Width =3240
                    ColumnWidth =4710
                    TabIndex =1
                    Name ="Title"
                    ControlSource ="Title"
                    FontName ="Arial"
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =4790
                    Top =43
                    Width =330
                    TabIndex =2
                    Name ="EditionNumber"
                    ControlSource ="EditionNumber"
                    FontName ="Arial"
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =5175
                    Top =43
                    Width =660
                    TabIndex =3
                    Name ="SuggPrice"
                    ControlSource ="SuggPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Suggested Retail Price"
                    FontName ="Arial"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =6440
                    Top =43
                    Width =645
                    Height =255
                    TabIndex =5
                    Name ="Discount"
                    ControlSource ="Discount"
                    Format ="Percent"
                    StatusBarText ="Possible discount % when ordering multiple copies (not all stores offer discount"
                        "s)"
                    FontName ="Arial"
                    ConditionalFormat = Begin
                        0x010000006c000000010000000000000004000000000000000500000001000000 ,
                        0xffffff0080008000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x30002e003000350000000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =7140
                    Top =45
                    Width =765
                    TabIndex =6
                    Name ="FullPrice"
                    ControlSource ="=Round(CCur(NZ(([Quantity]*[SuggPrice])*(1-[Discount]),0)),2)"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Extended price calculation"
                    FontName ="Arial"
                End
            End
        End
        Begin FormFooter
            Height =285
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =5785
                    Width =600
                    Height =255
                    Name ="SumQty"
                    ControlSource ="=Sum([Quantity])"
                    FontName ="Arial"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =6990
                    Width =915
                    TabIndex =1
                    Name ="SumPrice"
                    ControlSource ="=Sum([ExtPrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =4830
                    Width =630
                    Height =285
                    FontSize =9
                    Name ="Label20"
                    Caption ="Totals:"
                    FontName ="Times New Roman"
                End
            End
        End
    End
End
