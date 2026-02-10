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
    Width =12960
    DatasheetFontHeight =10
    ItemSuffix =75
    Left =180
    Top =90
    DatasheetGridlinesColor =12632256
    Toolbar ="Book Print Tools"
    RecSrcDt = Begin
        0x2da7dc6dcea7e140
    End
    RecordSource ="qryCustOrderDetails"
    Caption ="Customer Sales"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderColor =8388608
        End
        Begin Line
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            OldBorderStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin ListBox
            OldBorderStyle =0
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =2
            ControlSource ="CustName"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            GroupOn =4
            ControlSource ="OrderDate"
        End
        Begin BreakLevel
            ControlSource ="OrderDate"
        End
        Begin BreakLevel
            ControlSource ="ISBNNumber"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            OnFormat ="[Event Procedure]"
            Name ="PageHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Width =2805
                    Height =510
                    FontSize =20
                    Name ="Label40"
                    Caption ="Customer Sales"
                End
                Begin Line
                    BorderWidth =3
                    Width =12960
                    BorderColor =12632256
                    Name ="Line43"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =12960
                    BorderColor =12632256
                    Name ="Line44"
                End
                Begin Line
                    BorderWidth =3
                    Top =555
                    Width =12960
                    BorderColor =12632256
                    Name ="Line45"
                End
                Begin Line
                    BorderWidth =3
                    Top =585
                    Width =12960
                    BorderColor =12632256
                    Name ="Line46"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    Left =7365
                    Top =675
                    Width =2760
                    Height =330
                    FontSize =11
                    ForeColor =8388608
                    Name ="OrderDate by Month"
                    ControlSource ="=Format$([OrderDate],\"mmmm yyyy\")"
                    FontName ="Times New Roman"
                    EventProcPrefix ="OrderDate_by_Month"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =5790
                            Top =675
                            Width =1455
                            Height =330
                            Name ="OrderDate by Month Label"
                            Caption ="Order Month:"
                            EventProcPrefix ="OrderDate_by_Month_Label"
                        End
                    End
                End
                Begin Label
                    Left =1080
                    Top =1095
                    Width =1170
                    Height =300
                    Name ="OrderDate Label"
                    Caption ="Order Date"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="OrderDate_Label"
                End
                Begin Label
                    Left =2340
                    Top =1095
                    Width =1440
                    Height =300
                    Name ="ISBNNumber Label"
                    Caption ="ISBN Number"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="ISBNNumber_Label"
                End
                Begin Label
                    Left =3840
                    Top =1095
                    Width =2640
                    Height =300
                    Name ="Title Label"
                    Caption ="Title"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Title_Label"
                End
                Begin Label
                    TextAlign =2
                    Left =8115
                    Top =1095
                    Width =585
                    Height =300
                    Name ="SuggPrice Label"
                    Caption ="SRP"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SuggPrice_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =8850
                    Top =1095
                    Width =930
                    Height =300
                    Name ="Quantity Label"
                    Caption ="Quantity"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Quantity_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =9840
                    Top =1095
                    Width =1380
                    Height =300
                    Name ="Discount Label"
                    Caption ="Discount"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Discount_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =11610
                    Top =1095
                    Width =540
                    Height =300
                    Name ="ExtPrice Label"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="ExtPrice_Label"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =1095
                    Width =11070
                    Name ="Line47"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =1065
                    Width =11070
                    Name ="Line48"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =1395
                    Width =11070
                    Name ="Line49"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =1425
                    Width =11070
                    Name ="Line50"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontFamily =18
                    Left =1365
                    Top =675
                    Width =4065
                    Height =360
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    BackColor =12632256
                    Name ="CustName"
                    ControlSource ="CustName"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Top =675
                            Width =1200
                            Height =330
                            FontSize =12
                            BackColor =12632256
                            ForeColor =0
                            Name ="CustName Label"
                            Caption ="Customer:"
                            EventProcPrefix ="CustName_Label"
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =0
            Name ="GroupHeader4"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =810
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader5"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    Left =7365
                    Top =60
                    Width =2760
                    Height =330
                    FontSize =11
                    ForeColor =8388608
                    Name ="Text57"
                    ControlSource ="=Format$([OrderDate],\"mmmm yyyy\")"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =5790
                            Top =60
                            Width =1455
                            Height =330
                            Name ="Label58"
                            Caption ="Order Month:"
                        End
                    End
                End
                Begin Label
                    Left =1080
                    Top =480
                    Width =1170
                    Height =300
                    Name ="Label59"
                    Caption ="Order Date"
                End
                Begin Label
                    Left =2340
                    Top =480
                    Width =1440
                    Height =300
                    Name ="Label60"
                    Caption ="ISBN Number"
                End
                Begin Label
                    Left =3840
                    Top =480
                    Width =2640
                    Height =300
                    Name ="Label61"
                    Caption ="Title"
                End
                Begin Label
                    TextAlign =2
                    Left =8115
                    Top =480
                    Width =585
                    Height =300
                    Name ="Label62"
                    Caption ="SRP"
                End
                Begin Label
                    TextAlign =3
                    Left =8850
                    Top =480
                    Width =930
                    Height =300
                    Name ="Label63"
                    Caption ="Quantity"
                End
                Begin Label
                    TextAlign =3
                    Left =9840
                    Top =480
                    Width =1380
                    Height =300
                    Name ="Label64"
                    Caption ="Discount"
                End
                Begin Label
                    TextAlign =3
                    Left =11610
                    Top =480
                    Width =540
                    Height =300
                    Name ="Label65"
                    Caption ="Price"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =480
                    Width =11070
                    Name ="Line66"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =450
                    Width =11070
                    Name ="Line67"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =780
                    Width =11070
                    Name ="Line68"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =810
                    Width =11070
                    Name ="Line69"
                End
                Begin Line
                    BorderWidth =3
                    Width =12960
                    BorderColor =12632256
                    Name ="Line72"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =360
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =1335
                    Top =29
                    Width =780
                    Height =270
                    ColumnWidth =1185
                    Name ="OrderDate"
                    ControlSource ="OrderDate"
                    Format ="Short Date"
                    StatusBarText ="Date order was entered"
                    InputMask ="99/99/00"
                End
                Begin TextBox
                    Left =3900
                    Top =29
                    Width =3690
                    Height =270
                    ColumnWidth =4710
                    TabIndex =2
                    Name ="Title"
                    ControlSource ="Title"
                    StatusBarText ="Book title"
                End
                Begin TextBox
                    DecimalPlaces =2
                    Left =7965
                    Top =29
                    Width =900
                    Height =270
                    TabIndex =3
                    Name ="SuggPrice"
                    ControlSource ="SuggPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Suggested retail price"
                End
                Begin TextBox
                    Left =9300
                    Top =29
                    Width =540
                    Height =270
                    TabIndex =4
                    Name ="Quantity"
                    ControlSource ="Quantity"
                    StatusBarText ="Quantity in this order"
                End
                Begin TextBox
                    Left =9900
                    Top =29
                    Width =1380
                    Height =270
                    TabIndex =5
                    Name ="Discount"
                    ControlSource ="Discount"
                    Format ="Percent"
                    StatusBarText ="Possible discount % when ordering multiple copies (not all stores offer discount"
                        "s)"
                End
                Begin TextBox
                    Left =11340
                    Top =29
                    Width =870
                    Height =270
                    TabIndex =6
                    Name ="ExtPrice"
                    ControlSource ="=Round(CCur(([Quantity]*[SuggPrice])*(1-[Discount])),2)"
                    Format ="$#,##0.00;($#,##0.00)"
                End
                Begin TextBox
                    Left =2400
                    Top =29
                    Width =1380
                    Height =270
                    ColumnWidth =1650
                    TabIndex =1
                    Name ="ISBNNumber"
                    ControlSource ="ISBNNumber"
                    StatusBarText ="International Standard Book Number"
                End
                Begin Line
                    Visible = NotDefault
                    Left =1080
                    Top =360
                    Width =11070
                    Name ="linUp"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =915
            BreakLevel =1
            Name ="GroupFooter3"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =60
                    Top =29
                    Width =12030
                    Name ="Text7"
                    ControlSource ="=\"Summary for \" & Format$([OrderDate],\"mmmm yyyy\") & \" (\" & Count(*) & \" "
                        "\" & IIf(Count(*)=1,\"detail record\",\"detail records\") & \")\""
                End
                Begin Label
                    FontItalic = NotDefault
                    Left =60
                    Top =285
                    Width =525
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label8"
                    Caption ="Total"
                End
                Begin Label
                    FontItalic = NotDefault
                    Left =60
                    Top =570
                    Width =1920
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label9"
                    Caption ="Percent of Total Sales"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =9300
                    Top =285
                    Width =540
                    Height =285
                    TabIndex =1
                    Name ="Sum Of Quantity"
                    ControlSource ="=Sum([Quantity])"
                    EventProcPrefix ="Sum_Of_Quantity"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =9255
                    Top =570
                    Width =585
                    Height =285
                    TabIndex =2
                    Name ="Percent Of Quantity"
                    ControlSource ="=Sum([Quantity])/([Quantity Grand Total Sum])"
                    Format ="Percent"
                    EventProcPrefix ="Percent_Of_Quantity"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =11040
                    Top =285
                    Width =1170
                    Height =285
                    TabIndex =3
                    Name ="Sum Of ExtPrice"
                    ControlSource ="=Sum(Round(CCur(([Quantity]*[SuggPrice])*(1-[Discount])),2))"
                    Format ="$#,##0.00;($#,##0.00)"
                    EventProcPrefix ="Sum_Of_ExtPrice"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =11340
                    Top =570
                    Width =870
                    Height =285
                    TabIndex =4
                    Name ="Percent Of ExtPrice"
                    ControlSource ="=Sum(Round(CCur(([Quantity]*[SuggPrice])*(1-[Discount])),2))/([ExtPrice Grand To"
                        "tal Sum])"
                    Format ="Percent"
                    EventProcPrefix ="Percent_Of_ExtPrice"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Width =12840
                    BorderColor =12632256
                    Name ="Line51"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =975
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =60
                    Top =29
                    Width =12840
                    Name ="Text2"
                    ControlSource ="=\"Summary for \" & \"Customer:  \" & \" \" & [CustName] & \" (\" & Count(*) & \""
                        " \" & IIf(Count(*)=1,\"detail record\",\"detail records\") & \")\""
                End
                Begin Label
                    FontItalic = NotDefault
                    Left =60
                    Top =285
                    Width =525
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label3"
                    Caption ="Total"
                End
                Begin Label
                    FontItalic = NotDefault
                    Left =60
                    Top =570
                    Width =1920
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label4"
                    Caption ="Percent of Total Sales"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =9300
                    Top =285
                    Width =540
                    Height =285
                    TabIndex =1
                    Name ="Sum Of Quantity1"
                    ControlSource ="=Sum([Quantity])"
                    EventProcPrefix ="Sum_Of_Quantity1"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =9255
                    Top =570
                    Width =585
                    Height =285
                    TabIndex =2
                    Name ="Percent Of Quantity1"
                    ControlSource ="=Sum([Quantity])/([Quantity Grand Total Sum])"
                    Format ="Percent"
                    EventProcPrefix ="Percent_Of_Quantity1"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =11040
                    Top =285
                    Width =1170
                    Height =285
                    TabIndex =3
                    Name ="Sum Of ExtPrice1"
                    ControlSource ="=Sum(Round(CCur(([Quantity]*[SuggPrice])*(1-[Discount])),2))"
                    Format ="$#,##0.00;($#,##0.00)"
                    EventProcPrefix ="Sum_Of_ExtPrice1"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =11340
                    Top =570
                    Width =870
                    Height =285
                    TabIndex =4
                    Name ="Percent Of ExtPrice1"
                    ControlSource ="=Sum(Round(CCur(([Quantity]*[SuggPrice])*(1-[Discount])),2))/([ExtPrice Grand To"
                        "tal Sum])"
                    Format ="Percent"
                    EventProcPrefix ="Percent_Of_ExtPrice1"
                End
                Begin Line
                    BorderWidth =1
                    Width =12840
                    BorderColor =12632256
                    Name ="Line54"
                End
                Begin Line
                    BorderWidth =1
                    Top =900
                    Width =12840
                    BorderColor =12632256
                    Name ="Line55"
                End
                Begin Line
                    BorderWidth =1
                    Top =975
                    Width =12840
                    BorderColor =12632256
                    Name ="Line56"
                End
            End
        End
        Begin PageFooter
            Height =540
            Name ="PageFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Top =240
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text41"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =7860
                    Top =240
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text42"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"
                End
                Begin Line
                    BorderWidth =3
                    Width =12960
                    BorderColor =12632256
                    Name ="Line52"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =12960
                    BorderColor =12632256
                    Name ="Line53"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =390
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    Left =9300
                    Top =90
                    Width =540
                    Height =300
                    FontSize =10
                    Name ="Quantity Grand Total Sum"
                    ControlSource ="=Sum([Quantity])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Quantity_Grand_Total_Sum"
                End
                Begin Label
                    Left =60
                    Top =90
                    Width =1095
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label29"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    Left =11040
                    Top =90
                    Width =1170
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="ExtPrice Grand Total Sum"
                    ControlSource ="=Sum(Round(CCur(([Quantity]*[SuggPrice])*(1-[Discount])),2))"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Times New Roman"
                    EventProcPrefix ="ExtPrice_Grand_Total_Sum"
                End
                Begin Line
                    Width =12226
                    Name ="Line74"
                End
            End
        End
    End
End
CodeBehindForm
' See "rptCustSales.cls"
