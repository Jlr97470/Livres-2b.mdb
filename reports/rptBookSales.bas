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
    Width =10798
    DatasheetFontHeight =10
    ItemSuffix =55
    Left =180
    Top =90
    DatasheetGridlinesColor =12632256
    Toolbar ="Book Print Tools"
    RecSrcDt = Begin
        0x63cae9e71543e140
    End
    RecordSource ="qryCustOrderDetails"
    Caption ="Book Sales"
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
            ControlSource ="Title"
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
            ControlSource ="CustName"
        End
        Begin BreakLevel
            ControlSource ="OrderDate"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =2025
            OnFormat ="[Event Procedure]"
            Name ="PageHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =90
                    Top =45
                    Width =2205
                    Height =510
                    FontSize =20
                    Name ="Label16"
                    Caption ="Book Sales "
                End
                Begin Line
                    BorderWidth =3
                    Top =45
                    Width =10786
                    BorderColor =12632256
                    Name ="Line19"
                End
                Begin Line
                    BorderWidth =3
                    Top =75
                    Width =10786
                    BorderColor =12632256
                    Name ="Line20"
                End
                Begin Line
                    BorderWidth =3
                    Top =615
                    Width =10786
                    BorderColor =12632256
                    Name ="Line21"
                End
                Begin Line
                    BorderWidth =3
                    Top =645
                    Width =10786
                    BorderColor =12632256
                    Name ="Line22"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontFamily =18
                    Left =4050
                    Top =795
                    Width =4905
                    Height =360
                    ColumnWidth =4710
                    FontSize =12
                    FontWeight =700
                    BackColor =12632256
                    Name ="Title"
                    ControlSource ="Title"
                    StatusBarText ="Book title"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =3360
                            Top =795
                            Width =630
                            Height =360
                            FontSize =12
                            BackColor =12632256
                            ForeColor =0
                            Name ="Title Label"
                            Caption ="Title:"
                            EventProcPrefix ="Title_Label"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =1740
                    Top =795
                    Width =1545
                    Height =360
                    ColumnWidth =1650
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="ISBNNumber"
                    ControlSource ="ISBNNumber"
                    StatusBarText ="International Standard Book Number"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =75
                            Top =795
                            Width =1605
                            Height =360
                            FontSize =12
                            ForeColor =0
                            Name ="Label29"
                            Caption ="ISBN Number:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontFamily =18
                    Left =9705
                    Top =795
                    Width =1020
                    Height =360
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="SuggPrice"
                    ControlSource ="SuggPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Suggested retail price"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =9015
                            Top =795
                            Width =600
                            Height =360
                            FontSize =12
                            ForeColor =0
                            Name ="Label30"
                            Caption ="SRP:"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    Left =2280
                    Top =1260
                    Width =2760
                    Height =330
                    FontSize =11
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text41"
                    ControlSource ="=Format$([OrderDate],\"mmmm yyyy\")"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            Left =300
                            Top =1260
                            Width =1860
                            Height =345
                            Name ="Label42"
                            Caption ="Order Month:"
                        End
                    End
                End
                Begin Label
                    TextAlign =3
                    Left =2475
                    Top =1680
                    Width =1095
                    Height =300
                    Name ="Label43"
                    Caption ="Order Date"
                End
                Begin Label
                    Left =3630
                    Top =1680
                    Width =2310
                    Height =300
                    Name ="Label44"
                    Caption ="CustName"
                End
                Begin Label
                    TextAlign =3
                    Left =1380
                    Top =1680
                    Width =1020
                    Height =300
                    Name ="Label45"
                    Caption ="Order ID"
                End
                Begin Label
                    TextAlign =3
                    Left =6840
                    Top =1680
                    Width =930
                    Height =300
                    Name ="Label46"
                    Caption ="Quantity"
                End
                Begin Label
                    TextAlign =3
                    Left =7815
                    Top =1680
                    Width =1260
                    Height =300
                    Name ="Label47"
                    Caption ="Pot.Discount"
                End
                Begin Label
                    TextAlign =3
                    Left =9120
                    Top =1680
                    Width =1590
                    Height =300
                    Name ="Label48"
                    Caption ="Extended Price"
                End
                Begin Line
                    BorderWidth =1
                    Left =1380
                    Top =1680
                    Width =9360
                    Name ="Line49"
                End
                Begin Line
                    BorderWidth =1
                    Left =1380
                    Top =1650
                    Width =9360
                    Name ="Line50"
                End
                Begin Line
                    BorderWidth =1
                    Left =1380
                    Top =1980
                    Width =9360
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =1
                    Left =1380
                    Top =2010
                    Width =9360
                    Name ="Line52"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =0
            Name ="GroupHeader0"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =840
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    Left =2280
                    Top =60
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
                            Left =300
                            Top =60
                            Width =1860
                            Height =345
                            Name ="OrderDate by Month Label"
                            Caption ="Order Month:"
                            EventProcPrefix ="OrderDate_by_Month_Label"
                        End
                    End
                End
                Begin Label
                    TextAlign =3
                    Left =2475
                    Top =480
                    Width =1095
                    Height =300
                    Name ="OrderDate Label"
                    Caption ="Order Date"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="OrderDate_Label"
                End
                Begin Label
                    Left =3630
                    Top =480
                    Width =2310
                    Height =300
                    Name ="CustName Label"
                    Caption ="CustName"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="CustName_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =1380
                    Top =480
                    Width =1020
                    Height =300
                    Name ="OrderID Label"
                    Caption ="Order ID"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="OrderID_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =6840
                    Top =480
                    Width =930
                    Height =300
                    Name ="Quantity Label"
                    Caption ="Quantity"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Quantity_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =7815
                    Top =480
                    Width =1260
                    Height =300
                    Name ="Discount Label"
                    Caption ="Pot.Discount"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Discount_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =9120
                    Top =480
                    Width =1590
                    Height =300
                    Name ="ExtPrice Label"
                    Caption ="Extended Price"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="ExtPrice_Label"
                End
                Begin Line
                    BorderWidth =1
                    Left =1380
                    Top =480
                    Width =9360
                    Name ="Line23"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1380
                    Top =450
                    Width =9360
                    Name ="Line24"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1380
                    Top =780
                    Width =9360
                    Name ="Line25"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1380
                    Top =810
                    Width =9360
                    Name ="Line26"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =390
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =2475
                    Top =60
                    Width =1095
                    Height =270
                    ColumnWidth =1185
                    Name ="OrderDate"
                    ControlSource ="OrderDate"
                    Format ="Short Date"
                    StatusBarText ="Date order was entered"
                    InputMask ="99/99/00"
                End
                Begin TextBox
                    Left =3630
                    Top =60
                    Width =3240
                    Height =270
                    TabIndex =1
                    Name ="CustName"
                    ControlSource ="CustName"
                End
                Begin TextBox
                    Left =1380
                    Top =60
                    Width =1050
                    Height =270
                    ColumnWidth =945
                    TabIndex =2
                    Name ="OrderID"
                    ControlSource ="OrderID"
                    StatusBarText ="Unique Order ID"
                End
                Begin TextBox
                    Left =6915
                    Top =60
                    Width =855
                    Height =270
                    TabIndex =3
                    Name ="Quantity"
                    ControlSource ="Quantity"
                    StatusBarText ="Quantity in this order"
                End
                Begin TextBox
                    Left =7950
                    Top =60
                    Width =1125
                    Height =270
                    TabIndex =4
                    Name ="Discount"
                    ControlSource ="Discount"
                    Format ="Percent"
                    StatusBarText ="Possible discount % when ordering multiple copies (not all stores offer discount"
                        "s)"
                End
                Begin TextBox
                    Left =9165
                    Top =60
                    Width =1620
                    Height =270
                    TabIndex =5
                    Name ="ExtPrice"
                    ControlSource ="ExtPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                End
                Begin Line
                    Visible = NotDefault
                    Left =1380
                    Top =375
                    Width =9418
                    Name ="linUp"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =826
            BreakLevel =1
            Name ="GroupFooter3"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =105
                    Width =10650
                    Name ="Text7"
                    ControlSource ="=\"Summary for \" & Format$([OrderDate],\"mmmm yyyy\") & \" (\" & Count(*) & \" "
                        "\" & IIf(Count(*)=1,\"detail record\",\"detail records\") & \")\""
                End
                Begin Label
                    FontItalic = NotDefault
                    Left =105
                    Top =256
                    Width =525
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label8"
                    Caption ="Total"
                End
                Begin Label
                    FontItalic = NotDefault
                    Left =105
                    Top =541
                    Width =1920
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label9"
                    Caption ="Percent of Total Sales"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =7230
                    Top =255
                    Width =540
                    Height =285
                    TabIndex =1
                    Name ="Sum Of Quantity"
                    ControlSource ="=Sum([Quantity])"
                    EventProcPrefix ="Sum_Of_Quantity"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =7185
                    Top =540
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
                    Left =9690
                    Top =255
                    Width =1095
                    Height =285
                    TabIndex =3
                    Name ="Sum Of ExtPrice"
                    ControlSource ="=Sum([ExtPrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    EventProcPrefix ="Sum_Of_ExtPrice"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =9915
                    Top =540
                    Width =870
                    Height =285
                    TabIndex =4
                    Name ="Percent Of ExtPrice"
                    ControlSource ="=Sum([ExtPrice])/([ExtPrice Grand Total Sum])"
                    Format ="Percent"
                    EventProcPrefix ="Percent_Of_ExtPrice"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =826
            Name ="GroupFooter4"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =105
                    Width =10620
                    Name ="Text2"
                    ControlSource ="=\"Summary for \" & \"Book:  \" & \" \" & [ISBNNumber] & \"  \" & [Title] & \" ("
                        "\" & Count(*) & \" \" & IIf(Count(*)=1,\"detail record\",\"detail records\") & \""
                        ")\""
                End
                Begin Label
                    FontItalic = NotDefault
                    Left =105
                    Top =256
                    Width =525
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label3"
                    Caption ="Total"
                End
                Begin Label
                    FontItalic = NotDefault
                    Left =105
                    Top =541
                    Width =1920
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label4"
                    Caption ="Percent of Total Sales"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =7230
                    Top =256
                    Width =540
                    Height =285
                    TabIndex =1
                    Name ="Sum Of Quantity1"
                    ControlSource ="=Sum([Quantity])"
                    EventProcPrefix ="Sum_Of_Quantity1"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =7185
                    Top =541
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
                    Left =9690
                    Top =256
                    Width =1095
                    Height =285
                    TabIndex =3
                    Name ="Sum Of ExtPrice1"
                    ControlSource ="=Sum([ExtPrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    EventProcPrefix ="Sum_Of_ExtPrice1"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =9915
                    Top =541
                    Width =870
                    Height =285
                    TabIndex =4
                    Name ="Percent Of ExtPrice1"
                    ControlSource ="=Sum([ExtPrice])/([ExtPrice Grand Total Sum])"
                    Format ="Percent"
                    EventProcPrefix ="Percent_Of_ExtPrice1"
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
                    Name ="Text17"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5730
                    Top =240
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text18"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"
                End
                Begin Line
                    BorderWidth =3
                    Width =10786
                    BorderColor =12632256
                    Name ="Line27"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10786
                    BorderColor =12632256
                    Name ="Line28"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =405
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    Left =7230
                    Top =105
                    Width =540
                    Height =300
                    FontSize =10
                    Name ="Quantity Grand Total Sum"
                    ControlSource ="=Sum([Quantity])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Quantity_Grand_Total_Sum"
                End
                Begin Label
                    Left =105
                    Top =105
                    Width =1095
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label53"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    Left =9690
                    Top =105
                    Width =1095
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="ExtPrice Grand Total Sum"
                    ControlSource ="=Sum([ExtPrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Times New Roman"
                    EventProcPrefix ="ExtPrice_Grand_Total_Sum"
                End
                Begin Line
                    Width =10786
                    Name ="Line54"
                End
            End
        End
    End
End
CodeBehindForm
' See "rptBookSales.cls"
