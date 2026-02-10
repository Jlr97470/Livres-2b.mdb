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
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =89
    Left =180
    Top =90
    DatasheetGridlinesColor =12632256
    Toolbar ="Book Print Tools"
    RecSrcDt = Begin
        0xb9bf6651cea7e140
    End
    RecordSource ="qryOrderReport"
    Caption ="Orders"
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
            LabelAlign =3
            BackStyle =0
            LabelX =-144
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
            ControlSource ="OrderID"
        End
        Begin BreakLevel
            ControlSource ="Title"
        End
        Begin PageHeader
            Height =3345
            OnFormat ="[Event Procedure]"
            Name ="PageHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =2115
                    Height =510
                    FontSize =20
                    Name ="Label28"
                    Caption ="Book Order "
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =10800
                    BorderColor =12632256
                    Name ="Line31"
                End
                Begin Line
                    BorderWidth =3
                    Top =90
                    Width =10800
                    BorderColor =12632256
                    Name ="Line32"
                End
                Begin Line
                    BorderWidth =3
                    Top =645
                    Width =10800
                    BorderColor =12632256
                    Name ="Line33"
                End
                Begin Line
                    BorderWidth =3
                    Top =675
                    Width =10800
                    BorderColor =12632256
                    Name ="Line34"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =6570
                    Top =225
                    Width =1425
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text29"
                    ControlSource ="OrderDate"
                    Format ="Short Date"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =5130
                            Top =225
                            Width =1230
                            Height =300
                            Name ="Label58"
                            Caption ="Order Date:"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =9780
                    Top =225
                    Width =915
                    Height =300
                    ColumnWidth =945
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="OrderID"
                    ControlSource ="OrderID"
                    StatusBarText ="Unique Order ID"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =8595
                            Top =225
                            Width =1065
                            Height =300
                            Name ="OrderID Label"
                            Caption ="Order ID:"
                            EventProcPrefix ="OrderID_Label"
                        End
                    End
                End
                Begin TextBox
                    Left =1695
                    Top =1890
                    Width =3960
                    Height =225
                    TabIndex =2
                    ForeColor =8388608
                    Name ="txtName"
                    ControlSource ="=[FirstName] & \" \" & IIf(IsNull([MiddleInit]),Null,[MiddleInit] & \". \") & [L"
                        "astName]"
                    StatusBarText ="Customer Last Name"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =345
                            Top =1890
                            Width =1200
                            Height =300
                            Name ="LastName Label"
                            Caption ="Customer :"
                            EventProcPrefix ="LastName_Label"
                        End
                    End
                End
                Begin TextBox
                    Left =1695
                    Top =2149
                    Width =3960
                    Height =225
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Address"
                    ControlSource ="Address"
                    StatusBarText ="Billing address"
                End
                Begin TextBox
                    Left =1695
                    Top =2408
                    Width =3960
                    Height =225
                    TabIndex =4
                    ForeColor =8388608
                    Name ="CitySZ"
                    ControlSource ="=[City] & \", \" & [StateOrProvince] & \"  \" & [PostalCode]"
                    StatusBarText ="Billing city"
                End
                Begin Label
                    Left =1677
                    Top =3015
                    Width =3555
                    Height =300
                    Name ="Title Label"
                    Caption ="Title"
                    EventProcPrefix ="Title_Label"
                End
                Begin Label
                    Top =3015
                    Width =1440
                    Height =300
                    Name ="ISBNNumber Label"
                    Caption ="ISBN Number"
                    EventProcPrefix ="ISBNNumber_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =5514
                    Top =3015
                    Width =870
                    Height =300
                    Name ="Quantity Label"
                    Caption ="Quantity"
                    EventProcPrefix ="Quantity_Label"
                End
                Begin Label
                    TextAlign =2
                    Left =7848
                    Top =3015
                    Width =990
                    Height =300
                    Name ="Discount Label"
                    Caption ="Discount"
                    EventProcPrefix ="Discount_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =6621
                    Top =3015
                    Width =1080
                    Height =300
                    Name ="SuggPrice Label"
                    Caption ="SRP"
                    EventProcPrefix ="SuggPrice_Label"
                End
                Begin Line
                    BorderWidth =1
                    Top =3015
                    Width =10800
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =1
                    Top =2985
                    Width =10800
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =1
                    Top =3315
                    Width =10800
                    Name ="Line37"
                End
                Begin Line
                    BorderWidth =1
                    Top =3345
                    Width =10800
                    Name ="Line38"
                End
                Begin TextBox
                    Left =1695
                    Top =734
                    Width =3960
                    ColumnWidth =2850
                    TabIndex =5
                    Name ="StoreName"
                    ControlSource ="StoreName"
                    StatusBarText ="Name of the store"
                End
                Begin TextBox
                    Left =1695
                    Top =1009
                    Width =3960
                    TabIndex =6
                    Name ="StAddress"
                    ControlSource ="StAddress"
                    StatusBarText ="Store address"
                End
                Begin TextBox
                    Left =1695
                    Top =1283
                    Width =3960
                    TabIndex =7
                    Name ="StCSZ"
                    ControlSource ="=[StCity] & \", \" & [StStateOrProvince] & \"  \" & [StPostalCode]"
                    StatusBarText ="Store address"
                End
                Begin TextBox
                    Left =6759
                    Top =1890
                    Width =3960
                    Height =225
                    TabIndex =8
                    Name ="ShipName"
                    ControlSource ="ShipName"
                    StatusBarText ="Ship-to name"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =5745
                            Top =1890
                            Width =870
                            Height =300
                            Name ="Label65"
                            Caption ="Ship To:"
                        End
                    End
                End
                Begin TextBox
                    Left =6759
                    Top =2149
                    Width =3960
                    Height =225
                    ColumnWidth =2175
                    TabIndex =9
                    Name ="ShipAddress"
                    ControlSource ="ShipAddress"
                    StatusBarText ="Ship-to address"
                End
                Begin TextBox
                    Left =6765
                    Top =2408
                    Width =3960
                    Height =225
                    ColumnWidth =1005
                    TabIndex =10
                    Name ="ShipCSZ"
                    ControlSource ="=[ShipCity] & \", \" & [ShipStateOrProvince] & \"  \" & [ShipPostalCode]"
                    StatusBarText ="Ship-to city"
                End
                Begin TextBox
                    TextAlign =1
                    Left =8694
                    Top =1146
                    Width =2025
                    TabIndex =12
                    Name ="CCNumber"
                    ControlSource ="CCNumber"
                    StatusBarText ="Credit card number"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =6825
                            Top =1146
                            Width =1665
                            Height =300
                            Name ="Label69"
                            Caption ="Credit Card No.:"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    Left =8694
                    Top =735
                    Width =2025
                    TabIndex =11
                    Name ="txtPay"
                    ControlSource ="=Switch([PayBy]=1,\"Cash\",[PayBy]=2,\"Check\",[PayBy]=3,\"American Express\",[P"
                        "ayBy]=4,\"Visa\",[PayBy]=5,\"Master Card\",[PayBy]=6,\"Discover\")"
                    StatusBarText ="1 = Cash, 2 = Check, 3 = American Express, 4 = Visa, 5 = MasterCard, 6 = Discove"
                        "r"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =6795
                            Top =735
                            Width =1695
                            Height =300
                            Name ="Label68"
                            Caption ="Payment Method:"
                        End
                    End
                End
                Begin TextBox
                    Left =1695
                    Top =1556
                    Width =3960
                    TabIndex =13
                    Name ="StCountry"
                    ControlSource ="StCountry"
                    StatusBarText ="Store address"
                End
                Begin TextBox
                    Left =1695
                    Top =2667
                    Width =3960
                    Height =225
                    TabIndex =14
                    ForeColor =8388608
                    Name ="Country"
                    ControlSource ="Country"
                    StatusBarText ="Billing city"
                End
                Begin TextBox
                    Left =6765
                    Top =2667
                    Width =3960
                    Height =225
                    TabIndex =15
                    Name ="ShipCountry"
                    ControlSource ="ShipCountry"
                    StatusBarText ="Ship-to city"
                End
                Begin TextBox
                    Left =8694
                    Top =1557
                    Width =2025
                    TabIndex =16
                    Name ="ShipPhoneNumber"
                    ControlSource ="ShipPhoneNumber"
                    StatusBarText ="Ship-to phone"
                    InputMask ="!\\(999\") \"000\\-0000"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =7695
                            Top =1557
                            Width =795
                            Height =300
                            Name ="Label73"
                            Caption ="Phone:"
                        End
                    End
                End
                Begin Label
                    TextAlign =3
                    Top =735
                    Width =1545
                    Height =300
                    Name ="Label62"
                    Caption ="Ordered From:"
                End
                Begin TextBox
                    Left =1695
                    Top =734
                    Width =3960
                    TabIndex =17
                    Name ="txtMSPName"
                    ControlSource ="=\"Microsoft Press\""
                    StatusBarText ="Name of the store"
                End
                Begin TextBox
                    Left =1695
                    Top =1009
                    Width =3960
                    TabIndex =18
                    Name ="txtMSPAddress"
                    ControlSource ="=\"One Microsoft Way\""
                    StatusBarText ="Store address"
                End
                Begin TextBox
                    Left =1695
                    Top =1283
                    Width =3960
                    TabIndex =19
                    Name ="txtMSPCSZ"
                    ControlSource ="=\"Redmond, WA  98052\""
                    StatusBarText ="Store address"
                End
                Begin TextBox
                    Left =1695
                    Top =1556
                    Width =3960
                    TabIndex =20
                    Name ="txtMSPCountry"
                    ControlSource ="=\"USA\""
                    StatusBarText ="Store address"
                End
                Begin Label
                    TextAlign =3
                    Left =8985
                    Top =3015
                    Width =1530
                    Height =300
                    Name ="Label78"
                    Caption ="Extended Price"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =0
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =390
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =1677
                    Top =60
                    Width =3600
                    Height =270
                    ColumnWidth =4710
                    Name ="Title"
                    ControlSource ="Title"
                    StatusBarText ="Book title"
                End
                Begin TextBox
                    Left =5514
                    Top =60
                    Width =870
                    Height =270
                    TabIndex =2
                    Name ="Quantity"
                    ControlSource ="Quantity"
                    StatusBarText ="Quantity in this order"
                End
                Begin TextBox
                    Left =7938
                    Top =60
                    Width =900
                    Height =270
                    TabIndex =3
                    Name ="Discount"
                    ControlSource ="Discount"
                    Format ="Percent"
                    StatusBarText ="Possible discount % when ordering multiple copies (not all stores offer discount"
                        "s)"
                End
                Begin TextBox
                    DecimalPlaces =2
                    Left =6621
                    Top =60
                    Width =1080
                    Height =270
                    TabIndex =4
                    Name ="SuggPrice"
                    ControlSource ="SuggPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Suggested retail price"
                End
                Begin TextBox
                    Top =60
                    Height =270
                    ColumnWidth =1650
                    TabIndex =1
                    Name ="ISBNNumber"
                    ControlSource ="ISBNNumber"
                    StatusBarText ="International Standard Book Number"
                End
                Begin TextBox
                    Left =9075
                    Top =60
                    TabIndex =5
                    Name ="ExtPrice"
                    ControlSource ="=Round(CCur(([Quantity]*[SuggPrice])*(1-[Discount])),2)"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =435
            Name ="GroupFooter1"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    Left =3675
                    Top =90
                    Width =1245
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label9"
                    Caption ="Order Totals:"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =5364
                    Top =90
                    Width =1020
                    Height =285
                    Name ="Sum Of Quantity"
                    ControlSource ="=Sum([Quantity])"
                    EventProcPrefix ="Sum_Of_Quantity"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    Left =8730
                    Top =90
                    Width =1785
                    Height =285
                    TabIndex =1
                    Name ="Sum Of ExtPrice"
                    ControlSource ="=Sum(Round(CCur(([Quantity]*[SuggPrice])*(1-[Discount])),2))"
                    Format ="$#,##0.00;($#,##0.00)"
                    EventProcPrefix ="Sum_Of_ExtPrice"
                End
                Begin Line
                    Width =10800
                    BorderColor =12632256
                    Name ="Line39"
                End
                Begin Line
                    Left =3456
                    Top =14
                    Width =7200
                    Name ="Line79"
                End
            End
        End
        Begin PageFooter
            Height =435
            Name ="PageFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =8745
                    Top =135
                    Width =1980
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text30"
                    ControlSource ="=\"Page \" & [Page]"
                    FontName ="Times New Roman"
                End
                Begin Line
                    BorderWidth =3
                    Width =10800
                    BorderColor =12632256
                    Name ="Line40"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10800
                    BorderColor =12632256
                    Name ="Line41"
                End
            End
        End
    End
End
CodeBehindForm
' See "rptOrders.cls"
