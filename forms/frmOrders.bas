Version =19
VersionRequired =19
Begin Form
    AutoCenter = NotDefault
    DefaultView =0
    ScrollBars =1
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8654
    DatasheetFontHeight =10
    ItemSuffix =87
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    DatasheetGridlinesColor =12632256
    OnFilter ="[Event Procedure]"
    OnApplyFilter ="[Event Procedure]"
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0xbeaadb908aa7e140
    End
    RecordSource ="qryOrdersSorted"
    Caption ="Orders"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    MenuBar ="Book Menu Bar"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextAlign =3
            TextFontFamily =18
            FontSize =9
            FontWeight =700
            BackColor =-2147483633
            ForeColor =-2147483630
            FontName ="Times New Roman"
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
            Height =432
            FontName ="Arial"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            LabelX =-143
            LabelY =-30
        End
        Begin OptionGroup
            AutoLabel = NotDefault
            SpecialEffect =3
            Width =2475
            Height =1095
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
        End
        Begin TextBox
            SpecialEffect =2
            LabelAlign =3
            LabelX =-144
            BackColor =-2147483643
            ForeColor =-2147483640
            FontName ="Arial"
        End
        Begin ListBox
            SpecialEffect =2
            LabelAlign =3
            Width =2880
            LabelX =-144
            BackColor =-2147483643
            ForeColor =-2147483640
            FontName ="Arial"
        End
        Begin ComboBox
            SpecialEffect =2
            LabelAlign =3
            LabelX =-144
            BackColor =-2147483643
            ForeColor =-2147483640
            FontName ="Arial"
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            Height =360
            FontSize =8
            FontName ="Arial"
        End
        Begin FormHeader
            Height =446
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    AccessKey =65
                    Left =1454
                    Top =14
                    Name ="cmdShowAll"
                    Caption ="Show &All"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Show all records"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    Left =14
                    Top =14
                    TabIndex =1
                    Name ="cmdClose"
                    Caption ="&Close"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Save changes and close the form"
                End
                Begin CommandButton
                    OverlapFlags =215
                    AccessKey =65
                    Left =1454
                    Top =14
                    TabIndex =2
                    Name ="cmdAddNew"
                    Caption ="&Add New"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add a new order"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =69
                    Left =4334
                    Top =14
                    TabIndex =3
                    Name ="cmdSearch"
                    Caption ="S&earch..."
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Search for orders..."
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =2894
                    Top =14
                    TabIndex =4
                    Name ="cmdSave"
                    Caption ="&Save"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Save current record"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =78
                    Left =5774
                    Top =14
                    TabIndex =5
                    Name ="cmdCancel"
                    Caption ="Ca&ncel"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Cancel edits and close the form"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =80
                    Left =7214
                    Top =14
                    TabIndex =6
                    Name ="cmdPrint"
                    Caption ="&Print"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Print a Purchase Order"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =4485
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =1290
                    Top =75
                    Width =810
                    ColumnWidth =945
                    TabIndex =18
                    Name ="OrderID"
                    ControlSource ="OrderID"
                    StatusBarText ="Unique Order ID"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =390
                            Top =60
                            Width =855
                            Height =285
                            Name ="Label60"
                            Caption ="Order ID:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =4
                    ListWidth =5400
                    Left =1290
                    Top =420
                    Width =3024
                    ColumnWidth =2355
                    Name ="CustomerID"
                    ControlSource ="CustomerID"
                    RowSourceType ="Table/Query"
                    RowSource ="qryCustForCombo"
                    ColumnWidths ="0;2880;1440;720"
                    StatusBarText ="Customer for this order"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =255
                            Top =420
                            Width =990
                            Height =285
                            Name ="Label61"
                            Caption ="Customer:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1290
                    Top =726
                    Width =3024
                    TabIndex =1
                    Name ="Address"
                    ControlSource ="Address"
                    StatusBarText ="Billing address"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =420
                            Top =725
                            Width =825
                            Height =285
                            Name ="Label62"
                            Caption ="Address:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1290
                    Top =1032
                    Width =3024
                    TabIndex =2
                    Name ="CustCSZ"
                    ControlSource ="=[City] & \", \" & [StateOrProvince] & \"  \" & [PostalCode] & \"  \" & [Country"
                        "]"
                    StatusBarText ="Billing city"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =3
                    ListWidth =4320
                    Left =5325
                    Top =1730
                    Width =3240
                    TabIndex =14
                    Name ="StoreID"
                    ControlSource ="StoreID"
                    RowSourceType ="Table/Query"
                    RowSource ="qryStoresForComboWNull"
                    ColumnWidths ="0;2520;1440"
                    StatusBarText ="Optional Store for Purchase Order"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4155
                            Top =1730
                            Width =1095
                            Height =285
                            Name ="Label64"
                            Caption ="Store (opt.):"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =5325
                    Top =2035
                    Width =3240
                    TabIndex =15
                    Name ="StAddress"
                    ControlSource ="StAddress"
                    StatusBarText ="Store address"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4425
                            Top =2033
                            Width =825
                            Height =285
                            Name ="Label65"
                            Caption ="Address:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =5325
                    Top =2340
                    Width =3240
                    TabIndex =16
                    Name ="StCSZ"
                    ControlSource ="=IIf(IsNull([StoreID]),\"\",[StCity] & \", \" & [StStateOrProvince] & \"  \" & ["
                        "StPostalCode])"
                    StatusBarText ="Store city"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =5325
                    Top =420
                    Width =3240
                    TabIndex =7
                    Name ="ShipName"
                    ControlSource ="ShipName"
                    StatusBarText ="Ship-to name"
                    OnEnter ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4455
                            Top =420
                            Width =795
                            Height =285
                            Name ="Label68"
                            Caption ="Ship To:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =5325
                    Top =725
                    Width =3240
                    ColumnWidth =2175
                    TabIndex =8
                    Name ="ShipAddress"
                    ControlSource ="ShipAddress"
                    StatusBarText ="Ship-to address"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4485
                            Top =718
                            Width =750
                            Height =285
                            Name ="Label69"
                            Caption ="Address:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =5325
                    Top =1030
                    Width =1584
                    ColumnWidth =1005
                    TabIndex =9
                    Name ="ShipCity"
                    ControlSource ="ShipCity"
                    StatusBarText ="Ship-to city"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =6945
                    Top =1030
                    Width =720
                    TabIndex =10
                    Name ="ShipStateOrProvince"
                    ControlSource ="ShipStateOrProvince"
                    StatusBarText ="Ship-to state/province"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =7701
                    Top =1030
                    Width =864
                    TabIndex =11
                    Name ="ShipPostalCode"
                    ControlSource ="ShipPostalCode"
                    StatusBarText ="Ship-to postal code"
                    InputMask ="00000\\-9999"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =5325
                    Top =1335
                    TabIndex =12
                    Name ="ShipCountry"
                    ControlSource ="ShipCountry"
                    StatusBarText ="Ship-to country"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4455
                            Top =1335
                            Width =795
                            Height =285
                            Name ="Label73"
                            Caption ="Country:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1290
                    Top =1338
                    TabIndex =3
                    Name ="ShipPhoneNumber"
                    ControlSource ="ShipPhoneNumber"
                    StatusBarText ="Ship-to phone"
                    InputMask ="!\\(999\") \"000\\-0000"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =615
                            Top =1338
                            Width =630
                            Height =285
                            Name ="Label74"
                            Caption ="Phone:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    ColumnCount =2
                    Left =1290
                    Top =1734
                    TabIndex =4
                    Name ="PayBy"
                    ControlSource ="PayBy"
                    RowSourceType ="Value List"
                    RowSource ="1;\"Cash\";2;\"Check\";3;\"American Express\";4;\"Visa\";5;\"Master Card\";6;\"D"
                        "iscover/Novus\""
                    ColumnWidths ="0;1080"
                    StatusBarText ="1 = Cash, 2 = Check, 3 = American Express, 4 = Visa, 5 = MasterCard, 6 = Discove"
                        "r"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =105
                            Top =1763
                            Width =1140
                            Height =285
                            Name ="Label75"
                            Caption ="Pay Method:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1290
                    Top =2040
                    Width =1875
                    TabIndex =5
                    Name ="CCNumber"
                    ControlSource ="CCNumber"
                    StatusBarText ="Credit card number"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Top =2040
                            Width =1245
                            Height =285
                            Name ="Label76"
                            Caption ="Credit Card #:"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =3960
                    Top =120
                    TabIndex =19
                    Name ="Printed"
                    ControlSource ="Printed"
                    StatusBarText ="Order has been printed"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3060
                            Top =60
                            Width =810
                            Height =285
                            Name ="Label77"
                            Caption ="Printed?:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =75
                    Top =2640
                    Width =8460
                    Height =1785
                    TabIndex =17
                    Name ="fsubOrderBooks"
                    SourceObject ="Form.fsubOrderBooks"
                    LinkChildFields ="OrderID"
                    LinkMasterFields ="OrderID"
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =225
                    Top =30
                    Width =4095
                    Height =330
                    Name ="Box80"
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3405
                    Top =1710
                    Width =870
                    Height =315
                    FontSize =8
                    TabIndex =13
                    Name ="cmdStoreSearch"
                    Caption ="Search >"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1290
                    Top =2346
                    ColumnWidth =1185
                    TabIndex =6
                    Name ="OrderDate"
                    ControlSource ="OrderDate"
                    Format ="Short Date"
                    StatusBarText ="Date order was entered"
                    InputMask ="99/99/00"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =210
                            Top =2346
                            Width =1035
                            Height =285
                            Name ="Label85"
                            Caption ="Order Date:"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    OverlapFlags =85
                    SpecialEffect =0
                    Left =60
                    Top =1650
                    Width =8565
                    Name ="Line86"
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "frmOrders.cls"
