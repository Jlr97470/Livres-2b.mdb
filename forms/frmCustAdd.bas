Version =19
VersionRequired =19
Begin Form
    AllowFilters = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DataEntry = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    DatasheetFontHeight =10
    ItemSuffix =27
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    DatasheetGridlinesColor =12632256
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0xee28e26c8aa7e140
    End
    RecordSource ="tblCustomers"
    Caption ="Add A Customer"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
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
        Begin CommandButton
            FontName ="Arial"
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
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
            FontName ="Arial"
        End
        Begin ListBox
            FontName ="Arial"
        End
        Begin ComboBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
            FontName ="Arial"
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin ToggleButton
            FontSize =8
            FontName ="Arial"
        End
        Begin FormHeader
            Height =446
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    Left =14
                    Top =14
                    Height =432
                    Name ="cmdClose"
                    Caption ="&Close"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Save changes and close the form"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =1454
                    Top =14
                    Height =432
                    TabIndex =1
                    Name ="cmdSave"
                    Caption ="&Save"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Save current record"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =78
                    Left =2894
                    Top =14
                    Height =432
                    TabIndex =2
                    Name ="cmdCancel"
                    Caption ="Ca&ncel"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Cancel edits and close the form"
                End
            End
        End
        Begin Section
            Height =3585
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =5355
                    Top =150
                    Width =840
                    Height =270
                    ColumnWidth =1440
                    TabIndex =1
                    Name ="CustomerID"
                    ControlSource ="CustomerID"
                    StatusBarText ="Unique Customer ID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4110
                            Top =150
                            Width =1200
                            Height =300
                            Name ="CustomerID Label"
                            Caption ="Customer ID:"
                            EventProcPrefix ="CustomerID_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1575
                    Top =960
                    Width =2310
                    Height =270
                    ColumnWidth =2310
                    TabIndex =3
                    Name ="LastName"
                    ControlSource ="LastName"
                    StatusBarText ="Customer Last Name"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =510
                            Top =960
                            Width =1020
                            Height =285
                            Name ="LastName Label"
                            Caption ="Last Name:"
                            EventProcPrefix ="LastName_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1575
                    Top =150
                    Width =2310
                    Height =270
                    ColumnWidth =2310
                    Name ="FirstName"
                    ControlSource ="FirstName"
                    StatusBarText ="Customer First Name"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =465
                            Top =150
                            Width =1065
                            Height =285
                            Name ="FirstName Label"
                            Caption ="First Name:"
                            EventProcPrefix ="FirstName_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1575
                    Top =555
                    Width =240
                    Height =270
                    ColumnWidth =240
                    TabIndex =2
                    Name ="MiddleInit"
                    ControlSource ="MiddleInit"
                    StatusBarText ="Customer Middle Initial (if any)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1035
                            Top =555
                            Width =495
                            Height =285
                            Name ="MiddleInit Label"
                            Caption ="M.I.:"
                            EventProcPrefix ="MiddleInit_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    Left =1575
                    Top =1365
                    Width =3495
                    Height =480
                    ColumnWidth =3000
                    TabIndex =4
                    Name ="Address"
                    ControlSource ="Address"
                    StatusBarText ="Billing address"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =1365
                            Width =1290
                            Height =285
                            Name ="Address Label"
                            Caption ="Street Address:"
                            EventProcPrefix ="Address_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1575
                    Top =1980
                    Width =2310
                    Height =270
                    ColumnWidth =2310
                    TabIndex =5
                    Name ="City"
                    ControlSource ="City"
                    StatusBarText ="Billing city"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1005
                            Top =1980
                            Width =525
                            Height =285
                            Name ="City Label"
                            Caption ="City:"
                            EventProcPrefix ="City_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =5355
                    Top =1980
                    Width =1755
                    Height =270
                    ColumnWidth =1560
                    TabIndex =6
                    Name ="StateOrProvince"
                    ControlSource ="StateOrProvince"
                    StatusBarText ="Billing state or province"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4035
                            Top =1980
                            Width =1275
                            Height =285
                            Name ="StateOrProvince Label"
                            Caption ="State/Province:"
                            EventProcPrefix ="StateOrProvince_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1575
                    Top =2385
                    Width =1755
                    Height =270
                    ColumnWidth =1560
                    TabIndex =7
                    Name ="PostalCode"
                    ControlSource ="PostalCode"
                    StatusBarText ="Billing postal code"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =330
                            Top =2385
                            Width =1200
                            Height =300
                            Name ="PostalCode Label"
                            Caption ="Postal Code:"
                            EventProcPrefix ="PostalCode_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =5355
                    Top =2385
                    Width =1755
                    Height =270
                    ColumnWidth =2310
                    TabIndex =8
                    Name ="Country"
                    ControlSource ="Country"
                    StatusBarText ="Billing country"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4515
                            Top =2385
                            Width =795
                            Height =285
                            Name ="Country Label"
                            Caption ="Country:"
                            EventProcPrefix ="Country_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1575
                    Top =2790
                    Width =1755
                    Height =270
                    ColumnWidth =2310
                    TabIndex =9
                    Name ="PhoneNumber"
                    ControlSource ="PhoneNumber"
                    StatusBarText ="Phone number"
                    InputMask ="!\\(999\") \"000\\-0000"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =210
                            Top =2790
                            Width =1320
                            Height =285
                            Name ="PhoneNumber Label"
                            Caption ="Phone Number:"
                            EventProcPrefix ="PhoneNumber_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =5355
                    Top =2790
                    Width =1755
                    Height =270
                    ColumnWidth =2310
                    TabIndex =10
                    Name ="EmailAddress"
                    ControlSource ="EmailAddress"
                    StatusBarText ="Email address"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4050
                            Top =2790
                            Width =1260
                            Height =285
                            Name ="EmailAddress Label"
                            Caption ="Email Address:"
                            EventProcPrefix ="EmailAddress_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =5355
                    Top =960
                    Width =1755
                    Height =270
                    ColumnWidth =1260
                    TabIndex =13
                    Name ="Password"
                    ControlSource ="Password"
                    StatusBarText ="Logon password (optional)"
                    InputMask ="Password"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4110
                            Top =960
                            Width =1200
                            Height =300
                            Name ="Password Label"
                            Caption ="Password:"
                            EventProcPrefix ="Password_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    ColumnCount =2
                    Left =1575
                    Top =3195
                    Width =1755
                    TabIndex =11
                    Name ="PayBy"
                    ControlSource ="PayBy"
                    RowSourceType ="Value List"
                    RowSource ="1;\"Cash\";2;\"Check\";3;\"American Express\";4;\"Visa\";5;\"Master Card\";6;\"D"
                        "iscover\""
                    ColumnWidths ="0;1080"
                    StatusBarText ="1 = Cash, 2 = Check, 3 = American Express, 4 = Visa, 5 = MasterCard, 6 = Discove"
                        "r"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =60
                            Top =3195
                            Width =1470
                            Height =285
                            Name ="Label24"
                            Caption ="Payment Method:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =5355
                    Top =3195
                    Width =1755
                    TabIndex =12
                    Name ="CCNumber"
                    ControlSource ="CCNumber"
                    StatusBarText ="Credit card number"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =3915
                            Top =3195
                            Width =1395
                            Height =285
                            Name ="Label25"
                            Caption ="Credit Card No.:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5355
                    Top =1365
                    TabIndex =14
                    Name ="Admin"
                    ControlSource ="Admin"
                    StatusBarText ="User is authorized to edit book/author data"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =5630
                            Top =1365
                            Width =600
                            Height =285
                            Name ="Label26"
                            Caption ="Admin"
                        End
                    End
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
' See "frmCustAdd.cls"
