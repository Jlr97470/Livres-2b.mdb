Version =19
VersionRequired =19
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6315
    DatasheetFontHeight =10
    ItemSuffix =21
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    DatasheetGridlinesColor =12632256
    OnFilter ="[Event Procedure]"
    OnApplyFilter ="[Event Procedure]"
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x36f3204c8ba7e140
    End
    RecordSource ="qryStores"
    Caption ="Stores"
    MenuBar ="Book Menu Bar"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
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
            Height =446
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    AccessKey =65
                    Left =1454
                    Top =14
                    Height =432
                    FontSize =10
                    FontWeight =700
                    Name ="cmdShowAll"
                    Caption ="Show &All"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Show all records"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    Left =14
                    Top =14
                    Height =432
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdClose"
                    Caption ="&Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close the form"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =69
                    Left =4320
                    Top =14
                    Height =432
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="cmdSearch"
                    Caption ="S&earch..."
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Search for stores..."
                End
            End
        End
        Begin Section
            Height =4275
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =1680
                    Top =120
                    Height =270
                    ColumnWidth =1440
                    Name ="StoreID"
                    ControlSource ="StoreID"
                    StatusBarText ="Unique Store ID within this application"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =720
                            Top =120
                            Width =900
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="StoreID Label"
                            Caption ="Store ID:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="StoreID_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1680
                    Top =2115
                    Width =1560
                    Height =270
                    ColumnWidth =1560
                    TabIndex =5
                    Name ="PostalCode"
                    ControlSource ="PostalCode"
                    StatusBarText ="Store zip or postal code"
                    FontName ="Arial"
                    InputMask ="00000\\-9999"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =465
                            Top =2115
                            Width =1155
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="PostalCode Label"
                            Caption ="Postal Code:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="PostalCode_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1680
                    Top =477
                    Width =4560
                    Height =270
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="StoreName"
                    ControlSource ="StoreName"
                    StatusBarText ="Name of the store"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =450
                            Top =477
                            Width =1170
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="StoreName Label"
                            Caption ="Store Name:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="StoreName_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    Left =1680
                    Top =834
                    Width =4560
                    Height =480
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="Address"
                    ControlSource ="Address"
                    StatusBarText ="Store address"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =780
                            Top =834
                            Width =840
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="Address Label"
                            Caption ="Address:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="Address_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1680
                    Top =1401
                    Width =2310
                    Height =270
                    ColumnWidth =2310
                    TabIndex =3
                    Name ="City"
                    ControlSource ="City"
                    StatusBarText ="Store city"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =1140
                            Top =1401
                            Width =480
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="City Label"
                            Caption ="City:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="City_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1680
                    Top =1758
                    Width =1560
                    Height =270
                    ColumnWidth =1560
                    TabIndex =4
                    Name ="StateOrProvince"
                    ControlSource ="StateOrProvince"
                    StatusBarText ="Store state or province"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =255
                            Top =1758
                            Width =1365
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="StateOrProvince Label"
                            Caption ="State/Province:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="StateOrProvince_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1680
                    Top =2472
                    Width =2310
                    Height =270
                    ColumnWidth =2310
                    TabIndex =6
                    Name ="Country"
                    ControlSource ="Country"
                    StatusBarText ="Store Country"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =735
                            Top =2472
                            Width =885
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="Country Label"
                            Caption ="Country:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="Country_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1680
                    Top =2829
                    Width =1185
                    Height =270
                    ColumnWidth =1185
                    TabIndex =7
                    Name ="PhoneNumber"
                    ControlSource ="PhoneNumber"
                    StatusBarText ="Store phone number"
                    FontName ="Arial"
                    InputMask ="!\\(999\") \"000\\-0000;0;_"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =210
                            Top =2829
                            Width =1410
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="PhoneNumber Label"
                            Caption ="Phone Number:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="PhoneNumber_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1680
                    Top =3186
                    Width =2310
                    Height =270
                    ColumnWidth =2310
                    TabIndex =8
                    Name ="FaxNumber"
                    ControlSource ="FaxNumber"
                    StatusBarText ="Store fax number"
                    FontName ="Arial"
                    InputMask ="!\\(999\") \"000\\-0000"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =420
                            Top =3186
                            Width =1200
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="FaxNumber Label"
                            Caption ="Fax Number:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="FaxNumber_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    Left =1680
                    Top =3543
                    Width =4560
                    Height =270
                    ColumnWidth =3000
                    TabIndex =9
                    Name ="Email"
                    ControlSource ="Email"
                    StatusBarText ="Store email address"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =930
                            Top =3543
                            Width =690
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="Email Label"
                            Caption ="Email:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="Email_Label"
                        End
                    End
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =1680
                    Top =3900
                    Width =4560
                    Height =255
                    TabIndex =10
                    ForeColor =1279872587
                    Name ="WebAddress"
                    ControlSource ="WebAddress"
                    StatusBarText ="Store web address"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =330
                            Top =3900
                            Width =1290
                            Height =240
                            FontWeight =700
                            Name ="Label20"
                            Caption ="Web Address:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =446
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =82
                    Left =14
                    Top =14
                    Height =432
                    FontSize =10
                    FontWeight =700
                    Name ="cmdReports"
                    Caption ="&Reports"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="View store report options"
                End
            End
        End
    End
End
CodeBehindForm
' See "frmStores.cls"
