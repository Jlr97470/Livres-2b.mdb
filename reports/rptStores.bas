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
    Width =9360
    DatasheetFontHeight =10
    ItemSuffix =17
    Left =180
    Top =90
    DatasheetGridlinesColor =12632256
    Toolbar ="Book Print Tools"
    RecSrcDt = Begin
        0x34ec0af0cd3de140
    End
    RecordSource ="qryStoresSorted"
    Caption ="Stores"
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
            ControlSource ="Country"
        End
        Begin BreakLevel
            ControlSource ="StateOrProvince"
        End
        Begin BreakLevel
            ControlSource ="City"
        End
        Begin BreakLevel
            ControlSource ="StoreName"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =1020
            OnFormat ="[Event Procedure]"
            Name ="PageHeader"
            Begin
                Begin Label
                    Left =60
                    Top =660
                    Width =2310
                    Height =300
                    Name ="StoreName Label"
                    Caption ="Store Name"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="StoreName_Label"
                End
                Begin Label
                    Left =2430
                    Top =660
                    Width =4560
                    Height =300
                    Name ="Address Label"
                    Caption ="Address"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Address_Label"
                End
                Begin Label
                    Left =7050
                    Top =660
                    Width =2055
                    Height =300
                    Name ="PhoneNumber Label"
                    Caption ="Phone / Fax / Email"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="PhoneNumber_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =960
                    Width =9072
                    Name ="Line9"
                End
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =75
                    Width =1095
                    Height =510
                    FontSize =20
                    Name ="Label6"
                    Caption ="Stores"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =885
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =60
                    Top =60
                    Width =2310
                    Height =270
                    ColumnWidth =2850
                    Name ="StoreName"
                    ControlSource ="StoreName"
                    StatusBarText ="Name of the store"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    Left =2430
                    Top =60
                    Width =4560
                    Height =270
                    TabIndex =1
                    Name ="Address"
                    ControlSource ="Address"
                    StatusBarText ="Store address"
                End
                Begin TextBox
                    Left =7050
                    Top =60
                    Width =2160
                    Height =270
                    TabIndex =2
                    Name ="PhoneNumber"
                    ControlSource ="PhoneNumber"
                    StatusBarText ="Store phone number"
                    InputMask ="!\\(999\") \"000\\-0000;0;_"
                End
                Begin TextBox
                    Left =2430
                    Top =330
                    Width =4560
                    TabIndex =3
                    Name ="CSZ"
                    ControlSource ="=[City] & \", \" & [StateOrProvince] & \"  \" & [PostalCode]"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    Left =2430
                    Top =570
                    Width =4560
                    TabIndex =4
                    Name ="Country"
                    ControlSource ="Country"
                    StatusBarText ="Store Country"
                End
                Begin TextBox
                    Left =7050
                    Top =330
                    Width =2160
                    TabIndex =5
                    Name ="FaxNumber"
                    ControlSource ="FaxNumber"
                    StatusBarText ="Store fax number"
                    InputMask ="!\\(999\") \"000\\-0000"
                End
                Begin TextBox
                    Left =7050
                    Top =570
                    Width =2160
                    TabIndex =6
                    Name ="Email"
                    ControlSource ="Email"
                    StatusBarText ="Store email address"
                End
                Begin Line
                    Visible = NotDefault
                    Left =60
                    Top =878
                    Width =9072
                    Name ="linSep"
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
                    Width =4560
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text7"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =4740
                    Top =240
                    Width =4560
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text8"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"
                End
                Begin Line
                    BorderWidth =3
                    Left =60
                    Top =240
                    Width =9240
                    BorderColor =12632256
                    Name ="Line10"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "rptStores.cls"
