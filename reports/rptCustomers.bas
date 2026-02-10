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
    ItemSuffix =19
    Left =180
    Top =90
    DatasheetGridlinesColor =12632256
    Toolbar ="Book Print Tools"
    RecSrcDt = Begin
        0x136bcad31543e140
    End
    RecordSource ="qryCustomerSorted"
    Caption ="Customers"
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
            ControlSource ="LastName"
        End
        Begin BreakLevel
            ControlSource ="FirstName"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =750
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =1845
                    Height =510
                    FontSize =20
                    Name ="Label10"
                    Caption ="Customers"
                End
            End
        End
        Begin PageHeader
            Height =420
            OnFormat ="[Event Procedure]"
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =1305
                    Height =300
                    Name ="CustomerID Label"
                    Caption ="Customer ID"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="CustomerID_Label"
                End
                Begin Label
                    Left =1560
                    Top =60
                    Width =2310
                    Height =300
                    Name ="LastName Label"
                    Caption ="Name"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="LastName_Label"
                End
                Begin Label
                    Left =3930
                    Top =60
                    Width =4560
                    Height =300
                    Name ="Address Label"
                    Caption ="Address"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Address_Label"
                End
                Begin Label
                    Left =8550
                    Top =60
                    Width =2310
                    Height =300
                    Name ="PhoneNumber Label"
                    Caption ="Phone Number / Email"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="PhoneNumber_Label"
                End
                Begin Label
                    Left =10920
                    Top =60
                    Width =1260
                    Height =300
                    Name ="Password Label"
                    Caption ="Password?"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Password_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =360
                    Width =12120
                    Name ="Line13"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =634
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =300
                    Top =15
                    Width =915
                    Height =225
                    Name ="CustomerID"
                    ControlSource ="CustomerID"
                    StatusBarText ="Unique Customer ID"
                End
                Begin TextBox
                    Left =1560
                    Top =15
                    Width =2310
                    Height =270
                    TabIndex =1
                    Name ="CustName"
                    ControlSource ="=[LastName] & \", \" & [FirstName] & IIf(IsNull([MiddleInit]),\"\",\" \" & [Midd"
                        "leInit] & \".\")"
                    StatusBarText ="Customer Last Name"
                End
                Begin TextBox
                    Left =3930
                    Top =15
                    Width =4560
                    Height =270
                    TabIndex =2
                    Name ="Address"
                    ControlSource ="Address"
                    StatusBarText ="Billing address"
                End
                Begin TextBox
                    Left =8550
                    Top =15
                    Width =2310
                    Height =270
                    TabIndex =3
                    Name ="PhoneNumber"
                    ControlSource ="PhoneNumber"
                    StatusBarText ="Phone number"
                    InputMask ="!\\(999\") \"000\\-0000"
                End
                Begin TextBox
                    Left =11055
                    Top =15
                    Width =720
                    Height =270
                    TabIndex =4
                    Name ="PWD"
                    ControlSource ="=Not IsNothing([Password])"
                    Format ="Yes/No"
                    StatusBarText ="Logon password (optional)"
                    InputMask ="Password"
                End
                Begin TextBox
                    Left =3930
                    Top =285
                    Width =4560
                    TabIndex =5
                    Name ="CSZ"
                    ControlSource ="=[City] & \", \" & [StateOrProvince] & \"  \" & [PostalCode] & \"  \" & [Country"
                        "]"
                    StatusBarText ="Billing city"
                End
                Begin TextBox
                    Left =8550
                    Top =285
                    Width =2310
                    TabIndex =6
                    Name ="EmailAddress"
                    ControlSource ="EmailAddress"
                    StatusBarText ="Email address"
                End
                Begin Line
                    Visible = NotDefault
                    BorderWidth =1
                    Left =60
                    Top =585
                    Width =12120
                    Name ="linSeparator"
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
                    Name ="Text11"
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
                    Name ="Text12"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"
                End
                Begin Line
                    BorderWidth =3
                    Left =60
                    Top =240
                    Width =12840
                    BorderColor =12632256
                    Name ="Line14"
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
' See "rptCustomers.cls"
