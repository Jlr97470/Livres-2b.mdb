Version =19
VersionRequired =19
Begin Form
    AutoCenter = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =10
    ItemSuffix =23
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    DatasheetGridlinesColor =12632256
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0xbc10d345f8a6e140
    End
    RecordSource ="tblAuthors"
    Caption ="Add an author..."
    AfterUpdate ="[Event Procedure]"
    MenuBar ="Book Menu Bar"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
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
                    OverlapFlags =85
                    AccessKey =67
                    Left =14
                    Top =14
                    Height =432
                    FontSize =10
                    FontWeight =700
                    Name ="cmdClose"
                    Caption ="&Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Save changes and close the form"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =1454
                    Top =14
                    Height =432
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdSave"
                    Caption ="&Save"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Save current record"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =78
                    Left =2894
                    Top =14
                    Height =432
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="cmdCancel"
                    Caption ="Ca&ncel"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Cancel edits and close the form"
                End
            End
        End
        Begin Section
            Height =3180
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =1605
                    Top =60
                    Width =1080
                    Height =270
                    ColumnWidth =1440
                    TabIndex =6
                    Name ="AuthorID"
                    ControlSource ="AuthorID"
                    StatusBarText ="Unique Author ID"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =675
                            Top =60
                            Width =870
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="AuthorID Label"
                            Caption ="Author ID"
                            FontName ="Times New Roman"
                            EventProcPrefix ="AuthorID_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1605
                    Top =408
                    Width =3420
                    Height =270
                    ColumnWidth =2310
                    Name ="FirstName"
                    ControlSource ="FirstName"
                    StatusBarText ="Author first name"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =585
                            Top =408
                            Width =960
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="FirstName Label"
                            Caption ="First Name"
                            FontName ="Times New Roman"
                            EventProcPrefix ="FirstName_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1605
                    Top =756
                    Width =225
                    Height =270
                    ColumnWidth =225
                    TabIndex =1
                    Name ="MiddleInit"
                    ControlSource ="MiddleInit"
                    StatusBarText ="Author middle initial (if any)"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =1155
                            Top =756
                            Width =390
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="MiddleInit Label"
                            Caption ="M.I."
                            FontName ="Times New Roman"
                            EventProcPrefix ="MiddleInit_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1605
                    Top =1104
                    Width =3420
                    Height =270
                    ColumnWidth =2310
                    TabIndex =2
                    Name ="LastName"
                    ControlSource ="LastName"
                    StatusBarText ="Author last name"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =630
                            Top =1104
                            Width =915
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="LastName Label"
                            Caption ="Last Name"
                            FontName ="Times New Roman"
                            EventProcPrefix ="LastName_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1605
                    Top =1452
                    Width =3420
                    Height =270
                    ColumnWidth =2310
                    TabIndex =3
                    Name ="EmailAddress"
                    ControlSource ="EmailAddress"
                    StatusBarText ="Email address"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =345
                            Top =1467
                            Width =1200
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="EmailAddress Label"
                            Caption ="Email Address"
                            FontName ="Times New Roman"
                            EventProcPrefix ="EmailAddress_Label"
                        End
                    End
                End
                Begin BoundObjectFrame
                    OverlapFlags =85
                    Left =5100
                    Top =60
                    Width =1980
                    Height =1980
                    TabIndex =5
                    Name ="Photograph"
                    ControlSource ="Photograph"
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    Left =105
                    Top =2175
                    Width =6975
                    Height =900
                    ColumnWidth =3000
                    TabIndex =4
                    Name ="Bio"
                    ControlSource ="Bio"
                    StatusBarText ="Author short biography"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =105
                            Top =1845
                            Width =405
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="Bio Label"
                            Caption ="Bio:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="Bio_Label"
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
' See "frmAuthorAdd.cls"
