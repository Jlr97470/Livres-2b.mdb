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
    Width =8376
    DatasheetFontHeight =10
    ItemSuffix =26
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    DatasheetGridlinesColor =12632256
    OnFilter ="[Event Procedure]"
    OnApplyFilter ="[Event Procedure]"
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0xff8e0a58f8a6e140
    End
    RecordSource ="tblAuthors"
    Caption ="Authors"
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
            Height =432
            FontName ="Arial"
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
                    ControlTipText ="Add a new author"
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
                    ControlTipText ="Search for authors..."
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
            End
        End
        Begin Section
            Height =8640
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =2040
                    Top =60
                    Width =1080
                    Height =270
                    ColumnWidth =1440
                    Name ="AuthorID"
                    ControlSource ="AuthorID"
                    StatusBarText ="Unique Author ID"
                    OnEnter ="[Event Procedure]"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =1110
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
                    Left =2040
                    Top =408
                    Width =3420
                    Height =270
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="FirstName"
                    ControlSource ="FirstName"
                    StatusBarText ="Author first name"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =1020
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
                    Left =2040
                    Top =756
                    Width =225
                    Height =270
                    ColumnWidth =225
                    TabIndex =2
                    Name ="MiddleInit"
                    ControlSource ="MiddleInit"
                    StatusBarText ="Author middle initial (if any)"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =1590
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
                    Left =2040
                    Top =1104
                    Width =3420
                    Height =270
                    ColumnWidth =2310
                    TabIndex =3
                    Name ="LastName"
                    ControlSource ="LastName"
                    StatusBarText ="Author last name"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =1065
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
                    Left =2040
                    Top =1452
                    Width =3420
                    Height =270
                    ColumnWidth =2310
                    TabIndex =4
                    Name ="EmailAddress"
                    ControlSource ="EmailAddress"
                    StatusBarText ="Email address"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =780
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
                    Left =5535
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
                    Top =4725
                    Width =8145
                    Height =3840
                    ColumnWidth =3000
                    TabIndex =8
                    Name ="Bio"
                    ControlSource ="Bio"
                    StatusBarText ="Author short biography"
                    OnEnter ="[Event Procedure]"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =105
                            Top =4410
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
                Begin Subform
                    OverlapFlags =85
                    Left =105
                    Top =2160
                    Width =8145
                    Height =2100
                    TabIndex =6
                    Name ="fsubAuthorBooks"
                    SourceObject ="Form.fsubAuthorBooks"
                    LinkChildFields ="AuthorID"
                    LinkMasterFields ="AuthorID"
                    OnEnter ="[Event Procedure]"
                End
                Begin PageBreak
                    OverlapFlags =85
                    Top =4320
                    Name ="PageBreak16"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =4395
                    Top =4380
                    Width =3855
                    Height =270
                    TabIndex =7
                    Name ="txtDisplayName"
                    ControlSource ="=[FirstName] & IIf(IsNull([MiddleInit]),Null,\" \" & [MiddleInit] & \".\") & \" "
                        "\" & [LastName]"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =18
                    Left =2340
                    Top =1800
                    Width =2415
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label22"
                    Caption ="(Page down to see author bio.)"
                    FontName ="Times New Roman"
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
' See "frmAuthors.cls"
