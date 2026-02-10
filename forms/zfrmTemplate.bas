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
    Width =9000
    DatasheetFontHeight =10
    ItemSuffix =60
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    DatasheetGridlinesColor =12632256
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x973bf8efcd3de140
    End
    Caption ="Books"
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
            Height =432
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =65
                    Left =1440
                    Name ="cmdShowAll"
                    Caption ="Show &All"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Show all records"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    TabIndex =1
                    Name ="cmdClose"
                    Caption ="&Close"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Save changes and close the form"
                End
                Begin CommandButton
                    OverlapFlags =215
                    AccessKey =65
                    Left =1440
                    TabIndex =2
                    Name ="cmdAddNew"
                    Caption ="&Add New"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add a new author"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =69
                    Left =4320
                    TabIndex =3
                    Name ="cmdSearch"
                    Caption ="S&earch..."
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Search for books..."
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =2880
                    TabIndex =4
                    Name ="cmdSave"
                    Caption ="&Save"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Save current record"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =78
                    Left =5760
                    TabIndex =5
                    Name ="cmdCancel"
                    Caption ="Ca&ncel"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Cancel edits and close the form"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =4608
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =2
                    OverlapFlags =93
                    Width =9000
                    Height =4608
                    Name ="Box35"
                End
                Begin OptionGroup
                    OverlapFlags =215
                    Left =1080
                    Top =945
                    Name ="Frame38"
                    Begin
                        Begin OptionButton
                            OverlapFlags =215
                            Left =1380
                            Top =1125
                            OptionValue =1
                            Name ="Option40"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =0
                                    Left =1610
                                    Top =1095
                                    Width =915
                                    Height =285
                                    Name ="Label41"
                                    Caption ="Option40"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =215
                            Left =1365
                            Top =1575
                            OptionValue =2
                            Name ="Option42"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =0
                                    Left =1595
                                    Top =1545
                                    Width =915
                                    Height =285
                                    Name ="Label43"
                                    Caption ="Option42"
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    Left =2889
                    Top =270
                    TabIndex =1
                    Name ="Gorp"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =2010
                            Top =270
                            Width =735
                            Height =285
                            Name ="Label45"
                            Caption ="Text44:"
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =215
                    Left =4920
                    Top =930
                    TabIndex =2
                    Name ="Toggle46"
                    Caption ="Toggle"
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =5063
                    Top =1425
                    TabIndex =3
                    Name ="Check51"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4035
                            Top =1395
                            Width =885
                            Height =285
                            Name ="Label52"
                            Caption ="Check51:"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =215
                    Left =5115
                    Top =1965
                    TabIndex =4
                    Name ="Combo55"
                    RowSourceType ="Table/Query"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4041
                            Top =1965
                            Width =930
                            Height =285
                            Name ="Label56"
                            Caption ="Combo55:"
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =215
                    Left =1329
                    Top =2385
                    TabIndex =5
                    Name ="List57"
                    RowSourceType ="Table/Query"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =510
                            Top =2385
                            Width =675
                            Height =285
                            Name ="Label58"
                            Caption ="List57:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =4860
                    Top =300
                    TabIndex =6
                    Name ="Command59"
                    Caption ="Command59"
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
' See "zfrmTemplate.cls"
