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
    Width =4334
    DatasheetFontHeight =10
    ItemSuffix =30
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    DatasheetGridlinesColor =12632256
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0xbab4a5efcd3de140
    End
    RecordSource ="qryCategories"
    Caption ="Add A Category..."
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
            Height =1200
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1320
                    Top =120
                    Width =825
                    TabIndex =2
                    Name ="CategoryID"
                    ControlSource ="CategoryID"
                    StatusBarText ="Unique ID for this category"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =120
                            Top =120
                            Width =1095
                            Height =285
                            Name ="Label26"
                            Caption ="Category ID:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1320
                    Top =810
                    Width =2880
                    ColumnWidth =2055
                    TabIndex =1
                    Name ="Category"
                    ControlSource ="Category"
                    StatusBarText ="Category description"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =360
                            Top =810
                            Width =855
                            Height =285
                            Name ="Label28"
                            Caption ="Category:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =2
                    Left =1320
                    Top =465
                    Width =2880
                    ColumnWidth =1950
                    Name ="ClassID"
                    ControlSource ="ClassID"
                    RowSourceType ="Table/Query"
                    RowSource ="tblClasses"
                    ColumnWidths ="0;1440"
                    StatusBarText ="1 = Op Sys; 2 = Application; 3 = Audience; 4 = Book Type"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =645
                            Top =465
                            Width =570
                            Height =285
                            Name ="Label29"
                            Caption ="Class:"
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
' See "frmCategoryAdd.cls"
