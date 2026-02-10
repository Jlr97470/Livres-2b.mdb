Version =19
VersionRequired =19
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    Width =5040
    ItemSuffix =18
    Left =210
    Top =1230
    Right =10590
    Bottom =7005
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0xcf1c7fefcd3de140
    End
    Caption ="Select Authors"
    MenuBar ="Book Menu Bar"
    OnLoad ="[Event Procedure]"
    Begin
        Begin CommandButton
            FontName ="Arial"
        End
        Begin OptionGroup
            SpecialEffect =3
        End
        Begin BoundObjectFrame
            SpecialEffect =2
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
        End
        Begin ListBox
            AutoLabel = NotDefault
            Width =4680
            Height =2130
            FontName ="Arial"
        End
        Begin ComboBox
            SpecialEffect =2
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin Section
            Height =4320
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =86
                    Left =180
                    Top =3015
                    TabIndex =1
                    Name ="cmdSome"
                    Caption ="&View"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="View selected author(s)"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    AccessKey =78
                    Left =1800
                    Top =3442
                    TabIndex =4
                    Name ="cmdNew"
                    Caption ="&New"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Create a new author record"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =3420
                    Top =3015
                    TabIndex =3
                    Name ="cmdSearch"
                    Caption ="&Search..."
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Perform an author search"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =246
                    Top =55
                    Width =1320
                    Height =405
                    FontSize =14
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8355711
                    Name ="Text12"
                    Caption ="Authors"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =210
                    Top =15
                    Width =1320
                    Height =405
                    FontSize =14
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="Text13"
                    Caption ="Authors"
                End
                Begin ListBox
                    OverlapFlags =85
                    MultiSelect =2
                    ColumnCount =2
                    Left =180
                    Top =525
                    Height =2055
                    Name ="lstAName"
                    RowSourceType ="Table/Query"
                    RowSource ="qryAuthForCombo"
                    ColumnWidths ="0;4320"
                    StatusBarText ="List of available authors"
                    OnDblClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    Left =1800
                    Top =3869
                    TabIndex =5
                    Name ="Close"
                    Caption ="&Close"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Close this form"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =65
                    Left =1800
                    Top =3015
                    TabIndex =2
                    Name ="cmdAll"
                    Caption ="View &All"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="View all authors"
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =180
                    Top =2640
                    Width =4680
                    Height =270
                    BackColor =12632256
                    ForeColor =16711680
                    Name ="Label17"
                    Caption ="Use Ctrl or Shift key to select multiple authors with your mouse."
                    FontName ="Times New Roman"
                End
            End
        End
    End
End
CodeBehindForm
' See "frmAuthorList.cls"
