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
    ItemSuffix =19
    Left =210
    Top =1230
    Right =10590
    Bottom =7005
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0xb5fad6efcd3de140
    End
    Caption ="Select Stores"
    MenuBar ="Book Menu Bar"
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
        End
        Begin CommandButton
            TextFontFamily =2
        End
        Begin Section
            Height =5010
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =86
                    TextFontFamily =34
                    Left =825
                    Top =3015
                    Width =1590
                    TabIndex =2
                    Name ="cmdView"
                    Caption ="&View Selected"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Review selected stores(s)"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    TextFontFamily =34
                    Left =825
                    Top =3465
                    Width =1590
                    TabIndex =1
                    Name ="cmdSearch"
                    Caption ="&Search..."
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Perform a store search"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =231
                    Top =55
                    Width =1320
                    Height =405
                    FontSize =14
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8355711
                    Name ="Text12"
                    Caption ="Stores"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =195
                    Top =15
                    Width =1320
                    Height =405
                    FontSize =14
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="Text13"
                    Caption ="Stores"
                End
                Begin ListBox
                    OverlapFlags =85
                    MultiSelect =2
                    ColumnCount =3
                    Left =180
                    Top =525
                    Width =4680
                    Height =2055
                    Name ="lstSName"
                    RowSourceType ="Table/Query"
                    RowSource ="qryStoresForCombo"
                    ColumnWidths ="0;2880;1440"
                    StatusBarText ="List of stores"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    TextFontFamily =34
                    Left =2625
                    Top =3465
                    Width =1590
                    TabIndex =3
                    Name ="cmdClose"
                    Caption ="&Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close this form"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =65
                    TextFontFamily =34
                    Left =2625
                    Top =3000
                    Width =1590
                    TabIndex =4
                    Name ="cmdAll"
                    Caption ="View &All"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="View all stores"
                End
                Begin Label
                    SpecialEffect =2
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =360
                    Top =3915
                    Width =4320
                    Height =1020
                    FontSize =10
                    FontWeight =700
                    ForeColor =16711680
                    Name ="Label17"
                    Caption ="You can also find Microsoft Press books at major national chains such as B. Dalt"
                        "on, Barnes && Noble, Borders, CompUSA, ComputerLand, Egghead Software,  Software"
                        " Etc., and Waldenbooks."
                    FontName ="Times New Roman"
                End
                Begin Label
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =180
                    Top =2640
                    Width =4680
                    Height =270
                    BackColor =12632256
                    ForeColor =16711680
                    Name ="Label18"
                    Caption ="Use Ctrl or Shift key to select multiple stores with your mouse."
                    FontName ="Times New Roman"
                End
            End
        End
    End
End
CodeBehindForm
' See "frmStoreList.cls"
