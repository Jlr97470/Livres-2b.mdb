Version =19
VersionRequired =19
Begin Form
    AllowEditing = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    Cycle =1
    Width =5400
    ItemSuffix =15
    Left =210
    Top =1230
    Right =10590
    Bottom =7005
    OnUnload ="[Event Procedure]"
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x304fd4efcd3de140
    End
    RecordSource ="tblLocalValues"
    Caption ="Microsoft Press Sign On"
    MenuBar ="Book Menu Bar"
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
            TextFontFamily =2
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
        End
        Begin ComboBox
            SpecialEffect =2
        End
        Begin FormHeader
            Height =374
            BackColor =12632256
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    AccessKey =83
                    Left =15
                    Top =14
                    Name ="cmdSignOn"
                    Caption ="&Sign On"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =88
                    Left =1455
                    Top =14
                    TabIndex =1
                    Name ="cmdExit"
                    Caption ="E&xit"
                    OnClick ="[Event Procedure]"
                End
            End
        End
        Begin Section
            Height =4185
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =465
                    Top =105
                    Width =4545
                    Height =1710
                    BackColor =16777215
                    ForeColor =16711680
                    Name ="Text2"
                    Caption ="Please select your Customer Name for this session.\015\012Type in your name, las"
                        "t name first, if you're not yet registered as a Customer in this database.\015\012"
                        "Be sure to enter the correct password if the customer information is secured.  A"
                        "lso, decide whether you want to see all reports in Print Preview or send them di"
                        "rectly to your default printer."
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =4
                    Left =1680
                    Top =1980
                    Width =3240
                    Height =270
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="cmbUserID"
                    ControlSource ="LastUserName"
                    RowSourceType ="Table/Query"
                    RowSource ="qryCustForSignonCombo"
                    ColumnWidths ="2880;0;0;0"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =45
                            Top =1980
                            Width =1515
                            Height =285
                            Name ="Text6"
                            Caption ="Customer Name:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4095
                    Top =2865
                    TabIndex =2
                    Name ="PrintPreview"
                    ControlSource ="PrintPreview"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =705
                            Top =2790
                            Width =3285
                            Height =285
                            Name ="Label9"
                            Caption ="I want to see all reports in Print Preview:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1680
                    Top =2415
                    Width =2625
                    TabIndex =1
                    Name ="txtPassword"
                    InputMask ="Password"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =645
                            Top =2415
                            Width =915
                            Height =285
                            Name ="Label13"
                            Caption ="Password:"
                        End
                    End
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    Left =450
                    Top =3135
                    Width =4560
                    Height =990
                    ForeColor =65535
                    Name ="Label14"
                    Caption ="Note:  Password for all sample customers is \"password\".\015\012           Sign"
                        " on as \"Viescas, John L.\" or add yourself with\015\012           \"Admin\" tur"
                        "ned on to be able to edit  books and \015\012           authors."
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "frmSignOn.cls"
