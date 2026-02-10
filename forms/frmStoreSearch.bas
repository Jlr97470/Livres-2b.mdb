Version =19
VersionRequired =19
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    Width =6480
    ItemSuffix =59
    Left =210
    Top =1230
    Right =10590
    Bottom =7005
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x76dfe2efcd3de140
    End
    Caption ="Store Search"
    MenuBar ="Book Menu Bar"
    Begin
        Begin Label
            TextAlign =3
            FontWeight =700
            BackColor =12632256
        End
        Begin Rectangle
            SpecialEffect =2
        End
        Begin CommandButton
            TextFontFamily =2
        End
        Begin OptionButton
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            Width =187
            Height =187
            LabelX =-154
        End
        Begin CheckBox
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            Width =187
            Height =187
            LabelX =-154
        End
        Begin TextBox
            SpecialEffect =2
            LabelAlign =3
            Height =255
            LabelX =-174
        End
        Begin ListBox
            SpecialEffect =2
            LabelAlign =3
            LabelX =-154
        End
        Begin ComboBox
            SpecialEffect =2
            LabelAlign =3
            Height =255
            LabelX =-154
        End
        Begin FormHeader
            Height =479
            BackColor =12632256
            Name ="FormHeader1"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =111
                    Top =40
                    Width =2280
                    Height =390
                    FontSize =14
                    ForeColor =8355711
                    Name ="Text12"
                    Caption ="Store Search"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =75
                    Width =2280
                    Height =390
                    FontSize =14
                    ForeColor =16777215
                    Name ="Text13"
                    Caption ="Store Search"
                    FontName ="Arial"
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    AccessKey =83
                    TextFontFamily =34
                    Left =4320
                    Top =14
                    Width =1080
                    Height =465
                    Name ="cmdSearch"
                    Caption ="&Search"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Perform search"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =78
                    TextFontFamily =34
                    Left =5400
                    Top =14
                    Width =1080
                    Height =465
                    TabIndex =1
                    Name ="cmdCancel"
                    Caption ="Ca&ncel"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Exit this form"
                End
            End
        End
        Begin Section
            Height =3240
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =2490
                    Top =483
                    Width =1905
                    TabIndex =1
                    Name ="City"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =1350
                            Top =483
                            Width =1050
                            Height =255
                            FontSize =9
                            Name ="Text21"
                            Caption ="City:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2490
                    Top =120
                    Width =3360
                    ColumnWidth =2025
                    Name ="StoreName"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =1260
                            Top =120
                            Width =1140
                            Height =240
                            FontSize =9
                            Name ="Text23"
                            Caption ="Store Name:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin Label
                    SpecialEffect =2
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =900
                    Top =2340
                    Width =4440
                    Height =780
                    FontSize =10
                    BackColor =16777215
                    ForeColor =16711680
                    Name ="Text28"
                    Caption ="Note:  To search for a string in the middle of a name,  start your search text w"
                        "ith an asterisk.  \015\012    For example:     *son"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =105
                    Top =120
                    Width =990
                    Height =285
                    FontSize =9
                    ForeColor =255
                    Name ="Text24"
                    Caption ="Search for:"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2490
                    Top =846
                    Width =1905
                    TabIndex =2
                    Name ="StateOrProvince"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =1125
                            Top =846
                            Width =1275
                            Height =255
                            FontSize =9
                            Name ="Label52"
                            Caption ="State/Province:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2490
                    Top =1209
                    Width =1905
                    TabIndex =3
                    Name ="PostalCode"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =1335
                            Top =1209
                            Width =1065
                            Height =255
                            FontSize =9
                            Name ="Label54"
                            Caption ="Postal Code:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2490
                    Top =1572
                    Width =1905
                    TabIndex =4
                    Name ="Country"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =1350
                            Top =1572
                            Width =1050
                            Height =255
                            FontSize =9
                            Name ="Label56"
                            Caption ="Country:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2490
                    Top =1935
                    Width =645
                    TabIndex =5
                    Name ="PhoneCode"
                    FontName ="Arial"
                    InputMask ="\\(999\\);0;_"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =900
                            Top =1935
                            Width =1500
                            Height =255
                            FontSize =9
                            Name ="Label58"
                            Caption ="Phone Area Code:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =5940
                    Top =105
                    Width =435
                    Height =285
                    FontSize =9
                    ForeColor =255
                    Name ="Label51"
                    Caption ="And"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =5940
                    Top =471
                    Width =435
                    Height =285
                    FontSize =9
                    ForeColor =255
                    Name ="Label61"
                    Caption ="And"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =5940
                    Top =837
                    Width =435
                    Height =285
                    FontSize =9
                    ForeColor =255
                    Name ="Label62"
                    Caption ="And"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =5940
                    Top =1203
                    Width =435
                    Height =285
                    FontSize =9
                    ForeColor =255
                    Name ="Label63"
                    Caption ="And"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =5940
                    Top =1569
                    Width =435
                    Height =285
                    FontSize =9
                    ForeColor =255
                    Name ="Label64"
                    Caption ="And"
                    FontName ="Times New Roman"
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="FormFooter2"
        End
    End
End
CodeBehindForm
' See "frmStoreSearch.cls"
