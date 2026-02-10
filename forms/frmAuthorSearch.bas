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
    Width =6660
    ItemSuffix =54
    Left =210
    Top =1230
    Right =10590
    Bottom =7005
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x88aa85efcd3de140
    End
    Caption ="Author Search"
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
                    Caption ="Author Search"
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
                    Caption ="Author Search"
                    FontName ="Arial"
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    AccessKey =83
                    TextFontFamily =34
                    Left =4500
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
                    Left =5580
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
            Height =2400
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =1440
                    Top =1200
                    Width =4305
                    TabIndex =5
                    Name ="TypeCodes"
                    FontName ="Arial"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2670
                    Top =435
                    Width =1905
                    TabIndex =1
                    Name ="FirstName"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =1380
                            Top =435
                            Width =1200
                            Height =240
                            FontSize =9
                            Name ="Text21"
                            Caption ="First Name:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2670
                    Top =120
                    Width =3360
                    ColumnWidth =2025
                    Name ="LastName"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =1365
                            Top =120
                            Width =1215
                            Height =240
                            FontSize =9
                            Name ="Text23"
                            Caption ="Last Name:"
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
                    Left =1335
                    Top =1515
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
                    Width =1020
                    Height =285
                    FontSize =9
                    ForeColor =255
                    Name ="Text24"
                    Caption ="Search for:"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    OverlapFlags =247
                    Left =1035
                    Top =1095
                    Width =5010
                    TabIndex =4
                    Name ="Types"
                    FontName ="Arial"
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    ColumnCount =3
                    ListWidth =3960
                    Left =2175
                    Top =780
                    Width =3865
                    TabIndex =2
                    BoundColumn =2
                    Name ="TypeList"
                    RowSourceType ="Table/Query"
                    RowSource ="qryCategoriesForCombo"
                    ColumnWidths ="1800;1800;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =735
                            Top =780
                            Width =1365
                            Height =240
                            FontSize =9
                            Name ="Text39"
                            Caption ="Book Types:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =195
                    Top =1065
                    Width =705
                    Height =300
                    FontSize =8
                    TabIndex =3
                    Name ="cmdTypeClear"
                    Caption ="Clear:"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Clear the Book Types list"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =6105
                    Top =120
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
                    Left =6105
                    Top =435
                    Width =435
                    Height =285
                    FontSize =9
                    ForeColor =255
                    Name ="Label52"
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
' See "frmAuthorSearch.cls"
