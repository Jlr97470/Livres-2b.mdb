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
    Width =6840
    ItemSuffix =69
    Left =210
    Top =1230
    Right =10590
    Bottom =7005
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0xdde89eefcd3de140
    End
    Caption ="Book Search"
    MenuBar ="Book Menu Bar"
    Begin
        Begin Label
            TextAlign =3
            TextFontFamily =18
            FontSize =9
            FontWeight =700
            BackColor =12632256
            FontName ="Times New Roman"
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
                    TextFontFamily =34
                    Left =111
                    Top =40
                    Width =2280
                    Height =390
                    FontSize =14
                    ForeColor =8355711
                    Name ="Text12"
                    Caption ="Book Search"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    TextFontFamily =34
                    Left =75
                    Width =2280
                    Height =390
                    FontSize =14
                    ForeColor =16777215
                    Name ="Text13"
                    Caption ="Book Search"
                    FontName ="Arial"
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    AccessKey =83
                    TextFontFamily =34
                    Left =4680
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
                    Left =5760
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
            Height =4200
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =1530
                    Top =2625
                    Width =4305
                    TabIndex =10
                    Name ="TypeCodes"
                    FontName ="Arial"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2760
                    Top =818
                    Width =1905
                    TabIndex =2
                    Name ="LastName"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1050
                            Top =818
                            Width =1620
                            Height =285
                            Name ="Text21"
                            Caption ="Author Last Name:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2760
                    Top =465
                    Width =3360
                    ColumnWidth =2025
                    TabIndex =1
                    Name ="Title"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2100
                            Top =465
                            Width =570
                            Height =285
                            Name ="Text23"
                            Caption ="Title:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =210
                    Top =120
                    Width =1005
                    Height =285
                    ForeColor =255
                    Name ="Text24"
                    Caption ="Search for:"
                End
                Begin TextBox
                    OverlapFlags =247
                    Left =1125
                    Top =2515
                    Width =5010
                    TabIndex =9
                    Name ="Types"
                    FontName ="Arial"
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    ColumnCount =3
                    ListWidth =3960
                    Left =2265
                    Top =2162
                    Width =3865
                    TabIndex =7
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
                            Left =1095
                            Top =2162
                            Width =1095
                            Height =285
                            Name ="Text39"
                            Caption ="Book Types:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =285
                    Top =2490
                    Width =705
                    Height =300
                    FontSize =8
                    TabIndex =8
                    Name ="cmdTypeClear"
                    Caption ="Clear:"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Clear the Book Types list"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2760
                    Top =1171
                    Width =1905
                    TabIndex =3
                    Name ="FirstName"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1005
                            Top =1171
                            Width =1665
                            Height =285
                            Name ="Label52"
                            Caption ="Author First Name:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2760
                    Top =1524
                    Width =870
                    TabIndex =4
                    Name ="LowPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1365
                            Top =1524
                            Width =1305
                            Height =285
                            Name ="Label54"
                            Caption ="Price Between:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =4290
                    Top =1524
                    Width =870
                    TabIndex =5
                    Name ="HighPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3765
                            Top =1524
                            Width =435
                            Height =285
                            Name ="Label56"
                            Caption ="And"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2760
                    Top =1875
                    TabIndex =6
                    Name ="Disk"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1350
                            Top =1845
                            Width =1290
                            Height =285
                            Name ="Label58"
                            Caption ="Includes Disk?"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2760
                    Top =120
                    Name ="ISBNNumber"
                    StatusBarText ="International Standard Book Number"
                    InputMask ="C\\-CCCCC\\-CCC\\-C;0;_"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1410
                            Top =120
                            Width =1260
                            Height =285
                            Name ="Label60"
                            Caption ="ISBN Number:"
                        End
                    End
                End
                Begin Label
                    SpecialEffect =2
                    BackStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    Left =945
                    Top =2865
                    Width =5415
                    Height =1260
                    FontSize =10
                    BackColor =16777215
                    ForeColor =16711680
                    Name ="Text28"
                    Caption ="Notes:  \015\012  To search for a string in the middle of a name or title, start"
                        " your search text with an asterisk.       For example:     *son\015\012  To sear"
                        "ch for unknown digits in the ISBN number, enter question marks in those position"
                        "s.   For example:    ?-55615-??*"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6270
                    Top =120
                    Width =435
                    Height =285
                    ForeColor =255
                    Name ="Label51"
                    Caption ="And"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6270
                    Top =471
                    Width =435
                    Height =285
                    ForeColor =255
                    Name ="Label61"
                    Caption ="And"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6270
                    Top =822
                    Width =435
                    Height =285
                    ForeColor =255
                    Name ="Label62"
                    Caption ="And"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6270
                    Top =1173
                    Width =435
                    Height =285
                    ForeColor =255
                    Name ="Label63"
                    Caption ="And"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6270
                    Top =1524
                    Width =435
                    Height =285
                    ForeColor =255
                    Name ="Label64"
                    Caption ="And"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6270
                    Top =1875
                    Width =435
                    Height =285
                    ForeColor =255
                    Name ="Label65"
                    Caption ="And"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5943
                    Top =1875
                    TabIndex =11
                    Name ="OutOfPrint"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3225
                            Top =1845
                            Width =2625
                            Height =285
                            Name ="Label68"
                            Caption ="Don't include books out of print:"
                        End
                    End
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
' See "frmBookSearch.cls"
