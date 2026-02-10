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
    Width =6120
    ItemSuffix =65
    Left =210
    Top =1230
    Right =10590
    Bottom =7005
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x83f0b6efcd3de140
    End
    Caption ="Customer Sales"
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
                    Width =2580
                    Height =390
                    FontSize =14
                    ForeColor =8355711
                    Name ="Text12"
                    Caption ="Customer Sales"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =75
                    Width =2580
                    Height =390
                    FontSize =14
                    ForeColor =16777215
                    Name ="Text13"
                    Caption ="Customer Sales"
                    FontName ="Arial"
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    AccessKey =80
                    TextFontFamily =34
                    Left =3960
                    Top =14
                    Width =1080
                    Height =465
                    Name ="cmdPrint"
                    Caption ="&Print"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Perform search"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =78
                    TextFontFamily =34
                    Left =5040
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
            Height =3720
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =2520
                    Top =135
                    Width =2880
                    Name ="LastName"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =1395
                            Top =150
                            Width =1035
                            Height =285
                            FontSize =9
                            Name ="Text21"
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
                    Left =855
                    Top =2850
                    Width =4395
                    Height =780
                    FontSize =10
                    BackColor =16777215
                    ForeColor =16711680
                    Name ="Text28"
                    Caption ="Note:  To search for a string in the middle of a name, start your search text wi"
                        "th an asterisk.  \015\012    For example:     *son"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =285
                    Top =120
                    Width =930
                    Height =285
                    FontSize =9
                    ForeColor =255
                    Name ="Text24"
                    Caption ="Sales for:"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2520
                    Top =530
                    Width =2880
                    TabIndex =1
                    Name ="FirstName"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =1305
                            Top =530
                            Width =1125
                            Height =285
                            FontSize =9
                            Name ="Label52"
                            Caption =" First Name:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2520
                    Top =910
                    Width =2880
                    TabIndex =2
                    Name ="City"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =1125
                            Top =910
                            Width =1305
                            Height =285
                            FontSize =9
                            Name ="Label54"
                            Caption ="City:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2520
                    Top =1290
                    Width =2880
                    TabIndex =3
                    Name ="State"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =1155
                            Top =1296
                            Width =1275
                            Height =285
                            FontSize =9
                            Name ="Label56"
                            Caption ="State/Province:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2520
                    Top =1670
                    Width =2880
                    TabIndex =4
                    Name ="Country"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =1155
                            Top =1676
                            Width =1275
                            Height =285
                            FontSize =9
                            Name ="Label60"
                            Caption ="Country:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2520
                    Top =2050
                    Height =270
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="FromDate"
                    Format ="Short Date"
                    DefaultValue ="=Date()-180"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =1455
                            Top =2050
                            Width =975
                            Height =285
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Text2"
                            Caption ="From Date:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2520
                    Top =2445
                    Height =270
                    TabIndex =6
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="ToDate"
                    Format ="Short Date"
                    DefaultValue ="=Date()"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =1665
                            Top =2445
                            Width =765
                            Height =285
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Text4"
                            Caption ="To Date:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =5520
                    Top =135
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
                    Left =5520
                    Top =518
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
                    Left =5520
                    Top =901
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
                    Left =5520
                    Top =1284
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
                    Left =5520
                    Top =1667
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
' See "frmCustSalesParm.cls"
