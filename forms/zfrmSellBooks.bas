Version =19
VersionRequired =19
Begin Form
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
    Width =5475
    ItemSuffix =38
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x1390f5efcd3de140
    End
    Caption ="Sell some books..."
    MenuBar ="Book Menu Bar"
    Begin
        Begin Label
            FontWeight =700
            BackColor =12632256
        End
        Begin CommandButton
            TextFontFamily =2
        End
        Begin TextBox
            SpecialEffect =2
        End
        Begin FormHeader
            Height =480
            BackColor =12632256
            Name ="FormHeader1"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =96
                    Top =40
                    Width =3030
                    Height =405
                    FontSize =14
                    ForeColor =8355711
                    Name ="Text12"
                    Caption ="Sell Some Books..."
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =60
                    Width =3030
                    Height =405
                    FontSize =14
                    ForeColor =16777215
                    Name ="Text13"
                    Caption ="Sell Some Books..."
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =3315
                    Top =15
                    Width =1080
                    Height =465
                    Name ="cmdSell"
                    Caption ="&Sell!"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =78
                    Left =4395
                    Top =15
                    Width =1080
                    Height =465
                    TabIndex =1
                    Name ="cmdCancel"
                    Caption ="Ca&ncel"
                    OnClick ="[Event Procedure]"
                End
            End
        End
        Begin Section
            Height =2175
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =2895
                    Top =150
                    Height =270
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="BeginDate"
                    Format ="Short Date"
                    ValidationRule =">=#1/1/96#"
                    ValidationText ="Beginning date must be later than Dec. 31, 1995"
                    DefaultValue ="=Date()-180"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =1200
                            Top =150
                            Width =1380
                            Height =285
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Text2"
                            Caption ="Beginning Date:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2895
                    Top =525
                    Width =720
                    Height =270
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="NumDays"
                    Format ="0"
                    ValidationRule ="<=365"
                    ValidationText ="You can't generate sales for more than one year at a time!"
                    DefaultValue ="60"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =1170
                            Top =525
                            Width =1410
                            Height =285
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Text4"
                            Caption ="Number of Days:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2895
                    Top =930
                    Width =720
                    Height =270
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="NumOrders"
                    Format ="0"
                    ValidationRule ="<=50"
                    ValidationText ="You can't generate more than 50 orders per day."
                    DefaultValue ="10"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =945
                            Top =930
                            Width =1635
                            Height =285
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label33"
                            Caption ="Max orders per day:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2895
                    Top =1335
                    Width =720
                    Height =270
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="NumBooks"
                    Format ="0"
                    DefaultValue ="20"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =840
                            Top =1335
                            Width =1740
                            Height =285
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label35"
                            Caption ="Max books per order:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2895
                    Top =1740
                    Width =720
                    Height =270
                    TabIndex =4
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="NumCopies"
                    Format ="0"
                    ValidationRule ="<=50"
                    ValidationText ="You can't order more than 50 copies of a book per order."
                    DefaultValue ="3"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =18
                            Left =855
                            Top =1740
                            Width =1725
                            Height =285
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label37"
                            Caption ="Max copies per book:"
                            FontName ="Times New Roman"
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
' See "zfrmSellBooks.cls"
