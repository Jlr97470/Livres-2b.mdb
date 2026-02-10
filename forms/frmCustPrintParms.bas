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
    ItemSuffix =32
    Left =210
    Top =1230
    Right =10590
    Bottom =7005
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x4e0eb3efcd3de140
    End
    Caption ="Customer Report Options"
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
                    Width =3615
                    Height =405
                    FontSize =14
                    ForeColor =8355711
                    Name ="Text12"
                    Caption ="Customer Report Options"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =60
                    Width =3615
                    Height =405
                    FontSize =14
                    ForeColor =16777215
                    Name ="Text13"
                    Caption ="Customer Report Options"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =80
                    Left =3960
                    Top =15
                    Width =1080
                    Height =465
                    Name ="cmdPrint"
                    Caption ="&Print"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =78
                    Left =5040
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
            Height =3060
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =2505
                    Top =2295
                    Height =270
                    TabIndex =1
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
                            TextAlign =3
                            TextFontFamily =18
                            Left =1215
                            Top =2295
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
                    Left =2505
                    Top =2670
                    Height =270
                    TabIndex =2
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
                            TextAlign =3
                            TextFontFamily =18
                            Left =1425
                            Top =2670
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
                Begin OptionGroup
                    SpecialEffect =3
                    OverlapFlags =85
                    Left =735
                    Top =135
                    Width =4395
                    Height =1884
                    Name ="optPrtOptions"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    Begin
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =1095
                            Top =240
                            OptionValue =1
                            Name ="Option30"
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    TextAlign =1
                                    TextFontFamily =18
                                    Left =1395
                                    Top =195
                                    Width =3480
                                    Height =285
                                    FontSize =9
                                    BackColor =-2147483633
                                    ForeColor =-2147483630
                                    Name ="Label31"
                                    Caption ="Listing of all customers currently displayed"
                                    FontName ="Times New Roman"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =1095
                            Top =604
                            OptionValue =2
                            Name ="Button18"
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    TextAlign =1
                                    TextFontFamily =18
                                    Left =1395
                                    Top =555
                                    Width =3510
                                    Height =285
                                    FontSize =9
                                    BackColor =-2147483633
                                    ForeColor =-2147483630
                                    Name ="Text19"
                                    Caption ="Labels for all customers currently displayed"
                                    FontName ="Times New Roman"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =1095
                            Top =968
                            OptionValue =3
                            Name ="Button16"
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    TextAlign =1
                                    TextFontFamily =18
                                    Left =1395
                                    Top =915
                                    Width =2745
                                    Height =285
                                    FontSize =9
                                    BackColor =-2147483633
                                    ForeColor =-2147483630
                                    Name ="Text17"
                                    Caption ="Envelope for the current customer"
                                    FontName ="Times New Roman"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =1095
                            Top =1332
                            OptionValue =4
                            Name ="Button11"
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    TextAlign =1
                                    TextFontFamily =18
                                    Left =1395
                                    Top =1290
                                    Width =2580
                                    Height =285
                                    FontSize =9
                                    BackColor =-2147483633
                                    ForeColor =-2147483630
                                    Name ="Text15"
                                    Caption ="Sales for the current customer"
                                    FontName ="Times New Roman"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =1095
                            Top =1696
                            OptionValue =5
                            Name ="Option28"
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    TextAlign =1
                                    TextFontFamily =18
                                    Left =1395
                                    Top =1650
                                    Width =3435
                                    Height =285
                                    FontSize =9
                                    BackColor =-2147483633
                                    ForeColor =-2147483630
                                    Name ="Label29"
                                    Caption ="Sales for all customers currently displayed"
                                    FontName ="Times New Roman"
                                End
                            End
                        End
                    End
                End
                Begin PageBreak
                    OverlapFlags =85
                    Top =2160
                    Name ="PageBreak20"
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
' See "frmCustPrintParms.cls"
