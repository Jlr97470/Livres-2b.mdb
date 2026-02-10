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
    ItemSuffix =34
    Left =210
    Top =1230
    Right =10590
    Bottom =7005
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x9cf5c6efcd3de140
    End
    Caption ="Order Print Options"
    OnOpen ="[Event Procedure]"
    MenuBar ="Book Menu Bar"
    OnLoad ="[Event Procedure]"
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
                    Caption ="Order Print Options"
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
                    Caption ="Order Print Options"
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
            Height =1740
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin OptionGroup
                    SpecialEffect =3
                    OverlapFlags =85
                    Left =950
                    Top =144
                    Width =4320
                    Height =1464
                    Name ="optPrtOptions"
                    DefaultValue ="1"
                    Begin
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =1215
                            Top =285
                            OptionValue =1
                            Name ="optOrder"
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    TextAlign =1
                                    TextFontFamily =18
                                    Left =1515
                                    Top =240
                                    Width =3480
                                    Height =285
                                    FontSize =9
                                    BackColor =-2147483633
                                    ForeColor =-2147483630
                                    Name ="Label31"
                                    Caption ="Print the current order"
                                    FontName ="Times New Roman"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =1215
                            Top =773
                            OptionValue =2
                            Name ="optCustomer"
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    TextAlign =1
                                    TextFontFamily =18
                                    Left =1515
                                    Top =720
                                    Width =3420
                                    Height =285
                                    FontSize =9
                                    BackColor =-2147483633
                                    ForeColor =-2147483630
                                    Name ="Text17"
                                    Caption ="Print a mailing label for the current order"
                                    FontName ="Times New Roman"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =1215
                            Top =1261
                            OptionValue =3
                            Name ="optStore"
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    TextAlign =1
                                    TextFontFamily =18
                                    Left =1515
                                    Top =1215
                                    Width =3435
                                    Height =285
                                    FontSize =9
                                    BackColor =-2147483633
                                    ForeColor =-2147483630
                                    Name ="Label29"
                                    Caption ="Print store mailing label for current order"
                                    FontName ="Times New Roman"
                                End
                            End
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
' See "frmOrderPrintParms.cls"
