Version =19
VersionRequired =19
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    Width =8310
    ItemSuffix =30
    Left =210
    Top =1230
    Right =10590
    Bottom =7005
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x6713c3efcd3de140
    End
    Caption ="Order Options"
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
        End
        Begin CommandButton
            TextFontFamily =2
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
        End
        Begin ListBox
            SpecialEffect =2
        End
        Begin ComboBox
            SpecialEffect =2
        End
        Begin Section
            Height =3330
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    TextFontFamily =34
                    Left =2565
                    Top =2325
                    Width =465
                    Height =375
                    TabIndex =3
                    Name ="cmd3"
                    Caption ="3."
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Start a new order"
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontFamily =34
                    Left =2565
                    Top =1185
                    Width =465
                    Height =375
                    TabIndex =1
                    Name ="cmd1"
                    Caption ="1."
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Edit all for current customer"
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontFamily =34
                    Left =2565
                    Top =1755
                    Width =465
                    Height =375
                    TabIndex =2
                    Name ="cmd2"
                    Caption ="2."
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Edit open orders for current customer"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =34
                    Left =171
                    Top =55
                    Width =2070
                    Height =405
                    FontSize =14
                    BackColor =12632256
                    ForeColor =8355711
                    Name ="Text12"
                    Caption ="Order Options"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    TextFontFamily =34
                    Left =135
                    Top =15
                    Width =2070
                    Height =405
                    FontSize =14
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="Text13"
                    Caption ="Order Options"
                    FontName ="MS Sans Serif"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    TextFontFamily =34
                    Left =3345
                    Top =2895
                    TabIndex =8
                    Name ="Close"
                    Caption ="&Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close this form"
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontFamily =34
                    Left =6750
                    Top =1185
                    Width =465
                    Height =375
                    TabIndex =5
                    Name ="cmd4"
                    Caption ="4."
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Review orders for customer above"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    ColumnCount =4
                    ListWidth =5400
                    Left =5145
                    Top =810
                    Width =3024
                    Height =270
                    TabIndex =4
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CustomerID"
                    RowSourceType ="Table/Query"
                    RowSource ="qryCustForCombo"
                    ColumnWidths ="0;2880;1440;720"
                    StatusBarText ="Customer for this order"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =4185
                            Top =555
                            Width =915
                            Height =525
                            Name ="Label61"
                            Caption ="Select A\015\012Customer:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =420
                    Top =1140
                    Width =2055
                    Height =525
                    Name ="Label18"
                    Caption ="1) Edit all orders \015\012for the current customer:"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =420
                    Top =1687
                    Width =2055
                    Height =525
                    Name ="Label19"
                    Caption ="2) Edit open orders \015\012for the current customer:"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =420
                    Top =2234
                    Width =2055
                    Height =525
                    Name ="Label20"
                    Caption ="3) Create a new order \015\012for the current customer:"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =4560
                    Top =2370
                    Width =2100
                    Height =285
                    Name ="Label21"
                    Caption ="6) Review all orders:"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =4560
                    Top =1140
                    Width =2100
                    Height =525
                    Name ="Label22"
                    Caption ="4) Review all orders \015\012for the selected customer:"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =4560
                    Top =1687
                    Width =2100
                    Height =525
                    Name ="Label23"
                    Caption ="5) Review open orders \015\012for the selected customer:"
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontFamily =34
                    Left =6750
                    Top =1755
                    Width =465
                    Height =375
                    TabIndex =6
                    Name ="cmd5"
                    Caption ="5."
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Review open orders for customer above"
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontFamily =34
                    Left =6750
                    Top =2325
                    Width =465
                    Height =375
                    TabIndex =7
                    Name ="cmd6"
                    Caption ="6."
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Review all orders"
                End
                Begin Rectangle
                    OverlapFlags =223
                    Left =60
                    Top =510
                    Width =4065
                    Height =2310
                    Name ="Box28"
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =4170
                    Top =510
                    Width =4065
                    Height =2310
                    Name ="Box29"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =215
                    Left =1035
                    Top =810
                    Width =3024
                    Height =270
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtCurrCust"
                    StatusBarText ="Customer for this order"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =90
                            Top =555
                            Width =915
                            Height =525
                            Name ="Label25"
                            Caption ="Current\015\012Customer:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "frmOrderOptions.cls"
