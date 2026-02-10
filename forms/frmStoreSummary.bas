Version =19
VersionRequired =19
Begin Form
    AllowEditing = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    ViewsAllowed =1
    Width =6450
    ItemSuffix =41
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x2120ae0e8ca7e140
    End
    RecordSource ="qryStoresSorted"
    Caption ="Store Search Summary"
    MenuBar ="Book Menu Bar"
    Begin
        Begin Label
            FontWeight =700
            BackColor =12632256
            FontName ="Arial"
        End
        Begin Rectangle
            SpecialEffect =2
        End
        Begin CommandButton
            TextFontFamily =2
        End
        Begin OptionButton
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =187
            Height =187
        End
        Begin CheckBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =187
            Height =187
        End
        Begin TextBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Height =255
        End
        Begin ListBox
            AutoLabel = NotDefault
            SpecialEffect =2
        End
        Begin ComboBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Height =255
        End
        Begin FormHeader
            DisplayWhen =2
            Height =825
            BackColor =12632256
            Name ="FormHeader1"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =468
                    Top =160
                    Width =3600
                    Height =375
                    FontSize =14
                    BackColor =8421376
                    ForeColor =8355711
                    Name ="Text7"
                    Caption ="Store Search Summary"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =432
                    Top =120
                    Width =3600
                    Height =375
                    FontSize =14
                    BackColor =8421376
                    ForeColor =16777215
                    Name ="Text8"
                    Caption ="Store Search Summary"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    TextFontFamily =18
                    Top =540
                    Width =922
                    Height =285
                    FontSize =9
                    Name ="Text10"
                    Caption ="Store ID"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =87
                    TextFontFamily =18
                    Left =1215
                    Top =540
                    Width =1020
                    Height =285
                    FontSize =9
                    Name ="Text12"
                    Caption ="Store Name"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    Left =4770
                    Top =570
                    Width =1035
                    Height =255
                    Name ="Label35"
                    Caption ="Location"
                End
            End
        End
        Begin PageHeader
            DisplayWhen =1
            Height =905
            Name ="PageHeader"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =261
                    Top =160
                    Width =3600
                    Height =375
                    FontSize =14
                    BackColor =8421376
                    ForeColor =8355711
                    Name ="Label22"
                    Caption ="Store Search Summary"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =225
                    Top =120
                    Width =3600
                    Height =375
                    FontSize =14
                    BackColor =8421376
                    Name ="Label23"
                    Caption ="Store Search Summary"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Top =615
                    Width =922
                    Height =285
                    FontSize =9
                    BackColor =16777215
                    Name ="Label36"
                    Caption ="Store ID"
                    FontName ="Times New Roman"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =1215
                    Top =615
                    Width =1020
                    Height =285
                    FontSize =9
                    BackColor =16777215
                    Name ="Label37"
                    Caption ="Store Name"
                    FontName ="Times New Roman"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =4230
                    Top =645
                    Width =1035
                    Height =255
                    BackColor =16777215
                    Name ="Label38"
                    Caption ="Location"
                End
            End
        End
        Begin Section
            Height =405
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =180
                    Top =75
                    Width =690
                    Name ="StoreID"
                    ControlSource ="StoreID"
                    StatusBarText ="Club ID"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =945
                    Top =75
                    Width =3075
                    TabIndex =1
                    Name ="StoreName"
                    ControlSource ="StoreName"
                    StatusBarText ="Club Name"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =4080
                    Top =75
                    Width =2265
                    TabIndex =2
                    Name ="Location"
                    ControlSource ="=[City] & \", \" & [StateOrProvince]"
                    StatusBarText ="Club Name"
                    FontName ="Arial"
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =360
            Name ="PageFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =4845
                    Top =45
                    Name ="Text29"
                    ControlSource ="=\"Page  \" & [Page]"
                    FontName ="Arial"
                End
            End
        End
        Begin FormFooter
            DisplayWhen =2
            Height =450
            BackColor =12632256
            Name ="FormFooter2"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =86
                    TextFontFamily =34
                    Left =43
                    Top =43
                    Width =1455
                    Name ="Details"
                    Caption ="&View Details"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="View details for current club"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    TextFontFamily =34
                    Left =4980
                    Top =45
                    TabIndex =1
                    Name ="Close"
                    Caption ="&Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close this form"
                End
            End
        End
    End
End
CodeBehindForm
' See "frmStoreSummary.cls"
