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
    ItemSuffix =35
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x0c5688efcd3de140
    End
    RecordSource ="qryAuthorsSorted"
    Caption ="Author Search Summary"
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
                    Caption ="Author Search Summary"
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
                    Caption ="Author Search Summary"
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
                    Caption ="Author ID"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =87
                    TextFontFamily =18
                    Left =1215
                    Top =540
                    Width =600
                    Height =285
                    FontSize =9
                    Name ="Text12"
                    Caption ="Name"
                    FontName ="Times New Roman"
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
                    Caption ="Author Search Summary"
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
                    Caption ="Author Search Summary"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =7
                    Top =620
                    Width =945
                    Height =285
                    FontSize =9
                    BackColor =16777215
                    Name ="Label24"
                    Caption ="Author ID"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =1230
                    Top =620
                    Width =615
                    Height =285
                    FontSize =9
                    BackColor =16777215
                    Name ="Label25"
                    Caption ="Name"
                    FontName ="Times New Roman"
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
                    Name ="AuthorID"
                    ControlSource ="AuthorID"
                    StatusBarText ="Club ID"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =945
                    Top =75
                    Width =5400
                    TabIndex =1
                    Name ="AName"
                    ControlSource ="=[FirstName] & IIf(IsNull([MiddleInit]),Null,\" \" & [MiddleInit] & \".\") & \" "
                        "\" & [LastName]"
                    StatusBarText ="Club Name"
                    OnDblClick ="[Event Procedure]"
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
' See "frmAuthorSummary.cls"
