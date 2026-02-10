Version =19
VersionRequired =19
Begin Form
    AutoCenter = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7667
    DatasheetFontHeight =10
    ItemSuffix =19
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    DatasheetGridlinesColor =12632256
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x6294a1efcd3de140
    End
    RecordSource ="qryBooksSorted"
    Caption ="Book Summary"
    MenuBar ="Book Menu Bar"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
        End
        Begin Line
            SpecialEffect =3
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
        End
        Begin TextBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ListBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin FormHeader
            DisplayWhen =2
            Height =885
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Top =600
                    Width =540
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="ISBNNumber Label"
                    Caption ="ISBN"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="ISBNNumber_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =1095
                    Top =600
                    Width =495
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Title Label"
                    Caption ="Title"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Title_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =18
                    Left =5235
                    Top =600
                    Width =450
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="SuggPrice Label"
                    Caption ="SRP"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SuggPrice_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =18
                    Left =5835
                    Top =600
                    Width =675
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="EditionNumber Label"
                    Caption ="Edition"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="EditionNumber_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =6600
                    Top =600
                    Width =480
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Disk Label"
                    Caption ="Disk"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Disk_Label"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =216
                    Top =115
                    Width =3600
                    Height =375
                    FontSize =14
                    FontWeight =700
                    BackColor =8421376
                    ForeColor =8355711
                    Name ="Text7"
                    Caption ="Book Search Summary"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =180
                    Top =75
                    Width =3600
                    Height =375
                    FontSize =14
                    FontWeight =700
                    BackColor =8421376
                    ForeColor =16777215
                    Name ="Text8"
                    Caption ="Book Search Summary"
                    FontName ="Arial"
                End
            End
        End
        Begin PageHeader
            DisplayWhen =1
            Height =945
            Name ="PageHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =75
                    Top =660
                    Width =555
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label10"
                    Caption ="ISBN"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =1110
                    Top =660
                    Width =510
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label11"
                    Caption ="Title"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =18
                    Left =5250
                    Top =660
                    Width =465
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label12"
                    Caption ="SRP"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =18
                    Left =5850
                    Top =660
                    Width =690
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label13"
                    Caption ="Edition"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =6615
                    Top =660
                    Width =495
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label14"
                    Caption ="Disk"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =276
                    Top =160
                    Width =3600
                    Height =375
                    FontSize =14
                    FontWeight =700
                    BackColor =8421376
                    ForeColor =8355711
                    Name ="Label15"
                    Caption ="Book Search Summary"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =240
                    Top =120
                    Width =3600
                    Height =375
                    FontSize =14
                    FontWeight =700
                    BackColor =8421376
                    ForeColor =0
                    Name ="Label16"
                    Caption ="Book Search Summary"
                    FontName ="Arial"
                End
            End
        End
        Begin Section
            Height =375
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =975
                    Height =255
                    ColumnWidth =1035
                    Name ="ISBNNumber"
                    ControlSource ="ISBNNumber"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                    InputMask ="0\\-00000\\-000\\-A;0;_"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1083
                    Top =60
                    Width =3660
                    Height =255
                    ColumnWidth =3000
                    TabIndex =1
                    Name ="Title"
                    ControlSource ="Title"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    Left =4791
                    Top =60
                    Width =900
                    Height =255
                    ColumnWidth =1560
                    TabIndex =2
                    Name ="SuggPrice"
                    ControlSource ="SuggPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =5739
                    Top =60
                    Width =720
                    Height =255
                    ColumnWidth =900
                    TabIndex =3
                    Name ="EditionNumber"
                    ControlSource ="EditionNumber"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =6507
                    Top =60
                    Width =1080
                    Height =255
                    ColumnWidth =1185
                    TabIndex =4
                    Name ="Disk"
                    ControlSource ="Disk"
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
                    Left =6105
                    Top =75
                    Height =255
                    BackColor =16777215
                    ForeColor =0
                    Name ="Text29"
                    ControlSource ="=\"Page  \" & [Page]"
                    FontName ="Arial"
                End
            End
        End
        Begin FormFooter
            Height =435
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =86
                    Left =14
                    Top =45
                    Width =1455
                    FontSize =10
                    FontWeight =700
                    Name ="Details"
                    Caption ="&View Details"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="View details for current book"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    Left =6227
                    Top =47
                    FontSize =10
                    FontWeight =700
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
' See "frmBookSummary.cls"
