Version =19
VersionRequired =19
Begin Form
    NavigationButtons = NotDefault
    KeyPreview = NotDefault
    ScrollBars =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7620
    DatasheetFontHeight =10
    ItemSuffix =24
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf42f358acea7e140
    End
    RecordSource ="qryAuthorBooks"
    Caption ="fsubAuthorBooks"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnKeyDown ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextAlign =1
            FontWeight =700
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
            Height =570
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =18
                    Left =930
                    Top =285
                    Width =1230
                    Height =285
                    FontSize =9
                    Name ="ISBNNumber Label"
                    Caption ="ISBN Number"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="ISBNNumber_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =6915
                    Top =45
                    Width =705
                    Height =525
                    FontSize =9
                    Name ="SeqNo Label"
                    Caption ="Author \015\012Seq."
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SeqNo_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =0
                    TextFontFamily =18
                    Left =2820
                    Top =285
                    Width =495
                    Height =285
                    FontSize =9
                    Name ="Label6"
                    Caption ="Title"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =0
                    TextFontFamily =18
                    Left =6135
                    Top =285
                    Width =675
                    Height =285
                    FontSize =9
                    Name ="Label9"
                    Caption ="Edition"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =87
                    TextFontFamily =18
                    Width =1815
                    Height =285
                    FontSize =9
                    Name ="Label10"
                    Caption ="Books by This Author:"
                    FontName ="Times New Roman"
                End
            End
        End
        Begin Section
            Height =375
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =3
                    ListWidth =5760
                    Left =870
                    Top =60
                    Width =1380
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="ISBNNumber"
                    ControlSource ="ISBNNumber"
                    RowSourceType ="Table/Query"
                    RowSource ="qryBooksForCombo"
                    ColumnWidths ="1296;3744;360"
                    StatusBarText ="International Standard Book Number"
                    FontName ="Arial"
                    InputMask ="0\\-00000\\-000\\-A;0;_"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =6990
                    Top =60
                    Width =495
                    Height =255
                    ColumnWidth =600
                    TabIndex =4
                    Name ="SeqNo"
                    ControlSource ="SeqNo"
                    StatusBarText ="Author sequence (multiple author books)"
                    FontName ="Arial"
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =2280
                    Top =60
                    Width =3960
                    ColumnWidth =4710
                    TabIndex =2
                    Name ="Title"
                    ControlSource ="Title"
                    FontName ="Arial"
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =6300
                    Top =60
                    Width =525
                    TabIndex =3
                    Name ="EditionNumber"
                    ControlSource ="EditionNumber"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =90
                    Top =45
                    Width =720
                    Height =300
                    Name ="cmdDetails"
                    Caption ="Details"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Show details for this book..."
                End
            End
        End
        Begin FormFooter
            Height =388
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =14
                    Top =14
                    Width =2085
                    Height =374
                    FontWeight =700
                    Name ="cmdOrder"
                    Caption ="Order > Selected Book"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Order book indicated by >"
                End
            End
        End
    End
End
CodeBehindForm
' See "fsubAuthorBooks.cls"
