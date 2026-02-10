Version =19
VersionRequired =19
Begin Form
    NavigationButtons = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5940
    DatasheetFontHeight =10
    ItemSuffix =7
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x02e2eaefcd3de140
    End
    RecordSource ="qryBookAuthors"
    Caption ="fsubBookAuthors"
    BeforeUpdate ="[Event Procedure]"
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
            Width =720
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
            Height =405
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =945
                    Top =120
                    Width =720
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="AuthorID Label"
                    Caption ="Author"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="AuthorID_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =18
                    Left =4737
                    Top =120
                    Width =1110
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="SeqNo Label"
                    Caption ="Author Seq."
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SeqNo_Label"
                End
            End
        End
        Begin Section
            Height =375
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =2
                    ListWidth =2520
                    Left =885
                    Top =60
                    Width =4230
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="AuthorID"
                    ControlSource ="AuthorID"
                    RowSourceType ="Table/Query"
                    RowSource ="qryAuthForCombo"
                    ColumnWidths ="0;2160"
                    StatusBarText ="Author"
                    FontName ="Arial"
                    OnNotInList ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =5256
                    Top =60
                    Width =510
                    Height =255
                    ColumnWidth =600
                    TabIndex =2
                    Name ="SeqNo"
                    ControlSource ="SeqNo"
                    StatusBarText ="Author sequence (multiple author books)"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Height =300
                    Name ="cmdDetails"
                    Caption ="Details"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Show details for this author..."
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "fsubBookAuthors.cls"
