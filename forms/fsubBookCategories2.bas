Version =19
VersionRequired =19
Begin Form
    NavigationButtons = NotDefault
    ScrollBars =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5940
    DatasheetFontHeight =10
    ItemSuffix =8
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd656ecefcd3de140
    End
    RecordSource ="qryBookCategories"
    Caption ="Book Categories"
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
            Height =390
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =180
                    Top =105
                    Width =1080
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Category ID Label"
                    Caption ="Category"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Category_ID_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =4148
                    Top =105
                    Width =555
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label6"
                    Caption ="Class"
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
                    ListWidth =4680
                    Left =150
                    Top =60
                    Width =3600
                    Height =255
                    ColumnWidth =2310
                    BoundColumn =2
                    Name ="CategoryID"
                    ControlSource ="CategoryID"
                    RowSourceType ="Table/Query"
                    RowSource ="qryCategoriesForCombo"
                    ColumnWidths ="2880;1440;0"
                    StatusBarText ="Category"
                    FontName ="Arial"
                    OnNotInList ="[Event Procedure]"
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =3945
                    Top =60
                    Width =1800
                    ColumnWidth =2100
                    TabIndex =1
                    Name ="ClassName"
                    ControlSource ="ClassName"
                    StatusBarText ="Operating System; Application; Audience; Book Type"
                    FontName ="Arial"
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
' See "fsubBookCategories2.cls"
