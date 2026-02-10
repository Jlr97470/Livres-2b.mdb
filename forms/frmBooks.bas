Version =19
VersionRequired =19
Begin Form
    AutoCenter = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8010
    DatasheetFontHeight =10
    ItemSuffix =52
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    DatasheetGridlinesColor =12632256
    OnFilter ="[Event Procedure]"
    OnApplyFilter ="[Event Procedure]"
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x69c5c1d1f8a6e140
    End
    RecordSource ="tblBooks"
    Caption ="Books"
    MenuBar ="Book Menu Bar"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextAlign =3
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
            LabelAlign =3
            Width =1230
            Height =255
            LabelX =-45
            BackColor =-2147483643
            ForeColor =-2147483640
            FontName ="Arial"
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
            Width =1230
            Height =405
            FontSize =8
            FontName ="Arial"
        End
        Begin FormHeader
            Height =403
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    AccessKey =65
                    Left =1454
                    Top =14
                    Height =389
                    FontSize =10
                    FontWeight =700
                    Name ="cmdShowAll"
                    Caption ="Show &All"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Show all records"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    Left =14
                    Top =14
                    Height =389
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdClose"
                    Caption ="&Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Save changes and close the form"
                End
                Begin CommandButton
                    OverlapFlags =215
                    AccessKey =65
                    Left =1454
                    Top =14
                    Height =389
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="cmdAddNew"
                    Caption ="&Add New"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Add a new book"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =69
                    Left =4334
                    Top =14
                    Height =389
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="cmdSearch"
                    Caption ="S&earch..."
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Search for books..."
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =2894
                    Top =14
                    Height =389
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="cmdSave"
                    Caption ="&Save"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Save current record"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =78
                    Left =5774
                    Top =14
                    Height =389
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="cmdCancel"
                    Caption ="Ca&ncel"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Cancel edits and close the form"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =4230
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =1035
                    Top =487
                    Width =1440
                    ColumnWidth =1035
                    TabIndex =1
                    Name ="ISBNNumber"
                    ControlSource ="ISBNNumber"
                    StatusBarText ="International Standard Book Number"
                    InputMask ="0\\-00000\\-000\\-A;0;_"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =375
                            Top =487
                            Width =615
                            Height =285
                            FontSize =9
                            Name ="ISBNNumber Label"
                            Caption ="ISBN:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="ISBNNumber_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1035
                    Top =60
                    Width =4680
                    ColumnWidth =3000
                    Name ="Title"
                    ControlSource ="Title"
                    StatusBarText ="Book title"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =420
                            Top =60
                            Width =570
                            Height =285
                            FontSize =9
                            Name ="Title Label"
                            Caption ="Title:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="Title_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1035
                    Top =914
                    Width =600
                    ColumnWidth =600
                    TabIndex =3
                    Name ="CopyrightYear"
                    ControlSource ="CopyrightYear"
                    StatusBarText ="Copyright year"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =15
                            Top =914
                            Width =975
                            Height =285
                            FontSize =9
                            Name ="CopyrightYear Label"
                            Caption ="Copyright:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="CopyrightYear_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    Left =1035
                    Top =1341
                    Width =1440
                    ColumnWidth =1560
                    TabIndex =5
                    Name ="SuggPrice"
                    ControlSource ="SuggPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Suggested retail price"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =465
                            Top =1341
                            Width =525
                            Height =285
                            FontSize =9
                            Name ="SuggPrice Label"
                            Caption ="SRP:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="SuggPrice_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =4935
                    Top =480
                    Width =780
                    ColumnWidth =900
                    TabIndex =2
                    Name ="EditionNumber"
                    ControlSource ="EditionNumber"
                    StatusBarText ="Edition number"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =4155
                            Top =480
                            Width =720
                            Height =240
                            FontSize =9
                            Name ="EditionNumber Label"
                            Caption ="Edition:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="EditionNumber_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =4935
                    Top =1334
                    Width =780
                    ColumnWidth =600
                    TabIndex =6
                    Name ="Pages"
                    ControlSource ="Pages"
                    StatusBarText ="Number of pages in the book"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =4215
                            Top =1334
                            Width =660
                            Height =240
                            FontSize =9
                            Name ="Pages Label"
                            Caption ="Pages:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="Pages_Label"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    ColumnCount =2
                    Left =1035
                    Top =1768
                    ColumnWidth =2310
                    TabIndex =7
                    Name ="Disk"
                    ControlSource ="Disk"
                    RowSourceType ="Table/Query"
                    RowSource ="ztblDiskValues"
                    ColumnWidths ="0;1440"
                    StatusBarText ="Type of disk included with the book (if any)"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =18
                            Left =405
                            Top =1768
                            Width =555
                            Height =285
                            FontSize =9
                            Name ="Disk Label"
                            Caption ="Disk:"
                            FontName ="Times New Roman"
                            EventProcPrefix ="Disk_Label"
                        End
                    End
                End
                Begin BoundObjectFrame
                    OverlapFlags =85
                    Left =5859
                    Top =60
                    Width =1995
                    Height =1965
                    TabIndex =9
                    Name ="CoverPix"
                    ControlSource ="CoverPix"
                    StatusBarText ="Picture of the book cover"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =3840
                    Top =1770
                    Width =1875
                    TabIndex =8
                    Name ="TOCDoc"
                    ControlSource ="TOCDoc"
                    StatusBarText ="Link to the book table of contents document"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2595
                            Top =1770
                            Width =1170
                            Height =240
                            Name ="Label37"
                            Caption ="Link to TOC:"
                        End
                    End
                End
                Begin Tab
                    OverlapFlags =85
                    Left =1035
                    Top =2055
                    Width =6810
                    Height =2145
                    FontWeight =700
                    TabIndex =10
                    TabFixedWidth =2226
                    Name ="tabDisplay"
                    FontName ="Arial"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =1166
                            Top =2475
                            Width =6544
                            Height =1590
                            Name ="Page44"
                            Caption ="Categories"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =1166
                                    Top =2520
                                    Width =6465
                                    Height =1530
                                    Name ="fsubBookCategories"
                                    SourceObject ="Form.fsubBookCategories"
                                    LinkChildFields ="ISBNNumber"
                                    LinkMasterFields ="ISBNNumber"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =1166
                            Top =2475
                            Width =6544
                            Height =1590
                            Name ="Page45"
                            Caption ="Authors"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =1166
                                    Top =2520
                                    Width =6465
                                    Height =1545
                                    Name ="fsubBookAuthors"
                                    SourceObject ="Form.fsubBookAuthors"
                                    LinkChildFields ="ISBNNumber"
                                    LinkMasterFields ="ISBNNumber"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =1166
                            Top =2475
                            Width =6544
                            Height =1590
                            Name ="Page48"
                            Caption ="Liner Notes"
                            Begin
                                Begin TextBox
                                    EnterKeyBehavior = NotDefault
                                    ScrollBars =2
                                    OverlapFlags =215
                                    Left =1166
                                    Top =2520
                                    Width =6465
                                    Height =1530
                                    Name ="LinerNotes"
                                    ControlSource ="LinerNotes"
                                    StatusBarText ="Book description"
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =4935
                    Top =914
                    Width =780
                    TabIndex =4
                    Name ="OutOfPrint"
                    ControlSource ="OutOfPrint"
                    Format ="Yes/No"
                    StatusBarText ="Is this book now out of print?"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3765
                            Top =914
                            Width =1110
                            Height =240
                            Name ="Label51"
                            Caption ="Out of print:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =403
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =82
                    Left =14
                    Top =14
                    Height =389
                    FontSize =10
                    FontWeight =700
                    Name ="cmdReports"
                    Caption ="&Reports"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="View order report options"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =79
                    Left =1454
                    Top =14
                    Height =389
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdOrder"
                    Caption ="&Order This"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Order this book"
                End
            End
        End
    End
End
CodeBehindForm
' See "frmBooks.cls"
