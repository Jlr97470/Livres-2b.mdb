Version =19
VersionRequired =19
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    Width =6960
    ItemSuffix =17
    Left =210
    Top =1230
    Right =10590
    Bottom =7005
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x90018befcd3de140
    End
    Caption ="Select Books"
    MenuBar ="Book Menu Bar"
    OnLoad ="[Event Procedure]"
    Begin
        Begin CommandButton
            TextFontFamily =2
        End
        Begin Section
            Height =4425
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =86
                    TextFontFamily =34
                    Left =1155
                    Top =3015
                    TabIndex =3
                    Name ="cmdSome"
                    Caption ="&View"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="View selected book(s)"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    AccessKey =78
                    TextFontFamily =34
                    Left =2745
                    Top =3495
                    TabIndex =1
                    Name ="cmdNew"
                    Caption ="&New"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Create a new book record"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    TextFontFamily =34
                    Left =4335
                    Top =3015
                    TabIndex =2
                    Name ="cmdSearch"
                    Caption ="&Search..."
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Perform a book search"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =216
                    Top =40
                    Width =1320
                    Height =405
                    FontSize =14
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8355711
                    Name ="Text12"
                    Caption ="Books"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =180
                    Width =1320
                    Height =405
                    FontSize =14
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="Text13"
                    Caption ="Books"
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    MultiSelect =2
                    ColumnCount =3
                    Left =135
                    Top =525
                    Width =6696
                    Height =2055
                    Name ="lstBName"
                    RowSourceType ="Table/Query"
                    RowSource ="qryBooksForCombo"
                    ColumnWidths ="1224;4680;432"
                    StatusBarText ="List of available books"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    TextFontFamily =34
                    Left =2745
                    Top =3975
                    TabIndex =4
                    Name ="cmdClose"
                    Caption ="&Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close this form"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =65
                    TextFontFamily =34
                    Left =2745
                    Top =3015
                    TabIndex =5
                    Name ="cmdAll"
                    Caption ="View &All"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="View all books"
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =1155
                    Top =2655
                    Width =4620
                    Height =270
                    BackColor =12632256
                    ForeColor =16711680
                    Name ="Label17"
                    Caption ="Use Ctrl or Shift key to select multiple books with your mouse."
                    FontName ="Times New Roman"
                End
            End
        End
    End
End
CodeBehindForm
' See "frmBookList.cls"
