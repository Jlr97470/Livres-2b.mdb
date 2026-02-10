Version =19
VersionRequired =19
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    Width =5430
    ItemSuffix =26
    Left =180
    Top =180
    Right =10560
    Bottom =5955
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0xfc6cd0efcd3de140
    End
    Caption ="Microsoft Press"
    MenuBar ="Book Menu Bar"
    Begin
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            TextFontFamily =2
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin Section
            Height =3585
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =65
                    TextFontFamily =34
                    Left =375
                    Top =690
                    Width =2055
                    Height =510
                    FontSize =12
                    Name ="cmdAuthors"
                    Caption ="&Authors"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Print list of all authors"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =66
                    TextFontFamily =34
                    Left =375
                    Top =1425
                    Width =2055
                    Height =510
                    FontSize =12
                    TabIndex =1
                    Name ="cmdBooks"
                    Caption ="&Books"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Print list of all books"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    TextFontFamily =34
                    Left =1860
                    Top =2910
                    Width =1728
                    Height =510
                    FontSize =12
                    TabIndex =6
                    Name ="Close"
                    Caption ="&Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close this form"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =336
                    Top =100
                    Width =4755
                    Height =390
                    FontSize =14
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="Text12"
                    Caption ="Book / Author Database - Reports"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    Left =300
                    Top =60
                    Width =4755
                    Height =390
                    FontSize =14
                    FontWeight =700
                    BackColor =12632256
                    Name ="Text13"
                    Caption ="Book / Author Database - Reports"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    TextFontFamily =34
                    Left =3015
                    Top =690
                    Width =2055
                    Height =510
                    FontSize =12
                    TabIndex =3
                    Name ="cmdCustomers"
                    Caption ="&Customers"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Print list of all customers"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    TextFontFamily =34
                    Left =3015
                    Top =2160
                    Width =2055
                    Height =510
                    FontSize =12
                    TabIndex =5
                    Name ="cmdStores"
                    Caption ="&Stores"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Print list of all stores"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =79
                    TextFontFamily =34
                    Left =375
                    Top =2160
                    Width =2055
                    Height =510
                    FontSize =12
                    TabIndex =2
                    Name ="cmdBookSales"
                    Caption ="B&ook Sales"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Analyze book sales"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =85
                    TextFontFamily =34
                    Left =3015
                    Top =1425
                    Width =2055
                    Height =510
                    FontSize =12
                    TabIndex =4
                    Name ="cmdCustSales"
                    Caption ="C&ustomer Sales"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Analyze customer sales"
                End
            End
        End
    End
End
CodeBehindForm
' See "frmReports.cls"
