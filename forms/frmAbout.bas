Version =19
VersionRequired =19
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    Width =5040
    ItemSuffix =16
    Left =210
    Top =1230
    Right =10590
    Bottom =7005
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0xf35078efcd3de140
    End
    Caption ="Microsoft Press Books Database"
    MenuBar ="Book Menu Bar"
    Begin
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin Section
            Height =2985
            Name ="Detail0"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =150
                    Top =180
                    Width =4740
                    Height =2415
                    FontSize =10
                    FontWeight =700
                    Name ="Text7"
                    Caption ="Book / Author Database\015\012\015\012Copyright © 1997, 1999 John L. Viescas.  D"
                        "ata contents copyright © 1997, 1999 Microsoft Press.  Permission to copy individ"
                        "ual code samples granted to purchasers of \015\012\015\012\015\012Commercial res"
                        "ale prohibited."
                    FontName ="Arial"
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =435
                    Top =1710
                    Width =4155
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text9"
                    Caption ="Running Microsoft Access 2000."
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =1725
                    Top =2490
                    FontSize =10
                    FontWeight =700
                    Name ="cmdOK"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                    FontName ="System"
                End
            End
        End
    End
End
CodeBehindForm
' See "frmAbout.cls"
