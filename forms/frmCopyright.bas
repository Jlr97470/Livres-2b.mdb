Version =19
VersionRequired =19
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =3
    Width =5040
    ItemSuffix =15
    Left =210
    Top =1230
    Right =10590
    Bottom =7005
    TimerInterval =5000
    Toolbar ="Book Form Bar"
    RecSrcDt = Begin
        0x6aeba6efcd3de140
    End
    Caption ="MS Press Book / Author Database"
    MenuBar ="Book Menu Bar"
    OnTimer ="[Event Procedure]"
    Begin
        Begin Section
            Height =2070
            Name ="Detail0"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =210
                    Top =180
                    Width =4605
                    Height =1725
                    FontSize =10
                    FontWeight =700
                    Name ="Text7"
                    Caption ="Copyright © 1997, 1999 John L. Viescas.  \015\012Data contents copyright © 1997,"
                        " 1999 Microsoft Press.  Permission to copy individual code samples granted to pu"
                        "rchasers of \015\012\015\012\015\012Commercial resale prohibited."
                    FontName ="Arial"
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =405
                    Top =1230
                    Width =4215
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text9"
                    Caption ="Running Microsoft Access 2000."
                    FontName ="Arial"
                End
            End
        End
    End
End
CodeBehindForm
' See "frmCopyright.cls"
