Version =19
VersionRequired =19
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =51
    Left =180
    Top =90
    DatasheetGridlinesColor =12632256
    Toolbar ="Book Print Tools"
    RecSrcDt = Begin
        0xe67cc5c21543e140
    End
    RecordSource ="qryBooksSorted"
    Caption ="Microsoft Press"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderColor =8388608
        End
        Begin Line
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            OldBorderStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin ListBox
            OldBorderStyle =0
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            KeepTogether =2
            ControlSource ="Title"
        End
        Begin PageHeader
            Height =1365
            Name ="PageHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =8010
                    Top =240
                    Width =1530
                    Height =510
                    FontSize =20
                    Name ="Label20"
                    Caption ="Books"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10800
                    BorderColor =12632256
                    Name ="Line23"
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =10800
                    BorderColor =12632256
                    Name ="Line24"
                End
                Begin Line
                    BorderWidth =3
                    Top =870
                    Width =10800
                    BorderColor =12632256
                    Name ="Line25"
                End
                Begin Line
                    BorderWidth =3
                    Top =900
                    Width =10800
                    BorderColor =12632256
                    Name ="Line26"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =1815
                    Top =975
                    Width =1395
                    Height =300
                    FontSize =11
                    ForeColor =8388608
                    Name ="ISBNNumber"
                    ControlSource ="ISBNNumber"
                    StatusBarText ="International Standard Book Number"
                    FontName ="Times New Roman"
                    InputMask ="0\\-00000\\-000\\-A;0;_"
                    Begin
                        Begin Label
                            Left =15
                            Top =975
                            Width =1635
                            Height =300
                            Name ="ISBNNumber Label"
                            Caption ="ISBN Number"
                            EventProcPrefix ="ISBNNumber_Label"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =4485
                    Top =975
                    Width =5460
                    Height =300
                    ColumnWidth =4710
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Title"
                    ControlSource ="Title"
                    StatusBarText ="Book title"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            Left =3630
                            Top =975
                            Width =720
                            Height =300
                            Name ="Title Label"
                            Caption ="Title"
                            EventProcPrefix ="Title_Label"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =1350
                    Width =10800
                    BorderColor =12632256
                    Name ="Line40"
                End
                Begin Image
                    BackStyle =0
                    Left =105
                    Top =135
                    Width =5820
                    Height =645
                    Name ="Image18"
                    PictureData = Begin
                        0x28000000840100002b0000000100040000000000ec200000ce0e0000d80e0000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff8877788fffffffffffffffffffffffffffffff ,
                        0xffffffffffffff777778fffffffffffffffffffffffff8877777077fffffffff ,
                        0xfffffffffffffff87777777fffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffff888ffffffffffffffffffffffff88888fffffffffffffff ,
                        0xfffff88888ffffffffff0000ffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffff700000000078ffffffffffffffffffff ,
                        0xffffffffffffffffffff70000000077fffffffffffffffffff77000000000000 ,
                        0x77fffffffffffffffffff8000000000078ffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffff711111177ffffffffffffffffff7711111117fffff ,
                        0xffffffffff711111111fffffffff00007000000000000ffffff000000000ffff ,
                        0xf7000000000008ff000000000008fffffffff800000000000000007fffffff00 ,
                        0x000000000ffffffffffffffff00000000000000000fffffffffffff700000000 ,
                        0x0000000000007ffffffffffffff00000000000000000ffffffffff7000000000 ,
                        0x00ffffffffff000000000007ffffffffffffffff11118fffffffffffffffffff ,
                        0xfffff11117ffffffffffffffffff11111111111117fffffffffffff711111111 ,
                        0x111117fffffffff811111111111111ffffff0000f000000000000ffffff00000 ,
                        0x0000ffffff000000000007ff000000000007ffffffff70000000000000000007 ,
                        0xffffff70000000000ffffffffffffff80000000000000000000fffffffffff70 ,
                        0x0000000000000000000007fffffffffff80000000000000000000fffffffff80 ,
                        0x0000000000ffffffff80000000000007ffffffffffffffff11117fffffffffff ,
                        0xfffffffffffff11117fffffffffffffffff1111111111111117fffffffffff11 ,
                        0x111111111111118ffffffff11111111111111118ffff0000f7000000000008ff ,
                        0xfff0000000008fffff000000000000ff700000000000fffffff0000000000000 ,
                        0x0000000007ffff800000000008ffffffffffff70000000000000000000000fff ,
                        0xffff7000000000000000000000000007ffffffff70000000000000000000000f ,
                        0xfffffff000000000007fffffff00000000000000ffffffffffffffff11111fff ,
                        0xfffffffffffffffffffff11111fffffffffffffff81111111111111111118fff ,
                        0xfffff1111111111111111117fffff71111111111111111118fff0000f8000000 ,
                        0x000000ffff800000000008ffff0000000000007ff000000000007fffff800000 ,
                        0x0000000000000000000ffff00000000007fffffffffff0000000000000000000 ,
                        0x0000008fffff00000000000000000000000000007ffffff70000000000000000 ,
                        0x000000008ffffff700000000000ffffff700000000000000ffffffffffffffff ,
                        0x711117fffffffffffffffffffffff11111fffffffffffffff111111fffffff81 ,
                        0x111118fffff81111117ffffffff111111fff8111117ffffffff711111fff0000 ,
                        0xff000000000000ffff800000000000ffff0000000000007ff000000000007fff ,
                        0xf80000000000000000000000000ffff70000000000fffffffffff00000000000 ,
                        0x0000000000000008fff7000000000000000000000000000008ffff8000000000 ,
                        0x000000000000000008fffff800000000000ffffff000000000000000ffffffff ,
                        0xffffffff811117fffffffffffffffffffffff71111ffffffffffffff7111118f ,
                        0xfffffff8111111fffff711111ffffffffff811111fff111111ffffffffff8111 ,
                        0x17ff0000ff000000000000ffff800000000000ffff7000000000000ff7000000 ,
                        0x00000ffff0000000000000000000000000000fff00000000008fffffffff0000 ,
                        0x0000000000000000000000008ff0000000000000000000000000000000ffff00 ,
                        0x000000000000000000000000008fffff00000000000ffffff000000000000000 ,
                        0x7ffffffffffffffff11111fffffffffffffffffffffff811118ffffffffffff7 ,
                        0x11111ffffffffffff711111ffff11111ffffffffffff811117ff11118fffffff ,
                        0xfffff711117f0000ff0000000000008fff8000000000000fff8000000000000f ,
                        0xff00000000000fff800000000000077700000000000008ff00000000007fffff ,
                        0xfff800000000000078800000000000000ff0000000000008ffffff8000000000 ,
                        0x007ff700000000000087770000000000000fffff00000000000ffffff0000000 ,
                        0x000000777ffffffffffffffff11111ffffffffffffffffffffffff71117fffff ,
                        0xfffffff11111ffffffffffffff711111ff711118fffffffffffff71111f71111 ,
                        0xffffffffffffff11111f0000ff0000000000008fff8000000000000ffff00000 ,
                        0x0000000fff00000000000fff7000000000007ffff0000000000000ff00000000 ,
                        0x000ffffffff7000000000007ffff800000000000078000000000000ffffffff0 ,
                        0x00000000000ff0000000000008ffff70000000000008ffff000000000008ffff ,
                        0xf000000000007ffffffffffffffffffff11111ffffffffffffffffffffffff71 ,
                        0x111ffffffffffff11111fffffffffffffff1111fff71117ffffffffffffff811 ,
                        0x11f11117ffffffffffffff11111f0000ff7000000000000fff80000000000008 ,
                        0xfff7000000000008ff000000000008ff000000000007ffffff0000000000008f ,
                        0x00000000000ffffffff000000000000fffffff0000000000000000000000007f ,
                        0xfffffff800000000000ff000000000007fffffff7000000000007fff70000000 ,
                        0x0007fffff70000000000fffffffffffffffffffff711118fffffffffffffffff ,
                        0xfffffff1111fffffffffff811117fffffffffffffff88ffffffff8ffffffffff ,
                        0xfffff81111fff88fffffffffffffff11111f0000fff0000000000008ff800000 ,
                        0x000000008fff000000000007ff000000000007ff00000000000ffffffff70000 ,
                        0x0000000870000000000fffffff7000000000008ffffffff70000000000000000 ,
                        0x0000007ffffffff700000000000ff00000000000fffffffff000000000000fff ,
                        0x800000000000fffff80000000000fffffffffffffffffffff811117fffffffff ,
                        0xfffffffffffffff1111fffffffffff711117ffffffffffffffffffffffffffff ,
                        0xfffffffffffff71111ffffffffffffffffffff11111f0000fff7000000000007 ,
                        0xff800000000000000fff000000000000ff700000000000ff00000000007fffff ,
                        0xffff000000000007700000000008ffffff700000000000fffffffff800777777 ,
                        0x777777777777777ffffffff000000000000ff00000000007fffffffff8000000 ,
                        0x00000ffff00000000000ffffff00000000008fffffffffffffffffffff11111f ,
                        0xfffffffffffffffffffffff1111fffffffffff711118ffffffffffffffffffff ,
                        0xfffffffffffffffffffff11111fffffffffffffffffff811111f0000fff70000 ,
                        0x00000000ff700000000000000fff000000000000fff00000000000ff00000000 ,
                        0x007ffffffffffffffffffffff00000000007ffffff700000000000ffffffffff ,
                        0x700007ffffffffffffffffffffff800000000000000ff00000000007ffffffff ,
                        0xff000000000000fff70000000000ffffff00000000007fffffffffffffffffff ,
                        0xff71111ffffffffffffffffffffffff11117ffffffffff71111fffffffffffff ,
                        0xfffffffffffffffffffffffffff7111111fffffffffffffffff87111111f0000 ,
                        0xffff000000000000ff7000000000000000ff000000000000fff00000000000ff ,
                        0x00000000007ffffffffffffffffffffff00000000000ffffff000000000000ff ,
                        0xfffffffff00000007ffffffffffffff77000000000000000000ff00000000007 ,
                        0xffffffffff800000000000fff80000000000ffffff00000000000fffffffffff ,
                        0xfffffffffff1111ffffffffffffffffffffffff71111ffffffffff711118ffff ,
                        0xfffffffffffffffffffffffffffffff7711111111fffffffffffffff77111111 ,
                        0x17ff0000ffff000000000000ff7000000000000000ff700000000000fff00000 ,
                        0x000000ff00000000007ffffffffffffffffffffff00000000000ffffff000000 ,
                        0x000000fffffffffff7000000007fffffffffff700000000000000000000ff000 ,
                        0x00000007fffffffffff70000000000ffff00000000007fffff00000000000fff ,
                        0xfffffffffffffffffff1111ffffffffffffffffffffffff811117fffffffff71 ,
                        0x1117fffffffffffffffffffffffffffffffff811111111111fffffffffffff81 ,
                        0x111111111fff0000ffff0000000000008f70000000000000008f700000000000 ,
                        0x7ff000000000007f000000000008fffffffffffffffffffff000000000008fff ,
                        0xff7000000000008fffffffffff00000000007ffffff700000000000000000000 ,
                        0x007ff00000000000fffffffffff800000000008fff00000000007fffff800000 ,
                        0x00000ffffffffffffffffffffff11117ffffffffffffffffffffffff11111fff ,
                        0xffffff8111117777777777777777777777ffffffff8111111111111fffffffff ,
                        0xfff1111111111111ffff0000ffff7000000000000f700000000000000008f000 ,
                        0x000000000ff700000000000f000000000007fffffffffffffffffffff8000000 ,
                        0x000008ffff8000000000000fffffffffff00000000000ffff800000000000000 ,
                        0x0000000000fff000000000007fffffffffff00000000007fff000000000008ff ,
                        0xfff00000000008fffffffffffffffffffff11111ffffffffffffffffffffffff ,
                        0x11111ffffffffff111111111111111111111111111ffffff71111111111178ff ,
                        0xfffffffff7111111111117ffffff0000ffff7000000000000f70000000000000 ,
                        0x0000f0000000000008f8000000000008700000000000ffffffffffffffffffff ,
                        0xff000000000007fffff000000000000fffffffffff000000000008fff0000000 ,
                        0x000000000000000007fff000000000007fffffffffff00000000007fff000000 ,
                        0x000007fffff70000000007fffffffffffffffffffff7111111111111111177ff ,
                        0xffffffff111118fffffffff111111111111111111111111111ffffff11111111 ,
                        0x1117ffffffffffff7111111111117fffffff0000fffff000000000000f700000 ,
                        0x000000000000f0000000000007ff000000000007f00000000000ffffffffffff ,
                        0xffffffffff700000000000fffff700000000000fffffffffff000000000008ff ,
                        0x700000000000000000000007fffff700000000000fffffffffff00000000007f ,
                        0xff700000000000fffff80000000007fffffffffffffffffffff8111111111111 ,
                        0x1111111117ffffff711111fffffffff111111111111111111111111111ffff81 ,
                        0x111111178ffffffffffffff81111111178ffffffffff0000fffff00000000000 ,
                        0x0f70000000070000000000000000000000ff700000000000f000000000007fff ,
                        0xffffffffffffffffff8000000000000fffff000000000007ffffffffff700000 ,
                        0x000008f70000000000000000000008fffffff8000000000008ffffffffff0000 ,
                        0x0000000ffff00000000000ffffff00000000008fffffffffffffffffffff1111 ,
                        0x11111111111111111118ffff811111fffffffff711117ffffffffffffffff111 ,
                        0x11ffff1111118ffffffffffffffffff1111118ffffffffffffff0000fffff000 ,
                        0x000000000870000000080000000007000000000000ff800000000000f0000000 ,
                        0x00000fffffffff00000000000ff0000000000000ffff000000000000ffffffff ,
                        0xff700000000008f000000000000000000078ffffffffff000000000000ffffff ,
                        0xffff00000000007ffff00000000000ffffff00000000007fffffffffffffffff ,
                        0xffff1111111111111111111111117ffff111117ffffffff811117fffffffffff ,
                        0xfffff71111ffff111117fffffffffffffffffff111117fffffffffffffff0000 ,
                        0xfffff0000000000000700000000f0000000008000000000000fff00000000000 ,
                        0xff000000000007ffffffff00000000000ff0000000000000008f700000000000 ,
                        0x7fffffffff00000000000ff0000000000000078fffffffffffffff0000000000 ,
                        0x00ffffffffff00000000007ffff00000000000ffffff00000000000fffffffff ,
                        0xffffffffffff11111ffffffffffff711111111fff7111117ffffffff711117ff ,
                        0xfffffffffffff71111ffff11111ffffffffffffffffffff11118ffffffffffff ,
                        0xffff0000fffff0000000000000700000000f0000000000000000000000fff000 ,
                        0x00000000ff7000000000008ffffff800000000000ff000000000000000007000 ,
                        0x000000000fffffffff00000000000ff0000000000000ffffffffffffffff8f80 ,
                        0x00000000000fffffffff00000000007ffff00000000000ffffff00000000000f ,
                        0xffffffffffffffffffff111118fffffffffffff11111117ff8111111ffffffff ,
                        0xf11111fffffffffffffff11111ffff11111ffffffffffffffffffff1111fffff ,
                        0xffffffffffff0000fffff7000000000000000000000ff0000000000000000000 ,
                        0x008ff000000000007ff0000000000008fffff000000000000ff0000000000000 ,
                        0x000000000000000000ffffffff00000000008ff000000000007fffffff000000 ,
                        0x00000ff8000000000000fffffff80000000000fffff700000000007fffff7000 ,
                        0x0000000fffffffffffffffffffff811117fffffffffffffff1111118ff111111 ,
                        0x1ffffffff11111fffffffffffffff11111ffff11111ffffffffffff718fffff1 ,
                        0x111fffffffffffff178f0000ffffff000000000000000000000fff0000000000 ,
                        0x00000000007ff700000000000fff00000000000007ff7000000000007ff70000 ,
                        0x0000000000000000000000000008ffff800000000000fff000000000008fffff ,
                        0xff00000000000fff00000000000000ffff700000000000077000000000000000 ,
                        0x000000000000000770007fffff7007fffffff11111ffffffffffffffff111111 ,
                        0xff1111f1118fffffff71111fffffffffffff811117ffff11111ffffffffffff1 ,
                        0x1111fff11117ffffffffffff11110000ffffff700000000000000000000fff00 ,
                        0x0000000000000000000ff700000000000fff0000000000000000000000000000 ,
                        0x8ff8000000000000000000000000000000000ff7000000000000fff700000000 ,
                        0x007ffffff700000000000fff800000000000007ff70000000000000000000000 ,
                        0x00000000000000000000000000007ffff87ff77ffffff711118fffffffffffff ,
                        0xff811111ff1111f7111fffffff811118ffffffffffff711118ffff711118ffff ,
                        0xffffff711117fff711118ffffffffff711110000ffffff800000000000000000 ,
                        0x000fff700000000000000000000fff00000000000ffff7000000000000000000 ,
                        0x00000000ffff7000000000000000000f0000000000000000000000000008ffff ,
                        0x70000000000077870000000000007ffff7000000000000000000000000000f00 ,
                        0x0000000000000000000000000000000000000fff777fff80fffff811117fffff ,
                        0xfffffffffff111117f71117f111118fffff811111ffffffffff111117ffffff7 ,
                        0x1111fffffffff711117fffff11111fffffffff1111170000fffffff000000000 ,
                        0x00000000000ffff00000000000000000000fff00000000000fffff7000000000 ,
                        0x0000000000000008ffff80000000000000000007f00000000000000000000000 ,
                        0x000ffffff000000000000000000000000000fffffff000000000000000000000 ,
                        0x00007f800000000000000000000000000000000000000fff0f7f0ff7ffffff11 ,
                        0x111ffffffffffffffff711111f81111ff111111111ff81111177ffff87111111 ,
                        0xffffffff1111117fff87111111ffffff71111117fff77111111f0000fffffff0 ,
                        0x0000000000000000000ffff000000000000000000008ff000000000008fffff0 ,
                        0x00000000000000000000008ffffff0000000007000000007f800000000000000 ,
                        0x00000000008ffffff000000000000000000000000007fffffff8000000000000 ,
                        0x000000000007ff800000000000000000000000000000000000000fff7f8708f7 ,
                        0xffffff11111ffffffffffffffff711111ff11118f711111111fff81111111171 ,
                        0x11111118ffffffff711111111111111117fffffff711111111111111118f0000 ,
                        0xfffffff00000000000000000000fffff00000000000000000007ff0000000000 ,
                        0x07fffffff000000000000000000008fffffff0000000000f700000008ff80000 ,
                        0x00000000000000000fffffffff7000000000000000000000007ffffffffff700 ,
                        0x0000000000000000007ffff700000000000000000000000000000000000007ff ,
                        0x0ff7f080ffffff11111ffffffffffffffff7111118f11117fff11111117ffff1 ,
                        0x11111111111111fffffffffff7111111111111118fffffffff71111111111111 ,
                        0x1fff0000fffffff70000000000000000000ffffff0000000000000000000ff70 ,
                        0x0000000000fffffffff70000000000000008fffffffff0000000000fff000000 ,
                        0x7ffff7000000000000000008ffffffffffff7000000000000000000078ffffff ,
                        0xfffffff700000000000000007ffffff800000000000000000000000000000000 ,
                        0x000007fff7ffff0fffffff711117fffffffffffffff7111117f71111ffff7711 ,
                        0x111ffffff71111111117fffffffffffffff81111111111ffffffffffffff8111 ,
                        0x1111117fffff0000fffffff70000000000000000000ffffff000000000000000 ,
                        0x00008f8700000000008fffffffff800000000000008ffffffffff7000000000f ,
                        0xfff870000ffffff000000000000000ffffffffffffffff000000000000000007 ,
                        0xffffffffffffffff8000000000000007fffffff8000000000000000070000000 ,
                        0x00000000000000ffff0780ffffffff711117fffffffffffffff7111117ffffff ,
                        0xffffff71111fffffff877111178ffffffffffffffffff87111178fffffffffff ,
                        0xffffff77111778ffffff0000ffffffff0000000000000000000ffffff7000000 ,
                        0x0000000000007fffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff877000778ffffffffffffffffffffff8770000 ,
                        0x00788fffffffffffffffffffffff887007778ffffffffffffffff70000000000 ,
                        0xfffffff0000000000ffffffffffffffffffffff11111fffffffffffffff11111 ,
                        0x18ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffff0000ffffffff0000000000000000000fffff ,
                        0xff0000000000000000000fffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff00 ,
                        0x000000007ffffff00000000008fffffffffffffffffffff11111ffffffffffff ,
                        0xfff111111fffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffff0000ffffffff0000000000000000 ,
                        0x000fffffff0000000000000000000fffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffff00000000000ffffff00000000007fffffffffffffffffffff11111ffff ,
                        0xffffffffff7111117fffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffff0000ffffffff00000000 ,
                        0x00000000000fffffff8000000000000000000fff77777777777fffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffff0000000000008ffff00000000000fffffffffffffffffffff1 ,
                        0x1111fffffffffffff8111111ffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffff0000ffffffff ,
                        0x7000000000000000007ffffffff8000000000000000008ff000000000008ffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffff800000000000007ff70000000000ffffffffffffff ,
                        0xfffffff8111178888888877111111117ffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff0000 ,
                        0xffffffff8000000000000000000fffffffff000000000000000007ff00000000 ,
                        0x0007ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffff70000000000007ff70000000000ffffff ,
                        0xfffffffffffffff811111111111111111111111fffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffff0000fffffffff000000000000000000fffffffff000000000000000000ff ,
                        0x700000000000ffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff0000000000000fff00000000 ,
                        0x00ffffffffffffffffffffff7111111111111111111117ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffff0000fffffffff700000000000000000ffffffffff00000000000 ,
                        0x000000fff00000000000ffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffff0000000000007ff ,
                        0xffffffffffffffffffffffffffffffff81111111111111111117ffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffff0000ffffffffff00000000000000007ffffffffff000 ,
                        0x00000000000000fff00000000000ffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffff8700000 ,
                        0x000007fffffffffffffffffffffffffffffffffff111111111111111118fffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffff0000ffffffffff00000000000000007fffff ,
                        0xfffff700000000000000008ff000000000008fffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffff8877707777ffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffff0000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    Picture ="H:\\RunAcc97\\Webstuff\\MSPLG16A.bmp"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =510
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader1"
            Begin
                Begin Line
                    BorderWidth =3
                    Width =10800
                    BorderColor =12632256
                    Name ="Line41"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10800
                    BorderColor =12632256
                    Name ="Line42"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =1815
                    Top =105
                    Width =1395
                    Height =300
                    FontSize =11
                    ForeColor =8388608
                    Name ="Text43"
                    ControlSource ="ISBNNumber"
                    StatusBarText ="International Standard Book Number"
                    FontName ="Times New Roman"
                    InputMask ="0\\-00000\\-000\\-A;0;_"
                    Begin
                        Begin Label
                            Left =15
                            Top =105
                            Width =1635
                            Height =300
                            Name ="Label44"
                            Caption ="ISBN Number"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =4485
                    Top =105
                    Width =5460
                    Height =300
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text45"
                    ControlSource ="Title"
                    StatusBarText ="Book title"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            Left =3630
                            Top =105
                            Width =720
                            Height =300
                            Name ="Label46"
                            Caption ="Title"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =480
                    Width =10800
                    BorderColor =12632256
                    Name ="Line47"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =3360
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin Label
                    Left =150
                    Top =2580
                    Width =5146
                    Height =300
                    Name ="Category Label"
                    Caption ="Categories"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Category_Label"
                End
                Begin Line
                    BorderWidth =1
                    Left =150
                    Top =2580
                    Width =5146
                    Name ="Line27"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =150
                    Top =2550
                    Width =5146
                    Name ="Line28"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =150
                    Top =2880
                    Width =5146
                    Name ="Line29"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =150
                    Top =2910
                    Width =5146
                    Name ="Line30"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =1995
                    Top =90
                    Width =1125
                    Height =300
                    FontSize =11
                    ForeColor =8388608
                    Name ="EditionNumber"
                    ControlSource ="EditionNumber"
                    StatusBarText ="Edition number"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            Left =15
                            Top =90
                            Width =990
                            Height =300
                            Name ="EditionNumber Label"
                            Caption ="Edition"
                            EventProcPrefix ="EditionNumber_Label"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =1995
                    Top =443
                    Width =765
                    Height =300
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="CopyrightYear"
                    ControlSource ="CopyrightYear"
                    StatusBarText ="Copyright year"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            Left =15
                            Top =443
                            Width =1710
                            Height =300
                            Name ="CopyrightYear Label"
                            Caption ="Copyright Year"
                            EventProcPrefix ="CopyrightYear_Label"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =1980
                    Top =796
                    Width =765
                    Height =300
                    FontSize =11
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Pages"
                    ControlSource ="Pages"
                    StatusBarText ="Number of pages in the book"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            Top =796
                            Width =855
                            Height =300
                            Name ="Pages Label"
                            Caption ="Pages"
                            EventProcPrefix ="Pages_Label"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =1980
                    Top =1149
                    Width =1410
                    Height =300
                    FontSize =11
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Disk"
                    ControlSource ="Disk"
                    StatusBarText ="Type of disk included with the book (if any)"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            Top =1149
                            Width =735
                            Height =300
                            Name ="Disk Label"
                            Caption ="Disk"
                            EventProcPrefix ="Disk_Label"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextFontFamily =18
                    Left =1980
                    Top =1502
                    Width =1860
                    Height =300
                    FontSize =11
                    TabIndex =4
                    ForeColor =8388608
                    Name ="SuggPrice"
                    ControlSource ="SuggPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Suggested retail price"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            Top =1502
                            Width =720
                            Height =300
                            Name ="SuggPrice Label"
                            Caption ="SRP"
                            EventProcPrefix ="SuggPrice_Label"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =150
                    Top =2164
                    Width =9780
                    Height =300
                    ColumnWidth =2040
                    FontSize =11
                    TabIndex =5
                    ForeColor =8388608
                    Name ="LinerNotes"
                    ControlSource ="LinerNotes"
                    StatusBarText ="Book description"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            Left =15
                            Top =1845
                            Width =1380
                            Height =300
                            Name ="LinerNotes Label"
                            Caption ="Liner Notes:"
                            EventProcPrefix ="LinerNotes_Label"
                        End
                    End
                End
                Begin BoundObjectFrame
                    Left =6480
                    Top =60
                    Width =2016
                    Height =2016
                    TabIndex =6
                    Name ="CoverPix"
                    ControlSource ="CoverPix"
                    StatusBarText ="Picture of the book cover"
                    Begin
                        Begin Label
                            Left =5550
                            Top =90
                            Width =810
                            Height =300
                            Name ="CoverPix Label"
                            Caption ="Cover"
                            EventProcPrefix ="CoverPix_Label"
                        End
                    End
                End
                Begin Label
                    Left =5445
                    Top =2580
                    Width =3616
                    Height =300
                    Name ="Label33"
                    Caption ="Authors"
                End
                Begin Line
                    BorderWidth =1
                    Left =5451
                    Top =2580
                    Width =3616
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =1
                    Left =5451
                    Top =2550
                    Width =3616
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =1
                    Left =5451
                    Top =2880
                    Width =3616
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =1
                    Left =5451
                    Top =2910
                    Width =3616
                    Name ="Line37"
                End
                Begin Subform
                    Left =150
                    Top =3000
                    Width =5146
                    Height =300
                    TabIndex =7
                    Name ="Child38"
                    SourceObject ="Report.rsubBookCategories"
                    LinkChildFields ="ISBNNumber"
                    LinkMasterFields ="ISBNNumber"
                End
                Begin Subform
                    Left =5451
                    Top =3000
                    Width =3616
                    Height =299
                    TabIndex =8
                    Name ="Child48"
                    SourceObject ="Report.rsubBookAuthors"
                    LinkChildFields ="ISBNNumber"
                    LinkMasterFields ="ISBNNumber"
                End
                Begin CheckBox
                    Left =3705
                    Top =120
                    TabIndex =9
                    Name ="OutOfPrint"
                    ControlSource ="OutOfPrint"
                    StatusBarText ="Is this book now out of print?"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =3935
                            Top =90
                            Width =1155
                            Height =300
                            Name ="Label50"
                            Caption ="Out of print"
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Height =540
            Name ="PageFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Top =240
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text21"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5685
                    Top =240
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text22"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"
                End
                Begin Line
                    BorderWidth =3
                    Width =10800
                    BorderColor =12632256
                    Name ="Line31"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10800
                    BorderColor =12632256
                    Name ="Line32"
                End
            End
        End
    End
End
CodeBehindForm
' See "rptBooks.cls"
