Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    DefaultView =2
    ViewsAllowed =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =1920
    DatasheetFontHeight =11
    ItemSuffix =3
    Left =6615
    Top =5070
    Right =9720
    Bottom =8220
    AfterInsert ="[Event Procedure]"
    BeforeDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0xda95e5dd8a52e640
    End
    RecordSource ="feUser"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Aptos"
    AllowFormView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =0
            FontSize =11
            FontName ="Aptos"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Aptos"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =390
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Height =300
                    ColumnWidth =2085
                    Name ="subUserID"
                    ControlSource ="UserID"

                    LayoutCachedWidth =1440
                    LayoutCachedHeight =300
                End
                Begin CheckBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =1620
                    ColumnWidth =705
                    TabIndex =1
                    Name ="CK"
                    ControlSource ="SQLAccountSetup"

                    LayoutCachedLeft =1620
                    LayoutCachedWidth =1880
                    LayoutCachedHeight =240
                End
            End
        End
    End
End
CodeBehindForm
' See "UserList.cls"
