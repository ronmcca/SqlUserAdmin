Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8160
    DatasheetFontHeight =11
    ItemSuffix =20
    Left =7110
    Top =2655
    Right =15810
    Bottom =7740
    RecSrcDt = Begin
        0xc9809ee98a52e640
    End
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Aptos"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
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
        Begin CommandButton
            TextFontFamily =0
            FontSize =11
            FontWeight =400
            FontName ="Aptos"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
        Begin ListBox
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Aptos"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =5100
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =240
                    Top =360
                    Width =3240
                    Height =300
                    Name ="txtServerTable"

                    LayoutCachedLeft =240
                    LayoutCachedTop =360
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =660
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3600
                    Top =360
                    Width =2640
                    Height =300
                    TabIndex =1
                    Name ="txtLocalTable"

                    LayoutCachedLeft =3600
                    LayoutCachedTop =360
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =660
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6360
                    Top =300
                    Width =1680
                    TabIndex =2
                    Name ="btnLinkTable"
                    Caption ="Link Table"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6360
                    LayoutCachedTop =300
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =660
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6360
                    Top =1020
                    Width =1680
                    TabIndex =4
                    Name ="btnAddUser"
                    Caption ="Add User"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6360
                    LayoutCachedTop =1020
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =1380
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6360
                    Top =1500
                    Width =1680
                    TabIndex =5
                    Name ="btnRemoveUser"
                    Caption ="Remove User"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6360
                    LayoutCachedTop =1500
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =1860
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6360
                    Top =1980
                    Width =1680
                    TabIndex =6
                    Name ="btnVerifyUser"
                    Caption ="Verify User"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6360
                    LayoutCachedTop =1980
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =2340
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =5460
                    Top =2040
                    Width =810
                    Height =315
                    Name ="lbVerified"
                    Caption ="Verified"
                    LayoutCachedLeft =5460
                    LayoutCachedTop =2040
                    LayoutCachedWidth =6270
                    LayoutCachedHeight =2355
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3600
                    Top =1080
                    Width =2640
                    Height =300
                    TabIndex =3
                    Name ="txtNewUserID"

                    LayoutCachedLeft =3600
                    LayoutCachedTop =1080
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =1380
                End
                Begin Subform
                    OverlapFlags =215
                    Left =300
                    Top =1080
                    Width =3120
                    Height =3900
                    TabIndex =7
                    Name ="UserList"
                    SourceObject ="Form.UserList"

                    LayoutCachedLeft =300
                    LayoutCachedTop =1080
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =4980
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =300
                            Top =840
                            Width =885
                            Height =315
                            Name ="UserList Label"
                            Caption ="UserList"
                            EventProcPrefix ="UserList_Label"
                            LayoutCachedLeft =300
                            LayoutCachedTop =840
                            LayoutCachedWidth =1185
                            LayoutCachedHeight =1155
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =4980
                    Top =4140
                    Width =3060
                    TabIndex =8
                    Name ="RedoUserCrypt"
                    Caption ="Update Database Encription"
                    StatusBarText ="Update Database Encription"
                    OnClick ="[Event Procedure]"
                    ShortcutMenuBar ="Update Database Encription"

                    LayoutCachedLeft =4980
                    LayoutCachedTop =4140
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =4500
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =4980
                    Top =3600
                    Width =3060
                    TabIndex =9
                    Name ="btnAddNewUsersFromList"
                    Caption ="Add New Users from List"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add all users from user list that are not configured in SQL server"

                    LayoutCachedLeft =4980
                    LayoutCachedTop =3600
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =3960
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =4980
                    Top =4620
                    Width =3060
                    TabIndex =10
                    Name ="btnCleanStart"
                    Caption ="Clean Start"
                    StatusBarText ="Update Database Encription"
                    OnClick ="[Event Procedure]"
                    ShortcutMenuBar ="Update Database Encription"

                    LayoutCachedLeft =4980
                    LayoutCachedTop =4620
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =4980
                End
            End
        End
    End
End
CodeBehindForm
' See "MainMenu.cls"
