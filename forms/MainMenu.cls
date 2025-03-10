Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const ModuleName As String = "MainMenu"

Private Sub Form_Close()
    On Error Resume Next
    DoCmd.Close acForm, "A"
End Sub

Private Sub Form_Open(Cancel As Integer)
    On Error Resume Next

    Me.txtServerTable = "dbo.sVar"
    Me.txtLocalTable = "sysVar"

    Me.Caption = Form_A.oAdm.ServerName & "." & Form_A.oAdm.DatabaseName & _
                 IIf(Form_A.oAdm.UsingTestDB, ".Testing", ".Production") & _
                 " LOG-IN: " & Form_A.oAdm.ODBCLogInName

End Sub

Private Sub btnLinkTable_Click()
    On Error Resume Next
    LinkSQL Me.txtServerTable, Me.txtLocalTable
End Sub

Private Sub LinkSQL(ByVal TableSQL As String, ByVal TableLocal As String)
    On Error GoTo errLinkSQL
    Dim NewLink As DAO.TableDef
    Set NewLink = Form_A.DB.CreateTableDef(TableLocal)
    NewLink.Connect = Form_A.oAdm.ODBCConnectionString
    NewLink.SourceTableName = TableSQL
    Form_A.DB.TableDefs.Append NewLink
doneLinkSQL:
    Set NewLink = Nothing
    Exit Sub
errLinkSQL:
    Logger ModuleName & ".LinkSQL", , Err.Number, Err.Description
    Resume doneLinkSQL
End Sub

Private Sub btnAddUser_Click()
    On Error GoTo errbtnAddUser_Click
    Dim NewUser As String
    NewUser = Trim$(Nz(Me.txtNewUserID))
    If NewUser <> vbNullString Then
        Form_A.AddUser NewUser
        Me.Refresh
    End If
donebtnAddUser_Click:
    Exit Sub
errbtnAddUser_Click:
    Logger ModuleName & ".btnAddUser_Click", _
          "Failed to add " & NewUser, _
           Err.Number, _
           Err.Description
    Resume donebtnAddUser_Click
End Sub

Private Sub btnRemoveUser_Click()
    On Error GoTo errbtnRemoveUser_Click
    Dim UserID As String
    UserID = Trim$(Nz(Me.txtNewUserID))
    If UserID <> vbNullString Then
        Form_A.RemoveUser UserID
        Me.Refresh
    End If
donebtnRemoveUser_Click:
    Exit Sub
errbtnRemoveUser_Click:
    Logger ModuleName & ".btnRemoveUser_Click", _
                 "Failed to remove " & UserID, _
                  Err.Number, _
                  Err.Description
    Resume donebtnRemoveUser_Click
End Sub

Private Sub btnVerifyUser_Click()
    On Error GoTo errbtnVerifyUser_Click
    Me.lbVerified.Visible = Form_A.ODBCTestLogin(Trim$(Nz(Me.txtNewUserID)))
donebtnVerifyUser_Click:
    Exit Sub
errbtnVerifyUser_Click:
    Logger ModuleName & ".btnVerifyUser_Click", , Err.Number, Err.Description
    Resume donebtnVerifyUser_Click
End Sub

Private Sub btnAddNewUsersFromList_Click()
    On Error GoTo errbtnAddNewUsersFromList_Click
    Dim NewUsers As Long
    With Form_A.DB.OpenRecordset("SELECT Count(*) AS NewUsers FROM feUser " & _
                                 "WHERE feUser.SQLAccountSetup=False;")
        NewUsers = CLng(![NewUsers])
        .Close
    End With
    If NewUsers > 0 Then
        If MsgBox("Add " & NewUsers & " new users to SQL Server?", vbYesNo, "Add User") = vbYes Then
            Form_A.AddUserList
            If Form_A.LastErrorNumber <> SystemErrorEnum.Noerror Then
                MsgBox "Failed to add all " & NewUsers & " new users from list"
            End If
            Me.Refresh
        End If
    End If
donebtnAddNewUsersFromList_Click:
    Exit Sub
errbtnAddNewUsersFromList_Click:
    Logger ModuleName & ".btnAddNewUsersFromList_Click", , Err.Number, Err.Description
    Resume donebtnAddNewUsersFromList_Click
End Sub

Private Sub RedoUserCrypt_Click()
    On Error GoTo errRedoUserCrypt_Click
    If MsgBox("Update Database Encription!" & vbCrLf & vbCrLf & _
              "This will update all user" & vbCrLf & _
              "login information stored on" & vbCrLf & _
              "the SQL server database: " & _
              Form_A.oAdm.DatabaseName, _
              vbYesNo + vbDefaultButton2, "Confirm") = vbYes Then
        Form_A.UpdateEncription
        Me.Refresh
    End If
doneRedoUserCrypt_Click:
    Exit Sub
errRedoUserCrypt_Click:
    Logger ModuleName & ".RedoUserCrypt_Click", , Err.Number, Err.Description
    Resume doneRedoUserCrypt_Click
End Sub

Private Sub btnCleanStart_Click()
    On Error GoTo errbtnCleanStart_Click
    If MsgBox("Update Database Encription!" & vbCrLf & vbCrLf & _
              "This will remove all user" & vbCrLf & _
              "login information stored on" & vbCrLf & _
              "the SQL server " & Form_A.oAdm.DatabaseName & vbCrLf & _
              "Then delete the application keys", _
              vbYesNo + vbDefaultButton2, "Confirm") = vbYes Then

        If Form_A.CleanStart = SystemErrorEnum.Noerror Then
            MsgBox "1) All user information has been removed from " & Form_A.oAdm.DatabaseName & "." & vbCrLf & _
                   "2) Applications keys have been deleted." & vbCrLf & vbCrLf & _
                   "Todo: Create a new book cipher for the application" & vbCrLf & vbCrLf & _
                   "1) In the immediate window execute 'GenerateBookCipher' " & vbCrLf & _
                   "2) Copy string into ClassAdm.AppKeyBookCipher property." & vbCrLf & _
                   "3) Close Access the reopen the FE.", , "Status"
        Else
            MsgBox "Clean start failed", , "Status"
        End If
        Me.Refresh
    End If
donebtnCleanStart_Click:
    Exit Sub
errbtnCleanStart_Click:
    Logger ModuleName & ".btnCleanStart_Click", , Err.Number, Err.Description
    Resume donebtnCleanStart_Click
End Sub

'-------------
' Sub form uses to fill-in current user
Public Property Let CurrentUserID(ByVal UserID As String)
    On Error Resume Next
    Me.txtNewUserID = Trim(UserID)
    Me.Refresh
End Property

'------------
' Sub form uses to remove user form SQL when deleted from application user list
Public Property Let RemoveUser(ByVal UserID As String)
    On Error Resume Next
    CurrentUserID = UserID
    btnRemoveUser_Click
    Me.Refresh
End Property
