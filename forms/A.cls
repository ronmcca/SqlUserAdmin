Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ---------------------- Start up form for Application -----------------

' Requires a reference to 'Microsoft ActiveX Data Objects 6.1 Library

Private Const ModuleName As String = "Form_A"

' System and VBA Error Codes
Public Enum SystemErrorEnum
    Noerror = 0
    OutOfRange = 5
    Overflow = 6
    PermissionDenied = 70
    NetworkPathNotFound = 76
    ActiveXComponentCannotCreateObject = 429
    DuplicateKey = 3022
    CannotOpenDatabase = 3049
    ODBCCallFailed = 3146                       ' VBA/Access err of ODBC error is noted
    ODBCConnectionFailed = 3151

    LastAccessError = vbObjectError             ' -2147221504 Represents the object error number.
                                                ' User-defined error numbers should be greater than this value.
    ObjectError = vbObjectError + 1
    CannotGenerateNewID = vbObjectError + 2
    IgnoreError = vbObjectError + 3
    SqlApplicationLoginNotFound = vbObjectError + 4
    ADAccountNotFound = vbObjectError + 5
    UserAlreadyAdded = vbObjectError + 6
    MissingInfo = vbObjectError + 14

    LastPAIError = vbObjectError + 15       ' Always last PAI error code
End Enum

Private Type SystemForm
    AppWorkSpace As DAO.Workspace  ' Current Workspace
    AppDatabase As DAO.Database    ' Current Database

    Cryptography As ClassCrypt     ' ClassAdm Uses Decrypt, Encrypt, Hash, Random, RandomInt

    FormError As Long              ' Save last error number
    FormErrorText As String        ' Save last error description
    FormErrorSource As String      ' Save last error source

    LoggedIntoSql As Boolean       ' True if successfull login to SQL server
    SQLUserType As AdmUserTypeEnum ' Type of user account, SQLNone if failed to login

    rsUserList As DAO.Recordset    ' Opened and closed by UpdateUserKeys and AddUserList
End Type
Private This As SystemForm

'---------------------------------
Private WithEvents clsAdm As ClassAdm
Attribute clsAdm.VB_VarHelpID = -1
'---------------------------------

'-------------
' Get user login ID from computer
' lpBuffer      Buffer to hold name, pre-fill with 254 "0" characters
' nSize         Buffer size 255
' Returns number of characters added to buffer
' Retrieves the name of the user associated with the current thread.
' Use the GetUserNameEx function to retrieve the user name in a specified format.
' Additional information is provided by the IADsADSystemInfo interface.
Private Declare PtrSafe Function apiGetUserName _
        Lib "advapi32.dll" _
        Alias "GetUserNameA" _
            (ByVal lpBuffer As String, _
             nSize As Long) _
        As Long

'-------------
' Log the module, function and optional message with error information
'
Private Sub LogMessage(ByVal SourceFunction As String, _
                     Optional ByVal errorMessage As String = vbNullString)
    This.FormErrorSource = ModuleName & "." & SourceFunction
    This.FormError = Err.Number
    This.FormErrorText = Err.Description
    Logger This.FormErrorSource, _
            errorMessage, _
            This.FormError, _
            This.FormErrorText
End Sub

Private Sub ClearLastError()
    On Error Resume Next
    This.FormError = SystemErrorEnum.Noerror
    This.FormErrorText = vbNullString
End Sub

'<><><><><><><> Start Handle ClassAdm Events <><><><><><><><><>

'------------------
' Return the current user's login ID
Private Sub clsAdm_LoginUserID(ByRef UserID)
    Dim logInShortName As String
    Dim lngLen As Long
    logInShortName = String$(254, 0)
    lngLen = 255
    If apiGetUserName(logInShortName, lngLen) > 0 Then
        UserID = Trim$(left$(logInShortName, lngLen - 1))
    End If
End Sub
'---------------------------
' Event to return User ID to ClassADM
' Pre-Open recordset with list of users
' Returns next selected user, vbNullString at end of list
Private Sub clsAdm_NextUserID(ByRef UserID As String)
    On Error GoTo errclsAdm_NextUserID

    If Not This.rsUserList.EOF Then
        UserID = Trim(Nz(This.rsUserList![UserID]))
        This.rsUserList.MoveNext
    Else
        UserID = vbNullString
    End If

doneclsAdm_NextUserID:
    Exit Sub
errclsAdm_NextUserID:
    LogMessage "clsAdm_NextUserID"
    UserID = vbNullString
    Resume doneclsAdm_NextUserID
End Sub
'--------------------------------
' In Production Check Active Directory for User ID
Private Sub clsAdm_VerifyUserID(ByVal UserID As String, _
                                ByRef Response As Boolean)
    Response = True
End Sub
'--------------------------------
' In Production Check Active Directory for user in group
Private Sub clsAdm_VerifyUserInGroup(ByVal UserID As String, _
                                     ByVal AccessType As AdmUserTypeEnum, _
                                     ByRef Response As Boolean)
    Response = True
End Sub
'----------------------------------
' In production get SQL server name from config file
Private Sub clsAdm_SqlServerName(ByRef Response As String)
    Response = "PACK04\JOBCOST"
End Sub
'----------------------------------
' In production get database name from config file
Private Sub clsAdm_SqlDatabaseName(ByRef Response As String)
    Response = "AdmSample"
End Sub
'----------------------------------
' In production get ODBC driver from config file
Private Sub clsAdm_SqlDriveODBC(ByRef Response As String)
    Response = "ODBC;DRIVER=ODBC Driver 18 for SQL Server"
End Sub
'----------------------------------
' In production get OLEDB driver from config file
Private Sub clsAdm_SqlDriveOLEDB(ByRef Response As String)
    Response = "Provider=SQLOLEDB"
End Sub
'-----------------------------------
' This will be "No" if using database user accounts
Private Sub clsAdm_SqlTrustedConnection(ByRef Response As String)
    Response = "No"
End Sub
'-----------------------------------
' In production get "Yes" or "No" from config file
Private Sub clsAdm_SqlEncrypt(ByRef Response As String)
    Response = "No"
End Sub

'<><><><><><><> End Handle ClassAdm Events <><><><><><><><><>

Private Sub Form_Activate()
    On Error Resume Next
    ' Hide the startup form
    DoCmd.Minimize
    Me.Visible = False
End Sub

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo errForm_Open

    ClearLastError

    Set This.Cryptography = New ClassCrypt
    This.Cryptography.LoggerFunctionName = ProjectName & ".Logger"

    Set clsAdm = New ClassAdm
    clsAdm.LoggerFunctionName = ProjectName & ".Logger"
    Set clsAdm.UseCryptography = This.Cryptography

    If Not clsAdm.SqlIsConfigured Then
        ' This can only be done once, on the 1st use of the class
        ' In production remove the UserAdmin password default and optionally the Guest password.
        ' To change the accounts, do a clean start.
        clsAdm.Configure InputBox("Enter User Admin ID", "Setup DB", "UserAdmin"), _
                         InputBox("Enter User Admin Password", "Setup DB", "S1mplePW"), _
                         InputBox("Enter Guest User ID", "Setup DB", "RDP"), _
                         InputBox("Enter Guest User Password", "Setup DB", "Today@1")
    End If

    ' Make 1st ODBC connection to the SQL server
    This.SQLUserType = clsAdm.ODBCConnectToSql()

    This.LoggedIntoSql = CBool(This.SQLUserType <> AdmUserTypeEnum.SQLNone)
    If Not This.LoggedIntoSql Or _
       This.SQLUserType = AdmUserTypeEnum.SQLNone Then
        ' This should cause the application to close and Cancel = True
        Logger ModuleName & ".Form_Open", _
                "Failed conneting to SQL server.", _
                clsAdm.LastErrorNumber, _
                clsAdm.LastErrorDescription
        MsgBox "Failed to connect and login to SQL server.", , "Warning"
    End If

    ' Open the main user form
    DoCmd.OpenForm "MainMenu"
    Forms("MainMenu").SetFocus
    Forms("MainMenu").txtNewUserID.SetFocus
    Forms("MainMenu").txtNewUserID = UserLoginID

doneForm_Open:
    On Error Resume Next
    ' ClassAdm is active but the ADODB connection is dropped until needed
    ' The connectin should be dropped after oAdm is used to manange accounts
    clsAdm.ReleaseADODB
    Exit Sub

errForm_Open:
    Logger ModuleName & ".Form_Open", "Error opening form", Err.Number, Err.Description
    MsgBox "Error opening PAI" & vbCrLf & Err.Description, vbCritical, "Error"
    On Error Resume Next
    Application.Quit acQuitSaveNone
    Resume doneForm_Open
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errForm_Unload

    If Not This.rsUserList Is Nothing Then
        This.rsUserList.Close
        Set This.rsUserList = Nothing
    End If
    clsAdm.ReleaseADODB
    Set clsAdm = Nothing
    Set This.AppDatabase = Nothing
    Set This.AppWorkSpace = Nothing
    Set This.Cryptography = Nothing

doneForm_Unload:
    On Error Resume Next
    ' Quit the applation to release the ODBC connection
'    Application.Quit acQuitSaveNone
    Exit Sub

errForm_Unload:
    Logger ModuleName & ".Form_Unload", _
            "PAI.form_unload", _
            Err.Number, _
            Err.Description
    Resume doneForm_Unload
End Sub

'<><><><><><><><> Start public objects <><><><><><><><><><><>

'-----------------
' Allow other FE modules to use the active ClassAdm
' Ex. Form_A.oADM.Databasename returns the database name open in the instance.
Public Property Get oAdm() As ClassAdm
    On Error Resume Next
    Set oAdm = clsAdm
End Property

'------------------
' Allow application to use already initiated Cryptography
Public Property Get oCrypt() As ClassCrypt
    On Error Resume Next
    Set oCrypt = This.Cryptography
End Property

'-------------
' Name of the current VBE project
' https://www.access-programmers.co.uk/forums/threads/how-to-get-the-project-name.332509/
Public Property Get ProjectName() As String
    On Error Resume Next
    Static PName As String
    If PName = vbNullString Then
        PName = Application.VBE.VBProjects(1).Name
    Else
        ProjectName = PName
    End If
End Property

Public Property Get UserLoginID() As String
    On Error Resume Next
    Dim UserID As String
    clsAdm_LoginUserID UserID
    UserLoginID = UserID
End Property

'--------------
' Attempt to login current user or the supplied user
Public Function ODBCTestLogin(Optional testLoginName As String = vbNullString) As Boolean
    On Error Resume Next
    ODBCTestLogin = CBool(clsAdm.ODBCConnectToSql(testLoginName) <> AdmUserTypeEnum.SQLNone)
    clsAdm.ReleaseADODB
End Function

'----------------
' Reference to the current workspace
Public Property Get WS() As DAO.Workspace
    On Error GoTo errWS
    If This.AppWorkSpace Is Nothing Then
        Set This.AppWorkSpace = DBEngine.Workspaces(0)
    End If
doneWS:
    Set WS = This.AppWorkSpace
    Exit Property
errWS:
    LogMessage "WS"
    Resume doneWS
End Property

'----------------
' Reference to the current database
Public Property Get DB() As DAO.Database
    On Error GoTo errDB
    Dim testWillFailIfDatabaseIsNotOpen As String
    If This.AppDatabase Is Nothing Then
        Set This.AppDatabase = WS.Databases(0)
        ' Test that a database is open
        testWillFailIfDatabaseIsNotOpen = This.AppWorkSpace.Name
    End If
doneDB:
    Set DB = This.AppDatabase
    Exit Property
errDB:
    LogMessage "DB"
    Set This.AppWorkSpace = Nothing
    Set This.AppDatabase = Nothing
    Resume doneDB
End Property

Public Property Get SqlServerLoggedIn() As Boolean
    SqlServerLoggedIn = This.LoggedIntoSql
End Property

Public Property Get SQLUserType() As AdmUserTypeEnum
    SQLUserType = This.SQLUserType
End Property

Public Property Get LastErrorNumber() As Long
    On Error Resume Next
    LastErrorNumber = This.FormError
End Property

Public Property Get LastErrorDescription() As String
    On Error Resume Next
    LastErrorDescription = This.FormErrorText
End Property

'--------------------
' Remove 1 user from SQL server and application
Public Function RemoveUser(ByVal UserID As String) As SystemErrorEnum
    On Error GoTo errRemoveUser
    ClearLastError
    If UserID <> vbNullString Then
        If oAdm.RemoveUser(UserID) Then
            With DB.CreateQueryDef(vbNullString, _
                                    "PARAMETERS RemoveUser Text ( 255 ); " & _
                                    "DELETE feUser.* FROM feUser " & _
                                    "WHERE feUser.UserID=[RemoveUser];")
                .Parameters(0) = UserID
                .Execute RecordsetOptionEnum.dbFailOnError
            End With
        End If
        RemoveUser = oAdm.LastErrorNumber
    Else
        RemoveUser = SystemErrorEnum.Noerror
    End If
doneRemoveUser:
    On Error Resume Next
    oAdm.ReleaseADODB
    Exit Function
errRemoveUser:
    RemoveUser = Err.Number
    LogMessage "RemoveUser"
    Resume doneRemoveUser
End Function

'------------
' Add 1 user to SQL server db and application user list
Public Function AddUser(ByVal UserID As String) As SystemErrorEnum
    On Error GoTo errAddUser
    Dim UserExists As Long
    ClearLastError
    If UserID <> vbNullString Then
        If oAdm.AddUser(UserID) Then
            With DB.CreateQueryDef(vbNullString, _
                                "PARAMETERS newUserID Text ( 255 ); " & _
                                "SELECT Count(*) AS UserExist FROM feUser " & _
                                "WHERE feUser.UserID=[newUserID];")
                .Parameters(0) = UserID
                With .OpenRecordset
                    UserExists = CLng(![UserExist])
                End With
            End With
            With DB.CreateQueryDef(vbNullString, _
                    IIf(UserExists, _
                        "PARAMETERS [newUserID] Text ( 255 ); " & _
                        "UPDATE feUser SET feUser.SQLAccountSetup = True " & _
                        "WHERE feUser.UserID=[newUserID];", _
 _
                        "PARAMETERS [newUserID] Text ( 255 ); " & _
                        "INSERT INTO feUser ( UserID, SQLAccountSetup ) " & _
                        "VALUES ([newUserID], True);"))
                .Parameters(0) = UserID
                .Execute RecordsetOptionEnum.dbFailOnError
            End With
        End If
        AddUser = oAdm.LastErrorNumber
    Else
        AddUser = SystemErrorEnum.Noerror
    End If

doneAddUser:
    On Error Resume Next
    oAdm.ReleaseADODB
    Exit Function
errAddUser:
    AddUser = Err.Number
    LogMessage "AddUser"
    Resume doneAddUser
End Function

'----------------
' Update database keys
Public Function AddUserList() As SystemErrorEnum
    On Error GoTo errAddUserList
    ClearLastError
    Set This.rsUserList = DB.OpenRecordset("SELECT feUser.UserID FROM feUser " & _
                                           "WHERE SQLAccountSetup = False;", _
                                            RecordsetTypeEnum.dbOpenDynaset)
    If This.rsUserList.RecordCount > 0 Then
        This.rsUserList.MoveFirst
        ' clsAdm_NextUserID event
        ' supplies the User ID and moves to next record.
        oAdm.AddUserFromNextUserID
        AddUserList = oAdm.LastErrorNumber
        If oAdm.LastErrorNumber = AdmErrorEnum.Noerror Then
            DB.Execute "UPDATE feUser SET feUser.SQLAccountSetup = True " & _
                       "WHERE feUser.SQLAccountSetup<>True;", _
                        RecordsetOptionEnum.dbFailOnError
        ElseIf Not This.rsUserList.EOF Then
            LogMessage "Failed to add user: " & This.rsUserList![UserID]
        End If
    End If

doneAddUserList:
    On Error Resume Next
    If Not This.rsUserList Is Nothing Then
        This.rsUserList.Close
        Set This.rsUserList = Nothing
    End If
    clsAdm.ReleaseADODB
    Exit Function
errAddUserList:
    AddUserList = Err.Number
    LogMessage "AddUserList"
    Resume doneAddUserList
End Function

'----------------
' Update database keys
Public Function UpdateEncription() As SystemErrorEnum
    On Error GoTo errUpdateEncription
    ClearLastError
    Set This.rsUserList = DB.OpenRecordset("SELECT feUser.UserID FROM feUser " & _
                                           "WHERE SQLAccountSetup <> False;", _
                                            RecordsetTypeEnum.dbOpenDynaset)
    If This.rsUserList.RecordCount > 0 Then
        This.rsUserList.MoveFirst
        ' clsAdm_NextUserID event
        ' supplies the User ID and moves to next record.
        oAdm.UpdateDbEncription
        UpdateEncription = oAdm.LastErrorNumber
    Else
        UpdateEncription = SystemErrorEnum.Noerror
    End If

doneUpdateEncription:
    On Error Resume Next
    clsAdm.ReleaseADODB
    Set clsAdm = Nothing
    Set clsAdm = New ClassAdm
    If Not This.rsUserList Is Nothing Then
        This.rsUserList.Close
        Set This.rsUserList = Nothing
    End If
    Exit Function
errUpdateEncription:
    UpdateEncription = Err.Number
    LogMessage "UpdateUserKeys"
    Resume doneUpdateEncription
End Function
