VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'------------------------------- ClassAdm -------------------------------
' (c) Ron McCarry, Pack Air, Inc
' Manange User rights on the SQL Server
'
' Requires a reference to 'Microsoft ActiveX Data Objects 6.1 Library
'
' Requires an instance of a Cryptography class (UseCryptography property).
' See the ClassCrypt module based on Gustav Brock bCrypt module.
'
' Optional logging function otherwise debug.print for logging.
' See LoggerFunctionName property.
'
' (c) Ron McCarry, Pack Air, Inc
' Please feel free to use this code
' without restriction in any application you develop.
' Keep this copyright notice with this code.
'--------------------------------------------------------------------------

Private Const ModuleName As String = "ClassAdm"

'--------------------------------------------------------------------------
' Setup SQL server
' 1) Add Database ([AdmSample])
'   a. Set Options/Containment Type: Partial
'   a. See AdmSample.txt for full script
'   c. See AdmSample.bak for full backup
' 2) Create Users, Role and Schema
'   a. See A1CreateUsersRoleSchema.txt for script
'       .. Edit script for database name
'   b. [jcsm] user admin Schema
'   c. [jcs_user] user role
'       .. Member of db_datareader, db_datawriter
'       .. Securables:
'           Deny: Select, Update, View definition: jcsm.UserID
'           Deny: Execute: 'Application' stored procedures
'       .. Note: FE adds users to this role.
'   d. (PaiJcsApp) user admin account
'       .. Password: S1mplePW
'       .. Default Schema: [jcsm]
'       .. No owned schemas
'       .. No membership of user roles
'       .. No database rights
'       .. Securables: Only execute rights on "Application" stored procedures
'   e. (RDP) Guest read-only accout
'       .. Password: Today@1
'       .. Default schema: [dbo]
'       .. No owned schemas
'       .. Membership: db_datareader
'       .. No Securables
' 3) Add user admin table  (ex. jcsm.UserID)
'   a. See A2CreateUserIdTable.txt for script
'       .. Edit script for database name
'   b. Columns
'       ID (PK, int, not null)
'       WinID (varchar(253), null)
'       AppPw (varchar(253), not null)
'   c. Index
'       PK_UserID (Clustered)
'       UC_WinID (Unique, Non-Clusted)
' 4) Add Stored procedures
'       See Application*.txt files, edit database name
'       jcsm.ApplicationAddUserID
'       jcsm.ApplicationCryptChange
'       jcsm.ApplicationGetPassword
'       jcsm.ApplicationRemoveUserID
'       jcsm.ApplicationUpdatePassword
'       jcsm.ApplicationUserAdd
'       jcsm.ApplicationUserRemove
'       jcsm.ApplicationTestDB
' 5) (Optional) Add sample table sVar
'   a. See CreateSampleSVar.text edit database name
'
' Setup FE
' The encripted application keys are stored in USysVar, normally not visable in the object viewer.
' 1) Open the "USysVar" table in the FE
'   a. Delete the data NOT the records
' 2) Create a new book cipher for the application
'   a. In the Immediate window execute
'       ?GenerateBookCipher
'       ex output: 2yskzj5vE7CKnw5zycBi/GsuVHQ+EEkjfmLvIXJIe/Ghm21MCKsQ0lbn3rr9ZDKzwudOUMln/sBPFf1ZDcvMo
'   b. Goto the "AppKeyBookCipher" property replace the exiting string with the new string.
' 3) Open the Form_A module to enter your connection string information in the following procedures
'       clsAdm_SqlServerName
'       clsAdm_SqlDriveODBC
'       clsAdm_SqlDriveOLEDB
'       clsAdm_SqlTrustedConnection
'       clsAdm_SqlEncrypt
' 4) Configure the application SQL server keys.
'   a. Open the "A" form. (as reference)
'       On 1st run you will be prompted for the user admin and guest accounts you setup in the server.
'       ( The defaults are set to the sample server's name and passwords in form_Open event. )
'       If succesfull the "MainMenu" form will open with the guest account logged in.
'       Your Windows log-in will be in the text box to the left of Add User.
'   b. Click the "Add User" button to add your account to the server
' 5) Close Access then reopen the FE
'   a. Open the "A" form. (as reference)
'   b. Your account should now be logged into the server.
'
' Distribute complied version of FE to your users.
'
' Update all encription keys without changing the book cipher
' 1) Open the "A" form. (as reference)
' 2) Click the "Update Database Encription"
'    button on the "MainMenu" form. (as Referece)
'   a. If it errors, goto Fresh start
'
' Fresh start
' 1) Open the "A" form and remove all users
'   a. If problem
'       .. Remove those users from the SQL server database
'       .. Delete all records in [jcsm].[UserID] table
'       .. Delete all records in FE [feUser]
' 2) Close all forms.
' 3) Open the "USysVar" table in the FE
'   a. Delete the data NOT the records
' 4) Generate new book cipher
' 5) Open the "A" form and re-enter the user admin and guest account information
' 6) Add your users
'
'---------------------------------------------------------------------------------

' Constants referencing 'Active DS Type Library
Private Const admVarChar As Long = 200   ' ADODB.DataTypeEnum.adVarChar
Private Const admInteger As Long = 3     ' ADODB.DataTypeEnum.adInteger
Private Const admParamInput As Long = 1  ' ADODB.ParameterDirectionEnum.adParamInput
Private Const admParamOutput As Long = 2 ' ADODB.ParameterDirectionEnum.adParamOutput

Private Const ALPHABETICCHARACTERS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const NUMERICCHARACTERS As String = "0123456789"
Private Const SINGLESPACE As String = " "
Private Const DOUBLEQUOTE As String = """"

Private Const PASSWORDSPECIALCHAR As String = "!@#$%^&*:.=+-_[]{}()<>"
' Characters allowed in a SQL server password
Private Const PASSWORDCHARACTERS As String = NUMERICCHARACTERS & _
                                             ALPHABETICCHARACTERS & _
                                             "abcdefghijklmnopqrstuvwxyz" & _
                                             PASSWORDSPECIALCHAR

Private Const HashAlgorithem As Long = 6 ' From ClassCrypt, BcHashAlgorithm.bcSha384

' USysVar row ID for application keys
Private Enum AdmKeyEnum
    admUserAdminKey = 1
    admUserAdminID = 3
    admUserAdminPW = 4
    admGuestID = 5
    admGuestLogin = 6
    admAppKeyList = 7
End Enum

Public Enum AdmDatabaseStatusEnum
    UnknownDbStatus
    ProductionDbStatus
    TestDbStatus
End Enum

'-------------------
' Type of SQL login provided   was SQLUserTypeEnum
Public Enum AdmUserTypeEnum
    SQLNone
    SQLGuest
    SQLNormal
    SQLProtected
End Enum

' ClassAdm and VBA Error Codes
Public Enum AdmErrorEnum
    Noerror = 0
    OutOfRange = 5
    PermissionDenied = 70
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

    LastADMError = vbObjectError + 15       ' Always last PAI error code
End Enum

'----------
' Query string used to update keys ins usysvar
Private Const QryUpdateUSysVar As String = "PARAMETERS [KeyType] Long, [ctKeyValue] Value; " & _
                                           "UPDATE usysvar SET usysvar.[sValue] = [ctKeyValue] " & _
                                           "WHERE usysvar.[ID]=[KeyType]"
' Instance varables
Private Type AdmType
    ADODBConnection As ADODB.Connection ' Opened by Class_Initialize and closed by Class_Terminate

    LoggerFunctionName As String        ' Handle error logging
    Cryptography As ClassCrypt          ' Uses Decrypt, Encrypt, Hash, Random, RandomInt

    ODBCLogInName As String             ' Use name logged into the ODBC connection

    ServerName As String
    DatabaseName As String
    DriverODBC As String
    DriverOLEDB As String
    TrustedConnection As String
    EncryptConnection As String

    DatabaseStatus As AdmDatabaseStatusEnum
    NamedUserAccount As Boolean        ' True if User account was a user AD login

    SQLUserType As AdmUserTypeEnum

    EncryptedUserAdminID As String
    EncryptedUserAdminLogin As String
    EncryptedGuestID As String
    EncryptedGuestLogin As String

    EncryptedUserAdminKey As String
    EncryptedApplicationKey As String

    ClassError As Long                 ' Save last error number
    ClassErrorText As String           ' Save last error description
    ClassErrorSource As String         ' Save last error source
End Type
Private This As AdmType

'<><><><><><><><> Events <><><><><><><><><><>

'-----------------
' The login ID of current user, default for ODBC login
Public Event LoginUserID(ByRef ClearTextUserID As String)
'-----------------
' Next userID to update, vbNullString when done
Public Event NextUserID(ByRef ClearTextUserID As String)
'----------
' Check user ID is configured
' Normaly check Active Directory users
Public Event VerifyUserID(ByVal ClearTextUserID As String, _
                          ByRef Resonse As Boolean)
'---------
' Check if User is in application group
Public Event VerifyUserInGroup(ByVal ClearTextUserID As String, _
                               ByVal AccessType As AdmUserTypeEnum, _
                               ByRef Resonse As Boolean)
' Provide SQL server connection string information
Public Event SqlServerName(ByRef Response As String)
Public Event SqlDatabaseName(ByRef Response As String)
Public Event SqlDriveODBC(ByRef Response As String)
Public Event SqlDriveOLEDB(ByRef Response As String)
Public Event SqlTrustedConnection(ByRef Response As String)
Public Event SqlEncrypt(ByRef Response As String)

'<><><><><><><><><> End Events <><><><><><><><>

'-------------
' Log the module, function and optional message with error information
'
Private Sub LogMessage(ByVal SourceFunction As String, _
                       Optional ByVal ErrorMessage As String = vbNullString)
    Dim Result As Variant
    With This
        .ClassErrorSource = ModuleName & "." & SourceFunction
        .ClassError = Err.Number
        .ClassErrorText = Err.Description
        On Error Resume Next
        If .LoggerFunctionName = vbNullString Then
            Debug.Print .ClassErrorSource, _
                         ErrorMessage, _
                        "error:( " & .ClassError & ")", _
                        .ClassErrorText
        Else
            ' This will cause an un-trapped error if LoggerFunctionName cannot be found!
            Result = Application.Run(.LoggerFunctionName, _
                                     .ClassErrorSource, _
                                      ErrorMessage, _
                                     .ClassError, _
                                    .ClassErrorText)
        End If
    End With
End Sub

Private Sub ClearLastError()
    On Error Resume Next
    This.ClassError = AdmErrorEnum.Noerror
    This.ClassErrorText = vbNullString
End Sub

Private Property Let ODBCLogInName(ByVal Value As String)
    On Error Resume Next
    This.ODBCLogInName = Value
End Property
Private Property Get ctLoginUserID() As String
    Dim Response As String
    RaiseEvent LoginUserID(Response)
    ctLoginUserID = Response
End Property
Private Property Get ODBCDriver() As String
    If This.DriverODBC = vbNullString Then
        RaiseEvent SqlDriveODBC(This.DriverODBC)
    End If
    ODBCDriver = This.DriverODBC
End Property
Private Property Get OLEDBDriver() As String
    If This.DriverOLEDB = vbNullString Then
        RaiseEvent SqlDriveOLEDB(This.DriverOLEDB)
    End If
    OLEDBDriver = This.DriverOLEDB
End Property
Private Property Get TrustedConnection() As String
    If This.TrustedConnection = vbNullString Then
        RaiseEvent SqlTrustedConnection(This.TrustedConnection)
    End If
    TrustedConnection = This.TrustedConnection
End Property
Private Property Get EncryptSQL() As String
    If This.EncryptConnection = vbNullString Then
        RaiseEvent SqlEncrypt(This.EncryptConnection)
    End If
    EncryptSQL = This.EncryptConnection
End Property

Private Sub Class_Initialize() '/////////////////////////////////////////
    On Error GoTo errClass_Initialize
    This.DatabaseStatus = AdmDatabaseStatusEnum.UnknownDbStatus
    NamedUserAccount = False
doneClass_Initialize:
    Exit Sub
errClass_Initialize:
    LogMessage "Class_Initialize"
    Resume doneClass_Initialize
End Sub
Private Sub class_Terminate()
    On Error GoTo errClass_Terminate
    ReleaseADODB
    Set This.Cryptography = Nothing
doneClass_Terminate:
    Exit Sub
errClass_Terminate:
    LogMessage "Class_Terminate"
    Resume doneClass_Terminate
End Sub '//////////////////////////////////////////////////////////////

'----------------
' Return the open ADODB connection
' If if not already open, then open it using the user admin account
Private Property Get ADODBConnection() As ADODB.Connection
    On Error GoTo errADODBConnection
    If This.ADODBConnection Is Nothing Then
        Set This.ADODBConnection = New ADODB.Connection
    End If
    If This.ADODBConnection.State <> ObjectStateEnum.adStateOpen Then
        This.ADODBConnection.connectionString = ADODBConnectString
        This.ADODBConnection.Open
    End If
doneADODBConnection:
    On Error Resume Next
    Set ADODBConnection = This.ADODBConnection
    Exit Property
errADODBConnection:
    LogMessage "ADODBConnection"
    Resume doneADODBConnection
End Property

'----------------
' Return the full ADODB connection string
Private Function ADODBConnectString() As String
    On Error GoTo errADODBConnectString
    ' Production pulls info from config file
    ADODBConnectString = OLEDBDriver & _
                         ";SERVER=" & ServerName & _
                         ";DATABASE=" & DatabaseName & _
                         ";Trusted_Connection=" & TrustedConnection & _
                         ";Encrypt=" & EncryptSQL & _
                         ";UID=" & ctUserAdminID & _
                         ";PWD=" & ctUserAdminLogin
doneADODBConnectString:
    Exit Function
errADODBConnectString:
    LogMessage "ADODBConnectString"
    ADODBConnectString = vbNullString
    Resume doneADODBConnectString
End Function

'----------------
' Login ID of read only guest user
Private Property Get ctGuestLoginID(Optional ByVal ReturnDecripted As Boolean = True) As String
    On Error GoTo errctGuestLoginID
    If This.EncryptedGuestID = vbNullString Then
        This.EncryptedGuestID = DLookup("[sValue]", "usysvar", "[ID]=" & AdmKeyEnum.admGuestID)
    End If
    If ReturnDecripted Then
        ctGuestLoginID = This.Cryptography.Decrypt(This.EncryptedGuestID, ctUserAdminKey)
    Else
        ctGuestLoginID = This.EncryptedGuestID
    End If
donectGuestLoginID:
    Exit Property
errctGuestLoginID:
    LogMessage "ctGuestLoginID"
    ctGuestLoginID = vbNullString
    Resume donectGuestLoginID
End Property
'--------
' Login password of read only guest user
Private Property Get ctGuestLogin(Optional ByVal ReturnDecripted As Boolean = True) As String
    On Error GoTo errctGuestLogin
    If This.EncryptedGuestLogin = vbNullString Then
        This.EncryptedGuestLogin = DLookup("[sValue]", "usysvar", "[ID]=" & AdmKeyEnum.admGuestLogin)
    End If
    If ReturnDecripted Then
        ctGuestLogin = This.Cryptography.Decrypt(This.EncryptedGuestLogin, ctUserAdminKey)
    Else
        ctGuestLogin = This.EncryptedGuestLogin
    End If
donectGuestLogin:
    Exit Property
errctGuestLogin:
    LogMessage "ctGuestLogin"
    ctGuestLogin = vbNullString
    Resume donectGuestLogin
End Property
'--------------
' Login in ID of user administrator account
Private Property Get ctUserAdminID(Optional ByVal ReturnDecripted As Boolean = True) As String
    On Error GoTo errctUserAdminID
    If This.EncryptedUserAdminID = vbNullString Then
        This.EncryptedUserAdminID = DLookup("[sValue]", "usysvar", "[ID]=" & AdmKeyEnum.admUserAdminID)
    End If
    If ReturnDecripted Then
        ctUserAdminID = This.Cryptography.Decrypt(This.EncryptedUserAdminID, ctUserAdminKey)
    Else
        ctUserAdminID = This.EncryptedUserAdminID
    End If
donectUserAdminID:
    Exit Property
errctUserAdminID:
    LogMessage "ctUserAdminID"
    ctUserAdminID = vbNullString
    Resume donectUserAdminID
End Property
'--------
' Login password of user administrator account
Private Property Get ctUserAdminLogin(Optional ByVal ReturnDecripted As Boolean = True) As String
    On Error GoTo errctUserAdminLogin
    If This.EncryptedUserAdminLogin = vbNullString Then
        This.EncryptedUserAdminLogin = DLookup("[sValue]", "usysvar", "[ID]=" & AdmKeyEnum.admUserAdminPW)
    End If
    If ReturnDecripted Then
        ctUserAdminLogin = This.Cryptography.Decrypt(This.EncryptedUserAdminLogin, ctUserAdminKey)
    Else
        ctUserAdminLogin = This.EncryptedUserAdminLogin
    End If
donectUserAdminLogin:
    Exit Property
errctUserAdminLogin:
    LogMessage "ctUserAdminLogin"
    ctUserAdminLogin = vbNullString
    Resume donectUserAdminLogin
End Property

'------------------------------------------
' Return decrypted user admin key
' issueNewKey If true generate new key and return that
'             If current key is null, fill current key
'             else fill new key
Private Property Get ctUserAdminKey(Optional ByVal IssueNewKey As Boolean = False) As String
    On Error GoTo errctUserAdminKey
    Dim CurrentKeyIsEmpty As Boolean

    If IssueNewKey Then
        This.EncryptedUserAdminKey = vbNullString
    Else
        This.EncryptedUserAdminKey = DLookup("[sValue]", "usysvar", "[ID]=" & AdmKeyEnum.admUserAdminKey)
    End If

    If This.EncryptedUserAdminKey = vbNullString Then
        This.EncryptedUserAdminKey = This.Cryptography.Encrypt(This.Cryptography.Random(12), ctApplicationKey)
        With CurrentDb.CreateQueryDef(vbNullString, QryUpdateUSysVar)
            .Parameters(0) = AdmKeyEnum.admUserAdminKey
            .Parameters(1) = This.EncryptedUserAdminKey
            .Execute dbFailOnError + dbSeeChanges
        End With
    End If

    ctUserAdminKey = This.Cryptography.Decrypt(This.EncryptedUserAdminKey, ctApplicationKey)

donectUserAdminKey:
    Exit Property
errctUserAdminKey:
    LogMessage "ctUserAdminKey"
    ctUserAdminKey = vbNullString
    Resume donectUserAdminKey
End Property

'------------------
' Return clear text application key
' The application key, book cipher stored withing this class
Private Property Get ctApplicationKey(Optional IssueNewKey As Boolean = False) As String
    On Error GoTo errctApplicationKey
    Dim ctDecriptedKey As String
    Dim KeyList() As String
    Dim BookLength As String
    Dim i As Long

    If IssueNewKey Then
        This.EncryptedApplicationKey = vbNullString
    Else
        This.EncryptedApplicationKey = Nz(DLookup("[sValue]", "usysvar", "[ID]=" & AdmKeyEnum.admAppKeyList))
    End If

    If This.EncryptedApplicationKey = vbNullString Then
        ' Create a new key using book cipher
        BookLength = Len(AppKeyBookCipher) - 1
        For i = 0 To 25
            This.EncryptedApplicationKey = This.EncryptedApplicationKey & _
                                           This.Cryptography.RandomInt(1, BookLength) & SINGLESPACE
        Next
        With CurrentDb.CreateQueryDef(vbNullString, QryUpdateUSysVar)
            .Parameters(0) = AdmKeyEnum.admAppKeyList
            .Parameters(1) = This.EncryptedApplicationKey
            .Execute dbFailOnError + dbSeeChanges
        End With
    End If

    ' Decript application key
    KeyList = Split(This.EncryptedApplicationKey, SINGLESPACE)
    ctDecriptedKey = vbNullString
    For i = 0 To UBound(KeyList) - 1
        ctDecriptedKey = ctDecriptedKey & Mid(AppKeyBookCipher, CInt(KeyList(i)), 1)
    Next

    ctApplicationKey = ctDecriptedKey

donectApplicationKey:
    Exit Property
errctApplicationKey:
    LogMessage "ctApplicationKey"
    Resume donectApplicationKey
End Property

'------------------------------------------------------------------------------
' Use ADODB and stored procedure to add user ID to SQL server application table
' Note: The user account is NOT added to the database
Private Sub AddEncryptedUserInfo(ByVal ecWinID As String, _
                                 ByVal ecPassword As String)
    On Error GoTo errAddEncryptedUserInfo

    Dim cmdSql As ADODB.Command
    Dim ParamIn As ADODB.Parameter

    ClearLastError

    Set cmdSql = New ADODB.Command
    cmdSql.ActiveConnection = ADODBConnection
    cmdSql.CommandType = adCmdStoredProc
    cmdSql.CommandText = DatabaseName & ".[jcsm].[ApplicationAddUserID]"

    Set ParamIn = cmdSql.CreateParameter("@EncodedWinID", admVarChar, admParamInput, 80)
    ParamIn.Value = ecWinID
    cmdSql.Parameters.Append ParamIn
    Set ParamIn = Nothing

    Set ParamIn = cmdSql.CreateParameter("@EncodedPassword", admVarChar, admParamInput, 252)
    ParamIn.Value = ecPassword
    cmdSql.Parameters.Append ParamIn
    Set ParamIn = Nothing

    cmdSql.Execute

doneAddEncryptedUserInfo:
    On Error Resume Next
    Set ParamIn = Nothing
    Exit Sub
errAddEncryptedUserInfo:
    LogMessage "AddEncryptedUserInfo"
    Resume doneAddEncryptedUserInfo
End Sub

'------------------------------------------------------------------------------
' Use ADODB and stored procedure to add user account to database
Private Sub AddUserToDB(ByVal ctWinIdUpperCase As String, _
                        ByVal ctPassword As String)
    On Error GoTo errAddUserToDB

    Dim cmdSql As ADODB.Command
    Dim ParamIn As ADODB.Parameter

    ClearLastError

    Set cmdSql = New ADODB.Command
    cmdSql.ActiveConnection = ADODBConnection
    cmdSql.CommandType = adCmdStoredProc
    cmdSql.CommandText = DatabaseName & ".[jcsm].[ApplicationUserAdd]"

    Set ParamIn = cmdSql.CreateParameter("@winID", admVarChar, admParamInput, 39)
    ParamIn.Value = ctWinIdUpperCase
    cmdSql.Parameters.Append ParamIn
    Set ParamIn = Nothing

    Set ParamIn = cmdSql.CreateParameter("@password", admVarChar, admParamInput, 61)
    ParamIn.Value = ctPassword
    cmdSql.Parameters.Append ParamIn
    Set ParamIn = Nothing

    cmdSql.Execute

doneAddUserToDB:
    On Error Resume Next
    Set ParamIn = Nothing
    Set cmdSql = Nothing
    Exit Sub
errAddUserToDB:
    LogMessage "AddUserToDB"
    Resume doneAddUserToDB
End Sub

'------------------------------------------------------------------------------
' Use ADODB and stored procedure to remove user ID to SQL server application table
' Note: The user account is NOT removed from the database
Private Sub RemoveEncryptedUserInfo(ByVal ecWinID As String)
    On Error GoTo errRemoveEncryptedUserInfo

    Dim cmdSql As ADODB.Command
    Dim ParamIn As ADODB.Parameter

    ClearLastError

    Set cmdSql = New ADODB.Command
    cmdSql.ActiveConnection = ADODBConnection
    cmdSql.CommandType = adCmdStoredProc
    cmdSql.CommandText = DatabaseName & ".[jcsm].[ApplicationRemoveUserID]"

    Set ParamIn = cmdSql.CreateParameter("@EncodedWinID", admVarChar, admParamInput, 80)
    ParamIn.Value = ecWinID
    cmdSql.Parameters.Append ParamIn
    Set ParamIn = Nothing

    cmdSql.Execute

doneRemoveEncryptedUserInfo:
    On Error Resume Next
    Set ParamIn = Nothing
    Set cmdSql = Nothing
    Exit Sub
errRemoveEncryptedUserInfo:
    LogMessage "RemoveEncryptedUserInfo"
    Resume doneRemoveEncryptedUserInfo
End Sub

'-------------------------------------------------------
' Key used to to Encrypt and Decrypt the application mananger login.
'
' The Encrypted password in stored in USysVar. It is used to decript
' the user admin and read only user login, also stored in USysVar.
' This Book is the only part of the system that can be seen by using a
' text editor to read your application code, if released compiled.
' To make it harder to find in a text editor, change the property name to something random.
'
' Max length 85
' Generate new book for your application in the immediate window before 1st run
' Ex.  ?GenerateBookCipher
Private Property Get AppKeyBookCipher() As String
    AppKeyBookCipher = "2yskzj5vE7CKnw5zycBi/GsuVHQ+EEkjfmLvIXJIe/Ghm21MCKsQ0lbn3rr9ZDKzwudOUMln/sBPFf1ZDcvMo"
End Property

'------------------------------------------------------------------------------
' Use ADODB and stored procedure to remove user account from database
Private Sub RemoveUserFromDB(ByVal ctWinIdUpperCase As String)
    On Error GoTo errRemoveUserFromDB

    Dim cmdSql As ADODB.Command
    Dim ParamIn As ADODB.Parameter

    ClearLastError

    Set cmdSql = New ADODB.Command
    cmdSql.ActiveConnection = ADODBConnection
    cmdSql.CommandType = adCmdStoredProc
    cmdSql.CommandText = DatabaseName & ".[jcsm].[ApplicationUserRemove]"

    Set ParamIn = cmdSql.CreateParameter("@winID", admVarChar, admParamInput, 39)
    ParamIn.Value = ctWinIdUpperCase
    cmdSql.Parameters.Append ParamIn
    Set ParamIn = Nothing

    cmdSql.Execute

doneRemoveUserFromDB:
    On Error Resume Next
    Set ParamIn = Nothing
    Set cmdSql = Nothing
    Exit Sub
errRemoveUserFromDB:
    LogMessage "RemoveUserFromDB"
    Resume doneRemoveUserFromDB
End Sub

'-------------
' Use stored procedure to get encrypted password for given encrypted user ID
Private Property Get GetEncryptedPassword(ByVal ecWinID As String) As String
    On Error GoTo errGetEncryptedPassword

    Dim cmdSql As ADODB.Command
    Dim ParamIn As ADODB.Parameter
    Dim ParamOut As ADODB.Parameter

    ClearLastError

    Set cmdSql = New ADODB.Command
    cmdSql.ActiveConnection = ADODBConnection
    cmdSql.CommandType = adCmdStoredProc
    cmdSql.CommandText = DatabaseName & ".[jcsm].[ApplicationGetPassword]"

    Set ParamIn = cmdSql.CreateParameter("@EncodedWinID", admVarChar, admParamInput, 80)
    ParamIn.Value = ecWinID
    cmdSql.Parameters.Append ParamIn
    Set ParamIn = Nothing

    ' Setup Return Parameters
    Set ParamOut = cmdSql.CreateParameter("@EncodedPassword", admVarChar, admParamOutput, 252)
    cmdSql.Parameters.Append ParamOut
    Set ParamOut = Nothing

    cmdSql.Execute

    GetEncryptedPassword = CStr(Nz(cmdSql.Parameters("@EncodedPassword").Value, vbNullString))

doneGetEncryptedPassword:
    On Error Resume Next
    Set ParamIn = Nothing
    Set ParamOut = Nothing
    Set cmdSql = Nothing
    Exit Property
errGetEncryptedPassword:
    GetEncryptedPassword = vbNullString
    LogMessage "GetEncryptedPassword"
    Resume doneGetEncryptedPassword
End Property

'------------------------------------------------------------------------------
' Use ADODB and stored procedure to change userId/Password Encryption
' New application key must exist in key table if using UsingThis.Cryptography
' Returns true if completed
Private Function ChangeCrypt(ByVal ctWinID As String, _
                             ByVal ctKeyValueOld As String, _
                             ByVal ctKeyValueNew As String) As Boolean
    On Error GoTo errChangeCrypt

    Dim cmdSql As ADODB.Command
    Dim ParamIn As ADODB.Parameter
    Dim ecPassword As String

    ctWinID = UCase(ctWinID) ' Store user ID in upper case

    ' Check for user account by looking for it's stored password
    ecPassword = GetEncryptedPassword(This.Cryptography.Hash(ctWinID, HashAlgorithem))

    If ecPassword <> vbNullString Then
        Set cmdSql = New ADODB.Command
        cmdSql.ActiveConnection = ADODBConnection
        cmdSql.CommandType = adCmdStoredProc
        cmdSql.CommandText = "[" & DatabaseName & "].[jcsm].[ApplicationCryptChange]"

        ' Current WinID hash
        Set ParamIn = cmdSql.CreateParameter("@CurrentWinIDEncoding", admVarChar, admParamInput, 80)
        ParamIn.Value = This.Cryptography.Hash(ctWinID, HashAlgorithem)
        cmdSql.Parameters.Append ParamIn
        Set ParamIn = Nothing

        ' New WinID hash
        Set ParamIn = cmdSql.CreateParameter("@EncodedWinID", admVarChar, admParamInput, 80)
        ParamIn.Value = This.Cryptography.Hash(ctWinID, HashAlgorithem)
        cmdSql.Parameters.Append ParamIn
        Set ParamIn = Nothing

        ' New PW encryption
        Set ParamIn = cmdSql.CreateParameter("@EncodedPassword", admVarChar, admParamInput, 255)
        ParamIn.Value = This.Cryptography.Encrypt( _
                        This.Cryptography.Decrypt(ecPassword, ctKeyValueOld), ctKeyValueNew)
        cmdSql.Parameters.Append ParamIn
        Set ParamIn = Nothing

        cmdSql.Execute
        ChangeCrypt = True
    Else
        LogMessage "ChangeCrypt", "User ID " & ctWinID & " not in SQL database"
        ChangeCrypt = False
    End If

doneChangeCrypt:
    On Error Resume Next
    Set ParamIn = Nothing
    Set cmdSql = Nothing
    Exit Function
errChangeCrypt:
    LogMessage "ChangeCrypt"
    ChangeCrypt = False
    Resume doneChangeCrypt
End Function

'------------
' Check user's access level
'  1) They have an AD account
'  2) Member of one of the application AD groups
' Return AdmUserTypeEnum.SQLNone if no access
'
Private Property Get UserAccessLevel(ByVal ctWinID As String) As AdmUserTypeEnum
    On Error GoTo errUserAccessLevel
    Dim ConfiguredAccessLevel As AdmUserTypeEnum
    Dim Response As Boolean

    ConfiguredAccessLevel = AdmUserTypeEnum.SQLNone

    ' Check if user has an Active Directory account.
    RaiseEvent VerifyUserID(ctWinID, Response)
    If Response Then
        ' Check if user is member of the protected application group
        RaiseEvent VerifyUserInGroup(ctWinID, AdmUserTypeEnum.SQLProtected, Response)
        If Response Then
            ConfiguredAccessLevel = AdmUserTypeEnum.SQLProtected
        Else
            ' Check if user is member of the normal application group
            RaiseEvent VerifyUserInGroup(ctWinID, AdmUserTypeEnum.SQLNormal, Response)
            If Response Then
                ConfiguredAccessLevel = AdmUserTypeEnum.SQLNormal
            Else
                ' Check if user is member of the guest application group
                RaiseEvent VerifyUserInGroup(ctWinID, AdmUserTypeEnum.SQLGuest, Response)
                If Response Then
                    ConfiguredAccessLevel = AdmUserTypeEnum.SQLGuest
                Else
                    ' User ID is not in Active Directory application group
                    ConfiguredAccessLevel = AdmUserTypeEnum.SQLNone
                End If
            End If
        End If
    Else
        ' User ID is Not in Active Directory
        ConfiguredAccessLevel = AdmUserTypeEnum.SQLNone
    End If

doneUserAccessLevel:
    UserAccessLevel = ConfiguredAccessLevel
    Exit Property
errUserAccessLevel:
    LogMessage "UserAccessLevel"
    ConfiguredAccessLevel = AdmUserTypeEnum.SQLNone
    Resume doneUserAccessLevel
End Property

'-------------
' Use stored procedure to check iValue for 'TestSQL' key
Private Sub CheckDataBaseStatus()
    On Error GoTo errCheckDataBaseStatus

    Dim cmdSql As ADODB.Command
    Dim ParamIn As ADODB.Parameter
    Dim ParamOut As ADODB.Parameter

    ClearLastError

    Set cmdSql = New ADODB.Command
    cmdSql.ActiveConnection = ADODBConnection
    cmdSql.CommandType = adCmdStoredProc
    cmdSql.CommandText = "[" & DatabaseName & "].[jcsm].[ApplicationUsingTestDB]"

    ' Setup Return Parameters
    Set ParamOut = cmdSql.CreateParameter("@Result", _
                                            admInteger, _
                                            admParamOutput)
    cmdSql.Parameters.Append ParamOut
    Set ParamOut = Nothing

    cmdSql.Execute

    If CBool(Nz(cmdSql.Parameters("@Result").Value, True)) Then
        This.DatabaseStatus = AdmDatabaseStatusEnum.TestDbStatus
    Else
        This.DatabaseStatus = AdmDatabaseStatusEnum.ProductionDbStatus
    End If

doneCheckDataBaseStatus:
    On Error Resume Next
    Set ParamIn = Nothing
    Set ParamOut = Nothing
    Set cmdSql = Nothing
    Exit Sub
errCheckDataBaseStatus:
    This.DatabaseStatus = AdmDatabaseStatusEnum.UnknownDbStatus
    LogMessage "CheckDataBaseStatus"
    Resume doneCheckDataBaseStatus
End Sub

'----------------
' Provide SQL connection string's user name and password.
'   usingNamedAccount: If True try using current user name.
'                      If False use read-only name and password.
' ctWinID           User name to get login information
' usingNamedAccount Use named user ID otherwise read only account
' Return
'   Example return:     ";UID=ClearTextUserID;PWD=ClearTextPassword"
' On error return empty string
Private Property Get ODBC_UID_And_PWD(ByVal ctWinID As String, _
                                      ByVal UsingNamedAccount As Boolean) As String
    On Error GoTo errODBC_UID_And_PWD
    Dim HashWinID As String
    Dim ecPassword As String
    Dim ctWinIdUpperCase As String

    If UsingNamedAccount Then
        ' User ID setup upper case in database
        ctWinIdUpperCase = UCase(ctWinID)
        If ctWinIdUpperCase <> vbNullString Then
            HashWinID = This.Cryptography.Hash(ctWinIdUpperCase, HashAlgorithem)
            ecPassword = GetEncryptedPassword(HashWinID)
            If ecPassword = vbNullString Then
                ' Revert to fallback account
                 ' Log missing SQL database account
                 Err.Raise AdmErrorEnum.SqlApplicationLoginNotFound, _
                           "ODBC_UID_And_PWD_ForCurrentUser", _
                           "User AD Name: " & ctWinIdUpperCase & _
                           " needs to be added, using 'RDP' read-only account"
            End If
            NamedUserAccount = True
        Else
            ' No user name found in AD, use fallback acount
            Err.Raise AdmErrorEnum.ADAccountNotFound, _
                      "ODBC_UID_And_PWD_ForCurrentUser", _
                      "User not in AD group, using 'RDP' read-only account"
        End If
    Else
        ' Use fallback read only account
        UsingNamedAccount = False
        ctWinIdUpperCase = UCase(ctGuestLoginID)
        HashWinID = This.Cryptography.Hash(ctWinIdUpperCase, HashAlgorithem)
        ecPassword = ctGuestLogin(ReturnDecripted:=False)
        NamedUserAccount = False
    End If

doneODBC_UID_And_PWD:
    On Error Resume Next
    This.ODBCLogInName = ctWinIdUpperCase
    ODBC_UID_And_PWD = ";UID=" & ctWinIdUpperCase & _
                       ";PWD=" & This.Cryptography.Decrypt(ecPassword, ctUserAdminKey)
    Exit Property
errODBC_UID_And_PWD:
    LogMessage "ODBC_UID_And_PWD_ForCurrentUser"
    UsingNamedAccount = False ' Cause a single level decrypt on return
    ctWinIdUpperCase = UCase(ctGuestLoginID)
    HashWinID = This.Cryptography.Hash(ctWinIdUpperCase, HashAlgorithem)
    ecPassword = ctGuestLogin(ReturnDecripted:=False)                                         '
    NamedUserAccount = False
    Resume doneODBC_UID_And_PWD
End Property

'------------------------------------------------------------------------------
'<><><><><><><><><><><><> Start of Public <><><><><><><><><><><><><>
'------------------------------------------------------------------------------

'--------------------------------------------------------
' Set the Crypt instace to use before using SQL functions
'--------------------------------------------------------
Public Property Set UseCryptography(ByRef UseInstance As ClassCrypt)
    On Error Resume Next
    If Not This.Cryptography Is Nothing Then Set This.Cryptography = Nothing
    Set This.Cryptography = UseInstance
End Property

'-----------------
' Pass a function name to handle the logging
' otherwise debug.print
Public Property Let LoggerFunctionName(ByVal LoggerFunction As String)
    On Error Resume Next
    ' LoggerFunctionName: expects the following paramiters
    'Public Sub Logger(ByVal SourceFunction As String, _
    '                  Optional ByVal logMessage As String = vbNullString, _
    '                  Optional ByVal errNumber As Long = 0, _
    '                  Optional ByVal errDescription As String = vbNullString)
    This.LoggerFunctionName = LoggerFunction
End Property
Public Property Get LastErrorNumber() As Long
    On Error Resume Next
    LastErrorNumber = This.ClassError
End Property
Public Property Get LastErrorDescription() As String
    On Error Resume Next
    LastErrorDescription = This.ClassErrorText
End Property

Public Property Get NamedUserAccount() As Boolean
    NamedUserAccount = This.NamedUserAccount
End Property
Private Property Let NamedUserAccount(ByVal newValue As Boolean)
    This.NamedUserAccount = newValue
End Property

Public Property Get SqlIsConfigured() As Boolean
    SqlIsConfigured = Not IsNull(DLookup("[sValue]", "usysvar", "[ID]=" & AdmKeyEnum.admUserAdminKey))
End Property

Public Property Get DatabaseStatus() As AdmDatabaseStatusEnum
    DatabaseStatus = This.DatabaseStatus
End Property

Public Property Get UsingTestDB() As Boolean
    On Error GoTo errUsingTestDB
    If This.DatabaseStatus = AdmDatabaseStatusEnum.UnknownDbStatus Then
        CheckDataBaseStatus
        ReleaseADODB
    End If
doneUsingTestDB:
    UsingTestDB = Not CBool((This.DatabaseStatus = AdmDatabaseStatusEnum.ProductionDbStatus))
    Exit Property
errUsingTestDB:
    LogMessage "UsingTestDB"
    Resume doneUsingTestDB
End Property

'-------------
' Add Admin user account info if missing
Public Sub Configure(ByVal ClearTextAdminID As String, _
                     ByVal ClearTextAdminPW As String, _
                     ByVal ClearTextGuestID As String, _
                     ByVal ClearTextGuestPW As String)
    On Error GoTo errConfigure
    Dim ctKeyValue As String

    If IsNull(DLookup("[sValue]", "usysvar", "[ID]=" & AdmKeyEnum.admUserAdminKey)) Then
        ' Note: Existing User accounts will need to be updated after setting the admin encription key
        If ClearTextAdminID = vbNullString Then
            Err.Raise AdmErrorEnum.MissingInfo, "Class_Initialize", "No admin ID provided"
        End If

        If ClearTextAdminPW = vbNullString Then
            Err.Raise AdmErrorEnum.MissingInfo, "Class_Initialize", "No admin password provided"
        End If

        If ClearTextGuestID = vbNullString Then
            Err.Raise AdmErrorEnum.MissingInfo, "Class_Initialize", "No read only ID provided"
        End If

        If ClearTextGuestPW = vbNullString Then
            Err.Raise AdmErrorEnum.MissingInfo, "Class_Initialize", "No read only password provided"
        End If

        ' Create new user administration key
        ctKeyValue = ctUserAdminKey(IssueNewKey:=True)

        With CurrentDb.CreateQueryDef(vbNullString, QryUpdateUSysVar)

            ' Save user admin account
            .Parameters(0) = AdmKeyEnum.admUserAdminID
            .Parameters(1) = This.Cryptography.Encrypt(ClearTextAdminID, ctKeyValue)
            .Execute dbFailOnError + dbSeeChanges

            ' Save user admin password
            .Parameters(0) = AdmKeyEnum.admUserAdminPW
            .Parameters(1) = This.Cryptography.Encrypt(ClearTextAdminPW, ctKeyValue)
            .Execute dbFailOnError + dbSeeChanges

            ' Save Read-only ID
            .Parameters(0) = AdmKeyEnum.admGuestID
            .Parameters(1) = This.Cryptography.Encrypt(ClearTextGuestID, ctKeyValue)
            .Execute dbFailOnError + dbSeeChanges

            ' Save Read-Only password
            .Parameters(0) = AdmKeyEnum.admGuestLogin
            .Parameters(1) = This.Cryptography.Encrypt(ClearTextGuestPW, ctKeyValue)
            .Execute dbFailOnError + dbSeeChanges
        End With
    End If

doneConfigure:
    Exit Sub
errConfigure:
    LogMessage "Configure"
    Resume doneConfigure
End Sub

'---------------
' Release user mangagement accoounts' ADODB connection to the SQL server
Public Sub ReleaseADODB()
    On Error GoTo errReleaseADODB
    If Not This.ADODBConnection Is Nothing Then
        If This.ADODBConnection.State = ObjectStateEnum.adStateOpen Then
            This.ADODBConnection.Close
        End If
    End If
doneReleaseADODB:
    On Error Resume Next
    Set This.ADODBConnection = Nothing
    Exit Sub
errReleaseADODB:
    LogMessage "ReleaseADODB"
    Resume doneReleaseADODB
End Sub

'----------------
' Generate a random password
' Test that it can be Encrypted and Decrypted
' https://learn.microsoft.com/en-us/sql/relational-databases/security/strong-passwords?view=sql-server-ver16
' https://stackoverflow.com/questions/2030610/password-limitations-in-sql-server-and-mysql
Public Function RandomPassword(Optional ByVal MinLength As Integer = 10, _
                               Optional ByVal MaxLength As Integer = 13) As String
    On Error GoTo errRandomPassword
    Dim ctPassword As String
    Dim i As Integer
    Dim TryCount As Integer
    Dim ctKeyValue As String

    ctKeyValue = ctUserAdminKey
    Do
        TryCount = TryCount + 1
        If TryCount > 20 Then
            Err.Raise AdmErrorEnum.ObjectError, "RandomPassword", _
                      "Failed to generate a password that can be Encripted/Decrypted"
        End If

        ctPassword = Mid(ALPHABETICCHARACTERS, This.Cryptography.RandomInt(1, Len(ALPHABETICCHARACTERS)), 1)
        For i = 2 To This.Cryptography.RandomInt(MinLength, MaxLength)
            ctPassword = ctPassword & _
                         Mid(PASSWORDCHARACTERS, This.Cryptography.RandomInt(1, Len(PASSWORDCHARACTERS)), 1)
        Next
        ' Try up to 20 times or until it can be Encrypted and then Decrypted
    Loop Until This.Cryptography.Decrypt( _
               This.Cryptography.Encrypt(ctPassword, ctKeyValue), ctKeyValue) = ctPassword
doneRandomPassword:
    RandomPassword = ctPassword
    Exit Function
errRandomPassword:
    LogMessage "RandomPassword"
    ctPassword = vbNullString
    Resume doneRandomPassword
End Function

'----------------
' Add encripted user/password to application user table
'     then the user account using the clearText ID and password
' clearTextWinID    User ID from windows
' clearTextPWadmin  Password for special (Admin) accounts, all other users accounts should pass empty string
' Return 'True' if both were done without error
Public Function AddUser(ByVal ClearTextWinID As String, _
                        Optional ByVal ClearTextPWadmin As String = vbNullString) As Boolean
    On Error GoTo errAddUser
    Dim UserWasAdded As Boolean
    Dim ctKeyValue As String

    ClearLastError
    UserWasAdded = False

    If InStr(1, ClearTextWinID, SINGLESPACE) > 0 Then
        Err.Raise AdmErrorEnum.OutOfRange, "AddUser", _
                  "User name (" & ClearTextPWadmin & ") cannot contain a space?"
    End If

    ClearTextWinID = UCase(ClearTextWinID) ' Store user ID in upper case
    ctKeyValue = ctUserAdminKey

    If UserAccessLevel(ClearTextWinID) <> AdmUserTypeEnum.SQLNone Then
        If GetEncryptedPassword(This.Cryptography.Hash(ClearTextWinID, HashAlgorithem)) = vbNullString Then
            ' New user account
            If ClearTextPWadmin = vbNullString Then
                ' Generate password for new user
                ClearTextPWadmin = RandomPassword()
            Else
                ' Verify supplied password can be encrypted and decrypted
                If ClearTextPWadmin <> This.Cryptography.Decrypt( _
                                       This.Cryptography.Encrypt(ClearTextPWadmin, ctKeyValue), ctKeyValue) Then
                    Err.Raise AdmErrorEnum.ObjectError, "AddUser", _
                              "User ID: " & ClearTextWinID & " supplied password could not be used"
                End If
            End If
            If ClearTextPWadmin <> vbNullString Then
                ' Add the encripted user and password to the application table on SQL server.
                AddEncryptedUserInfo This.Cryptography.Hash(ClearTextWinID, HashAlgorithem), _
                                     This.Cryptography.Encrypt(ClearTextPWadmin, ctKeyValue)
                If LastErrorNumber = AdmErrorEnum.Noerror Then
                    ' Add user to database
                    AddUserToDB ClearTextWinID, ClearTextPWadmin
                    If LastErrorNumber = AdmErrorEnum.Noerror Then
                        UserWasAdded = True
                    End If
                Else
                    RemoveEncryptedUserInfo This.Cryptography.Hash(ClearTextWinID, HashAlgorithem)
                    Err.Raise AdmErrorEnum.ObjectError, _
                              "User ID: " & ClearTextWinID & " failed save Encrypted name and password"
                End If
            Else
                Err.Raise AdmErrorEnum.ObjectError, "AddUser", _
                          "User ID: " & ClearTextWinID & " failed to generate password for new user"
            End If
        Else
            Err.Raise AdmErrorEnum.UserAlreadyAdded, "AddUser", _
                      "User ID: " & ClearTextWinID & " has already been added"
        End If
    Else
        Err.Raise AdmErrorEnum.SqlApplicationLoginNotFound, "AddUser", _
                  "User ID: " & ClearTextWinID & " must be in the 'JobCostSQL' or 'JobCost' AD group for database login."
    End If
    If UserWasAdded Then
        LogMessage "AddUser", _
                   "User: " & ClearTextWinID & _
                   " was added to SQL Server: " & ServerName & "." & DatabaseName
    End If
doneAddUser:
    AddUser = UserWasAdded
    Exit Function
errAddUser:
    LogMessage "AddUser"
    Resume doneAddUser
End Function

'----------------
' Remove encripted user/password to application user table
'        then remove user account from database
Public Function RemoveUser(ByVal ClearTextWinID As String) As Boolean
    On Error GoTo errRemoveUser
    Dim UserRemoved As Boolean

    ClearLastError
    UserRemoved = False
    ClearTextWinID = UCase(ClearTextWinID) ' Store user ID in upper case

    ' Remove the encripted user and password from the application table on SQL server.
    RemoveEncryptedUserInfo This.Cryptography.Hash(ClearTextWinID, HashAlgorithem)
    If LastErrorNumber = AdmErrorEnum.Noerror Then
        ' Remove user from the database
        RemoveUserFromDB ClearTextWinID
        If LastErrorNumber = AdmErrorEnum.Noerror Then
            UserRemoved = True
        End If
    End If
    If UserRemoved Then
        LogMessage "RemoveUser", _
                   "User: " & ClearTextWinID & _
                   " was removed from SQL Server: " & ServerName & "." & DatabaseName
    End If
doneRemoveUser:
    RemoveUser = UserRemoved
    Exit Function
errRemoveUser:
    LogMessage "RemoveUser"
    Resume doneRemoveUser
End Function

'---------------
' Add user getting UserID from NextUserID event
Public Sub AddUserFromNextUserID()
    On Error GoTo errAddUserFromNextUserID
    Dim ctWinID As String

    ClearLastError
    RaiseEvent NextUserID(ctWinID)
    Do Until ctWinID = vbNullString
        If AddUser(ctWinID) Then
            RaiseEvent NextUserID(ctWinID)
        Else
            Err.Raise LastErrorNumber, "AddUserFromNextUserID", "Failed to add: " & ctWinID
        End If
    Loop

doneAddUserFromNextUserID:
    Exit Sub
errAddUserFromNextUserID:
    LogMessage "AddUserFromNextUserID"
    Resume doneAddUserFromNextUserID
End Sub

'---------------
' Change User Admin Key then replace encription for
' each user returned by NextUserID event
' Caller supplies UserID though class event procedure
Public Sub UpdateDbEncription()
    On Error GoTo errUpdateDbEncription

    Dim ctWinID As String
    Dim ctGuestWinID As String
    Dim ctGuestPassword As String
    Dim ctApplicationKeyNew As String
    Dim ctAdminKeyOld As String
    Dim ctAdminKeyNew As String
    Dim ctAdminDatabaseID As String
    Dim ctAdminPassword As String

    ClearLastError

    If This.ADODBConnection Is Nothing Then
        Set This.ADODBConnection = ADODBConnection
    End If

    ctAdminKeyOld = ctUserAdminKey
    ctAdminDatabaseID = ctUserAdminID
    ctAdminPassword = ctUserAdminLogin
    ctGuestWinID = ctGuestLoginID
    ctGuestPassword = ctGuestLogin

    ' Generate new application key in key table
    ctApplicationKeyNew = ctApplicationKey(IssueNewKey:=True)
    If ctApplicationKeyNew = vbNullString Then
        Err.Raise LastErrorNumber, "UpdateUserInDb", "Failed to get new application key"
    End If

    ' Generate new admin key in key table
    ctAdminKeyNew = ctUserAdminKey(IssueNewKey:=True)
    If ctAdminKeyNew = vbNullString Then
        Err.Raise LastErrorNumber, "UpdateUserInDb", "Failed to get new admin key"
    End If

    With CurrentDb.CreateQueryDef(vbNullString, QryUpdateUSysVar)
        ' Save user admin account ID
        .Parameters(0) = AdmKeyEnum.admUserAdminID
        .Parameters(1) = This.Cryptography.Encrypt(ctAdminDatabaseID, ctAdminKeyNew)
        .Execute dbFailOnError + dbSeeChanges

        ' Save new Admin account password
        .Parameters(0) = AdmKeyEnum.admUserAdminPW
        .Parameters(1) = This.Cryptography.Encrypt(ctAdminPassword, ctAdminKeyNew)
        .Execute dbFailOnError + dbSeeChanges

        ' Save Read only account ID
        .Parameters(0) = AdmKeyEnum.admGuestID
        .Parameters(1) = This.Cryptography.Encrypt(ctGuestWinID, ctAdminKeyNew)
        .Execute dbFailOnError + dbSeeChanges

        ' Save Read only account password
        .Parameters(0) = AdmKeyEnum.admGuestLogin
        .Parameters(1) = This.Cryptography.Encrypt(ctGuestPassword, ctAdminKeyNew)
        .Execute dbFailOnError + dbSeeChanges
    End With

    ' Update user accounts using old and new user admin keys
    RaiseEvent NextUserID(ctWinID)
    Do Until ctWinID = vbNullString
Debug.Print ctWinID
        If Not ChangeCrypt(ctWinID, ctAdminKeyOld, ctAdminKeyNew) Then
            If LastErrorNumber <> AdmErrorEnum.Noerror Then
                Err.Raise AdmErrorEnum.ObjectError, "UpdateUserInDb", "Failed to update user: " & ctWinID
            End If
        End If
        RaiseEvent NextUserID(ctWinID)
    Loop

doneUpdateDbEncription:
    ReleaseADODB
    Exit Sub
errUpdateDbEncription:
    LogMessage "UpdateUserInDB"
    Resume doneUpdateDbEncription
End Sub

Public Property Get ServerName() As String
    If This.ServerName = vbNullString Then
        RaiseEvent SqlServerName(This.ServerName)
    End If
    ServerName = This.ServerName
End Property
Public Property Get DatabaseName() As String
    If This.DatabaseName = vbNullString Then
        RaiseEvent SqlDatabaseName(This.DatabaseName)
    End If
    DatabaseName = This.DatabaseName
End Property

'----------------
' Return the ODBC connection string, without username and password
Public Property Get ODBCConnectionString() As String
    On Error GoTo errODBCConnectionString
    ' Production pulls info from config file
    ODBCConnectionString = ODBCDriver & _
                           ";SERVER=" & ServerName & _
                           ";DATABASE=" & DatabaseName & _
                           ";Trusted_Connection=" & TrustedConnection & _
                           ";Encrypt=" & EncryptSQL
doneODBCConnectionString:
    Exit Property
errODBCConnectionString:
    LogMessage "ODBCConnectionString'"
    Resume doneODBCConnectionString:
End Property

Public Property Get ODBCLogInName() As String
    On Error Resume Next
    ODBCLogInName = This.ODBCLogInName
End Property

'---------------------
' Create a Cached logged in connection to SQL server
' https://www.microsoft.com/en-us/microsoft-365/blog/2011/04/08/power-tip-improve-the-security-of-database-connections/
' https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16
' ClearTextLoginName User ID to test ability to login otherwise currently login user ID
Public Function ODBCConnectToSql(Optional ClearTextLoginName As String = vbNullString) As AdmUserTypeEnum
    On Error GoTo errODBCConnectToSql
    Dim rst As DAO.Recordset
    Dim UseNamedAccount As Boolean
    Dim ctUserIdToConnect As String

    ClearLastError
    UseNamedAccount = True
    This.SQLUserType = AdmUserTypeEnum.SQLNone
    This.ODBCLogInName = vbNullString

    If ClearTextLoginName = vbNullString Then
        ctUserIdToConnect = ctLoginUserID
    Else
        ctUserIdToConnect = ClearTextLoginName
    End If

    With DBEngine(0)(0)
retryConnectUsingReadOnly:
        With .CreateQueryDef(vbNullString)
            .Connect = ODBCConnectionString & _
                       ODBC_UID_And_PWD(ctUserIdToConnect, UseNamedAccount)

            If ClearTextLoginName <> vbNullString And InStr(1, .Connect, ClearTextLoginName) = 0 Then
                This.SQLUserType = AdmUserTypeEnum.SQLNone
            ElseIf InStr(1, .Connect, "PWD=") = 0 Then
                This.SQLUserType = AdmUserTypeEnum.SQLNone
            ElseIf Trim$(Mid(.Connect, InStr(1, .Connect, "PWD="))) = "PWD=" Then
                This.SQLUserType = AdmUserTypeEnum.SQLNone
            Else
                .SQL = "SELECT @@OPTIONS;" ' Simple query to cause login
                Set rst = .OpenRecordset(dbOpenSnapshot, dbSQLPassThrough)
                This.SQLUserType = AdmUserTypeEnum.SQLNormal
            End If
        End With
    End With

doneODBCConnectToSql:
    ODBCConnectToSql = This.SQLUserType
    On Error Resume Next
    If Not rst Is Nothing Then
        rst.Close
        Set rst = Nothing
    End If
    Exit Function

errODBCConnectToSql:
    LogMessage "ConnectToSqlThoughODBC", _
               IIf(UseNamedAccount, "Retry using using read only account. ", vbNullString) & _
               "LastError (" & LastErrorNumber & ") " & LastErrorDescription & _
               ", " & ODBCConnectionString
    If UseNamedAccount And ClearTextLoginName = vbNullString Then
        ' Retry using read-only acccount
        UseNamedAccount = False
        Resume retryConnectUsingReadOnly
    End If
    ' Failed to login to SQL server
    This.SQLUserType = AdmUserTypeEnum.SQLNone
    This.ODBCLogInName = vbNullString
    Resume doneODBCConnectToSql
End Function
