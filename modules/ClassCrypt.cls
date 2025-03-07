VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassCrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Binary  '*** Binary required for this code to work ***'
Option Explicit

'------------------------------- ClassBCrypt -------------------------------
' Manange User rights on the SQL Server
'
' Requires a reference to 'Microsoft ActiveX Data Objects 6.1 Library

Private Const ModuleName As String = "ClassBCrypt"

' Changed to Class RAM 3-3-2025

'-----------------------------------------------
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Cryptography     with some modifications
' License for this module is MIT, see
' https://github.com/GustavBrock/VBA.Cryptography/blob/main/LICENSE
'-----------------------------------------------
'https://learn.microsoft.com/en-us/windows/win32/seccng/cng-portal
'https://learn.microsoft.com/en-us/windows/win32/seccng/cng-reference
'https://www.experts-exchange.com/articles/37113/Encryption-in-VBA-using-the-Microsoft-NG-Cryptography-CNG-API.html

'The single algorithm offered here is AES (Advanced Encryption Standard)

'Encrypting Text
'A set of functions is used to encrypt the input and return the encrypted data:
'
'Encrypt
'    EncryptData
'        HashData
'            CngHash
'        RandomData
'            CngRandom
'        CngEncrypt
'    ByteBase64

' https://github.com/GustavBrock/VBA.Cryptography
'
' BCrypt V1.1.2
' Hashing, encrypting, and decrypting of text using the BCrypt API.
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Cryptography
'
'
' Credit:
'
' Based on code at Stack Overflow published 2021-04-28 (with later edits):
'   https://stackoverflow.com/questions/67294035/basic-encrypting-of-a-text-file/67294779#comment122708972_67294779
' by Stack Overflow user Erik A:
'   https://stackoverflow.com/users/7296893/erik-a
'
' CNG API documentation:
'   Cryptography API: Next Generation
'   https://docs.microsoft.com/en-us/windows/win32/seccng/cng-portal

'<><><><><><><<><><><><><><>
' API declarations.
'
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "BCrypt.dll" ( _
    ByRef phAlgorithm As LongPtr, _
    ByVal pszAlgId As LongPtr, _
    ByVal pszImplementation As LongPtr, _
    ByVal dwFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "BCrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByVal dwFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptGetProperty Lib "BCrypt.dll" ( _
    ByVal hObject As LongPtr, _
    ByVal pszProperty As LongPtr, _
    ByRef pbOutput As Any, _
    ByVal cbOutput As Long, _
    ByRef pcbResult As Long, _
    ByVal dfFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptSetProperty Lib "BCrypt.dll" ( _
    ByVal hObject As LongPtr, _
    ByVal pszProperty As LongPtr, _
    ByRef pbInput As Any, _
    ByVal cbInput As Long, _
    ByVal dfFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptCreateHash Lib "BCrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByRef phHash As LongPtr, _
    ByRef pbHashObject As Any, _
    ByVal cbHashObject As Long, _
    ByVal pbSecret As LongPtr, _
    ByVal cbSecret As Long, _
    ByVal dwFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptHashData Lib "BCrypt.dll" ( _
    ByVal hHash As LongPtr, _
    ByRef pbInput As Any, _
    ByVal cbInput As Long, _
    Optional ByVal dwFlags As Long = 0) _
    As Long

Private Declare PtrSafe Function BCryptFinishHash Lib "BCrypt.dll" ( _
    ByVal hHash As LongPtr, _
    ByRef pbOutput As Any, _
    ByVal cbOutput As Long, _
    ByVal dwFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptDestroyHash Lib "BCrypt.dll" ( _
    ByVal hHash As LongPtr) _
    As Long

Private Declare PtrSafe Function BCryptGenRandom Lib "BCrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByRef pbBuffer As Any, _
    ByVal cbBuffer As Long, _
    ByVal dwFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptGenerateSymmetricKey Lib "BCrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByRef hKey As LongPtr, _
    ByRef pbKeyObject As Any, _
    ByVal cbKeyObject As Long, _
    ByRef pbSecret As Any, _
    ByVal cbSecret As Long, _
    ByVal dwFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptEncrypt Lib "BCrypt.dll" ( _
    ByVal hKey As LongPtr, _
    ByRef pbInput As Any, _
    ByVal cbInput As Long, _
    ByRef pPaddingInfo As Any, _
    ByRef pbIV As Any, _
    ByVal cbIV As Long, _
    ByRef pbOutput As Any, _
    ByVal cbOutput As Long, _
    ByRef pcbResult As Long, _
    ByVal dwFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptDecrypt Lib "BCrypt.dll" ( _
    ByVal hKey As LongPtr, _
    ByRef pbInput As Any, _
    ByVal cbInput As Long, _
    ByRef pPaddingInfo As Any, _
    ByRef pbIV As Any, _
    ByVal cbIV As Long, _
    ByRef pbOutput As Any, _
    ByVal cbOutput As Long, _
    ByRef pcbResult As Long, _
    ByVal dwFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptDestroyKey Lib "BCrypt.dll" ( _
    ByVal hKey As LongPtr) _
    As Long

Private Declare PtrSafe Sub RtlMoveMemory Lib "Kernel32.dll" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As LongPtr)

' Constants for API.
'
Private Const BcryptBlockPadding    As Long = 1
Private Const StatusNtSuccess       As Long = 0

'<><><><><><><<><><><><><>

' User defined types.
'
Private Type QuadSextet
    Sextet1 As Byte
    Sextet2 As Byte
    Sextet3 As Byte
    Sextet4 As Byte
End Type

' Enums.
'
' Allowed BCrypt hash algorithms.
Public Enum cBcHashAlgorithm
    [_First] = 1
    bcMd2 = 1
    bcMd4 = 2
    bcMd5 = 3
    bcSha1 = 4
    bcSha256 = 5
    bcSha384 = 6
    bcSha512 = 7
    [_Last] = 7
End Enum

' Allowed BCrypt random algorithms.
Public Enum cBcRandomAlgorithm
    [_First] = 1
    bcRng = 1
    bcFips186DsaRng = 2
    [_Last] = 2
End Enum

' Default application values.
'

' Used to find random Long
Private Const sc_MAXLONG As Long = 2147483647
' Used to find random Integer
Private Const sc_MAXINT As Integer = 32767

' Default encryption algorithm.
'Private Const DefaultEncryptionAlgorithm    As Long = BcEncryptionAlgorithm.bcAes
' Use string value instead of ENUM RAM 10-10-24
Private Const DefaultEncryptionAlgorithm    As String = "AES"
' Default hash algorithm.
Private Const DefaultcBcHashAlgorithm        As Long = cBcHashAlgorithm.bcSha256
' Default random algorithm.
Private Const DefaultcBcRandomAlgorithm      As Long = cBcRandomAlgorithm.bcRng

' Application contants.
'
' Maximum size (byte count) of a Binary field of an Access table.
Private Const MaximumBinaryFieldSize        As Integer = 510
' Maximum size (character count) of a Short Text field of an Access table.
Private Const MaximumTextFieldSize          As Integer = 255

Private Type CryptType
    LoggerFunctionName As String       ' Handle error logging

    ClassError As Long                 ' Save last error number
    ClassErrorText As String           ' Save last error description
    ClassErrorSource As String         ' Save last error source
End Type
Private This As CryptType

Private Sub LogMessage(ByVal FunctionName As String, _
                       Optional ByVal ErrorMessage As String = vbNullString)
    With This
        .ClassError = Err.Number
        .ClassErrorText = Err.Description
        .ClassErrorSource = ModuleName & "." & .ClassErrorSource
        On Error Resume Next
        If This.LoggerFunctionName = vbNullString Then
            Debug.Print .ClassErrorSource, _
                        ErrorMessage, _
                       "error (" & .ClassError & ")", _
                       .ClassErrorText
        Else
            ' This will cause an un-trapped error if LoggerFunctionName cannot be found!
            Application.Run .LoggerFunctionName, _
                            .ClassErrorSource, _
                             ErrorMessage, _
                            .ClassError, _
                            .ClassErrorText
        End If
    End With
End Sub

Private Sub ClearLastError()
    On Error Resume Next
    This.ClassError = 0 ' ClassAdmErrorEnum.Noerror
    This.ClassErrorText = vbNullString
End Sub

' Convert and return a Base64 encoded string as a byte array.
' Generic function.
'
' To be called from function Decrypt.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Base64Bytes( _
    ByVal Text64 As String) _
    As Byte()
    On Error GoTo errBase64Bytes

    Dim Index           As Long
    Dim Tail            As Long
    Dim Sextets         As QuadSextet
    Dim Data()          As Byte
    Dim DataLength      As Long

    If right(Text64, 2) = "==" Then
        Tail = 2
    ElseIf right(Text64, 1) = "=" Then
        Tail = 1
    End If
    DataLength = Len(Text64) * 3 \ 4 - Tail

    ReDim Data(0 To DataLength - 1)
    For Index = LBound(Data) To UBound(Data)
        Select Case Index Mod 3
            Case 0
                Sextets = Base64QuadSextet(Mid(Text64, (Index \ 3) * 4 + 1, 4))
                Data(Index) = Sextets.Sextet1 * 2 ^ 2 + (Sextets.Sextet2 \ 2 ^ 4)
            Case 1
                Data(Index) = (Sextets.Sextet2 * 2 ^ 4 And 255) + Sextets.Sextet3 \ 2 ^ 2
            Case 2
                Data(Index) = (Sextets.Sextet3 * 2 ^ 6 And 255) + Sextets.Sextet4
        End Select
    Next

    Base64Bytes = Data
DoneBase64Bytes:
    Exit Function
errBase64Bytes:
    LogMessage "Base64Bytes"
    Resume DoneBase64Bytes
End Function

' Convert and return a Base64 encoded string as a QuadSextet.
' Generic function.
'
' To be called from function Base64Bytes.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
' Made Private RAM 10-10-24
Private Function Base64QuadSextet( _
    ByVal Text64 As String) _
    As QuadSextet
    On Error GoTo errBase64QuadSextet

    Const Base64Table   As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

    Dim Sextets         As QuadSextet

    Sextets.Sextet1 = InStr(Base64Table, Mid(Text64, 1, 1)) - 1
    Sextets.Sextet2 = InStr(Base64Table, Mid(Text64, 2, 1)) - 1
    Sextets.Sextet3 = InStr(Base64Table, Mid(Text64, 3, 1)) - 1
    Sextets.Sextet4 = InStr(Base64Table, Mid(Text64, 4, 1)) - 1

    Base64QuadSextet = Sextets
doneBase64QuadSextet:
    Exit Function
errBase64QuadSextet:
    LogMessage "Base64QuadSextet"
    Resume doneBase64QuadSextet
End Function

' Return the literal hash algorithm name determined by
' the passed value of enum cBcHashAlgorithm.
'
' To be called from functions CngHash and IsBcryptHashAlgorithm.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
' Made private RAM 10-10-24
Private Function BcryptHashAlgorithm( _
    ByVal BcryptHashAlgorithmId As cBcHashAlgorithm) _
    As String
    On Error GoTo ErrBcryptHashAlgorithm

    Dim HashAlgorithmName       As String

    ' Note: HashAlgorithmName must be in UPPERCASE.
    Select Case BcryptHashAlgorithmId
        Case cBcHashAlgorithm.bcMd2
            HashAlgorithmName = "MD2"
        Case cBcHashAlgorithm.bcMd4
            HashAlgorithmName = "MD4"
        Case cBcHashAlgorithm.bcMd5
            HashAlgorithmName = "MD5"
        Case cBcHashAlgorithm.bcSha1
            HashAlgorithmName = "SHA1"
        Case cBcHashAlgorithm.bcSha256
            HashAlgorithmName = "SHA256"
        Case cBcHashAlgorithm.bcSha384
            HashAlgorithmName = "SHA384"
        Case cBcHashAlgorithm.bcSha512
            HashAlgorithmName = "SHA512"
    End Select

    BcryptHashAlgorithm = HashAlgorithmName
doneBcryptHashAlgorithm:
    Exit Function
ErrBcryptHashAlgorithm:
    LogMessage "BcryptHashAlgorithm"
    Resume doneBcryptHashAlgorithm
End Function

' Return the literal random algorithm name determined by
' the passed value of enum cBcRandomAlgorithm.
'
' To be called from functions CngRandom and IsBcryptRandomAlgorithm.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
' Made private RAM 10-10-24
Private Function BcryptRandomAlgorithm( _
    ByVal BcryptRandomAlgorithmId As cBcRandomAlgorithm) _
    As String
    On Error GoTo errBcryptRandomAlgorithm
    Dim RandomAlgorithmName     As String

    ' Note: RandomAlgorithmName must be in UPPERCASE.
    Select Case BcryptRandomAlgorithmId
        Case cBcRandomAlgorithm.bcRng
            RandomAlgorithmName = "RNG"
        Case cBcRandomAlgorithm.bcFips186DsaRng
            RandomAlgorithmName = "FIPS186DSARNG"
    End Select

    BcryptRandomAlgorithm = RandomAlgorithmName
doneBcryptRandomAlgorithm:
    Exit Function
errBcryptRandomAlgorithm:
    LogMessage "BcryptRandomAlgorithm"
    Resume doneBcryptRandomAlgorithm
End Function

' Convert and return a byte array as a Base64 encoded string.
' Generic function.
'
' To be called from functions Encrypt, Hash, and TextBase64.
'
' 2022-02-20. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ByteBase64( _
    ByRef Data() As Byte) _
    As String
    On Error GoTo errBcryptRandomAlgorithm

    Const Base64Table   As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

    Dim Sextets         As QuadSextet
    Dim Index           As Long
    Dim Text            As String
    Dim Number          As Long

    If StrPtr(Data) = 0 Then
        ' No data. Nothing to do.
    Else
        Number = (1 + (UBound(Data) \ 3)) * 4
        Text = String(Number, "=")

        For Index = 0 To (UBound(Data) - 2) \ 3
            Select Case (UBound(Data) - 2)
                Case -2
                    Sextets = BytesQuadSextet(Data(Index * 3))
                Case -1
                    Sextets = BytesQuadSextet(Data(Index * 3), Data(Index * 3 + 1))
                Case Else
                    Sextets = BytesQuadSextet(Data(Index * 3), Data(Index * 3 + 1), Data(Index * 3 + 2))
            End Select
            Mid(Text, (Index * 4) + 1, 1) = Mid(Base64Table, Sextets.Sextet1 + 1, 1)
            Mid(Text, (Index * 4) + 2, 1) = Mid(Base64Table, Sextets.Sextet2 + 1, 1)
            If UBound(Data) >= 1 Then
                Mid(Text, (Index * 4) + 3, 1) = Mid(Base64Table, Sextets.Sextet3 + 1, 1)
            End If
            If UBound(Data) >= 2 Then
                Mid(Text, (Index * 4) + 4, 1) = Mid(Base64Table, Sextets.Sextet4 + 1, 1)
            End If
        Next

        If UBound(Data) >= 2 Then
            Select Case (1 + UBound(Data)) Mod 3
                Case Is = 2
                    ' Leave one trailing equal sign.
                    Sextets = BytesQuadSextet(Data(Index * 3), Data(Index * 3 + 1))
                    Mid(Text, (Index * 4) + 1, 1) = Mid(Base64Table, Sextets.Sextet1 + 1, 1)
                    Mid(Text, (Index * 4) + 2, 1) = Mid(Base64Table, Sextets.Sextet2 + 1, 1)
                    Mid(Text, (Index * 4) + 3, 1) = Mid(Base64Table, Sextets.Sextet3 + 1, 1)
                Case Is = 1
                    ' Leave two trailing equal signs.
                    Sextets = BytesQuadSextet(Data(Index * 3))
                    Mid(Text, (Index * 4) + 1, 1) = Mid(Base64Table, Sextets.Sextet1 + 1, 1)
                    Mid(Text, (Index * 4) + 2, 1) = Mid(Base64Table, Sextets.Sextet2 + 1, 1)
                Case Is = 0
                    ' Leave no trailing equal sign.
            End Select
        End If
    End If

    ByteBase64 = Text
doneBcryptRandomAlgorithm:
    Exit Function
errBcryptRandomAlgorithm:
    LogMessage "BcryptRandomAlgorithm"
    Resume doneBcryptRandomAlgorithm
End Function

' Convert and return one, two, or three byte arrays as a QuadSextet.
' Generic function.
'
' To be called from function ByteBase64.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
' Made private RAM 10-10-24
Private Function BytesQuadSextet( _
    ByVal Byte1 As Byte, _
    Optional ByVal Byte2 As Byte, _
    Optional ByVal Byte3 As Byte) _
    As QuadSextet
    On Error GoTo errBytesQuadSextet
    Dim Sextets         As QuadSextet

    Sextets.Sextet1 = Byte1 \ 4
    Sextets.Sextet2 = (((Byte1 * 2 ^ 6) And 255) \ 4) + Byte2 \ (2 ^ 4)
    Sextets.Sextet3 = (((Byte2 * 2 ^ 4) And 255) \ 4) + Byte3 \ (2 ^ 6)
    Sextets.Sextet4 = (((Byte3 * 2 ^ 2) And 255) \ 4)

    BytesQuadSextet = Sextets
doneBytesQuadSextet:
    Exit Function
errBytesQuadSextet:
    LogMessage "BytesQuadSextet"
    Resume doneBytesQuadSextet
End Function

' Decrypt data at pointer Data using AES encryption, IVectorInput and Secret.
'
' To be called from function DecryptData.
'
'   Arguments:
'       Data:               Memory pointer to data.
'       DataLength:         Length of byte array to decrypt.
'       IVectorInput:       Memory pointer to IVector.
'       IVectorInputLength: Length of byte array for the IVector.
'       Secret:             Memory pointer to 128-bits secret.
'       SecretLength:       Length of byte array for the secret for AES encrypting.
'   Output:
'       Byte array containing decrypted data.
'
' Original code by Erik A, 2019.
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Private Function CngDecrypt( _
    ByRef Data As LongPtr, _
    ByRef DataLength As Long, _
    ByRef IVectorInput As LongPtr, _
    ByVal IVectorInputLength As Long, _
    ByRef Secret As LongPtr, _
    ByRef SecretLength As Long) _
    As Byte()
    On Error GoTo errCngDecrypt

    Const Implementation    As LongPtr = 0
    Const Flags             As Long = 0
    Const Result            As Long = 0

    Dim Algorithm           As LongPtr
    Dim AlgorithmId         As String
    Dim Property            As String
    Dim KeyObject()         As Byte
    Dim KeyObjectLength     As Long
    Dim IVector()           As Byte
    Dim IVectorLength       As Long
    Dim Value               As String
    Dim Key                 As LongPtr
    Dim PlainText()         As Byte
    Dim PlainTextLength     As Long
    Dim ResultLength        As Long     ' Not used.
    Dim Status              As Long

    ' Open algorithm provider.
    ' Use constant instead of function RAM 10-10-24
    AlgorithmId = DefaultEncryptionAlgorithm & vbNullChar
    Status = BCryptOpenAlgorithmProvider(Algorithm, StrPtr(AlgorithmId), Implementation, Flags)

    If Status = StatusNtSuccess Then
        ' Allocate memory to hold KeyObject.
        Property = "ObjectLength" & vbNullChar
        Status = BCryptGetProperty(Algorithm, StrPtr(Property), KeyObjectLength, LenB(KeyObjectLength), Result, Flags)
    End If

    If Status = StatusNtSuccess Then
        ReDim KeyObject(0 To KeyObjectLength - 1)
        ' Calculate the block length for the IVector.
        Property = "BlockLength" & vbNullChar
        Status = BCryptGetProperty(Algorithm, StrPtr(Property), IVectorLength, LenB(IVectorLength), Result, Flags)
        If IVectorLength > IVectorInputLength Then
            Debug.Print "IVector lengths:", IVectorLength, IVectorInputLength
            Status = Not StatusNtSuccess
        End If
    End If

    If Status = StatusNtSuccess Then
        ' Resize and copy the initialization vector.
        ReDim IVector(0 To IVectorLength - 1)
        RtlMoveMemory IVector(0), ByVal IVectorInput, IVectorLength

        ' Set chaining mode.
        Property = "ChainingMode" & vbNullChar
        Value = "ChainingModeCBC" & vbNullChar
        Status = BCryptSetProperty(Algorithm, StrPtr(Property), ByVal StrPtr(Value), LenB(Value), Flags)
    End If

    If Status = StatusNtSuccess Then
        ' Create KeyObject using secret.
        If BCryptGenerateSymmetricKey(Algorithm, Key, KeyObject(1), KeyObjectLength, ByVal Secret, SecretLength, Flags) = StatusNtSuccess Then
            ' Calculate output buffer size and allocate output buffer.
            If BCryptDecrypt(Key, ByVal Data, DataLength, ByVal 0, IVector(0), IVectorLength, ByVal 0, 0, PlainTextLength, BcryptBlockPadding) = StatusNtSuccess Then
                ReDim PlainText(0 To PlainTextLength - 1)
                ' Decrypt the data.
                If BCryptDecrypt(Key, ByVal Data, DataLength, ByVal 0, IVector(0), IVectorLength, PlainText(0), PlainTextLength, ResultLength, BcryptBlockPadding) = StatusNtSuccess Then
                    CngDecrypt = PlainText
                End If
            End If
        End If
    End If

doneCngDecrypt:
    On Error Resume Next
    ' Clean up.
    If Algorithm <> 0 Then
        BCryptCloseAlgorithmProvider Algorithm, Flags
    End If
    If Key <> 0 Then
        BCryptDestroyKey Key
    End If
    Exit Function
errCngDecrypt:
    LogMessage "CngDecrypt"
    Resume doneCngDecrypt
End Function

' Encrypt data at pointer Data using AES encryption, IVectorInput and Secret.
'
' To be called from functions CngEncryptW and EncryptData.
'
'   Arguments:
'       Data:               Memory pointer to data.
'       DataLength:         Length of byte array to encrypt.
'       IVectorInput:       Memory pointer to IVector.
'       IVectorInputLength: Length of byte array for the IVector.
'       Secret:             Memory pointer to 128-bits secret.
'       SecretLength:       Length of byte array for the secret for AES encrypting.
'   Output:
'       Byte array containing encrypted data.
'
' Original code by Erik A, 2019.
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Private Function CngEncrypt( _
    ByVal Data As LongPtr, _
    ByVal DataLength As Long, _
    ByVal IVectorInput As LongPtr, _
    ByVal IVectorInputLength As Long, _
    ByVal Secret As LongPtr, _
    ByVal SecretLength As Long) _
    As Byte()
    On Error GoTo errCngEncrypt

    Const Implementation    As LongPtr = 0
    Const Flags             As Long = 0
    Const Result            As Long = 0

    Dim Algorithm           As LongPtr
    Dim AlgorithmId         As String
    Dim Property            As String
    Dim KeyObject()         As Byte
    Dim KeyObjectLength     As Long
    Dim IVector()           As Byte
    Dim IVectorLength       As Long
    Dim Value               As String
    Dim Key                 As LongPtr
    Dim CipherText()        As Byte
    Dim CipherTextLength    As Long
    Dim ResultLength        As Long     ' Not used.
    Dim Status              As Long

    ' Open algorithm provider.
    ' Use constant instead of function RAM 10-10-24
    AlgorithmId = DefaultEncryptionAlgorithm & vbNullChar
    Status = BCryptOpenAlgorithmProvider(Algorithm, StrPtr(AlgorithmId), Implementation, Flags)

    If Status = StatusNtSuccess Then
        ' Allocate memory to hold KeyObject.
        Property = "ObjectLength" & vbNullChar
        Status = BCryptGetProperty(Algorithm, StrPtr(Property), KeyObjectLength, LenB(KeyObjectLength), Result, Flags)
    End If

    If Status = StatusNtSuccess Then
        ReDim KeyObject(0 To KeyObjectLength - 1)
        ' Check that block length = 128 bits.
        Property = "BlockLength" & vbNullChar
        Status = BCryptGetProperty(Algorithm, StrPtr(Property), IVectorLength, LenB(IVectorLength), Result, Flags)
        If IVectorLength > IVectorInputLength Then
            Debug.Print "IVector lengths:", IVectorLength, IVectorInputLength
            Status = Not StatusNtSuccess
        End If
    End If

    If Status = StatusNtSuccess Then
        ' Resize and copy the initialization vector.
        ReDim IVector(0 To IVectorLength - 1)
        RtlMoveMemory IVector(0), ByVal IVectorInput, IVectorLength

        ' Set chaining mode.
        Property = "ChainingMode" & vbNullChar
        Value = "ChainingModeCBC" & vbNullChar
        Status = BCryptSetProperty(Algorithm, StrPtr(Property), ByVal StrPtr(Value), LenB(Value), Flags)
    End If

    If Status = StatusNtSuccess Then
        ' Create KeyObject using secret.
        If BCryptGenerateSymmetricKey(Algorithm, Key, KeyObject(0), KeyObjectLength, ByVal Secret, SecretLength, Flags) = StatusNtSuccess Then
            ' Calculate output buffer size and allocate output buffer.
            If BCryptEncrypt(Key, ByVal Data, DataLength, ByVal 0, IVector(0), IVectorLength, ByVal 0, 0, CipherTextLength, BcryptBlockPadding) = StatusNtSuccess Then
                ReDim CipherText(0 To CipherTextLength - 1)
                ' Encrypt the data.
                If BCryptEncrypt(Key, ByVal Data, DataLength, ByVal 0, IVector(0), IVectorLength, CipherText(0), CipherTextLength, ResultLength, BcryptBlockPadding) = StatusNtSuccess Then
                    ' Output the encrypted data.
                    CngEncrypt = CipherText
                End If
            End If
        End If
    End If

doneCngEncrypt:
    On Error Resume Next
    ' Clean up.
    If Algorithm <> 0 Then
        BCryptCloseAlgorithmProvider Algorithm, Flags
    End If
    If Key <> 0 Then
        BCryptDestroyKey Key
    End If
    Exit Function
errCngEncrypt:
    LogMessage "CngEncrypt"
    Resume doneCngEncrypt
End Function

' Create a hash of the data at pointer Data by using the Next Generation Cryptography API
' and returns the hash as a byte array.
'
' To be called from functions DecryptData and HashData.
'
' Allowed algorithms (NB: Hash algorithms only; check OS support):
'   https://docs.microsoft.com/en-us/windows/desktop/SecCNG/cng-algorithm-identifiers
'
' Loosely based on the method described at:
'   https://docs.microsoft.com/en-us/windows/desktop/SecCNG/creating-a-hash-with-cng
'
' Original code by Erik A, 2019.
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Private Function CngHash( _
    ByRef Data As LongPtr, _
    ByVal DataLength As Long, _
    Optional ByVal BcryptHashAlgorithmId As cBcHashAlgorithm = DefaultcBcHashAlgorithm) _
    As Byte()
    On Error GoTo errCngHash

    Const Implementation    As LongPtr = 0
    Const Flags             As Long = 0
    Const Secret            As Long = 0
    Const SecretLength      As Long = 0
    Const Result            As Long = 0

    Dim Algorithm           As LongPtr
    Dim AlgorithmId         As String
    Dim Property            As String
    Dim HashObject()        As Byte
    Dim HashObjectLength    As Long
    Dim HashDigestLength    As Long
    Dim Hash                As LongPtr
    Dim Values()            As Byte
    Dim Status              As Long

    ' Validate requested hash algorithm.
    If Not IsBcryptHashAlgorithmId(BcryptHashAlgorithmId) Then
        ' Invalid hash algorithm id. Use default hash algorithm.
        BcryptHashAlgorithmId = DefaultcBcHashAlgorithm
    End If

    ' Open algorithm provider.
    AlgorithmId = BcryptHashAlgorithm(BcryptHashAlgorithmId) & vbNullChar
    Status = BCryptOpenAlgorithmProvider(Algorithm, StrPtr(AlgorithmId), Implementation, Flags)

    If Status = StatusNtSuccess Then
        ' Determine hash object size and allocate memory.
        Property = "ObjectLength" & vbNullChar
        Status = BCryptGetProperty(Algorithm, StrPtr(Property), HashObjectLength, LenB(HashObjectLength), Result, Flags)
    End If

    If Status = StatusNtSuccess Then
        ReDim HashObject(0 To HashObjectLength - 1)
        ' Determine hash digest size and allocate memory.
        Property = "HashDigestLength" & vbNullChar
        Status = BCryptGetProperty(Algorithm, StrPtr(Property), HashDigestLength, LenB(HashDigestLength), Result, Flags)
    End If

    If Status = StatusNtSuccess Then
        ReDim Values(0 To HashDigestLength - 1)
        ' Create the hash object.
        If BCryptCreateHash(Algorithm, Hash, HashObject(0), HashObjectLength, Secret, SecretLength, Flags) = StatusNtSuccess Then
            ' Hash the data.
            If BCryptHashData(Hash, ByVal Data, DataLength) = StatusNtSuccess Then
                ' Get the hash.
                If BCryptFinishHash(Hash, Values(0), HashDigestLength, Flags) = StatusNtSuccess Then
                    ' Return the hash.
                    CngHash = Values
                End If
            End If
        End If
    End If

doneCngHash:
    On Error Resume Next
    ' Clean up.
    If Algorithm <> 0 Then
        BCryptCloseAlgorithmProvider Algorithm, Flags
    End If
    If Hash <> 0 Then
        BCryptDestroyHash Hash
    End If
    Exit Function
errCngHash:
    LogMessage "CngHash"
    Resume doneCngHash
End Function

' Fill data at pointer Data with random bytes.
'
' To be called from function RandomData.
'
' Allowed algorithms (NB: Random algorithms only; check OS support):
'   https://docs.microsoft.com/en-us/windows/desktop/SecCNG/cng-algorithm-identifiers
'
' Original code by Erik A, 2019.
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Private Sub CngRandom( _
    ByRef Data As LongPtr, _
    ByRef BufferSize As Long, _
    Optional ByVal BcryptRandomAlgorithmId As cBcRandomAlgorithm = DefaultcBcRandomAlgorithm)
    On Error GoTo errCngRandom

    Const Implementation    As LongPtr = 0
    Const Flags             As Long = 0

    Dim Algorithm           As LongPtr
    Dim AlgorithmId         As String

    ' Validate requested randomise algorithm.
    If Not IsBcryptRandomAlgorithmId(BcryptRandomAlgorithmId) Then
        ' Invalid randomise algorithm id. Use default randomise algorithm.
        BcryptRandomAlgorithmId = DefaultcBcRandomAlgorithm
    End If

    ' Open crypto provider.
    AlgorithmId = BcryptRandomAlgorithm(BcryptRandomAlgorithmId) & vbNullChar
    BCryptOpenAlgorithmProvider Algorithm, StrPtr(AlgorithmId), Implementation, Flags

    ' Fill byte array with random data and return size of buffer.
    BCryptGenRandom Algorithm, ByVal Data, BufferSize, Flags

doneCngRandom:
    On Error Resume Next
    ' Clean up.
    BCryptCloseAlgorithmProvider Algorithm, Flags
    Exit Sub
errCngRandom:
    LogMessage "CngRandom"
    Resume doneCngRandom
End Sub

' Decrypt an AES encrypted byte array using a key passed as another byte array.
' Return by reference the decrypted data as a byte array.
' Return True if success.
'
' To be called from function Decrypt.
'
' Original code by Erik A, 2019.
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
' Made private RAM 10-10-24
Private Function DecryptData( _
    ByRef EncryptedData() As Byte, _
    ByRef KeyData() As Byte, _
    ByRef DecryptedData() As Byte) _
    As Boolean
    On Error GoTo errDecryptData

    Const SizeLength        As Long = 4
    Const IVectorLength     As Long = 16
    Const SecretLength      As Long = 16

    Dim DataLength          As Long
    Dim DataHash()          As Byte
    Dim KeyHash()           As Byte
    Dim IVector             As LongPtr
    Dim Data()              As Byte
    Dim Index               As Long
    Dim HashError           As Boolean
    Dim Result              As Boolean

    If StrPtr(EncryptedData) = 0 Or StrPtr(KeyData) = 0 Then
        ' Either no data or no key. Nothing to do.
    ElseIf EncryptedByteLength(UBound(EncryptedData) - LBound(EncryptedData) + 1) = 0 Then
        ' Data length is below minimum.
    Else
        ' Get the SHA1 hash of the key.
        KeyHash = HashData(KeyData, bcSha1)
        ' Get the pointer to IVector. The last 16 bytes is IVector.
        IVector = VarPtr(EncryptedData(UBound(EncryptedData) - LBound(EncryptedData) + 1 - IVectorLength))
        ' Decrypt the data.
        Data = CngDecrypt( _
            VarPtr(EncryptedData(0)), UBound(EncryptedData) - LBound(EncryptedData) + 1 - IVectorLength, _
            IVector, IVectorLength, _
            VarPtr(KeyHash(0)), SecretLength)

        ' Check that CngDecrypt did return some data.
        If StrPtr(Data) = 0 Then
            ' No data. Most likely, a wrong key was passed.
        ElseIf UBound(Data) <= SizeLength Then
            ' No useful data returned. A minimum of five bytes is expected.
        Else
            ' Get the data length.
            RtlMoveMemory DataLength, Data(0), SizeLength

            ' Check if length is valid, with invalid key length = random data.
            If DataLength > (UBound(Data) - LBound(Data) + 1 - SizeLength) Or DataLength < 0 Then
                ' No useful data.
                ' Leave DecryptedData empty, and return False.
            Else
                ' Get the hash of the decrypted data.
                DataHash = CngHash(VarPtr(Data(SizeLength)), DataLength, bcSha1)
                ' Verify the hash.
                For Index = LBound(DataHash) To UBound(DataHash)
                    If DataHash(Index) <> Data(SizeLength + DataLength + Index) Then
                        ' Stored hash not equal to hash with decrypted data, key incorrect, or encrypted data has been tampered with.
                        ' Leave DecryptedData empty, and return False.
                        HashError = True
                        Exit For
                    End If
                Next
                If Not HashError Then
                    ' Initialize output array.
                    ReDim DecryptedData(0 To DataLength - 1)
                    ' Copy data to output array.
                    RtlMoveMemory DecryptedData(0), Data(SizeLength), DataLength
                    Result = True
                End If
            End If
        End If
    End If

    DecryptData = Result
doneDecryptData:
    Exit Function
errDecryptData:
    LogMessage "DecryptData"
    Resume doneDecryptData
End Function

' Encrypt a byte array using AES encryption and a KeyData passed as another byte array.
' Return by reference the encrypted data as a byte array.
' Return True if success.
'
' To be called from function Encrypt.
'
' NOTE:
'   Even when passed the same arguments (TextData and KeyData), the returned and
'   encrypted data will be unique for every call.
'
' Original code by Erik A, 2019.
' 2022-04-03. Gustav Brock, Cactus Data ApS, CPH.
'
' Made private RAM 10-10-24
Private Function EncryptData( _
    ByRef TextData() As Byte, _
    ByRef KeyData() As Byte, _
    ByRef EncryptedData() As Byte) _
    As Boolean
    On Error GoTo errEncryptData

    Const SizeLength        As Long = 4
    Const IVectorLength     As Long = 16
    Const SecretLength      As Long = 16

    Dim KeyHash()           As Byte
    Dim InputHash()         As Byte
    Dim Data()              As Byte
    Dim DataLength          As Long
    Dim IVector()           As Byte
    Dim InputHashLength     As Long
    Dim Result              As Boolean

    ' Get SHA1 hash of the data and of the KeyData.
    InputHash = HashData(TextData, bcSha1)
    If StrPtr(InputHash) = 0 Then
        ' Empty data. Exit with success.
        Result = True
    Else
        InputHashLength = UBound(InputHash) + 1
        KeyHash = HashData(KeyData, bcSha1)
        ReDim Preserve KeyHash(0 To SecretLength)
    End If

    If StrPtr(InputHash) = 0 Or StrPtr(KeyHash) = 0 Then
        ' Either no data or no KeyData. Nothing to do.
    Else
        DataLength = UBound(TextData) - LBound(TextData) + 1

        ' Data size is: Long (4 bytes) + DataLength + SHA1 (20 bytes)
        ReDim Data(0 To SizeLength + DataLength + InputHashLength - 1)
        ' Append length (in bytes) to start of array.
        RtlMoveMemory Data(0), DataLength, SizeLength
        ' Append data.
        RtlMoveMemory Data(SizeLength), TextData(LBound(TextData)), DataLength
        ' Append hash of the data.
        RtlMoveMemory Data(SizeLength + DataLength), InputHash(0), InputHashLength

        ' Generate IVector.
        IVector = RandomData(IVectorLength)
        ' Encrypt data.
        EncryptedData = CngEncrypt( _
            VarPtr(Data(0)), SizeLength + DataLength + InputHashLength, _
            VarPtr(IVector(0)), IVectorLength, _
            VarPtr(KeyHash(0)), SecretLength)
        ' Deallocate copy made to encrypt.
        Erase Data
        ' Extend encrypted data to append IVector.
        ReDim Preserve EncryptedData(LBound(EncryptedData) To UBound(EncryptedData) + IVectorLength)
        ' Append IVector.
        RtlMoveMemory EncryptedData(UBound(EncryptedData) - LBound(EncryptedData) + 1 - IVectorLength), IVector(0), IVectorLength
        Result = True
    End If

    EncryptData = Result
doneEncryptData:
    Exit Function
errEncryptData:
    LogMessage "EncryptData"
    Resume doneEncryptData
End Function

' Create and return the hash of a byte array as another byte array
' using the specified hash algorithm.
' By default, hash algorithm SHA256 is used.
'
' To be called from functions EncryptData, DecryptData, and Hash.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
' Made private RAM 10-10-24
Private Function HashData( _
    ByRef TextData() As Byte, _
    Optional ByVal BcryptHashAlgorithmId As cBcHashAlgorithm = DefaultcBcHashAlgorithm) _
    As Byte()
    On Error GoTo errHashData
    Dim Data()          As Byte
    Dim DataLength      As Long

    If StrPtr(TextData) = 0 Then
        ' Array TextData is empty.
    Else
        DataLength = UBound(TextData) - LBound(TextData) + 1
        Data = CngHash(VarPtr(TextData(LBound(TextData))), DataLength, BcryptHashAlgorithmId)
    End If

    HashData = Data
doneHashData:
    Exit Function
errHashData:
    LogMessage "HashData"
    Resume doneHashData
End Function

' Return True if the passed text value represents a value of
' enum cBcHashAlgorithm.
' Note: To validate, HashAlgorithm must be in UPPERCASE.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
' Made private RAM 10-10-24
Private Function IsBcryptHashAlgorithm( _
    ByVal HashAlgorithm As String) _
    As Boolean
    On Error GoTo errIsBcryptHashAlgorithm
    Dim Index           As cBcHashAlgorithm
    Dim Result          As Boolean

    For Index = cBcHashAlgorithm.[_First] To cBcHashAlgorithm.[_Last]
        If BcryptHashAlgorithm(Index) = HashAlgorithm Then
            Result = True
            Exit For
        End If
    Next

    IsBcryptHashAlgorithm = Result
doneIsBcryptHashAlgorithm:
    Exit Function
errIsBcryptHashAlgorithm:
    LogMessage "IsBcryptHashAlgorithm"
    Resume doneIsBcryptHashAlgorithm
End Function

' Return True if the passed value of enum cBcHashAlgorithm is valid.
'
' To be called from function CngHash.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
' Made private RAM 10-10-24
Private Function IsBcryptHashAlgorithmId( _
    ByVal HashAlgorithmId As cBcHashAlgorithm) _
    As Boolean
    On Error GoTo errIsBcryptHashAlgorithmId
    Dim Result          As Boolean

    If cBcHashAlgorithm.[_First] <= HashAlgorithmId And HashAlgorithmId <= cBcHashAlgorithm.[_Last] Then
        Result = True
    End If

    IsBcryptHashAlgorithmId = Result
doneIsBcryptHashAlgorithmId:
    Exit Function
errIsBcryptHashAlgorithmId:
    LogMessage "IsBcryptHashAlgorithmId"
    Resume doneIsBcryptHashAlgorithmId
End Function

' Return True if the passed text value represents a value of
' enum cBcRandomAlgorithm.
' Note: To validate, RandomAlgorithm must be in UPPERCASE.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
' Made private RAM 10-10-24
Private Function IsBcryptRandomAlgorithm( _
    ByVal RandomAlgorithm As String) _
    As Boolean
    On Error GoTo errIsBcryptRandomAlgorithm
    Dim Index           As cBcRandomAlgorithm
    Dim Result          As Boolean

    For Index = cBcRandomAlgorithm.[_First] To cBcRandomAlgorithm.[_Last]
        If BcryptRandomAlgorithm(Index) = RandomAlgorithm Then
            Result = True
            Exit For
        End If
    Next

    IsBcryptRandomAlgorithm = Result
doneIsBcryptRandomAlgorithm:
    Exit Function
errIsBcryptRandomAlgorithm:
    LogMessage "IsBcryptRandomAlgorithm"
    Resume doneIsBcryptRandomAlgorithm
End Function

' Return True if the passed value of enum cBcRandomAlgorithm
' is valid.
'
' To be called from function CngRandom.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
' Made private RAM 10-10-24
Private Function IsBcryptRandomAlgorithmId( _
    ByVal RandomAlgorithmId As cBcRandomAlgorithm) _
    As Boolean
    On Error GoTo errIsBcryptRandomAlgorithmId
    Dim Result          As Boolean

    If cBcRandomAlgorithm.[_First] <= RandomAlgorithmId And RandomAlgorithmId <= cBcRandomAlgorithm.[_Last] Then
        Result = True
    End If

    IsBcryptRandomAlgorithmId = Result
doneIsBcryptRandomAlgorithmId:
    Exit Function
errIsBcryptRandomAlgorithmId:
    LogMessage "IsBcryptRandomAlgorithmId"
    Resume doneIsBcryptRandomAlgorithmId
End Function

' Return a byte array with random bytes using the specified random algorithm.
' By default, random algorithm RNG is used.
'
' To be called from function Random or EncryptData.
'
' Allowed algorithms (NB: Random algorithms only; check OS support):
'   https://docs.microsoft.com/en-us/windows/desktop/SecCNG/cng-algorithm-identifiers
'
' 2022-02-20. Gustav Brock, Cactus Data ApS, CPH.
'
' Made private RAM 10-10-24
Private Function RandomData( _
    ByVal DataLength As Long, _
    Optional ByVal BcryptRandomAlgorithmId As cBcRandomAlgorithm = DefaultcBcRandomAlgorithm)
    On Error GoTo errRandomData

    Dim Data()      As Byte

    If DataLength <= 0 Then
        ' Nothing to do.
    Else
        ReDim Data(0 To DataLength - 1)
        CngRandom VarPtr(Data(LBound(Data))), DataLength, BcryptRandomAlgorithmId
    End If

    RandomData = Data
doneRandomData:
    Exit Function
errRandomData:
    LogMessage "RandomData"
    Resume doneRandomData
End Function

'<><><><><><><><><><><><><><><><>
'       Start of Public Code
'<><><><><><><><><><><><><><><><>

'-----------------  Handle error logging RAM 3-4-2025
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

' Decrypt a Base64 encoded string encrypted using AES and a key.
' Return the decrypted and decoded text as a plain string.
'
' Example:
'   EncryptedText = "6uLffExuQmAi/oI3AzCLZTRZfv1XL6kl01z4hJ5y1MWXHgFACj3XhvboF/rNU89znrX1d5btmCbRK9dAjjjlKxTDJMImQr3YGiscMDvn/YtjKmc8nFuR65IU9vEn4a0Rca72k55cZXjKzOGMpbZ/6A=="
'   Key = "Have a Cigar"
'   Text = Decrypt(EncryptedText, Key)
'   Text -> Careful with that axe, Eugene!
'
' Original code by Erik A, 2019.
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Decrypt( _
    ByVal EncryptedText As String, _
    ByVal Key As String) _
    As String
    On Error GoTo errDecrypt

    Dim EncryptedData()     As Byte
    Dim TextData()          As Byte

    ClearLastError

    If EncryptedText = "" Or Key = "" Then
        ' Nothing to do.
    Else
        ' Convert the Base64 encoded string to a byte array.
        EncryptedData = Base64Bytes(EncryptedText)
        If DecryptData(EncryptedData, (Key), TextData) = True Then
            ' Success.
        Else
            ' Invalid EncryptedData or wrong key.
        End If
    End If

    Decrypt = TextData
doneDecrypt:
    Exit Function
errDecrypt:
    LogMessage "Decrypt"
    Resume doneDecrypt
End Function

' Return for an avaliable length of an encrypted byte array
' the possible maximum length of a byte array to be encrypted.
'
' To be called from function FitByteField.
'
' Examples:
'   Available encrypted     Length plain maximum
'   0                         0
'  47                         0
'  48                         1
'  63                         1
'  64                         4
'  79                         4
'  80                        12
'  95                        12
'  96                        20
' 111                        20
' 112                        28
' 127                        28
' 128                        36
' 143                        36
' 144                        44
' 159                        44
' 160                        52
' 175                        52
' 176                        60
' 191                        60
' 192                        68
' 207                        68
' 208                        76
' 223                        76
' 224                        84
' 239                        84
' 240                        92
' 255                        92
' 256                       100
' 480                       212
' 495                       212
' 496                       220
' Largest option for a Binary field of Access
' holding the encrypted byte array:
' 510                       220
' 511                       220
' 512                       228
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DecryptedByteLength( _
    ByVal EncryptedByteLength As Long) _
    As Long
    On Error GoTo errDecryptedByteLength
    Dim Length      As Long

    ClearLastError

    Select Case EncryptedByteLength
        Case Is < 48
            Length = 0
        Case Is < 64
            Length = 1
        Case Else
            Length = Int(EncryptedByteLength / 16) * 8 - 28
    End Select

    DecryptedByteLength = Length
doneDecryptedByteLength:
    Exit Function
errDecryptedByteLength:
    LogMessage "DecryptedByteLength"
    Resume doneDecryptedByteLength
End Function

' Return for an available length of an encrypted and Base64 encoded string
' the possible maximum length of a string to be encrypted.
'
' To be called from function FitTextField.
'
' Examples:
'   Available encrypted     Length plain maximum
'     0                       0
'    63                       0
'    64                       3
'    87                       3
'    88                      11
'   107                      11
'   108                      19
'   127                      19
'   128                      27
'   151                      27
'   152                      35
'   171                      35
'   172                      43
'   191                      43
'   192                      51
'   215                      51
'   216                      59
'   235                      59
' Largest option for a Short Text field of Access
' holding the decrypted string:
'   236                      67
'   255                      67
'   256                      75
'   747                     251
' Largest option for a Short Text field of Access
' holding the encrypted string:
'   748                     259
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DecryptedTextLength( _
    ByVal EncryptedTextLength As Long) _
    As Long
    On Error GoTo errDecryptedTextLength
    Dim Length      As Long
    Dim Length64    As Long
    Dim Delta20     As Long
    Dim Delta       As Long

    ClearLastError

    Select Case EncryptedTextLength
        Case Is < 64
            Length = 0
        Case Is < 64 + 24
            Length = 3
        Case Else
            Length64 = -Int(-EncryptedTextLength / 64)
            Delta20 = -Int(-((Length64 * 64 - EncryptedTextLength) / 20))
            Delta = (Delta20 \ 4) * 3 + Delta20 Mod 4
            Length = (Length64 - 1) * 24 - (Delta * 8) - 4 + 7
    End Select

    DecryptedTextLength = Length
doneDecryptedTextLength:
    Exit Function
errDecryptedTextLength:
    LogMessage "DecryptedTextLength"
    Resume doneDecryptedTextLength
End Function

' Encrypt a string using AES and a key.
' Return the encrypted text as a Base64 encoded string.
'
' Example:
'   Text = "Careful with that axe, Eugene!"
'   Key = "Have a Cigar"
'   EncryptedText = Encrypt(Text, Key)
'   EncryptedText -> 6uLffExuQmAi/oI3AzCLZTRZfv1XL6kl01z4hJ5y1MWXHgFACj3XhvboF/rNU89znrX1d5btmCbRK9dAjjjlKxTDJMImQr3YGiscMDvn/YtjKmc8nFuR65IU9vEn4a0Rca72k55cZXjKzOGMpbZ/6A==
'
' Note: Length of the encrypted string can be predetermined by the function EncryptedTextLength:
'   ' Use Text from example above.
'   Length = EncryptedTextLength(Len(Text))
'   Length -> 152
'
' Original code by Erik A, 2019.
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Encrypt( _
    ByVal Text As String, _
    ByVal Key As String) _
    As String

    Dim EncryptedData()     As Byte
    Dim EncryptedText       As String

    ClearLastError

    If Text = "" Or Key = "" Then
        ' Nothing to do.
    Else
        If EncryptData((Text), (Key), EncryptedData) = True Then
            ' Success.
            ' Convert the byte array to a Base64 encoded string.
            EncryptedText = ByteBase64(EncryptedData)
        Else
            ' Missing Text or Key.
        End If
    End If

    Encrypt = EncryptedText
doneEncrypt:
    Exit Function
errEncrypt:
    LogMessage "Encrypt"
    Resume doneEncrypt
End Function

' Return the byte length of a string of the length DecryptedTextLength
' encrypted with function EncryptData.
'
' To be called from function FitByteField.
'
' Example:
'   Text = "Careful with that axe, Eugene!"
'   DecryptedTextLength = Len(Text)     ' = 30
'   Length = EncryptedByteLength(DecryptedTextLength)
'   Length -> 112
'
' Example data:
'
' Length plain  Length encrypted
'   0             0
'   1            48
'   4            64
'  12            80
'  20            96
'  28           112
'  36           128
'  44           144
'  52           160
'  60           176
'  68           192
'  76           208
'  84           224
'  92           240
' 100           256
' 108           272
' 116           288
' 124           304
' 132           320
' 140           336
' 148           352
' 156           368
' 164           384
' 172           400
' 180           416
' 188           432
' 196           448
' 204           464
' 212           480
' 220           496
' 227           496
' 227 characters is the largest string to encrypt, if the
' encrypted byte array must fit a Binary field of Access.
' 228           512
' 236           528
' 244           544
' 252           560
' The maximum length of an Access Short Text field is 255 characters.
' 255           560
' 260           576
'
' 2022-02-18. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function EncryptedByteLength( _
    ByVal DecryptedTextLength As Long) _
    As Long
    On Error GoTo errEncryptedByteLength
    Dim Length  As Long

    ClearLastError

    If DecryptedTextLength > 0 Then
        Length = 48 + ((DecryptedTextLength + 4) \ 8) * 16
    End If

    EncryptedByteLength = Length
doneEncryptedByteLength:
    Exit Function
errEncryptedByteLength:
    LogMessage "EncryptedByteLength"
    Resume doneEncryptedByteLength
End Function

' Return the length of a string encrypted and Base64 encoded with function Encrypt.
'
' To be called from function FitTextField.
'
' Example:
'   Text = "Careful with that axe, Eugene!"
'   DecryptedTextLength = Len(Text)     ' = 30
'   Length = EncryptedTextLength(DecryptedTextLength)
'   Length -> 152
'
' Example data:
'
' Length plain  Length encrypted
'   0             0
'   1            64
'   4            88
'  12           108
'  20           128
'  28           152
'  36           172
'  44           192
'  52           216
'  60           236
'  67           236
' 67 characters is the largest string to encrypt, if the
' encrypted string must fit a Short Text field of Access.
'  68           256
'  76           280
'  84           300
'  92           320
' 100           344
' 108           364
' 116           384
' 124           408
' 132           428
' 140           448
' 148           472
' 156           492
' 164           512
' 172           536
' 180           556
' 188           576
' 196           600
' 204           620
' 212           640
' 220           664
' 228           684
' 236           704
' 244           728
' 252           748
' The maximum length of an Access Short Text field is 255 characters.
' 255           748
' 260           768
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function EncryptedTextLength( _
    ByVal DecryptedTextLength As Long) _
    As Long
    On Error GoTo errDecryptedTextLength
    Dim Length      As Long

    ClearLastError

    If DecryptedTextLength > 0 Then
        Length = 64 - Int(-(DecryptedTextLength - 3) / 8) * 20 - Int(-(DecryptedTextLength - 3) / 24) * 4
    End If

    EncryptedTextLength = Length
doneDecryptedTextLength:
    Exit Function
errDecryptedTextLength:
    LogMessage "DecryptedTextLength"
    Resume doneDecryptedTextLength
End Function

' Check if the length of the encrypted value of argument Text will fit in
' a table field (of data type Binary) having the size of argument FieldSize.
' Return True if it will fit, False if not.
'
' Optionally, if argument ChopText is True:
'   - chop Text to fit the field size
'   - return the chopped text by reference
'   - return True for any length of Text
'
' 2022-02-18. Gustav Brock, Cactus Data ApS, CPH.
'
' Remove ChopText paramiter RAM 10-10-24
Public Function FitByteField( _
    ByRef Text As String, _
    Optional ByVal FieldSize As Integer = MaximumBinaryFieldSize) _
    As Boolean
    On Error GoTo errFitByteField

    Dim Length      As Long
    Dim Success     As Boolean

    ClearLastError

    If Text = "" Then
        Success = True
    Else
        Length = EncryptedByteLength(Len(Text))
        If Length < FieldSize Then
            Success = True
        End If
    End If

    FitByteField = Success
doneFitByteField:
    Exit Function
errFitByteField:
    LogMessage "FitByteField"
    Resume doneFitByteField
End Function

' Check if the length of the encrypted value of argument Text will fit in
' a table field (of data type Text) having the size of argument FieldSize.
' Return True if it will fit, False if not.
'
' Optionally, if argument ChopText is True:
'   - chop Text to fit the field size
'   - return the chopped text by reference
'   - return True for any length of Text
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
' Remove ChopText paramiter RAM 10-10-24
Public Function FitTextField( _
    ByRef Text As String, _
    Optional ByVal FieldSize As Integer = MaximumTextFieldSize) _
    As Boolean

    Dim Length      As Long
    Dim Success     As Boolean

    ClearLastError

    If Text = "" Then
        Success = True
    Else
        Length = EncryptedTextLength(Len(Text))
        If Length < FieldSize Then
            Success = True
        End If
    End If

    FitTextField = Success
doneFitTextField:
    Exit Function
errFitTextField:
    LogMessage "FitTextField"
    Resume doneFitTextField
End Function

' Convert and return a plain string as a Base64 encoded string.
' Generic function.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TextBase64( _
    ByVal Text As String) _
    As String
    On Error GoTo errTextBase64
    Dim Text64          As String

    ClearLastError

    Text64 = ByteBase64(StrConv(Text, vbFromUnicode))

    TextBase64 = Text64
doneTextBase64:
    Exit Function
errTextBase64:
    LogMessage "TextBase64"
    Resume doneTextBase64
End Function

' Return a Base64 encoded hash of a string using the specified hash algorithm.
' By default, hash algorithm SHA256 is used.
'
' Example:
'   Text = "Get your filthy hands off my desert."
'   Value = Hash(Text)
'   Value -> "AIPgWDlQLv7bvLdg7Oa78dyRbC0tStuEXJRk0MMehOc="
'
' Length of the generated Base64 encoded hash string:
'
'   Encoding    Length
'   MD2         24
'   MD4         24
'   MD5         24
'   SHA1        28
'   SHA256      44      ' Default.
'   SHA384      64
'   SHA512      88
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Hash( _
    ByVal Text As String, _
    Optional ByVal BcryptHashAlgorithmId As cBcHashAlgorithm = DefaultcBcHashAlgorithm) _
    As String
    On Error GoTo errHash

    Dim HashBase64          As String

    ClearLastError

    If Text = "" Then
        ' No data. Nothing to do.
    Else
        HashBase64 = ByteBase64(HashData((Text), BcryptHashAlgorithmId))
    End If

    Hash = HashBase64
doneHash:
    Exit Function
errHash:
    LogMessage "Hash"
    Resume doneHash
End Function

' Return the byte count of a hash from function HashData.
'
' Default count is 32.
' Maximum count is 64.
' Minimum count is 16.
'
' 2022-02-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function HashByteLength( _
    Optional ByVal BcryptHashAlgorithmId As cBcHashAlgorithm = bcSha256) _
    As Integer
    On Error GoTo errHashByteLength
    Dim Length  As Integer

    ClearLastError

    Select Case BcryptHashAlgorithmId
        Case bcMd2, bcMd4, bcMd5
            Length = 16
        Case bcSha1
            Length = 20
        Case bcSha256
            Length = 32
        Case bcSha384
            Length = 48
        Case bcSha512
            Length = 64
    End Select

    HashByteLength = Length
doneHashByteLength:
    Exit Function
errHashByteLength:
    LogMessage "HashByteLength"
    Resume doneHashByteLength
End Function

' Return the fixed length of a Base64 encoded hash from function Hash
' using the specified hash algorithm.
'
' Default length is 44.
' Maximum length is 88.
' Minimum length is 24.
'
' 2022-02-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function HashTextLength( _
    Optional ByVal BcryptHashAlgorithmId As cBcHashAlgorithm = bcSha256) _
    As Integer
    On Error GoTo errHashTextLength

    Dim Length  As Integer

    ClearLastError

    Select Case BcryptHashAlgorithmId
        Case bcMd2, bcMd4, bcMd5
            Length = 24
        Case bcSha1
            Length = 28
        Case bcSha256
            Length = 44
        Case bcSha384
            Length = 64
        Case bcSha512
            Length = 88
    End Select

    HashTextLength = Length
doneHashTextLength:
    Exit Function
errHashTextLength:
    LogMessage "HashTextLength"
    Resume doneHashTextLength
End Function

' Convert and return a plain string as a Base64 encoded string.
' Generic function.
'
' 2024-10-10. RAM based on TextBase64
'
Public Function Base64Text( _
    ByVal Text64 As String) _
    As String
    On Error GoTo errBase64Text
    Dim Data()          As Byte

    ClearLastError

    Data = Base64Bytes(Text64)

    Base64Text = StrConv(Data, vbUnicode)
doneBase64Text:
    Exit Function
errBase64Text:
    LogMessage "Base64Text"
    Resume doneBase64Text
End Function

' Return a random string of the length specified using the
' specified random algorithm.
' By default, random algorithm RNG is used.
' Optionally, random algorithm Fips186DSARng can be specified.
'
' Optionally, if argument TrueBase64 is True, a random byte array of
' the length specified will be created and returned Base64 encoded.
'
' Examples:
'   Value = Random(23)
'   Value -> "ZXq2fue8QZ+d83Lw0T4Mi68"
'
'   Value = Random(23, , True)
'   Value -> "M/KlMXO8l2FP9Kh0GtWcAr63a4UVqZw="
'
' 2022-02-20. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Random( _
    ByVal TextSize As Long, _
    Optional ByVal BcryptRandomAlgorithmId As cBcRandomAlgorithm = DefaultcBcRandomAlgorithm, _
    Optional ByVal TrueBase64 As Boolean) _
    As String
    On Error GoTo errRandom

    Dim Text    As String

    ClearLastError

    If TextSize <= 0 Then
        ' Nothing to do.
    ElseIf TrueBase64 Then
        Text = ByteBase64(RandomData(TextSize, BcryptRandomAlgorithmId))
    ElseIf TextSize < 4 Then
        Text = left(ByteBase64(RandomData(TextSize, BcryptRandomAlgorithmId)), TextSize)
    Else
        Text = left(ByteBase64(RandomData(TextSize * 0.75, BcryptRandomAlgorithmId)), TextSize)
    End If

    Random = Text
doneRandom:
    Exit Function
errRandom:
    LogMessage "Random"
    Resume doneRandom
End Function

'-----------------
' Return a random long
' See https://stackoverflow.com/questions/44396784/vba-bit-converter
' maxVal <= minVal for full range
' 10-15-24 RAM
Public Function RandomLng( _
    Optional ByVal minVal As Long = 0, _
    Optional ByVal maxVal As Long = 0, _
    Optional ByVal BcryptRandomAlgorithmId As cBcRandomAlgorithm = DefaultcBcRandomAlgorithm) _
    As Long
    On Error GoTo errRandomLng
    Dim randomVal() As Byte
    Dim longValue As Long

    ClearLastError

    randomVal = RandomData(4, BcryptRandomAlgorithmId)
    RtlMoveMemory longValue, randomVal(LBound(randomVal)), 4
    If maxVal <= minVal Then
        RandomLng = longValue
    Else
        longValue = CLng((maxVal - minVal + 1) * (Abs(longValue) / sc_MAXLONG)) + minVal
        If longValue > maxVal Then
            ' Slight overload to maxVal
            longValue = maxVal
        End If
        RandomLng = longValue
    End If
doneRandomLng:
    Exit Function
errRandomLng:
    LogMessage "RandomLng"
    Resume doneRandomLng
End Function

'-----------------
' Return a random integer
' See https://stackoverflow.com/questions/44396784/vba-bit-converter
' maxVal <= minVal for full range
' 10-15-24 RAM
Public Function RandomInt( _
    Optional ByVal minVal As Integer = 0, _
    Optional ByVal maxVal As Integer = 0, _
    Optional ByVal BcryptRandomAlgorithmId As cBcRandomAlgorithm = DefaultcBcRandomAlgorithm) _
    As Integer
    On Error GoTo errRandomInt
    Dim randomVal() As Byte
    Dim intValue As Integer

    ClearLastError

    randomVal = RandomData(2, BcryptRandomAlgorithmId)
    RtlMoveMemory intValue, randomVal(LBound(randomVal)), 2
    If maxVal <= minVal Then
        RandomInt = intValue
    Else
        ' Slight overload to minVal
        intValue = Fix((maxVal - minVal + 1) * (Abs(intValue) / sc_MAXINT)) + minVal
        RandomInt = intValue
    End If
doneRandomInt:
    Exit Function
errRandomInt:
    LogMessage "RandomInt"
    Resume doneRandomInt
End Function

'----------------
' Return double between given values
' See https://stackoverflow.com/questions/44396784/vba-bit-converter
' 0 to .179769313486232 if range
' maxVal <= minVal for full range
' 10-15-24 RAM
Public Function RandomDbl( _
    Optional ByVal minVal As Double = 0, _
    Optional ByVal maxVal As Double = 0, _
    Optional ByVal BcryptRandomAlgorithmId As cBcRandomAlgorithm = DefaultcBcRandomAlgorithm) _
    As Double
    On Error GoTo errRandomDbl
    Dim randomVal() As Byte
    Dim strRandom As String
    Dim ePos As Long
    Dim dblRandom As Double

    ClearLastError

    randomVal = RandomData(8, BcryptRandomAlgorithmId)
    RtlMoveMemory dblRandom, randomVal(LBound(randomVal)), 8

    If maxVal <= minVal Then
        RandomDbl = dblRandom
    Else
        strRandom = CStr(Abs(dblRandom))
        ePos = InStr(1, strRandom, "E")
        If ePos > 0 Then
            ' Eliminate the expo
            dblRandom = CDbl(left(strRandom, ePos - 1)) / 10#
        Else
            ' Move decimal to front of string
            ePos = InStr(1, strRandom, ".")
            If ePos > 1 Then
                dblRandom = CDbl(left(strRandom, 1) & "." & _
                                 Mid(strRandom, 1, ePos - 1) & _
                                 Mid(strRandom, ePos + 1)) / 10#
            Else
                dblRandom = CDbl(left(strRandom, 1) & "." & _
                                 Mid(strRandom, 1)) / 10#
            End If
        End If
        RandomDbl = CDbl((maxVal - minVal) * dblRandom + minVal)
    End If

doneRandomDbl:
    Exit Function
errRandomDbl:
    LogMessage "RandomDbl"
    Resume doneRandomDbl
End Function
