Attribute VB_Name = "modLog"
Option Compare Database
Option Explicit
'-----------------------------------------------
'           Logging functions
'-----------------------------------------------
Private Const ModuleName As String = "modUtility"

'-------------
' Log messages and errors from modules.             4/15/24
' Output using modUtility.Logger if it's initiated otherwise to the immediate window using debug.print
' Paramiters        Note: same order and type as ClassWMI.WriteLog
' sourceFunction    MODULENAME.FunctionName
' logMessage        Extra message after sourceFuntion do not include the err.Description
' errNumber         err.Number, else SystemErrorEnum.IgnoreError or SystemErrorEnum.Noerror for only sourceFunction and logMessage
' errDescription    err.Description
' saveError         Referece to a module variable to hold original error code
' saveErrorText     Referece to a module variable to hold original error text
' errLastDllError   err.LastDllError (only needed if function using external calls)
Public Sub Logger(ByVal SourceFunction As String, _
                  Optional ByVal LogMessage As String = vbNullString, _
                  Optional ByVal errNumber As Long = SystemErrorEnum.Noerror, _
                  Optional ByVal errDescription As String = vbNullString, _
                  Optional ByRef saveError As Long, _
                  Optional ByRef saveErrorText As String, _
                  Optional ByVal errLastDllError As Long = 0)
    On Error Resume Next

    If errNumber <> SystemErrorEnum.IgnoreError Then
        Debug.Print SourceFunction & ", " & _
                    LogMessage & _
                    IIf(errNumber = SystemErrorEnum.Noerror, vbNullString, _
                        ", error:( " & errNumber & _
                        " ) " & errDescription & _
                        IIf(errLastDllError = 0, vbNullString, _
                            ", DLL error:( " & errLastDllError & " )" _
                        ) _
                    )
        If Not IsMissing(saveError) And _
           Not IsMissing(saveErrorText) Then
            saveError = Err.Number
            saveErrorText = Err.Description
        End If
    End If
End Sub

'-------------
' Code to run before 1st start to genereate new application BookCipher for classAdm.AppKeyBookCipher
Public Function GenerateBookCipher() As String
    On Error Resume Next
    Dim Crypto As New ClassCrypt
    GenerateBookCipher = Crypto.Random(85)
    Set Crypto = Nothing
End Function
