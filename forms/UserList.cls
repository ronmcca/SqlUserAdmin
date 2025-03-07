Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim m_UserID As String

Private Sub Form_AfterInsert()
    On Error Resume Next
    m_UserID = Trim(Nz(Me.subUserID))
    Me.Form.Parent.CurrentUserID = m_UserID
End Sub

Private Sub Form_AfterUpdate()
    On Error Resume Next
    m_UserID = Trim(Nz(Me.subUserID))
    Me.Form.Parent.CurrentUserID = m_UserID
End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
    On Error Resume Next
    If CBool(Nz(Me.CK, False)) Then
        Response = acDataErrContinue
    End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error Resume Next
    If CBool(Nz(Me.CK, False)) Then
        Cancel = True
        MsgBox "Can not edit user name after it was added to SQL.", , "Cannot Edit"
    End If
End Sub

Private Sub Form_Delete(Cancel As Integer)
    On Error Resume Next
    If CBool(Nz(Me.CK, False)) Then
        Cancel = True
    End If
End Sub

Private Sub Form_Current()
    On Error Resume Next
    m_UserID = Trim(Nz(Me.subUserID))
    Me.Form.Parent.CurrentUserID = m_UserID
End Sub
