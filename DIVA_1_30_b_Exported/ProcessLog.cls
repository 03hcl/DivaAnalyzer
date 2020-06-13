VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private m_�o�̓t�@�C�� As TextStream
Private m_�o�̓t�H�[�� As LogForm

Public Function �t�@�C���o�͊J�n(Optional ByVal �t�@�C���p�X As String = "") As Long
    
    Dim myFso As New FileSystemObject
    
    If �t�@�C���p�X = "" Then
        �t�@�C���p�X = "log_" & Format(Now, "yyyy-mm-dd-hhnnss") & ".txt"
    End If
        
    On Error GoTo ���O�t�@�C���̍쐬�Ɏ��s�����ꍇ
    Set m_�o�̓t�@�C�� = myFso.CreateTextFile(�t�@�C���p�X, True)
'    Set m_�o�̓t�@�C�� = myFso.OpenTextFile(�t�@�C���p�X, ForWriting, True)
    On Error GoTo 0
    
    If Is�t�@�C���o�� Then
        �t�@�C���p�X = myFso.GetFile(�t�@�C���p�X).Path
        Debug.Print "���O�� " & �t�@�C���p�X & " �ɏo�͂���܂��"
    Else
        �t�@�C���p�X = ""
        Debug.Print "���O�̓t�@�C���o�͂���܂���B"
    End If
    
    ThisWorkbook.Names("logFilePath").RefersToRange.value = �t�@�C���p�X
    
    Set myFso = Nothing
    
    �t�@�C���o�͊J�n = 0
    Exit Function
    
���O�t�@�C���̍쐬�Ɏ��s�����ꍇ:
    
    Dim answer As Long
    answer = MsgBox("���O�t�@�C�����쐬�ł��܂���B" & vbCrLf & "���O�̃t�@�C���o�͂��s�킸�ɏ����𑱍s���܂����H", vbOKCancel + vbExclamation)
    
    If answer = vbOK Then
        Set m_�o�̓t�@�C�� = Nothing
        Resume Next
    End If
    
    �t�@�C���o�͊J�n = -1
    Exit Function
    
End Function

Public Sub �t�@�C���o�͏I��()
    If Me.Is�t�@�C���o�� Then
        m_�o�̓t�@�C��.Close
        Set m_�o�̓t�@�C�� = Nothing
    End If
End Sub

Public Function �t�H�[���o�͊J�n()
    Set m_�o�̓t�H�[�� = New LogForm
    m_�o�̓t�H�[��.Show
End Function

Public Sub �t�H�[���o�͏I��()
    If Me.Is�t�H�[���o�� Then
        Unload m_�o�̓t�H�[��
        Set m_�o�̓t�H�[�� = Nothing
    End If
End Sub

Public Sub �o��(ByVal �o�͕����� As String, Optional ByVal is�X�e�[�^�X�o�[�o�� As Boolean = True, Optional ByVal �o�͍s�ԍ� As Long = -1)
    Debug.Print �o�͕�����
    If is�X�e�[�^�X�o�[�o�� Then
        Application.StatusBar = Replace(Replace(�o�͕�����, vbCrLf, " "), vbTab, " ")
    End If
    If Me.Is�t�@�C���o�� Then
        m_�o�̓t�@�C��.WriteLine �o�͕�����
    End If
    If Is�t�H�[���o�� Then
        m_�o�̓t�H�[��.�o�� �o�͕�����, �o�͍s�ԍ�
    End If
End Sub

Public Sub �t�H�[��������폜(ByVal �s�ԍ� As Long)
    If Is�t�H�[���o�� Then
        m_�o�̓t�H�[��.�o�� "", �s�ԍ�
    End If
End Sub

Public Property Get Is�t�@�C���o��() As Boolean
    Is�t�@�C���o�� = (Not m_�o�̓t�@�C�� Is Nothing)
End Property

Public Property Get Is�t�H�[���o��() As Boolean
    Is�t�H�[���o�� = (Not m_�o�̓t�H�[�� Is Nothing)
End Property

Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()
    Me.�t�@�C���o�͏I��
    Me.�t�H�[���o�͏I��
End Sub