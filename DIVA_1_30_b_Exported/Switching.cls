VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Switching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private m_���ʃe�[�u�� As IChartTable

Private m_�ؑ֕����� As String
Private m_�ؑփt���O() As Boolean

Public Function �I�u�W�F�N�g�ݒ�(ByVal ���ʃe�[�u�� As IChartTable)
    
    Set m_���ʃe�[�u�� = ���ʃe�[�u��
    
    m_�ؑ֕����� = ""
    ReDim m_�ؑփt���O(m_���ʃe�[�u��.�f�[�^�s��)
    
    Dim �s As Long
    For �s = 1 To ���ʃe�[�u��.�f�[�^�s��
        Me.�ؑփt���O(�s) = False
    Next �s
    
End Function

Public Function �I�u�W�F�N�g���()
    
    Set m_���ʃe�[�u�� = Nothing
    
    m_�ؑ֕����� = ""
    ReDim m_�ؑփt���O(LBound(m_�ؑփt���O))
    m_�ؑփt���O(LBound(m_�ؑփt���O)) = False
    
End Function

Public Function �֑ؑ��x��񔽉f(Optional ByVal �����Čv�Z As Boolean = False) As OutputString
    
    If �����Čv�Z Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    
    Application.StatusBar = "�ؑւƑ��x�̏��𕈖ʃe�[�u���ɔ��f���Ă��܂�......"
    
    Dim �s As Long
    For �s = 1 To ���ʃe�[�u��.�f�[�^�s��
        ���ʃe�[�u��.�ؑ֔����(�s) = Me.�ؑփt���O(�s)
    Next
    
    ���ʃe�[�u��.�Čv�Z 1, ���ʃe�[�u��.�f�[�^�s��, �����Čv�Z
    
    Dim �J�n�s As Long
    �J�n�s = 1
    Do While ���ʃe�[�u��.�X�R�A��(�J�n�s) = ���ʃe�[�u��.�X�R�A��(1)
        �J�n�s = �J�n�s + 1
    Loop
    
    Dim �I���s As Long
    �I���s = ���ʃe�[�u��.�f�[�^�s��
    Do While ���ʃe�[�u��.�X�R�A��(�I���s - 1) = ���ʃe�[�u��.�X�R�A��(���ʃe�[�u��.�f�[�^�s��)
        �I���s = �I���s - 1
    Loop
    
    Set �֑ؑ��x��񔽉f = Analyzing.�w��͈͂̑��x��������(�J�n�s, �I���s, �����Čv�Z)
    
    Application.StatusBar = "�ؑւƑ��x�̏��𕈖ʃe�[�u���ɔ��f���܂����B"
    
End Function

' **************************************************************************************************
' �v���p�e�B
' **************************************************************************************************

Public Property Get �ؑ֕�����() As String
    �ؑ֕����� = m_�ؑ֕�����
End Property

Public Property Let �ؑ֕�����(ByVal value As String)
    m_�ؑ֕����� = value
End Property

Public Property Get �ؑփt���O(ByVal index As Long) As Boolean
    �ؑփt���O = m_�ؑփt���O(index)
End Property

Public Property Let �ؑփt���O(ByVal index As Long, ByVal value As Boolean)
    m_�ؑփt���O(index) = value
End Property

