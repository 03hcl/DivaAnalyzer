VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LogForm 
   Caption         =   "���O��� (�~�{�^���ňꎞ��~)"
   ClientHeight    =   7236
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   12828
   OleObjectBlob   =   "LogForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "LogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'Public �e�o�̓��O As OutputString
'Private ���O���x��() As Control
Private ���O���x���� As Long

Public Sub �o��(ByVal �o�͕����� As String, Optional ByVal �s�ԍ� As Long = -1)
    ���O���x���쐬 �s�ԍ�
    If �s�ԍ� > 0 Then
        Me.Controls(���O���x�����擾(�s�ԍ�)).Caption = �o�͕�����
        DoEvents
    End If
End Sub

Public Sub �t�H�[���o�͏I��()
    
End Sub

Private Function ���O���x���쐬(ByVal �s�ԍ� As Long)
    Dim �s As Long
    For �s = ���O���x���� + 1 To �s�ԍ�
'    For �s = UBound(���O���x��) + 1 To �s�ԍ�
'        ReDim ���O���x��(�s)
        ���O���x�������� �s
    Next �s
End Function

Private Function ���O���x��������(ByVal �s�ԍ� As Long)
    
    Dim ���݃R���g���[�� As Control
'    If ���O���x��(�s�ԍ�) Is Nothing Then
'        Set ���O���x��(�s�ԍ�) = Me.Controls.Add("Forms.Label.1", ���O���x�����擾(�s�ԍ�), False)
'    End If
    If ���O���x���� < �s�ԍ� Then
        ���O���x���� = �s�ԍ�
        Set ���݃R���g���[�� = Me.Controls.Add("Forms.Label.1", ���O���x�����擾(�s�ԍ�), False)
    End If
    
'    ���O���x��(�s�ԍ�).Width = 600
'    ���O���x��(�s�ԍ�).Height = 12
'    ���O���x��(�s�ԍ�).Left = 0
'    ���O���x��(�s�ԍ�).Top = ���O���x��(�s�ԍ�).Height * (�s�ԍ� - 1)
'
''    ���O���x��(�s�ԍ�).Font.Name = "Migu 1M"
'    ���O���x��(�s�ԍ�).Font.Name = ThisWorkbook.Theme.ThemeFontScheme.MajorFont(msoThemeEastAsian).Name
'    ���O���x��(�s�ԍ�).Font.Size = ���O���x��(�s�ԍ�).Height
'
'    ���O���x��(�s�ԍ�).BackColor = RGB(255, 255, 221)
'
'    ���O���x��(�s�ԍ�).Visible = True
    
    ���݃R���g���[��.Width = 600
    ���݃R���g���[��.Height = 12
    ���݃R���g���[��.Left = 0
    ���݃R���g���[��.Top = ���݃R���g���[��.Height * (�s�ԍ� - 1)
    
    ���݃R���g���[��.Font.name = ThisWorkbook.Theme.ThemeFontScheme.MajorFont(msoThemeEastAsian).name
    ���݃R���g���[��.Font.Size = ���݃R���g���[��.Height
    ���݃R���g���[��.Caption = ""
    
    ���݃R���g���[��.BackColor = RGB(255, 255, 221)
    
    ���݃R���g���[��.Visible = True
    
End Function

Private Function ���O���x�����擾(ByVal �s�ԍ� As Long) As String
    ���O���x�����擾 = "LogLabel" & �s�ԍ�
End Function

Private Sub UserForm_Initialize()
    ���O���x���� = 0
'    ReDim ���O���x��(1)
'    ���O���x�������� 1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        MsgBox "�ꎞ��~���Ă��܂��BOK�������ƍĊJ���܂��B"
        Cancel = 1
    End If
End Sub

Private Sub UserForm_Terminate()
'    Set �e�o�̓��O = Nothing
'    Erase ���O���x��
End Sub
