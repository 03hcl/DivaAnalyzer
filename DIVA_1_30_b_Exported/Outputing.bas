Attribute VB_Name = "Outputing"
Option Explicit
Option Base 1

Public Sub ���ʃf�[�^���e�L�X�g�`���ŊO���o��()
    
    Def.�}�[�N�����ݒ�
    Def.�����萔�ݒ�
    
    If Def.���ʃe�[�u���ݒ�() < 0 Then
        Exit Sub
    End If
    
    If MsgBox("���݂̃e�[�u���ŏ������J�n���܂��B��낵���ł����H" & vbCrLf & _
        "�e�[�u����: " & Def.���ʃe�[�u��.OwnTable.name, vbOKCancel + vbInformation) <> vbOK Then
        MsgBox "�����𒆎~���܂����B", vbCritical
        Exit Sub
    End If
    
    Dim �o�͗p As ProcessLog
    
    Set �o�͗p = New ProcessLog
    If �o�͗p.�t�@�C���o�͊J�n("Chart_" & Def.���ʃe�[�u��.OwnTable.name & "_" & Format(Now, "yyyy-mm-dd-hhnnss") & ".txt") < 0 Then
        MsgBox "�t�@�C���ɏo�͂ł��܂���B"
        Exit Sub
    End If
    
    �o�͗p.�o�� "Difficulty=" & Def.���ʃe�[�u��.�����V�[�g.Names("Difficulty").RefersToRange.value
    �o�͗p.�o�� "Notes=" & Def.���ʃe�[�u��.�m�[�c�ԍ���(Def.���ʃe�[�u��.�f�[�^�s��)
    �o�͗p.�o�� "Duration=" & Def.���ʃe�[�u��.�����V�[�g.Names("Duration").RefersToRange.value
    
    Dim �t���[������ As Long
    
    For �t���[������ = Def.���ʃe�[�u��.�ő�SAD�t���[�� To Def.���ʃe�[�u��.�ŒxSAD�t���[��
        If Def.���ʃe�[�u��.�t���[������ʕ]��(�t���[������) <> "" Then
            �o�͗p.�o�� "FrameGapRating=" & �t���[������ & "," & Def.���ʃe�[�u��.�t���[������ʕ]��(�t���[������)
        End If
    Next
    
    Dim ���݃m�[�c�ԍ� As Long
    ���݃m�[�c�ԍ� = 0
    
    Dim �s As Long
    Dim �}�[�N As Long
    Dim �o�͕��ʕ����� As String
    
    For �s = 1 To Def.���ʃe�[�u��.�f�[�^�s��
        
        If Def.���ʃe�[�u��.�m�[�c�ԍ���(�s) > ���݃m�[�c�ԍ� Then
            ���݃m�[�c�ԍ� = Def.���ʃe�[�u��.�m�[�c�ԍ���(�s)
            �o�͕��ʕ����� = Def.���ʃe�[�u��.�t���[����(�s) & ","
            For �}�[�N = 1 To Def.�}�[�N��
                �o�͕��ʕ����� = �o�͕��ʕ����� & Def.���ʃe�[�u��.�m�[�c��(�}�[�N, �s) & ","
            Next �}�[�N
            For �}�[�N = 1 To Def.�X���C�h�}�[�N��
                �o�͕��ʕ����� = �o�͕��ʕ����� & Def.���ʃe�[�u��.�X���C�h�m�[�c��(�}�[�N, �s) & ","
            Next �}�[�N
'            �o�͗p.�o�� "Note" & ���݃m�[�c�ԍ� & "=" & �o�͕��ʕ�����
            �o�͗p.�o�� "Note=" & �o�͕��ʕ�����
        End If
        
    Next �s
    
    �o�͗p.�t�@�C���o�͏I��
    
    MsgBox "�O���o�͂��������܂����B"
    
    Rescue
    
End Sub

