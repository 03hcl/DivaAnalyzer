Attribute VB_Name = "Arranging"
Option Explicit
Option Base 1

Public Sub �ؑ֌��ʂ��瑁�x�Ɛؑւ𕈖ʂɔ��f()
    
    Def.�}�[�N�����ݒ�
    Def.�����萔�ݒ�
    
    Set Def.�ؑ֌��ʃe�[�u�� = New SwitchingTable
    If Def.�ؑ֌��ʃe�[�u��.�I�u�W�F�N�g�ݒ�() < 0 Then
        GoTo �ؑ֌��ʃe�[�u���̐ݒ�Ɏ��s�����ꍇ
    End If
    
    Dim ���ʃe�[�u���� As String
    ���ʃe�[�u���� = Def.�ؑ֌��ʃe�[�u��.OwnTable.name
    ���ʃe�[�u���� = Left(���ʃe�[�u����, Application.WorksheetFunction.Max(0, InStr(���ʃe�[�u����, "_�ؑ�") - 1))
    If ���ʃe�[�u�������ݒ�(���ʃe�[�u����) < 0 Then
        Exit Sub
    End If
    
    ReDim ���x�蓮�w��t���O(Def.���ʃe�[�u��.�f�[�^�s��)
    ReDim MAX�\���t���O(Def.���ʃe�[�u��.�f�[�^�s��)
    
    Application.StatusBar = "�ؑւƑ��x�̏��𕈖ʃe�[�u���ɔ��f���Ă��܂�......"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim �őP�ؑ֏�� As Switching
    Set �őP�ؑ֏�� = Def.�ؑ֌��ʃe�[�u��.�őP�ؑ֏��擾(���ʃe�[�u��)
    Debug.Print �őP�ؑ֏��.�ؑ֕�����
    Dim MAX�\�� As OutputString
    Set MAX�\�� = �őP�ؑ֏��.�֑ؑ��x��񔽉f()
    Debug.Print MAX�\��.�\��������
    
    MsgBox "����ɏI�����܂����B", vbInformation
    
    Rescue
    
    Exit Sub
    
�ؑ֌��ʃe�[�u���̐ݒ�Ɏ��s�����ꍇ:
    
    MsgBox "ERR:�ؑ֌��ʃe�[�u���̐ݒ�Ɏ��s���܂����B" & vbCrLf & "�������I�����܂��B", vbCritical
    Exit Sub
    
End Sub

Public Sub ���x�Ɛؑւ�ʃV�[�g�ɏo��()
    
    Def.�}�[�N�����ݒ�
    Def.�����萔�ݒ�
    
    If Def.���ʃe�[�u���ݒ�() < 0 Then
        Exit Sub
    End If
    
    If Def.���x�ؑֈꗗ�e�[�u���ݒ�(Def.���ʃe�[�u��) < 0 Then
        Exit Sub
    End If
    
    If MsgBox("���݂̃e�[�u���ŏ������J�n���܂��B��낵���ł����H" & vbCrLf & _
        "�e�[�u����: " & Def.���ʃe�[�u��.OwnTable.name, vbOKCancel + vbInformation) <> vbOK Then
        MsgBox "�����𒆎~���܂����B", vbCritical
        Exit Sub
    End If
    
    Application.StatusBar = "���x�Ɛؑւ̈ꗗ�𕈖ʂ���V�[�g�ɏo�͂��Ă��܂�......"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Def.���x�ؑֈꗗ�e�[�u��.���x�ؑ֏��ǂݍ��� Def.���ʃe�[�u��
    
    MsgBox "����ɏI�����܂����B", vbInformation
    
    Rescue
    
End Sub

Public Sub ���x�Ɛؑւ𕈖ʂɔ��f()
    
    Def.�}�[�N�����ݒ�
    Def.�����萔�ݒ�
    
    Set Def.���x�ؑֈꗗ�e�[�u�� = New ElSwTable
    If Def.���x�ؑֈꗗ�e�[�u��.�I�u�W�F�N�g�ݒ�() < 0 Then
        GoTo ���x�ؑֈꗗ�e�[�u���̐ݒ�Ɏ��s�����ꍇ
    End If
    
    Dim ���ʃe�[�u���� As String
    ���ʃe�[�u���� = Def.���x�ؑֈꗗ�e�[�u��.OwnTable.name
    ���ʃe�[�u���� = Left(���ʃe�[�u����, Application.WorksheetFunction.Max(0, InStr(���ʃe�[�u����, "_���x�ؑփ��X�g") - 1))
    If ���ʃe�[�u�������ݒ�(���ʃe�[�u����) < 0 Then
        Exit Sub
    End If
    
    Application.StatusBar = "���x�Ɛؑւ̈ꗗ���V�[�g���畈�ʂɔ��f���Ă��܂�......"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Def.���x�ؑֈꗗ�e�[�u��.���x�ؑ֏�񏑂��o�� Def.���ʃe�[�u��
    
    MsgBox "����ɏI�����܂����B", vbInformation
    
    Rescue
    Exit Sub
    
���x�ؑֈꗗ�e�[�u���̐ݒ�Ɏ��s�����ꍇ:
    
    MsgBox "ERR:���x�Ɛؑւ̈ꗗ�e�[�u���̐ݒ�Ɏ��s���܂����B" & vbCrLf & "�������I�����܂��B", vbCritical
    
    Rescue
    Exit Sub
    
End Sub

Public Sub ���݂̏�Ԃł̃X�R�A�^���[�g������擾()
    
    Def.�}�[�N�����ݒ�
    Def.�����萔�ݒ�
    
    If Def.�X�R�A�^��͗p�萔�ݒ�() < 0 Then
        Exit Sub
    End If
    
    If Def.���ʃe�[�u���ݒ�() < 0 Then
        Exit Sub
    End If
    
    Dim �X�R�A�^���[�g������ As String
    �X�R�A�^���[�g������ = Analyzing.�X�R�A�^���[�g������擾()
    
    Dim cb As New dataobject
    cb.SetText �X�R�A�^���[�g������
    cb.PutInClipboard
    Set cb = Nothing
    
    MsgBox "���݂̏�Ԃł̃X�R�A�^���[�g������͈ȉ��̒ʂ�ł��B" & vbCrLf & _
        "(���̕�����̓N���b�v�{�[�h�ɃR�s�[����Ă��܂��B)" & vbCrLf & vbCrLf & _
        �X�R�A�^���[�g������, _
        vbInformation
    
    Rescue
    
    Exit Sub
End Sub
