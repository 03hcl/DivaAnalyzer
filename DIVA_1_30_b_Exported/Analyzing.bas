Attribute VB_Name = "Analyzing"
Option Explicit
Option Base 1

#Const �ؑ֏ڍ׃��O = False
#Const �X�R�A�^�ڍ׃��O = False
#Const �X�R�A�^�ʏ탍�O = True
Private �X�R�A�^�l�X�g�J�E���^ As Long

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Sub �X�R�A�^���()
    
    Application.StatusBar = "�X�R�A�^��͂��J�n���܂�......"
    
    Application.Calculation = xlCalculationAutomatic
    DoEvents
    
    #If �X�R�A�^�ڍ׃��O Then
        �ڍ׃��O�o�� "�X�R�A�^��͏����ݒ���s" & vbTab & "�J�n"
    #End If
    
    If Def.�X�R�A�^��͏����ݒ���s() <> 0 Then
        Rescue
        Exit Sub
    End If
    
    #If �X�R�A�^�ڍ׃��O Then
        �ڍ׃��O�o�� "�X�R�A�^��͏����ݒ���s" & vbTab & "�I��"
    #End If
    
    Def.�X�R�A�^�e�[�u��.Is�ȈՔ� = (MsgBox("�ȈՔłɂ��܂����H" & vbCrLf & _
        "(�\���ɗ��ꂽ�����̃��[�g�𕪗����Č������܂��񂪁A���x�ؑֈꗗ�e�[�u���̏o�͂Ɏ��Ԃ��g���܂���B)", _
        vbYesNo + vbInformation) = vbYes)
'    Def.�X�R�A�^�e�[�u��.Is�ȈՔ� = True
    
    Dim ��ʍX�V As Boolean
    Dim �����Čv�Z As Boolean
    
    Dim �J�n���� As Long
    Dim �I������ As Long
    
    ��ʍX�V = MsgBox("���s���[�g���Ƃ̉�ʍX�V���s���܂����H", vbYesNo + vbInformation) = vbYes
    �����Čv�Z = MsgBox("�u�b�N�̎����Čv�Z���s���܂����H", vbYesNo + vbInformation) = vbYes
    
    �J�n���� = Def.GetTickCount
    
    #If �X�R�A�^�ʏ탍�O Then
        Def.�������O.�t�H�[���o�͊J�n
    #End If
    
    �����X�R�A�^���[�g���� ��ʍX�V, �����Čv�Z ', 670
    
    #If �X�R�A�^�ʏ탍�O Then
        Def.�������O.�t�H�[���o�͏I��
    #End If
    
    �I������ = Def.GetTickCount
    
    #If �X�R�A�^�ڍ׃��O Then
        �ڍ׃��O�o�� "�X�R�A�^��͏I���ݒ���s" & vbTab & "�J�n"
    #End If
    
    Def.�X�R�A�^��͏I���������s �J�n����, �I������, ��ʍX�V, �����Čv�Z
    
    #If �X�R�A�^�ڍ׃��O Then
        �ڍ׃��O�o�� "�X�R�A�^��͏I���ݒ���s" & vbTab & "�I��"
    #End If
    
    MsgBox "����ɏI�����܂����B���肪�ف[�B" & vbCrLf & "��͎���: " & CDbl(�I������ - �J�n����) / 1000 & " �b"
    
    Rescue
    
End Sub

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Sub �ؑ։��()
    
    Application.StatusBar = "�ؑ։�͂��J�n���܂�......"
    
    Application.Calculation = xlCalculationAutomatic
    DoEvents
    
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�ؑ։�͏����ݒ���s" & vbTab & "�J�n"
    #End If
    
    If Def.�ؑ։�͏����ݒ���s() <> 0 Then
        Rescue
        Exit Sub
    End If
    
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�ؑ։�͏����ݒ���s" & vbTab & "�I��"
    #End If
    
    Dim ��ʍX�V As Boolean
    Dim �����Čv�Z As Boolean
    
    Dim �J�n���� As Long
    Dim �I������ As Long
    
    ��ʍX�V = MsgBox("���s���[�g���Ƃ̉�ʍX�V���s���܂����H", vbYesNo + vbInformation) = vbYes
    �����Čv�Z = MsgBox("�u�b�N�̎����Čv�Z���s���܂����H", vbYesNo + vbInformation) = vbYes
    
    �J�n���� = Def.GetTickCount
    
    �����ؑ֔��� ��ʍX�V, �����Čv�Z
    
    �I������ = Def.GetTickCount
    
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�ؑ։�͏I���ݒ���s" & vbTab & "�J�n"
    #End If
    
    Def.�ؑ։�͏I���ݒ���s �J�n����, �I������, ��ʍX�V, �����Čv�Z
    
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�ؑ։�͏I���ݒ���s" & vbTab & "�I��"
    #End If
    
    MsgBox "����ɏI�����܂����B���肪�ف[�B"
    
    Rescue
    
End Sub

' ======================================================================================================================================================================================================
'
' �� (�J�n�s - 1) �s�ڂ��]���ΏۂɂȂ�܂��B�� �J�n�s = 1 �̓_��
' ======================================================================================================================================================================================================

Public Function �����X�R�A�^���[�g����( _
    Optional ByVal ��ʍX�V As Boolean = False, _
    Optional ByVal �����Čv�Z As Boolean = False, _
    Optional ByVal ����J�n�s As Long = 1, _
    Optional ByVal ����I���s As Long = -1, _
    Optional ByVal ���t���[�h As Boolean = False)
    
    ' �J�n���� -------------------------------------------------------------------------------------
    
    #If �X�R�A�^�ڍ׃��O Then
        �ڍ׃��O�o�� "�����X�R�A�^���[�g����" & vbTab & "�����ݒ�J�n", True
    #End If
    
    Application.ScreenUpdating = False
    
    If �����Čv�Z Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    
    If ����I���s = -1 Then
        ����I���s = Def.���ʃe�[�u��.�f�[�^�s��
    End If
    
    ' ���x�𔽉f���Ă��̊m��X�R�A��z��Ɋi�[
    
    Def.���x�ؑֈꗗ�e�[�u��.���x�ؑ֏�񏑂��o�� Def.���ʃe�[�u��
    Def.���ʃe�[�u��.�Čv�Z ����J�n�s, ����I���s, �����Čv�Z
    Dim �s�X�R�A() As Long
    �s�X�R�A = �s�X�R�A�ݒ�(����J�n�s, ����I���s)
    
    ' ���C������ -----------------------------------------------------------------------------------
    
    #If �X�R�A�^�ڍ׃��O Then
        �ڍ׃��O�o�� "�����X�R�A�^���[�g����" & vbTab & "�J�n", True
    #End If
    
    Application.StatusBar = "�X�R�A�^��͂��J�n���܂�......"
    
    �X�R�A�^�l�X�g�J�E���^ = -1
    �w��͈͂̃��[�g���� ����J�n�s, ����I���s, �s�X�R�A, ���t���[�h, ��ʍX�V, �����Čv�Z
    Def.�X�R�A�^�e�[�u��.�g�ݍ��킹���[�g�o�� Def.���ʃe�[�u��, Def.���x�ؑֈꗗ�e�[�u��, ���t���[�h
    
    ' �I������ ------------------------------------------------------------------------------------
    
    Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True
    
    #If �X�R�A�^�ڍ׃��O Then
        �ڍ׃��O�o�� "�����X�R�A�^���[�g����" & vbTab & "�I��", True
    #End If
    
End Function

Private Function �s�X�R�A�ݒ�(ByVal �J�n�s As Long, ByVal �I���s As Long) As Long()
    Dim �s�X�R�A() As Long
    ReDim �s�X�R�A(Def.���ʃe�[�u��.�f�[�^�s��)
    Dim �s As Long
    For �s = �J�n�s To �I���s
        �s�X�R�A(�s) = ���ݍs�X�R�A�擾(�s)
    Next �s
    �s�X�R�A�ݒ� = �s�X�R�A
End Function

Private Function ���ݍs�X�R�A�擾(ByVal �s As Long) As Long
    If �s = 1 Then
        ���ݍs�X�R�A�擾 = Def.���ʃe�[�u��.�X�R�A��(�s)
    ElseIf Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�s - 1) = 0 Then
        ���ݍs�X�R�A�擾 = Def.���ʃe�[�u��.�X�R�A��(�s)
    ElseIf Def.���ʃe�[�u��.�z�[���h�{�[�i�X��(�s) > 0 Then
        ���ݍs�X�R�A�擾 = Def.���ʃe�[�u��.�X�R�A��(�s)
    Else
        ���ݍs�X�R�A�擾 = 0
    End If
End Function

' ======================================================================================================================================================================================================
'
' �߂�l�̓��[�g�����������ꍇ�̂݁A���̃��[�g�̍s�X�R�A�ƂȂ�
' ======================================================================================================================================================================================================

Private Function �w��͈͂̃��[�g����( _
    ByVal �J�n�s As Long, _
    ByVal �I���s As Long, _
    ByRef ���݃��[�g�s�X�R�A() As Long, _
    Optional ByVal ���t���[�h As Boolean = False, _
    Optional ByVal ��ʍX�V As Boolean = False, _
    Optional ByVal �����Čv�Z As Boolean = False, _
    Optional ByVal �e���J�n�s As Long = 0) _
    As Def.�s�X�R�A���
'    Optional ByRef ��r���[�g�s�X�R�A As Def.�s�X�R�A���, _

    #If �X�R�A�^�ڍ׃��O Then
        �ڍ׃��O�o�� "�w��͈͂̃��[�g����" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�J�n", True
    #End If
    
    #If �X�R�A�^�ʏ탍�O Then
        �X�R�A�^�l�X�g�J�E���^ = �X�R�A�^�l�X�g�J�E���^ + 1
        If �X�R�A�^�l�X�g�J�E���^ > 0 Then
            �ڍ׃��O�o�� String(�X�R�A�^�l�X�g�J�E���^ - 1, "��") & "��" & �J�n�s & ",", True, �X�R�A�^�l�X�g�J�E���^ + 1
        Else
            �ڍ׃��O�o�� �J�n�s & ",", True, �X�R�A�^�l�X�g�J�E���^ + 1
        End If
    #End If
    
    Dim �Ώۃ��[�g�s�X�R�A As Def.�s�X�R�A���
    �Ώۃ��[�g�s�X�R�A = �����[�g�L������Əo��(�J�n�s, �I���s, ���݃��[�g�s�X�R�A, ���t���[�h, ��ʍX�V, �����Čv�Z, �e���J�n�s)
    
    If �Ώۃ��[�g�s�X�R�A.�ő�e���s = 0 Then
        �Ώۃ��[�g�s�X�R�A.�s�X�R�A = ���݃��[�g�s�X�R�A
    End If
    
    Dim is�z�[���h�J�n(Def.�}�[�N��) As Boolean
    
    Dim ���O�z�[���h�J�n�t���[�� As Long
    Dim �ؑ֍Čv�Z�J�n�s As Long
    Dim �J�n�����ꋖ�e�t���[�� As Long
    
    Dim is���[�g�J�n�s As Boolean
    
    Dim �ؑ։\�� As Boolean
    
    Dim �s As Long
    Dim �}�[�N As Long
    
    For �s = �J�n�s To �I���s
        
        Def.���ʃe�[�u��.�Čv�Z �s, , �����Čv�Z
        
        ' ���݂̔�r�Ώۂ̍s�X�R�A�����ȏ�X�R�A���������ꍇ�͏I��
        Dim ���݃X�R�A As Long
        ���݃X�R�A = ���ݍs�X�R�A�擾(�s)
        If ���݃��[�g�s�X�R�A(�s) > 0 And ���݃X�R�A > 0 Then
            If ���݃X�R�A < ���݃��[�g�s�X�R�A(�s) + Def.�X�R�A�^�X�L�b�v�_ Then
                Exit For
            End If
        End If
        
        '���t���[�h�łȂ��ꍇ�A���C�t��0�Ȃ�I��
        If (Not ���t���[�h) And Def.���ʃe�[�u��.���C�t��(�s) = 0 Then
            Exit For
        End If
        
        '50�R���{�ȏ�Ń��C�t���ő�Ȃ�A�ȑO�ɐݒ肳�ꂽ���݃��[�g�s�X�R�A�Ɠ����X�R�A�㏸�y�[�X�ɂȂ�̂ŏI��
        If Not (Def.�X�R�A�^�e�[�u��.Is�ȈՔ�) And �e���J�n�s > 0 Then
            If Def.���ʃe�[�u��.�R���{��(�s) >= 50 And Def.���ʃe�[�u��.���C�t��(�s) = Def.�ő僉�C�t�� Then
                Exit For
            End If
        End If
        
        ' ���[�g�J�n�ʒu�ɂȂ�邩�ǂ����𔻒�
        is���[�g�J�n�s = False
        For �}�[�N = 1 To Def.�}�[�N��
            If Def.���ʃe�[�u��.�m�[�c��(�}�[�N, �s) = Def.HOLD���� Then
                is�z�[���h�J�n(�}�[�N) = True
                is���[�g�J�n�s = True
            Else
                is�z�[���h�J�n(�}�[�N) = False
            End If
        Next �}�[�N
        
        If is���[�g�J�n�s Then
            ' ���[�g�J�n�ʒu�ɂȂ��ꍇ�̏���
            
            ' �J�n�s�̐ؑ֐ݒ�(�ؑւ͋����ύX)
            �ؑ։\�� = False
            For �}�[�N = 1 To Def.�}�[�N��
                If Def.���ʃe�[�u��.�z�[���h�ʊJ�n�t���[����(�}�[�N, �s - 1) > 0 Then
                    �ؑ։\�� = True
                End If
            Next
            If �ؑ։\�� And Not (Def.���ʃe�[�u��.�z�[���h�{�[�i�X��(�s) > 0) Then
                Def.���ʃe�[�u��.�ؑ֔����(�s) = True
            End If
            
            ' ���O�u���b�N�̐ؑւƑ��x�̍Đݒ�
            �ؑ֍Čv�Z�J�n�s = �w��s���O�܂ł̐ؑ֍Čv�Z(�J�n�s, �s, ��ʍX�V, �����Čv�Z)
            
            ' �J�n�s�ƂȂ錻�݂̍s�̃t���[���O���ꋖ�e�t���[����ݒ肵�āA���x�t���[���ƕ]�������Z�b�g
            ' �����O��C-Sd��MAX�̃��[�g�ȂǁA�ő�COOL��MAX�ɂȂ�Ȃ��ꍇ�̓��Z�b�g���Ȃ�
            �J�n�����ꋖ�e�t���[�� = Def.���ʃe�[�u��.�f�t�H���g�����ꋖ�e�t���[��
            If Def.���ʃe�[�u��.�z�[���h�{�[�i�X��(�s) > 0 Then
                ' MAX������ꍇ�̂݃f�t�H���g����ύX
                If Def.���ʃe�[�u��.�z�[���h�t���[����(�s) > Def.��MAX�t���[���ő�l Then
                    �J�n�����ꋖ�e�t���[�� = Application.WorksheetFunction.Max(�J�n�����ꋖ�e�t���[��, _
                        Def.��MAX�t���[���ő�l - Def.���ʃe�[�u��.�z�[���h�t���[����(�s) + Def.���ʃe�[�u��.���x�t���[����(�s) + 1)
                End If
            End If
            
            ���x�t���[���ƕ]���̉��� �s, , Not (�J�n�����ꋖ�e�t���[�� > Def.���ʃe�[�u��.�ő�COOL�t���[��)
            Def.���ʃe�[�u��.�Čv�Z �s, , �����Čv�Z
            
            '���ݍs�̎��̍s����̃��[�g�T��
            If �e���J�n�s = 0 Then
                �w��s����̃��[�g�T�� �s, �I���s, is�z�[���h�J�n, �Ώۃ��[�g�s�X�R�A.�s�X�R�A, �J�n�����ꋖ�e�t���[��, , ���t���[�h, ��ʍX�V, �����Čv�Z, �ؑ֍Čv�Z�J�n�s
            Else
                �w��s����̃��[�g�T�� �s, �I���s, is�z�[���h�J�n, �Ώۃ��[�g�s�X�R�A.�s�X�R�A, �J�n�����ꋖ�e�t���[��, , ���t���[�h, ��ʍX�V, �����Čv�Z, �e���J�n�s
            End If
            
            '�ؑւƑ��x�����ɖ߂�
            Def.���x�ؑֈꗗ�e�[�u��.���x�ؑ֏�񏑂��o�� Def.���ʃe�[�u��, �ؑ֍Čv�Z�J�n�s, �s, False
            Def.���ʃe�[�u��.�Čv�Z �ؑ֍Čv�Z�J�n�s, �s, �����Čv�Z
            
        End If
        
    Next �s
    
    ' �߂�l��ݒ�
    �w��͈͂̃��[�g���� = �Ώۃ��[�g�s�X�R�A
    
    ' ���x���Z�b�g�͂Ȃ��Ă����v?
    
    #If �X�R�A�^�ڍ׃��O Then
        �ڍ׃��O�o�� "�w��͈͂̃��[�g����" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�I��", True
    #End If
    
    #If �X�R�A�^�ʏ탍�O Then
        Def.�������O.�t�H�[��������폜 �X�R�A�^�l�X�g�J�E���^ + 1
        �X�R�A�^�l�X�g�J�E���^ = �X�R�A�^�l�X�g�J�E���^ - 1
    #End If
    
End Function

Private Function �����[�g�L������Əo��( _
    ByVal �J�n�s As Long, _
    ByVal �I���s As Long, _
    ByRef ���݃��[�g�s�X�R�A() As Long, _
    Optional ByVal ���t���[�h As Boolean = False, _
    Optional ByVal ��ʍX�V As Boolean = False, _
    Optional ByVal �����Čv�Z As Boolean = False, _
    Optional ByVal �e���J�n�s As Long = 0) _
    As Def.�s�X�R�A���
'    Optional ByRef ��r���[�g�s�X�R�A As Def.�s�X�R�A���, _

    #If �X�R�A�^�ڍ׃��O Then
        �ڍ׃��O�o�� "�����[�g�L������Əo��" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�J�n", True
    #End If
    
    Dim is�L�����[�g As Boolean
    is�L�����[�g = True
    
    Dim �s As Long
    
    For �s = �J�n�s To �I���s
        
        Def.���ʃe�[�u��.�Čv�Z �s, , �����Čv�Z
        
        If is�L�����[�g Then
            
            ' ���݂̔�r�Ώۂ̍s�X�R�A�����ȏ�X�R�A���������ꍇ�͍Ō�܂Ōv�Z�����ɏI��
            Dim ���݃X�R�A As Long
            ���݃X�R�A = ���ݍs�X�R�A�擾(�s)
            If ���݃��[�g�s�X�R�A(�s) > 0 And ���݃X�R�A > 0 Then
                If ���݃X�R�A < ���݃��[�g�s�X�R�A(�s) + Def.�X�R�A�^�X�L�b�v�_ Then
                    is�L�����[�g = False
                    Exit For
                End If
            End If
        
            '���t���[�h�łȂ��ꍇ�A���C�t��0�Ȃ�I��
            If (Not ���t���[�h) And Def.���ʃe�[�u��.���C�t��(�s) = 0 Then
                is�L�����[�g = False
                Exit For
            End If
            
            '50�R���{�ȏ�Ń��C�t���ő�Ȃ�A�ȑO�ɐݒ肳�ꂽ���݃��[�g�s�X�R�A������邱�Ƃ͂Ȃ��̂Ŋm��
'            If �s > ��r���[�g�s�X�R�A.�ő�e���s Then
            If Not (Def.�X�R�A�^�e�[�u��.Is�ȈՔ�) And �e���J�n�s > 0 Then
                If Def.���ʃe�[�u��.�R���{��(�s) >= 50 And Def.���ʃe�[�u��.���C�t��(�s) = Def.�ő僉�C�t�� Then
                    Exit For
                End If
            End If
            
        End If
        
    Next �s
    
    Dim ���U���g As Def.���ʃZ�b�g
    
    ' ���t���[�h�łȂ��ꍇ�AMISS�~TAKE(�N���A�Q�[�W�������Ȃ�)�Ȃ�I��
    If is�L�����[�g Then
        'Def.���ʃe�[�u��.�Čv�Z �s + 1, �I���s, �����Čv�Z
        Def.���ʃe�[�u��.�Čv�Z �s + 1, Def.���ʃe�[�u��.�f�[�^�s��, �����Čv�Z
        ���U���g = Def.���ʃe�[�u��.���U���g�Čv�Z()
        If (Not ���t���[�h) And ���U���g.�N���A�����N = Def.MISSTAKE���� Then
            is�L�����[�g = False
        End If
    End If
    
    Dim ���� As Def.�s�X�R�A���
    
    ' �߂�l�ƂȂ郋�[�g�s�X�R�A��ݒ�
    If is�L�����[�g Then
        
        ����.�ő�e���s = �ؑ։e���I���s�擾(�s)
        
        ����.�s�X�R�A = �s�X�R�A�ݒ�(�J�n�s, �I���s)
        
        �����[�g�L������Əo�� = ����
    Else
'        �����[�g�L������Əo�� = ���݃��[�g�s�X�R�A
'        �����[�g�L������Əo�� = ����
        #If �X�R�A�^�ڍ׃��O Then
            �ڍ׃��O�o�� "�����[�g�L������Əo��" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�I��(����)", True
        #End If
        Exit Function
    End If
    
    ' �L�����[�g�łȂ��ꍇ�͂����ŏI��
    ' �L�����[�g�ł���ꍇ�̂ݏo�͏���
    
    #If �X�R�A�^�ڍ׃��O Then
        �ڍ׃��O�o�� "�����[�g�L������Əo��" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�o�͏����J�n", True
    #End If
    
    Dim �X�R�A�^���[�g������ As String
    �X�R�A�^���[�g������ = �X�R�A�^���[�g������擾()
    
    ' �X�R�A�^�e�[�u��(�ƐV�������x�ؑֈꗗ�e�[�u��)�Ɍ��ʏo��
    If �e���J�n�s = 0 Then
        �e���J�n�s = �J�n�s
    End If
    
    Def.�X�R�A�^�e�[�u��.���݃��[�g�o�� Def.���ʃe�[�u��, �X�R�A�^���[�g������, �e���J�n�s, ����.�ő�e���s, ���U���g.�B����, ���U���g.�X�R�A
    
    #If �X�R�A�^�ڍ׃��O Then
        �ڍ׃��O�o�� "�����[�g�L������Əo��" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�I��(�L��)", True
    #End If
    
End Function

Public Function �X�R�A�^���[�g������擾() As String
    
    Dim �X�R�A�^������ As String
    Dim ���݃��[�g�������� As String
    �X�R�A�^������ = ""
    
    Dim is�z�[���h�J�n(Def.�}�[�N��) As Boolean
    
    Dim is���[�g�J�n�s As Boolean
    
    Dim �s As Long
    Dim �}�[�N As Long
    
    Dim is���[�g�� As Boolean
    Dim is���[�g�I�� As Boolean
    is���[�g�� = False
    is���[�g�I�� = False
    
    Dim is�����[�g�� As Boolean
    is�����[�g�� = False
        
    Dim ���O�z�[���h�J�n�s As Long
    Dim w�� As Long
    Dim is�z�[���h��(Def.�}�[�N��) As Boolean
    
    Dim is�z�[���h�J�n�s As Boolean
    
    Dim �J�n�t���[������ As Long
    Dim �I���t���[������ As Long
    Dim �J�n�t���[���]�T As Long
    Dim �I���t���[���]�T As Long
    
    For �s = 1 To Def.���ʃe�[�u��.�f�[�^�s��
        
        ' ���[�g�I�����ǂ����̔���
        If Def.���ʃe�[�u��.�z�[���h�{�[�i�X��(�s) > 0 Then
            If is���[�g�� Then
                is���[�g�I�� = True
            ElseIf Def.���ʃe�[�u��.�z�[���h�t���[����(�s) > Def.��MAX�t���[���ő�l Then
                ' ���[�g���łȂ��Ă��]�T�̂Ȃ�MAX�̏ꍇ�͑k���ă��[�g�Ƃ��ĕ�����ǉ�
                If Def.���ʃe�[�u��.�z�[���h�I���������������(�s) And _
                    Def.���ʃe�[�u��.�z�[���h�t���[����(�s) < Def.��MAX�t���[���ő�l + 1 + Def.CC�ő�]�T�t���[���� Then
                    ���[�g�J�n���ϐ��ݒ� �s - 1, True, ���O�z�[���h�J�n�s, ���݃��[�g��������, w��, is�z�[���h��
                    is���[�g�I�� = True
                End If
            End If
        End If
        
        '���[�g�I���̏ꍇ�̕����񏈗�
        If is���[�g�I�� Then
            
            ���݃��[�g�������� = ���݃��[�g�������� & " ���y"
            
            Dim �{�^���� As Long
            �{�^���� = 0
            For �}�[�N = 1 To �}�[�N��
                If is�z�[���h��(�}�[�N) Then
                    �{�^���� = �{�^���� + 1
                End If
            Next �}�[�N
            If �{�^���� > 1 Then
                ���݃��[�g�������� = ���݃��[�g�������� & �{�^����
            End If
            
            If Def.���ʃe�[�u��.�z�[���h�t���[����(�s) > Def.��MAX�t���[���ő�l Then
                ���݃��[�g�������� = ���݃��[�g�������� & "MAX�z"
            Else
                ���݃��[�g�������� = ���݃��[�g�������� & "HOLD�z"
            End If
            
            Dim ���m�[�c�s As Long
            Dim �����m�[�c������ As String
            �����m�[�c������ = ""
            For ���m�[�c�s = �s To Def.���ʃe�[�u��.�f�[�^�s��
                For �}�[�N = 1 To Def.�}�[�N��
                    If Def.���ʃe�[�u��.�m�[�c��(�}�[�N, ���m�[�c�s) <> "" Then
                        �����m�[�c������ = �����m�[�c������ & Def.�}�[�N����(�}�[�N)
                    End If
                Next �}�[�N
                If �����m�[�c������ <> "" Then
                    Exit For
                End If
            Next
            If �����m�[�c������ <> "" Then
                ���݃��[�g�������� = ���݃��[�g�������� & "�� " & �m�[�c�ԍ�������̎擾(�s) & �����m�[�c������
            End If
            
            If Def.���ʃe�[�u��.�z�[���h�t���[����(�s) > Def.��MAX�t���[���ő�l And _
                Def.���ʃe�[�u��.�z�[���h�t���[����(�s) < Def.��MAX�t���[���ő�l + 1 + Def.CC�ő�]�T�t���[���� _
                And Def.���ʃe�[�u��.�z�[���h�I���������������(�s) = True Then
                
                '�]�T�̂Ȃ�MAX�ł���ꍇ�͏ڍׂ����݃��[�g��������ɒǋL
                ���݃��[�g�������� = ���݃��[�g�������� & "�s"
                
                �J�n�t���[������ = Def.���ʃe�[�u��.���x�t���[����(���O�z�[���h�J�n�s)
                �I���t���[������ = Def.���ʃe�[�u��.���x�t���[����(�s)
                
                If �J�n�t���[������ < 0 Then
                    �J�n�t���[���]�T = -1
                ElseIf �J�n�t���[������ > 0 Then
                    If Def.���ʃe�[�u��.�]����(���O�z�[���h�J�n�s) <> Def.���ʃe�[�u��.�t���[������ʕ]��(0) Then
                        ���݃��[�g�������� = ���݃��[�g�������� & "�x"
                    End If
                    �J�n�t���[���]�T = 0
                Else
                    �J�n�t���[���]�T = 0
                End If
                ���݃��[�g�������� = ���݃��[�g�������� & Def.�]������(Def.���ʃe�[�u��.�]����(���O�z�[���h�J�n�s)) & "-"
                Do Until Def.���ʃe�[�u��.�t���[������ʕ]��(�J�n�t���[������ - �J�n�t���[���]�T) = Def.���ʃe�[�u��.�t���[������ʕ]��(�J�n�t���[������)
                    �J�n�t���[���]�T = �J�n�t���[���]�T + 1
                Loop
                
                If �I���t���[������ < 0 Then
                    If Def.���ʃe�[�u��.�]����(�s) <> Def.���ʃe�[�u��.�t���[������ʕ]��(0) Then
                        ���݃��[�g�������� = ���݃��[�g�������� & "��"
                    End If
                    �I���t���[���]�T = 0
                ElseIf �I���t���[������ > 0 Then
                    �I���t���[���]�T = -1
                Else
                    �I���t���[���]�T = 0
                End If
                ���݃��[�g�������� = ���݃��[�g�������� & Def.�]������(Def.���ʃe�[�u��.�]����(�s)) & ",�P�\:"
                Do Until Def.���ʃe�[�u��.�t���[������ʕ]��(�I���t���[������ + �I���t���[���]�T) = Def.���ʃe�[�u��.�t���[������ʕ]��(�I���t���[������)
                    �I���t���[���]�T = �I���t���[���]�T + 1
                Loop
                
                ���݃��[�g�������� = ���݃��[�g�������� & (�J�n�t���[���]�T + �I���t���[���]�T) & "�t"
                
            End If
            
            If w�� > 0 Then
                ���݃��[�g�������� = ���݃��[�g�������� & " (W" & w�� & ")"
            End If
            
            �X�R�A�^������ = Def.������A��(�X�R�A�^������, ���݃��[�g��������, vbCrLf)
            
            is���[�g�� = False
            is���[�g�I�� = False
            
            is�����[�g�� = False
            
        End If
        
        ' �ؑւ�����ꍇ�̕����񏈗�
        If Def.���ʃe�[�u��.�ؑ֔����(�s) And (Not is���[�g��) Then
            �X�R�A�^������ = Def.������A��(�X�R�A�^������, �ؑ֕�����擾(�s), vbCrLf)
        End If
        
        '�����[�g���̏ꍇ
        If is�����[�g�� Then
            
            is�z�[���h�J�n�s = False
            For �}�[�N = 1 To Def.�}�[�N��
                If Def.���ʃe�[�u��.�m�[�c��(�}�[�N, �s) = Def.HOLD���� Then
                    is�z�[���h�J�n�s = True
                    Exit For
                End If
            Next �}�[�N
            
            If is�z�[���h�J�n�s Then
                If Def.���ʃe�[�u��.�z�[���h�t���[����(�s) > Def.��MAX�t���[���ő�l - 1 - Def.CC�ő�]�T�t���[���� Then
                    is���[�g�� = True
                    ���[�g�J�n���ϐ��ݒ� �s - 1, True, ���O�z�[���h�J�n�s, ���݃��[�g��������, w��, is�z�[���h��
                End If
                is�����[�g�� = False
            End If
            
        End If
        
        '���[�g�J�n���ǂ����̔���
        If Not is���[�g�� Then
            
            If Def.���ʃe�[�u��.�]����(�s) = Def.��WRONG���� Or Def.���ʃe�[�u��.�]����(�s) = Def.WORST���� Then
                
                '���݂̍s����WRONG��WORST�������ꍇ��(�J�n�s��k���Č�������)���[�g�J�n
                is���[�g�� = True
                ���[�g�J�n���ϐ��ݒ� �s, True, ���O�z�[���h�J�n�s, ���݃��[�g��������, w��, is�z�[���h��
                is�����[�g�� = False
                
            ElseIf Def.���ʃe�[�u��.�]����(�s) <> Def.���ʃe�[�u��.�t���[������ʕ]��(0) Then
'                Def.���ʃe�[�u��.�z�[���h�t���[����(�s) < Def.��MAX�t���[���ő�l + 1 + Def.CC�ő�]�T�t���[����
                
                If Def.���ʃe�[�u��.�z�[���h�{�[�i�X��(�s) = 0 Then
                    '���݂̍s�̃m�[�c���̂ĂĂ��炸COOL�ȊO�ŁA���z�[���h�I���s�łȂ��ꍇ�̓��[�g�J�n
                    is���[�g�� = True
                    ���[�g�J�n���ϐ��ݒ� �s, False, ���O�z�[���h�J�n�s, ���݃��[�g��������, w��, is�z�[���h��
                    is�����[�g�� = False
                Else
                    '�z�[���h�I���s���z�[���h�J�n�s�������ꍇ�͉����[�g�J�n
                    For �}�[�N = 1 To Def.�}�[�N��
                        If Def.���ʃe�[�u��.�m�[�c��(�}�[�N, �s) = Def.HOLD���� Then
                            is�����[�g�� = True
                            Exit For
                        End If
                    Next �}�[�N
                End If
                
            End If
            
        End If
        
        '���[�g���̏ꍇ�̕����񏈗�
        If is���[�g�� Then
            
            '�̂ĂĂ���w���J�E���g����
            If Def.���ʃe�[�u��.�]����(�s) = Def.��WRONG���� Or Def.���ʃe�[�u��.�]����(�s) = Def.WORST���� Then
                w�� = w�� + 1
            End If
            
            Dim cool�\�� As Boolean
            cool�\�� = COOL�\���擾(�s, is�z�[���h��)
            
            If cool�\�� Then
                
                is�z�[���h�J�n�s = False
                For �}�[�N = 1 To Def.�}�[�N��
                    If Def.���ʃe�[�u��.�m�[�c��(�}�[�N, �s) = Def.HOLD���� Then
                        is�z�[���h�J�n(�}�[�N) = True
                        is�z�[���h�J�n�s = True
                    Else
                        is�z�[���h�J�n(�}�[�N) = False
                    End If
                Next �}�[�N
                
                If is�z�[���h�J�n�s Then
                    
                    If Def.���ʃe�[�u��.�]����(�s) = Def.��WRONG���� Or Def.���ʃe�[�u��.�]����(�s) = Def.WORST���� Then
                        'COOL�̉\��������z�[���h���̂ĂĂ���ꍇ�͌��݃��[�g��������ɒǋL
                        
                        ���݃��[�g�������� = ���݃��[�g�������� & " (" & �m�[�c�ԍ�������̎擾(�s, False)
                        For �}�[�N = 1 To Def.�}�[�N��
                            If is�z�[���h�J�n(�}�[�N) Then
                                ���݃��[�g�������� = ���݃��[�g�������� & �}�[�N����(�}�[�N)
                            End If
                        Next �}�[�N
                        ���݃��[�g�������� = ���݃��[�g�������� & Def.�]������(Def.���ʃe�[�u��.�]����(�s)) & ")"
                    Else
                        '�z�[���h��COOL�łƂ��Ă���ꍇ�̓z�[���h����True�ɂ��Č��݃��[�g��������ɒǋL
                        
                        ���݃��[�g�������� = ���݃��[�g�������� & " " & �m�[�c�ԍ�������̎擾(�s)
                        For �}�[�N = 1 To Def.�}�[�N��
                            If is�z�[���h�J�n(�}�[�N) Then
                                is�z�[���h��(�}�[�N) = True
                                ���݃��[�g�������� = ���݃��[�g�������� & �}�[�N����(�}�[�N)
                            End If
                        Next
                        
                        If Def.���ʃe�[�u��.�z�[���h�t���[����(�s) > Def.��MAX�t���[���ő�l - 1 - Def.CC�ő�]�T�t���[���� Then
                            
                            '�]�T�̂Ȃ��ڑ��ł���ꍇ�͏ڍׂ����݃��[�g��������ɒǋL
                            ���݃��[�g�������� = ���݃��[�g�������� & "�s"
                            
                            �J�n�t���[������ = Def.���ʃe�[�u��.���x�t���[����(���O�z�[���h�J�n�s)
                            �I���t���[������ = Def.���ʃe�[�u��.���x�t���[����(�s)
                            
                            If �J�n�t���[������ < 0 Then
                                �J�n�t���[���]�T = 0
                            ElseIf �J�n�t���[������ > 0 Then
                                If Def.���ʃe�[�u��.�]����(���O�z�[���h�J�n�s) <> Def.���ʃe�[�u��.�t���[������ʕ]��(0) Then
                                    ���݃��[�g�������� = ���݃��[�g�������� & "�x"
                                End If
                                �J�n�t���[���]�T = -1
                            Else
                                �J�n�t���[���]�T = 0
                            End If
                            ���݃��[�g�������� = ���݃��[�g�������� & Def.�]������(Def.���ʃe�[�u��.�]����(���O�z�[���h�J�n�s)) & "-"
                            Do Until Def.���ʃe�[�u��.�t���[������ʕ]��(�J�n�t���[������ + �J�n�t���[���]�T) = Def.���ʃe�[�u��.�t���[������ʕ]��(�J�n�t���[������)
                                �J�n�t���[���]�T = �J�n�t���[���]�T + 1
                            Loop
                            
                            If �I���t���[������ < 0 Then
                                �I���t���[���]�T = -1
                            ElseIf �I���t���[������ > 0 Then
                                If Def.���ʃe�[�u��.�]����(�s) <> Def.���ʃe�[�u��.�t���[������ʕ]��(0) Then
                                    ���݃��[�g�������� = ���݃��[�g�������� & "�x"
                                End If
                                �I���t���[���]�T = 0
                            Else
                                �I���t���[���]�T = 0
                            End If
                            ���݃��[�g�������� = ���݃��[�g�������� & Def.�]������(Def.���ʃe�[�u��.�]����(�s)) & ",�P�\:"
                            Do Until Def.���ʃe�[�u��.�t���[������ʕ]��(�I���t���[������ - �I���t���[���]�T) = Def.���ʃe�[�u��.�t���[������ʕ]��(�I���t���[������)
                                �I���t���[���]�T = �I���t���[���]�T + 1
                            Loop
                            
                            ���݃��[�g�������� = ���݃��[�g�������� & (�J�n�t���[���]�T + �I���t���[���]�T) & "�t"
                            
                            ���O�z�[���h�J�n�s = �s
                            
                        End If
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
    Next �s
    
    �X�R�A�^���[�g������擾 = �X�R�A�^������
    
End Function

Private Function ���[�g�J�n���ϐ��ݒ�( _
    ByVal ���ݍs As Long, _
    ByVal ���O���[�g�J�n�s�T�� As Boolean, _
    ByRef ���O�z�[���h�J�n�s As Long, _
    ByRef ���݃��[�g�������� As String, _
    ByRef w�� As Long, _
    ByRef is�z�[���h��() As Boolean)
    
    Dim �}�[�N As Long
    
    ' ���[�g�J�n�s
    ���O�z�[���h�J�n�s = ���ݍs
    If ���O���[�g�J�n�s�T�� Then
        Do Until Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(���O�z�[���h�J�n�s - 1) < Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(���O�z�[���h�J�n�s)
            ���O�z�[���h�J�n�s = ���O�z�[���h�J�n�s - 1
        Loop
    End If
    
    ' ���݃��[�g��������
    ���݃��[�g�������� = �m�[�c�ԍ�������̎擾(���O�z�[���h�J�n�s)
    For �}�[�N = 1 To Def.�}�[�N��
        If Def.���ʃe�[�u��.�z�[���h�ʊJ�n�t���[����(�}�[�N, ���O�z�[���h�J�n�s) > Def.���ʃe�[�u��.�z�[���h�ʊJ�n�t���[����(�}�[�N, ���O�z�[���h�J�n�s - 1) Then
            ���݃��[�g�������� = ���݃��[�g�������� & �}�[�N����(�}�[�N)
        End If
    Next �}�[�N
    
    ' w��
    w�� = 0
    
    ' is�z�[���h��
    For �}�[�N = 1 To Def.�}�[�N��
        is�z�[���h��(�}�[�N) = (Def.���ʃe�[�u��.�z�[���h�ʊJ�n�t���[����(�}�[�N, ���O�z�[���h�J�n�s) > 0)
    Next �}�[�N
    
End Function

Private Function ���x�t���[���ƕ]���̉���(ByVal �J�n�s As Long, Optional ByVal �I���s As Long = -1, Optional ByVal �]���̉��� As Boolean = True, Optional ByVal �ؑւ̉��� As Boolean = False)
    
    If �I���s = -1 Then
        �I���s = �J�n�s
    End If
    
    Dim �s As Long
    For �s = �J�n�s To �I���s
        If �]���̉��� And (Not Def.���x�蓮�w��t���O(�s)) Then
            If Def.���ʃe�[�u��.�]����(�s) <> Def.���ʃe�[�u��.�t���[������ʕ]��(0) Then
                Def.���ʃe�[�u��.�]����(�s) = Def.���ʃe�[�u��.�t���[������ʕ]��(0)
            End If
            If �ؑւ̉��� And (Not Def.���ʃe�[�u��.�ؑ֔����(�s)) Then
                Def.���ʃe�[�u��.�ؑ֔����(�s) = False
            End If
            If Def.���ʃe�[�u��.���x�蓮�w���(�s) <> "" Then
                Def.���ʃe�[�u��.���x�蓮�w���(�s) = ""
            End If
            If Def.���ʃe�[�u��.���x�t���[���蓮�w���(�s) <> "" Then
                Def.���ʃe�[�u��.���x�t���[���蓮�w���(�s) = ""
            End If
        End If
    Next
    
End Function

Private Function �ؑ։e���J�n�s�擾(ByVal �J�n�s As Long, ByVal �w��s As Long)

    Dim �ؑ֌��ʍs As Long
    Dim �ؑ֌��ʃu���b�N�J�n�s As Long
    Dim �ؑ֍Čv�Z�J�n�s As Long
    
    For �ؑ֌��ʍs = 1 To Def.�ؑ֌��ʃe�[�u��.�f�[�^�s��
        
        �ؑ֌��ʃu���b�N�J�n�s = Def.�ؑ֌��ʃe�[�u��.�u���b�N�J�n��(�ؑ֌��ʍs)
        If �ؑ֌��ʃu���b�N�J�n�s = �w��s Then
            �ؑ֍Čv�Z�J�n�s = �w��s
            Exit For
        ElseIf �ؑ֌��ʃu���b�N�J�n�s < �w��s Then
            �ؑ֍Čv�Z�J�n�s = Application.WorksheetFunction.Max(�ؑ֍Čv�Z�J�n�s, �ؑ֌��ʃu���b�N�J�n�s)
        End If
        
    Next �ؑ֌��ʍs
    
    �ؑ։e���J�n�s�擾 = Application.WorksheetFunction.Max(�ؑ֍Čv�Z�J�n�s, �J�n�s)
    
End Function

Private Function �w��s���O�܂ł̐ؑ֍Čv�Z(ByVal �J�n�s As Long, ByVal �w��s As Long, Optional ByVal ��ʍX�V As Boolean = False, Optional ByVal �����Čv�Z As Boolean = False) As Long
    
    Dim �ؑ֍Čv�Z�J�n�s As Long
    �ؑ֍Čv�Z�J�n�s = �ؑ։e���J�n�s�擾(�J�n�s, �w��s)
    
    If �ؑ֍Čv�Z�J�n�s < �w��s Then
        
        ���x�t���[���ƕ]���̉��� �ؑ֍Čv�Z�J�n�s, �w��s - 1, �ؑւ̉���:=True
        Def.���ʃe�[�u��.�Čv�Z �ؑ֍Čv�Z�J�n�s, �w��s - 1, �����Čv�Z
        
        Dim �z�[���h�v�Z�J�n�s As Long
        �z�[���h�v�Z�J�n�s = �ؑ֍Čv�Z�J�n�s
        Do Until Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�z�[���h�v�Z�J�n�s) <> Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�z�[���h�v�Z�J�n�s - 1) Or _
            Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�z�[���h�v�Z�J�n�s) = 0
            �z�[���h�v�Z�J�n�s = �z�[���h�v�Z�J�n�s + 1
        Loop
        
        Dim �ؑփf�[�^ As Def.�ʐؑփf�[�^
        �ؑփf�[�^ = �w��͈͂̃z�[���h�v�Z(�z�[���h�v�Z�J�n�s, �w��s, ��ʍX�V, �����Čv�Z, �ؑ֌��ʏo��:=False)
        If Not Not �ؑփf�[�^.�ؑ֍s���X�g Then
            Dim index As Long
            For index = 1 To UBound(�ؑփf�[�^.�ؑ֍s���X�g)
                Def.���ʃe�[�u��.�ؑ֔����(�ؑփf�[�^.�ؑ֍s���X�g(index)) = True
            Next index
            Def.���ʃe�[�u��.�Čv�Z �ؑ֍Čv�Z�J�n�s, �w��s - 1, �����Čv�Z
        End If
        
        �w��͈͂̑��x�������� �z�[���h�v�Z�J�n�s, �w��s, �����Čv�Z
        'Set ���������MAX������͓����� = �w��͈͂̑��x��������(�ؑ֍Čv�Z�J�n�s, �w��s, �����Čv�Z)
    End If
    
    �w��s���O�܂ł̐ؑ֍Čv�Z = �ؑ֍Čv�Z�J�n�s
    
End Function

Private Function �ؑ։e���I���s�擾(ByVal �w��s As Long)
    
    Dim �ؑ֌��ʍs As Long
    Dim �ؑ֌��ʃu���b�N�I���s As Long
    Dim �ؑ֍Čv�Z�I���s As Long
    �ؑ֍Čv�Z�I���s = Def.�ؑ֌��ʃe�[�u��.�ő�u���b�N�I���s
    
    Dim �ؑ֍Čv�Z�J�n�s As Long
    �ؑ֍Čv�Z�J�n�s = Def.�ؑ֌��ʃe�[�u��.�ő�u���b�N�J�n�s
    
    For �ؑ֌��ʍs = 1 To Def.�ؑ֌��ʃe�[�u��.�f�[�^�s��
        
        �ؑ֌��ʃu���b�N�I���s = Def.�ؑ֌��ʃe�[�u��.�u���b�N�I����(�ؑ֌��ʍs)
        If �ؑ֌��ʃu���b�N�I���s = �w��s Then
            �ؑ֍Čv�Z�I���s = �w��s
            Exit For
        ElseIf �ؑ֌��ʃu���b�N�I���s > �w��s Then
            �ؑ֍Čv�Z�I���s = Application.WorksheetFunction.Min(�ؑ֍Čv�Z�I���s, �ؑ֌��ʃu���b�N�I���s)
            �ؑ֍Čv�Z�J�n�s = Application.WorksheetFunction.Min(�ؑ֍Čv�Z�J�n�s, Def.�ؑ֌��ʃe�[�u��.�u���b�N�J�n��(�ؑ֌��ʍs))
        End If
        
    Next �ؑ֌��ʍs
    
    If �ؑ֍Čv�Z�J�n�s > �w��s Then
        �ؑ։e���I���s�擾 = �w��s
    Else
        �ؑ։e���I���s�擾 = �ؑ֍Čv�Z�I���s
    End If
    
End Function

Private Function �w��s���ォ��̐ؑ֍Čv�Z(ByVal �w��s As Long, Optional ByVal ��ʍX�V As Boolean = False, Optional ByVal �����Čv�Z As Boolean = False) As Long
    
    Dim �ؑ֍Čv�Z�I���s As Long
    �ؑ֍Čv�Z�I���s = �ؑ։e���I���s�擾(�w��s)
    
    If �ؑ֍Čv�Z�I���s > �w��s Then
        
        ���x�t���[���ƕ]���̉��� �w��s + 1, �ؑ֍Čv�Z�I���s, �ؑւ̉���:=True
        Def.���ʃe�[�u��.�Čv�Z �w��s + 1, �ؑ֍Čv�Z�I���s, �����Čv�Z
        
        Dim �z�[���h�v�Z�J�n�s As Long
        �z�[���h�v�Z�J�n�s = �w��s
        Do Until Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�z�[���h�v�Z�J�n�s) <> Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�z�[���h�v�Z�J�n�s - 1) Or _
            Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�z�[���h�v�Z�J�n�s) = 0
            �z�[���h�v�Z�J�n�s = �z�[���h�v�Z�J�n�s + 1
        Loop
        
        Dim �ؑփf�[�^ As Def.�ʐؑփf�[�^
        �ؑփf�[�^ = �w��͈͂̃z�[���h�v�Z(�z�[���h�v�Z�J�n�s, �ؑ֍Čv�Z�I���s, ��ʍX�V, �����Čv�Z, �ؑ֌��ʏo��:=False)
        If Not Not �ؑփf�[�^.�ؑ֍s���X�g Then
            Dim index As Long
            For index = 1 To UBound(�ؑփf�[�^.�ؑ֍s���X�g)
                Def.���ʃe�[�u��.�ؑ֔����(�ؑփf�[�^.�ؑ֍s���X�g(index)) = True
            Next index
            Def.���ʃe�[�u��.�Čv�Z �w��s + 1, �ؑ֍Čv�Z�I���s, �����Čv�Z
        End If
        
        �w��͈͂̑��x�������� �z�[���h�v�Z�J�n�s, �ؑ֍Čv�Z�I���s + 1, �����Čv�Z, True
        
    Else
        
        �ؑ֍Čv�Z�I���s = �w��s
        
    End If
    
    �w��s���ォ��̐ؑ֍Čv�Z = �ؑ֍Čv�Z�I���s
    
End Function

Private Function COOL�\���擾(ByVal �s As Long, is�z�[���h��() As Boolean) As Boolean
    COOL�\���擾 = True
    Dim �}�[�N As Long
    For �}�[�N = 1 To Def.�}�[�N��
        If Def.���ʃe�[�u��.�m�[�c��(�}�[�N, �s) <> "" And is�z�[���h��(�}�[�N) Then
            COOL�\���擾 = False
            Exit For
        End If
    Next �}�[�N
End Function

' ======================================================================================================================================================================================================
'
' ���J�n�s���z�[���h�J�n�s�ƈ�v����͂�
' ���J�n�s�̕]�������͗\�ߐݒ肳��Ă���\��������(����ȊO=���x�Ȃǂ͐ݒ肳��Ă��Ȃ��͂�)
' �߂�l�͍ő�e���s
' ======================================================================================================================================================================================================

Public Function �w��s����̃��[�g�T��( _
    ByVal �J�n�s As Long, _
    ByVal �I���s As Long, _
    ByRef is�z�[���h��() As Boolean, _
    ByRef ���݃��[�g�s�X�R�A() As Long, _
    Optional ByVal �J�n�����ꋖ�e�t���[�� As Long = -Def.��MAX�t���[���ő�l, _
    Optional ByVal �J�n�x���ꋖ�e�t���[�� As Long = Def.��MAX�t���[���ő�l, _
    Optional ByVal ���t���[�h As Boolean = False, _
    Optional ByVal ��ʍX�V As Boolean = False, _
    Optional ByVal �����Čv�Z As Boolean = False, _
    Optional ByVal �e���J�n�s As Long = 0) ', _
    Optional ByVal ���Ow�� As Long = 0)
    
    #If �X�R�A�^�ڍ׃��O Then
        �ڍ׃��O�o�� "�w��s����̃��[�g�T��" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�J�n", True
    #End If
    
    #If �X�R�A�^�ʏ탍�O Then
        �X�R�A�^�l�X�g�J�E���^ = �X�R�A�^�l�X�g�J�E���^ + 1
        �ڍ׃��O�o�� String(�X�R�A�^�l�X�g�J�E���^ - 1, "��") & "��" & �J�n�s & "�`", True, �X�R�A�^�l�X�g�J�E���^ + 1
    #End If
    
    Dim �t���[���� As Long
    
    Dim is�z�[���h�J�n(Def.�}�[�N��) As Boolean
    Dim ���[�g�J�n�\�� As Boolean
    
    Dim cool�\�� As Boolean
    Dim wrong�\�� As Boolean
    Dim w�� As Long
    w�� = 0 '���Ow��
    
    Dim ���[�g�ʕ]��() As Def.�]���Z�b�g
    Dim �b��]�� As String
    
    Dim �s As Long
    Dim �}�[�N As Long
    Dim �]��index As Long
    
    For �s = �J�n�s + 1 To �I���s
        
        ���x�t���[���ƕ]���̉��� �s
        Def.���ʃe�[�u��.�Čv�Z �s, , �����Čv�Z
        �t���[���� = Def.���ʃe�[�u��.�t���[����(�s) - Def.���ʃe�[�u��.�t���[����(�J�n�s)
        
        ���[�g�J�n�\�� = False
        '�ǂ��撣���Ă��z�[���h��ڑ��ł��Ȃ��s�ɂȂ�ΏI��
        If �t���[���� > Def.��MAX�t���[���ő�l + Def.���ʃe�[�u��.�ő�z�[���h�ڑ��t���[������ + 1 _
            - Application.WorksheetFunction.Max(Def.���ʃe�[�u��.�f�t�H���g�x���ꋖ�e�t���[�� - �J�n�x���ꋖ�e�t���[��, 0) Then
            Exit For
        End If
        
        '���ݍs�̑��x�t���[���ƕ]�������Z�b�g
        ���x�t���[���ƕ]���̉��� �s
        Def.���ʃe�[�u��.�Čv�Z �s, , �����Čv�Z
        
        'COOL�ɂȂ��=�{�^����������\���𔻒�
        cool�\�� = COOL�\���擾(�s, is�z�[���h��)
        
        If cool�\�� Then
            
            ' �V���Ƀ��[�g�J�n�ʒu�ƂȂ邱�Ƃ��\���ǂ����𔻒�
            ���[�g�J�n�\�� = False
            For �}�[�N = 1 To Def.�}�[�N��
                If Def.���ʃe�[�u��.�m�[�c��(�}�[�N, �s) = Def.HOLD���� Then
                    is�z�[���h�J�n(�}�[�N) = True
                    ���[�g�J�n�\�� = True
                ElseIf is�z�[���h��(�}�[�N) Then
                    is�z�[���h�J�n(�}�[�N) = True
                Else
                    is�z�[���h�J�n(�}�[�N) = False
                End If
            Next �}�[�N
            
            '���[�g�J�n�ʒu�ɂȂ��ꍇ�́A�z�[���h���J�n���Ď��̍s����V���Ƀ��[�g�T��
            If ���[�g�J�n�\�� Then
                
                Erase ���[�g�ʕ]��
                �b��]�� = Def.���ʃe�[�u��.�]����(�J�n�s)
                
                ' ���O�̃z�[���h�J�n�s�ƌ��ݍs�̕]���ƃt���[�������ݒ�
                If �t���[���� - Application.WorksheetFunction.Min(�J�n�x���ꋖ�e�t���[�� - Def.���ʃe�[�u��.�ő�COOL�t���[��, 0) > Def.��MAX�t���[���ő�l Then
                    '�J�n�s���ő�COOL(�x���ꋖ�e�t���[�����ő�COOL�ȑO�̏ꍇ�͂��̃t���[��)�ɂ����ꍇ�ɁA���ݍs�̑��x�t���[�����ő�COOL�Őڑ��ł��Ȃ��ꍇ
                    ���[�g�ʕ]�� = Def.���ʃe�[�u��.�]�����X�g�擾(Def.��MAX�t���[���ő�l - �t���[����, �J�n�����ꋖ�e�t���[��, �J�n�x���ꋖ�e�t���[��)
                Else
                    '�J�n�s���ő�COOL(�x���ꋖ�e�t���[�����ő�COOL�ȑO�̏ꍇ�͂��̃t���[��)�ɂ����ꍇ�ɁA���ݍs�̑��x�t���[�����ő�COOL�Őڑ��ł���ꍇ
                    ReDim ���[�g�ʕ]��(1)
                    ���[�g�ʕ]��(1).�J�n�]�� = Def.���ʃe�[�u��.�]����(�s)
                    ���[�g�ʕ]��(1).�J�n�t���[������ = Application.WorksheetFunction.Max(�J�n�����ꋖ�e�t���[��, _
                        Application.WorksheetFunction.Min(�J�n�x���ꋖ�e�t���[��, Def.���ʃe�[�u��.�ő�COOL�t���[��))
                End If
                
                Dim ���ݑ����ꋖ�e�t���[�� As Long
                Dim ���ݒx���ꋖ�e�t���[�� As Long
                
                ' �]�����ƂɊJ�n�s(�ƌ��ݍs)��ݒ肵�āA�V���Ƀ��[�g�T��
                If Not Not ���[�g�ʕ]�� Then
                    
                    For �]��index = 1 To UBound(���[�g�ʕ]��)
                    
                        Def.���ʃe�[�u��.�]����(�J�n�s) = ���[�g�ʕ]��(�]��index).�J�n�]��
                        Def.���ʃe�[�u��.���x�t���[���蓮�w���(�J�n�s) = ���[�g�ʕ]��(�]��index).�J�n�t���[������
                        
                        If ���[�g�ʕ]��(�]��index).�I���]�� = "" Then
                            Def.���ʃe�[�u��.�]����(�s) = Def.���ʃe�[�u��.�t���[������ʕ]��(0)
                            ���ݑ����ꋖ�e�t���[�� = Def.���ʃe�[�u��.�f�t�H���g�����ꋖ�e�t���[��
                            ���ݒx���ꋖ�e�t���[�� = Application.WorksheetFunction.Min(Def.���ʃe�[�u��.�f�t�H���g�x���ꋖ�e�t���[��, _
                            Def.��MAX�t���[���ő�l + Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�J�n�s) - Def.���ʃe�[�u��.�t���[����(�s) + 1)
                        Else
                            Def.���ʃe�[�u��.�]����(�s) = ���[�g�ʕ]��(�]��index).�I���]��
                            ���ݑ����ꋖ�e�t���[�� = ���[�g�ʕ]��(�]��index).�I���t���[������
                            ���ݒx���ꋖ�e�t���[�� = ���[�g�ʕ]��(�]��index).�I���t���[������
                        End If
                        
                        Dim ���ݍs�ؑ� As String
                        ���ݍs�ؑ� = Def.���ʃe�[�u��.�ؑ֔��蕶����(�s)
                        If ���ݍs�ؑ� <> "" Then
                            Def.���ʃe�[�u��.�ؑ֔����(�s) = False
                        End If
                        Def.���ʃe�[�u��.�Čv�Z �J�n�s, �s, �����Čv�Z
                        
                        ' �V���Ƀ��[�g�T��
                        �w��s����̃��[�g�T�� �s, �I���s, is�z�[���h�J�n, ���݃��[�g�s�X�R�A, ���ݑ����ꋖ�e�t���[��, ���ݒx���ꋖ�e�t���[��, ���t���[�h, ��ʍX�V, �����Čv�Z, �e���J�n�s ', w��
                        
                        If ���ݍs�ؑ� <> "" Then
                            Def.���ʃe�[�u��.�ؑ֔��蕶����(�s) = ���ݍs�ؑ�
                        End If
                        
                    Next �]��index
                    
                End If
                
                ' ���O�̃z�[���h�J�n�s�̃��Z�b�g
                ���x�t���[���ƕ]���̉��� �J�n�s
                Def.���ʃe�[�u��.�]����(�J�n�s) = �b��]��
                ���x�t���[���ƕ]���̉��� �s
                Def.���ʃe�[�u��.�Čv�Z �J�n�s, �s, �����Čv�Z
                
                ' �Ȍ�A���̃z�[���h��ׂ��\��������������
                cool�\�� = False
                
            End If
            
        End If
        
        ' COOL�ɂȂ�Ȃ��\��������ꍇ(��L��COOL�ɂȂ��\���̂���z�[���h��ׂ��\�����܂߂�)
        If Not cool�\�� Then
            
            ' WRONG�Œׂ���WORST���𔻒肳����
            wrong�\�� = False
            For �}�[�N = 1 To Def.�}�[�N��
                If Def.���ʃe�[�u��.�m�[�c��(�}�[�N, �s) = "" And (Not is�z�[���h��(�}�[�N)) Then
                    wrong�\�� = True
                    Exit For
                End If
            Next �}�[�N
            
            ' ��WRONG�܂���WORST�����E�Čv�Z
            If wrong�\�� Then
                Def.���ʃe�[�u��.�]����(�s) = Def.��WRONG����
            Else
                Def.���ʃe�[�u��.�]����(�s) = Def.WORST����
            End If
            Def.���ʃe�[�u��.�Čv�Z �s, , �����Čv�Z
            
            w�� = w�� + 1
            
        End If
        
    Next
    
    If �s > �I���s Then
        �s = �I���s
    End If
    
    ' ���݂̍s�ł͊���MAX�������Ă���(�͂���)���߁AMAX�����钼��̍s�܂ők��
    Do Until Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�s - 1) = Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�J�n�s)
        If Def.���ʃe�[�u��.�]����(�s) <> Def.���ʃe�[�u��.�t���[������ʕ]��(0) Then
            Def.���ʃe�[�u��.�]����(�s) = Def.���ʃe�[�u��.�t���[������ʕ]��(0)
            w�� = w�� - 1
        End If
        Def.���x�ؑֈꗗ�e�[�u��.���x�ؑ֏�񏑂��o�� Def.���ʃe�[�u��, �s, �s, False
        �s = �s - 1
    Loop
    
    Dim �J�n�s���x�t���[�� As Long
    ' �J�n�s�̑��x�t���[����ڑ�������ꍇ�͍ŒxCOOL�A�Ȃ��ꍇ�͍ő�COOL�ɐݒ�
    If isMAX�^�C�~���O�s��(�J�n�s) Then
        �J�n�s���x�t���[�� = Application.WorksheetFunction.Max(�J�n�����ꋖ�e�t���[��, _
            Application.WorksheetFunction.Min(�J�n�x���ꋖ�e�t���[��, Def.���ʃe�[�u��.�ő�COOL�t���[��))
    Else
        �J�n�s���x�t���[�� = Application.WorksheetFunction.Max(�J�n�����ꋖ�e�t���[��, _
            Application.WorksheetFunction.Min(�J�n�x���ꋖ�e�t���[��, Def.���ʃe�[�u��.�ŒxCOOL�t���[��))
    End If
    Def.���ʃe�[�u��.���x�t���[���蓮�w���(�J�n�s) = �J�n�s���x�t���[��
    Def.���ʃe�[�u��.�Čv�Z �J�n�s, �s, �����Čv�Z
    
'    Dim �ő�e���s As Long
'    �ő�e���s = �s
'    �w��s����̃��[�g�T�� = �s
    
    ' ���Ȍ�Čv�Z������
    
    Dim �ؑ֍Čv�Z�I���s As Long
'    Dim �Ώۃ��[�g�s�X�R�A As Def.�s�X�R�A���
    
    ' �̂ėʂ��ő�̎��́A���̍s�ȍ~�̍s�X�R�A����x����
    �b��]�� = Def.���ʃe�[�u��.�]����(�s)
    If �b��]�� <> Def.���ʃe�[�u��.�t���[������ʕ]��(0) Then
        Def.���ʃe�[�u��.�]����(�s) = Def.���ʃe�[�u��.�t���[������ʕ]��(0)
        w�� = w�� - 1
        If Def.���ʃe�[�u��.�z�[���h�t���[����(�s) + Def.���ʃe�[�u��.�ŒxCOOL�t���[�� <= Def.��MAX�t���[���ő�l Then
            ' COOL�ɂ����ꍇ�ɍŒx�ł�MAX������Ȃ��Ȃ�ꍇ(����? ��WRONG��301F,COOL�ŉ���������300F�ɂȂ�A���ŒxCOOL��0F�̏ꍇ�͂��蓾��?)
            ' ���̂Ƃ��͐�WRONG(�܂���WORST?)�̂܂܃��[�g���m�肳����
            Def.���ʃe�[�u��.�]����(�s) = �b��]��
            w�� = w�� + 1
        End If
    End If
    
    If w�� > 0 Then
'        �Ώۃ��[�g�s�X�R�A = ���[�g�m��㏈��(�J�n�s, �s, �I���s, ���݃��[�g�s�X�R�A, ���t���[�h, ��ʍX�V, �����Čv�Z, �e���J�n�s)
        ���[�g�m��㏈�� �J�n�s, �s, �I���s, ���݃��[�g�s�X�R�A, ���t���[�h, ��ʍX�V, �����Čv�Z, �e���J�n�s
    End If
    
    ' ��������s��k���āA��s����WRONG��WORST���O�������̍s�X�R�A������
    ' ���X�^�[�g�s�̕]�������WRONG��WORST�̉\������
    �b��]�� = Def.���ʃe�[�u��.�]����(�J�n�s)
    
    For �s = �s To �J�n�s + 1 Step -1
        
        If Def.���ʃe�[�u��.�]����(�s) <> Def.���ʃe�[�u��.�t���[������ʕ]��(0) Then
            
            Def.���ʃe�[�u��.�]����(�s) = Def.���ʃe�[�u��.�t���[������ʕ]��(0)
            
            Erase ���[�g�ʕ]��
            ���[�g�ʕ]�� = Def.���ʃe�[�u��.�]�����X�g�擾( _
                Def.��MAX�t���[���ő�l + 3 + Def.���ʃe�[�u��.�t���[����(�J�n�s) - Def.���ʃe�[�u��.�t���[����(�s), _
                �J�n�����ꋖ�e�t���[��, �J�n�x���ꋖ�e�t���[��)
            
            ' MAX�����邱�Ƃ��o����ꍇ�A���̏�Ԃł̍s�X�R�A������
            If Not Not ���[�g�ʕ]�� Then
                For �]��index = 1 To UBound(���[�g�ʕ]��)
                    Def.���ʃe�[�u��.�]����(�J�n�s) = ���[�g�ʕ]��(�]��index).�J�n�]��
                    Def.���ʃe�[�u��.���x�t���[���蓮�w���(�J�n�s) = ���[�g�ʕ]��(�]��index).�J�n�t���[������
                    Def.���ʃe�[�u��.�]����(�s) = ���[�g�ʕ]��(�]��index).�I���]��
                    Def.���ʃe�[�u��.���x�t���[���蓮�w���(�s) = ���[�g�ʕ]��(�]��index).�I���t���[������
                    
'                    ���[�g�m��㏈�� �J�n�s, �s, �I���s, �Ώۃ��[�g�s�X�R�A, ���t���[�h, ��ʍX�V, �����Čv�Z, �e���J�n�s
                    ���[�g�m��㏈�� �J�n�s, �s, �I���s, ���݃��[�g�s�X�R�A, ���t���[�h, ��ʍX�V, �����Čv�Z, �e���J�n�s
                Next
            End If
            
            ' MAX�����Ȃ��ꍇ�̍s�X�R�A������
            Def.���ʃe�[�u��.�]����(�J�n�s) = �b��]��
            Def.���ʃe�[�u��.���x�t���[���蓮�w���(�J�n�s) = �J�n�s���x�t���[��
            Def.���ʃe�[�u��.�]����(�s) = Def.���ʃe�[�u��.�t���[������ʕ]��(0)
            Def.���ʃe�[�u��.���x�t���[���蓮�w���(�s) = Def.���ʃe�[�u��.�ŒxCOOL�t���[��
            
            w�� = w�� - 1
            If w�� > 0 Then
'                ���[�g�m��㏈�� �J�n�s, �s, �I���s, �Ώۃ��[�g�s�X�R�A, ���t���[�h, ��ʍX�V, �����Čv�Z, �e���J�n�s
                ���[�g�m��㏈�� �J�n�s, �s, �I���s, ���݃��[�g�s�X�R�A, ���t���[�h, ��ʍX�V, �����Čv�Z, �e���J�n�s
            End If
            
        End If
        
        Def.���x�ؑֈꗗ�e�[�u��.���x�ؑ֏�񏑂��o�� Def.���ʃe�[�u��, �s, �s, False
        
    Next �s
    
    '�ؑւƑ��x�����ɖ߂� �� ���x�ؑւ�1�s���ɖ߂��A�Čv�Z�͂���Ȃ�?
'    Def.���x�ؑֈꗗ�e�[�u��.���x�ؑ֏�񏑂��o�� Def.���ʃe�[�u��, �J�n�s + 1, �ؑ֍Čv�Z�I���s, False
'    Def.���ʃe�[�u��.�Čv�Z �J�n�s + 1, �ؑ֍Čv�Z�I���s, �����Čv�Z
    
    #If �X�R�A�^�ڍ׃��O Then
        �ڍ׃��O�o�� "�w��s����̃��[�g�T��" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�I��", True
    #End If
    
    #If �X�R�A�^�ʏ탍�O Then
        Def.�������O.�t�H�[��������폜 �X�R�A�^�l�X�g�J�E���^ + 1
        �X�R�A�^�l�X�g�J�E���^ = �X�R�A�^�l�X�g�J�E���^ - 1
    #End If
    
End Function

Private Function ���[�g�m��㏈��( _
    ByVal �J�n�s As Long, _
    ByVal ���ݍs As Long, _
    ByVal �I���s As Long, _
    ByRef ���݃��[�g�s�X�R�A() As Long, _
    Optional ByVal ���t���[�h As Boolean = False, _
    Optional ByVal ��ʍX�V As Boolean = False, _
    Optional ByVal �����Čv�Z As Boolean = False, _
    Optional ByVal �e���J�n�s As Long = 0) _
    As Def.�s�X�R�A���
'    Optional ByRef ��r���[�g�s�X�R�A As Def.�s�X�R�A���, _

    Dim �ؑ֍Čv�Z�I���s As Long
    
    Def.���ʃe�[�u��.�Čv�Z �J�n�s, ���ݍs, �����Čv�Z
    �ؑ֍Čv�Z�I���s = �w��s���ォ��̐ؑ֍Čv�Z(���ݍs, ��ʍX�V, �����Čv�Z)
    DoEvents
    ���[�g�m��㏈�� = �w��͈͂̃��[�g����(���ݍs, �I���s, ���݃��[�g�s�X�R�A, ���t���[�h, ��ʍX�V, �����Čv�Z, �e���J�n�s)
    Def.���x�ؑֈꗗ�e�[�u��.���x�ؑ֏�񏑂��o�� Def.���ʃe�[�u��, ���ݍs + 1, �ؑ֍Čv�Z�I���s, False
'    Def.���ʃe�[�u��.�Čv�Z ���ݍs + 1, �ؑ֍Čv�Z�I���s, �����Čv�Z
    
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function �����ؑ֔���( _
    Optional ByVal ��ʍX�V As Boolean = False, _
    Optional ByVal �����Čv�Z As Boolean = False, _
    Optional ByVal ����J�n�s As Long = 1, _
    Optional ByVal ����I���s As Long = -1)
    
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�����ؑ֔���" & vbTab & "�J�n"
    #End If
    
    Application.ScreenUpdating = False
    
    If �����Čv�Z Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    
    If ����I���s = -1 Then
        ����I���s = Def.���ʃe�[�u��.�f�[�^�s��
    End If
    
    ' ��ʍX�VON���̃t�B���^�X�V�p�����ݒ� ---------------------------------------------------------
    
    If ��ʍX�V Then
        
        Dim block As Long
        Dim hBlock As Long
        Dim hBlockColor As Long
        
        block = Def.���ʃe�[�u��.OwnTable.ListColumns("block").index
        hBlock = Def.���ʃe�[�u��.OwnTable.ListColumns("HBlock").index
        hBlockColor = RGB(146, 208, 80)
        
        Def.���ʃe�[�u��.OwnTable.Range.AutoFilter hBlock, hBlockColor, Operator:=xlFilterCellColor
        
    End If
    
    ' ��̓u���b�N���� -----------------------------------------------------------------------------
    
    Dim �s As Long
    Dim �J�n�s As Long
    Dim �I���s As Long
    
    Dim ���ݐؑ֌����� As Long
    Dim ���ʍs As Long
    ���ݐؑ֌����� = Def.�ؑ֌��ʃe�[�u��.�f�[�^�s��
    ���݃z�[���h�u���b�N = 0
    
    For �s = ����J�n�s To ����I���s
        
        If Def.���ʃe�[�u��.�z�[���h�u���b�N��(�s) <> ���݃z�[���h�u���b�N Then
            
            If ���݃z�[���h�u���b�N > 0 Then
                
                �I���s = �s - 1
                
                If ��ʍX�V Then
                    Def.���ʃe�[�u��.OwnTable.Range.AutoFilter block, ���݃z�[���h�u���b�N
                    Def.���ʃe�[�u��.OwnTable.ListRows(�s).Range.Rows.Hidden = False
                End If
                
                Def.�������O.�o�� "�y" & ���݃z�[���h�u���b�N & "�u���b�N�ډ�̓X�^�[�g�z(" & �J�n�s & "�`" & �I���s & "�s��)"
                
                Application.ScreenUpdating = ��ʍX�V
                DoEvents
                Application.ScreenUpdating = False
                
                ���ݐؑ֌����� = Def.�ؑ֌��ʃe�[�u��.�f�[�^�s��
                
                �w��͈͂̃z�[���h�v�Z �J�n�s, �I���s, ��ʍX�V, �����Čv�Z
                
                ' �u���b�N���ʏo��
                For ���ʍs = ���ݐؑ֌����� + 1 To Def.�ؑ֌��ʃe�[�u��.�f�[�^�s��
                    Def.�ؑ֌��ʃe�[�u��.�u���b�N�J�n��(���ʍs) = �J�n�s
                    Def.�ؑ֌��ʃe�[�u��.�u���b�N�I����(���ʍs) = �I���s
                Next ���ʍs
                
                'Application.ScreenUpdating = False
                
                If ��ʍX�V Then
                    Def.���ʃe�[�u��.OwnTable.Range.AutoFilter block
                End If
                
            End If
            
            ���݃z�[���h�u���b�N = Def.���ʃe�[�u��.�z�[���h�u���b�N��(�s)
            �J�n�s = �s
            
        End If
        
    Next �s
    
    If ��ʍX�V Then
        Def.���ʃe�[�u��.OwnTable.Range.AutoFilter hBlock
    End If
    
    Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True
    
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�����ؑ֔���" & vbTab & "�I��"
    #End If
    
End Function

' ======================================================================================================================================================================================================
'
' �ؑ֌��ʏo�� �� False �̎��̂ݍő�X�R�A�ɂȂ�ؑփf�[�^���i�[�����I�u�W�F�N�g���Ԃ���܂��B
' �� (�J�n�s - 1) �s�ڂ��]���ΏۂɂȂ�܂��B�� �J�n�s = 1 �̓_��
' �� ���x����̂��߂� (�I���s + 1) �s�ڂ܂ŕ]���ΏۂɂȂ�܂��B�� �I���s = �e�[�u���ŏI�s �̓_��
' �� �܂����x����̉e���ɂ���āA�]���Ώۂ̍s�̖���������ɒǉ������\��������܂��B
' ======================================================================================================================================================================================================

Public Function �w��͈͂̃z�[���h�v�Z( _
    ByVal �J�n�s As Long, _
    ByRef �I���s As Long, _
    Optional ByVal ��ʍX�V As Boolean = False, _
    Optional ByVal �����Čv�Z As Boolean = False, _
    Optional ByVal �u���b�N�J�n�s As Long = 0, _
    Optional ByRef MAX������ As OutputString = Nothing, _
    Optional ByVal �ؑ֌��ʏo�� As Boolean = True) _
    As Def.�ʐؑփf�[�^
    
    ' �����ݒ� -------------------------------------------------------------------------------------
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�w��͈͂̃z�[���h�v�Z" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�����ݒ�J�n"
    #End If
    
    If �u���b�N�J�n�s = 0 Then
        �u���b�N�J�n�s = �J�n�s
    End If
    
    If MAX������ Is Nothing Then
        Set MAX������ = New OutputString
        MAX������.�\�������� = ""
        MAX������.���O�o�͗p������ = ""
    End If
    
    Dim �z�[���h�J�n�t���[��(�}�[�N��) As Double
    Dim �z�[���h�I���s(�}�[�N��) As Long
    
    Dim �}�[�N As Long
    For �}�[�N = 1 To �}�[�N��
        �z�[���h�J�n�t���[��(�}�[�N) = 0
    Next
    
    Dim �ؑ։\�� As Boolean
    
    If Not �ؑ֌��ʏo�� Then
        Dim �ő�X�R�A�ؑփf�[�^ As Def.�ʐؑփf�[�^
        �ő�X�R�A�ؑփf�[�^.�X�R�A = 0
        Dim ���ݐؑփf�[�^ As Def.�ʐؑփf�[�^
    End If
    
    ' �����ؑ֔��� ---------------------------------------------------------------------------------
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�w��͈͂̃z�[���h�v�Z" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�ؑ֔���J�n"
    #End If
    
    Dim �s As Long
    'Dim �}�[�N As Long
    Dim �t���[�� As Double
    
    Dim ���݃z�[���h�I���s As Long
    Dim ���x����I���s As Long
    Dim ����MAX������ As OutputString
    
'    Dim ���x����J�n�s As Long
'    ���x����J�n�s = �J�n�s
    
    For �s = �J�n�s To �I���s
        
        �ؑ։\�� = False
        
        ' 1. �}�[�N���ƂɃz�[���h�J�n�t���[���ƃz�[���h�I���s�����̍s�̂��̂ɐݒ�
        
        For �}�[�N = 1 To �}�[�N��
            
            �t���[�� = Def.���ʃe�[�u��.�z�[���h�\�������(�}�[�N, �s)
            
            If �t���[�� > �z�[���h�J�n�t���[��(�}�[�N) Then
                
                �z�[���h�J�n�t���[��(�}�[�N) = �t���[��
                
                ���݃z�[���h�I���s = �s
                
                Do While �z�[���h�J�n�t���[��(�}�[�N) = Def.���ʃe�[�u��.�z�[���h�\�������(�}�[�N, ���݃z�[���h�I���s)
                    ���݃z�[���h�I���s = ���݃z�[���h�I���s + 1
                Loop
                
                ' �z�[���h�I���s��(�K�؂ȏI���s�ݒ�ł�) �I���s + 1 �ɂȂ�\��������
                �z�[���h�I���s(�}�[�N) = ���݃z�[���h�I���s
                
                �ؑ։\�� = True
                                
            ElseIf �t���[�� = 0 And �z�[���h�J�n�t���[��(�}�[�N) > 0 Then
                
                �z�[���h�J�n�t���[��(�}�[�N) = 0
                �z�[���h�I���s(�}�[�N) = 0
                
            End If
            
        Next �}�[�N
        
        ' 2. �ؑ։\��������A�ؑւ��蓮�w�肳��Ă��Ȃ��s�ł���ꍇ�A�ؑւ��w�肵�A������̓u���b�N���̑������ċA�ŒT��
        
        If �ؑ։\�� And Def.���ʃe�[�u��.�ؑ֔��蕶����(�s) = "" Then
            
            For �}�[�N = 1 To �}�[�N��
            
                If �z�[���h�I���s(�}�[�N) < ���݃z�[���h�I���s _
                    And �z�[���h�J�n�t���[��(�}�[�N) = Def.���ʃe�[�u��.�z�[���h�ʊJ�n�t���[����(�}�[�N, �s) _
                    And �z�[���h�J�n�t���[��(�}�[�N) > 0 Then
                    
                    Def.���ʃe�[�u��.�ؑ֔����(�s) = True
                    
                    'Def.���ʃe�[�u��.�Čv�Z �s, �I���s + 1, �����Čv�Z
                    'DoEvents
                    
                    ���x����I���s = �s
                    
                    Set ����MAX������ = �w��͈͂̑��x��������(�J�n�s, ���x����I���s, �����Čv�Z)
                    
                    If ���x����I���s <> �s Then
                        MsgBox "�\�z�O�̂��Ƃ��������܂����B�Ȃ��ł����B�����ĉ������B�ł��܂������͑����܂���B"
                    End If
                    
                    Def.���ʃe�[�u��.�Čv�Z �s, �I���s + 1, �����Čv�Z
                    DoEvents
                    
                    ����MAX������.�\�������� = Def.������A��(MAX������.�\��������, ����MAX������.�\��������, vbCrLf)
                    ����MAX������.���O�o�͗p������ = Def.������A��(MAX������.���O�o�͗p������, ����MAX������.���O�o�͗p������, ", ")
                    ���ݐؑփf�[�^ = �w��͈͂̃z�[���h�v�Z(�s, �I���s, ��ʍX�V, �����Čv�Z, �u���b�N�J�n�s, ����MAX������, �ؑ֌��ʏo��)
                    If Not �ؑ֌��ʏo�� Then
                        If ���ݐؑփf�[�^.�X�R�A > �ő�X�R�A�ؑփf�[�^.�X�R�A Then
                            �ő�X�R�A�ؑփf�[�^.�X�R�A = ���ݐؑփf�[�^.�X�R�A
                            �ő�X�R�A�ؑփf�[�^.�ؑ֍s���X�g = ���ݐؑփf�[�^.�ؑ֍s���X�g
                        End If
                    End If
                    ' ��ʍă��b�N
                    Application.ScreenUpdating = False
                    
                    �w��͈͂̑��x�w��폜 �J�n�s, ���x����I���s, �����Čv�Z
                    
                    'Def.���ʃe�[�u��.�Čv�Z �s, �I���s + 1, �����Čv�Z
                    
                    Def.���ʃe�[�u��.�ؑ֔����(�s) = False
                    
                    Def.���ʃe�[�u��.�Čv�Z �s, �I���s + 1, �����Čv�Z
                    'DoEvents
                    
                    Exit For
                    
                End If
                
            Next
            
        End If
        
    Next �s
    
    ' �����[�g�ɂ�錋�ʂ̕\�� ---------------------------------------------------------------------
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�w��͈͂̃z�[���h�v�Z" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "���ʕ\���J�n"
    #End If
    
    'Dim ����MAX������ As String
    
    ���x����I���s = �I���s + 1
    Set ����MAX������ = �w��͈͂̑��x��������(�J�n�s, ���x����I���s, �����Čv�Z)
    DoEvents
    �I���s = ���x����I���s - 1
    
    ����MAX������.�\�������� = Def.������A��(MAX������.�\��������, ����MAX������.�\��������, vbCrLf)
    ����MAX������.���O�o�͗p������ = Def.������A��(MAX������.���O�o�͗p������, ����MAX������.���O�o�͗p������, ", ")
    
    ' �ؑ֌��ʏo��
    If �ؑ֌��ʏo�� Then
        �ؑ֌��ʃe�[�u��.�o�͍s�ǉ�
    Else
'        ���ݐؑփf�[�^.�X�R�A = 0
        Erase ���ݐؑփf�[�^.�ؑ֍s���X�g
    End If
    
    Dim ���O�o�͗p�ؑ֕����� As String
    Dim �ؑ֕����� As String
    Dim �ؑփm�[�c�ԍ� As Long
    Dim �ؑփR���{�� As Long
    ���O�o�͗p�ؑ֕����� = ""
    �ؑ֕����� = ""
    
    Dim �ؑ֐� As Long
    �ؑ֐� = 0
    
    For �s = �u���b�N�J�n�s To �I���s
        
        If Def.���ʃe�[�u��.�ؑ֔����(�s) Then
            
            �ؑ֐� = �ؑ֐� + 1
            
            �ؑ֕����� = Def.������A��(�ؑ֕�����, �ؑ֕�����擾(�s), vbCrLf)
            
            If �ؑ֌��ʏo�� Then
                Def.�ؑ֌��ʃe�[�u��.�ؑ֍s���� �s, �ؑ֐�
            Else
                ReDim Preserve ���ݐؑփf�[�^.�ؑ֍s���X�g(�ؑ֐�)
                ���ݐؑփf�[�^.�ؑ֍s���X�g(�ؑ֐�) = �s
            End If
            
            ���O�o�͗p�ؑ֕����� = Def.������A��(���O�o�͗p�ؑ֕�����, �s & "�s��", ", ")
            
        End If
        
    Next �s
        
    Dim �X�R�A As Long
    �X�R�A = Def.���ʃe�[�u��.�z�[���h�X�R�A��(���x����I���s) - Def.���ʃe�[�u��.�z�[���h�X�R�A��(�u���b�N�J�n�s)
    
    If �ؑ֌��ʏo�� Then
        Def.�ؑ֌��ʃe�[�u��.��̓u���b�N��(Def.�ؑ֌��ʃe�[�u��.�f�[�^�s��) = ���݃z�[���h�u���b�N
        Def.�ؑ֌��ʃe�[�u��.�u���b�N�X�R�A��(Def.�ؑ֌��ʃe�[�u��.�f�[�^�s��) = �X�R�A
        Def.�ؑ֌��ʃe�[�u��.�ؑ֕�����(Def.�ؑ֌��ʃe�[�u��.�f�[�^�s��) = �ؑ֕�����
        Def.�ؑ֌��ʃe�[�u��.MAX�\��������(Def.�ؑ֌��ʃe�[�u��.�f�[�^�s��) = ����MAX������.�\��������
    Else
'        ���ݐؑփf�[�^.�X�R�A = �X�R�A
        If �X�R�A > �ő�X�R�A�ؑփf�[�^.�X�R�A Then
            �ő�X�R�A�ؑփf�[�^.�X�R�A = �X�R�A
            �ő�X�R�A�ؑփf�[�^.�ؑ֍s���X�g = ���ݐؑփf�[�^.�ؑ֍s���X�g
        End If
        �w��͈͂̃z�[���h�v�Z = �ő�X�R�A�ؑփf�[�^
    End If
    
    If �ؑ֌��ʏo�� Then
        
        If ���O�o�͗p�ؑ֕����� = "" Then
            ���O�o�͗p�ؑ֕����� = "(�Ȃ�)"
        End If
        
        Dim ���O�o�͕����� As String
        ���O�o�͕����� = ���݃z�[���h�u���b�N & "�u���b�N�� / �X�R�A: " & �X�R�A & " / �ؑ�: " & ���O�o�͗p�ؑ֕�����
        
        If ����MAX������.���O�o�͗p������ <> "" Then
            ���O�o�͕����� = ���O�o�͕����� & " / MAX�\��: " & ����MAX������.���O�o�͗p������
        End If
        Def.�������O.�o�� ���O�o�͕�����
        
    End If
    
    ' �V�[�g�Ɍ��ʕ\��
    Application.ScreenUpdating = ��ʍX�V
    DoEvents
    Application.ScreenUpdating = False
    
    �w��͈͂̑��x�w��폜 �J�n�s, ���x����I���s, �����Čv�Z
    
    Def.���ʃe�[�u��.�Čv�Z �J�n�s, ���x����I���s, �����Čv�Z
    
    'DoEvents
    
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�w��͈͂̃z�[���h�v�Z" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�I��"
    #End If
    
End Function

' ======================================================================================================================================================================================================
'
' MAX�\���t���O(�J�n�s to �I���s)���ݒ肳��܂��B
' �߂�l��MAX�\���̂���m�[�c�̏���\��������
' �� (�J�n�s - 1) �s�ڂ��]���ΏۂɂȂ�܂��B�� �J�n�s = 1 �̓_��
' �� �J�n�s���z�[���h���̍s������_��(�z�[���h���łȂ���΃z�[���h�u���b�N�r���̍s����ł�OK??)
' �܂��A�I���s�t�߂�MAX�����ނ��ƂŁAMAX�I���s�܂ŕ]���Ώۂ̖����̍s���ǉ������\��������܂��B
' �I���s���ύX���ꂽ�ꍇ�A�����̎Q�Ɠn���ɂ���Č��̕ϐ����ύX����܂��B
' ======================================================================================================================================================================================================

Public Function �w��͈͂̑��x��������( _
    ByVal �J�n�s As Long, _
    ByRef �I���s As Long, _
    Optional ByVal �����Čv�Z As Boolean = False, _
    Optional ByVal ���OMAX�m�F As Boolean = False) _
    As OutputString
    
    ' �����ݒ� -------------------------------------------------------------------------------------
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�w��͈͂̑��x��������" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�����ݒ�J�n"
    #End If
    
    Def.���ʃe�[�u��.�Čv�Z �J�n�s, �I���s, �����Čv�Z
    
    ' MAX�̉\�������݂���z�[���h�ӏ������� ------------------------------------------------------
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�w��͈͂̑��x��������" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "MAX�����J�n"
    #End If
    
    
    Dim �s As Long
    
    Dim ���݃z�[���h�J�n�s As Long
    Dim ���݃z�[���h�J�n�t���[������� As Double
    
    'Dim �s As Long
    
    For �s = �J�n�s To �I���s
        
        If Def.���ʃe�[�u��.�z�[���h�{�[�i�X��(�s) > 0 And �s > �J�n�s Then
            
            If Def.���ʃe�[�u��.�z�[���h�t���[����(�s) - (Def.���ʃe�[�u��.���x�t���[����(�s) - ���݃z�[���h�J�n�t���[�������) > _
                Def.��MAX�t���[���ő�l - (Def.���ʃe�[�u��.�ŒxCOOL�t���[�� - Def.���ʃe�[�u��.�ő�COOL�t���[��) Then
                MAX�\���t���O(���݃z�[���h�J�n�s) = True
            End If
            
            ���݃z�[���h�J�n�s = 0
            ���݃z�[���h�J�n�t���[������� = 0
            
        End If
        
        If Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�s) > Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�s - 1) Then
            
            ���݃z�[���h�J�n�s = �s
            ���݃z�[���h�J�n�t���[������� = Def.���ʃe�[�u��.���x�t���[����(�s)
            
        End If
    
    Next
    
    ' �ォ�珇�Ƀu���b�N���Ƃɑ��x��������(MAX�����ޏꍇ�̂ݑk���ĕύX�̉\������) ------------------------------
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�w��͈͂̑��x��������" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "���x�ݒ�J�n"
    #End If
    
    Dim �}�[�N As Long
    
    'Dim ���݃z�[���h�J�n�s As Long
    Dim ���݃z�[���h�{�^���� As Long
    Dim ����MAX�\�� As Boolean
    
    ���݃z�[���h�J�n�s = 0
    ���݃z�[���h�{�^���� = 0
    For �}�[�N = 1 To �}�[�N��
        If Def.���ʃe�[�u��.�z�[���h�ʊJ�n�t���[����(�}�[�N, �J�n�s - 1) > 0 Then
            ���݃z�[���h�{�^���� = ���݃z�[���h�{�^���� + 1
        End If
    Next �}�[�N
    
    If ���OMAX�m�F Then
        ����MAX�\�� = Def.���ʃe�[�u��.�z�[���h�t���[����(�J�n�s) > Def.��MAX�t���[���ő�l
        If ����MAX�\�� Then
            For ���݃z�[���h�J�n�s = �J�n�s - 1 To 2 Step -1
                If Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(���݃z�[���h�J�n�s) > Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(���݃z�[���h�J�n�s - 1) Then
                    Exit For
                End If
            Next
            If ���݃z�[���h�J�n�s = 1 Then
                ���݃z�[���h�J�n�s = 0
            End If
        End If
    Else
        ����MAX�\�� = False
    End If
    
    Dim ���z�[���h�u���b�N As Long
    Dim �u���b�N�J�n�s As Long
    
    ���z�[���h�u���b�N = Def.���ʃe�[�u��.�z�[���h�u���b�N��(�J�n�s)
    �u���b�N�J�n�s = �J�n�s
    
    For �s = �J�n�s To �I���s - 1
        If Def.���ʃe�[�u��.�z�[���h�u���b�N��(�s + 1) <> ���z�[���h�u���b�N Then
            �u���b�N�����x�������� �u���b�N�J�n�s, �s, ���݃z�[���h�J�n�s, ���݃z�[���h�{�^����, ����MAX�\��, �����Čv�Z
            DoEvents
            ���z�[���h�u���b�N = Def.���ʃe�[�u��.�z�[���h�u���b�N��(�s + 1)
            �u���b�N�J�n�s = �s + 1
        End If
    Next �s
    
    ' (���X�g�܂ő��x�v�Z)
    �u���b�N�����x�������� �u���b�N�J�n�s, �I���s, ���݃z�[���h�J�n�s, ���݃z�[���h�{�^����, ����MAX�\��, �����Čv�Z
    
    ' MAX�\���̂���ӏ��̃f�o�b�O�o�� ------------------------------------------------------------
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�w��͈͂̑��x��������" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "MAX�o�͊J�n"
    #End If
    
    Dim ���O�o�͗pMAX������ As String
    Dim MAX������ As String
    'Dim ����MAX�\�� As Boolean
    
    ���O�o�͗pMAX������ = ""
    MAX������ = ""
    ����MAX�\�� = False
    
    'Dim �s As Long
    'Dim �}�[�N As Long
    Dim �{�^���� As Long
    
    For �s = �J�n�s To �I���s
        
        If ����MAX�\�� And Def.���ʃe�[�u��.�z�[���h�{�[�i�X��(�s) > 0 Then
            
            �{�^���� = 0
            For �}�[�N = 1 To �}�[�N��
                If Def.���ʃe�[�u��.�z�[���h�ʊJ�n�t���[����(�}�[�N, �s - 1) > 0 Then
                    �{�^���� = �{�^���� + 1
                End If
            Next �}�[�N
            
            MAX������ = MAX������ & "�y"
            If �{�^���� > 1 Then
                MAX������ = MAX������ & �{�^����
            End If
            MAX������ = MAX������ & "MAX�z"
            
            ���O�o�͗pMAX������ = ���O�o�͗pMAX������ & �s & ���x�����܂��̓t���[���̎擾(�s) & "�s�� ("
            
            If Def.���ʃe�[�u��.�z�[���h�I���������������(�s) Then
                
                ���O�o�͗pMAX������ = ���O�o�͗pMAX������ & Def.���ʃe�[�u��.�z�[���h�t���[����(�s)
                
                MAX������ = MAX������ & "�� " & �m�[�c�ԍ�������̎擾(�s)
                
                For �}�[�N = 1 To �}�[�N��
                    If Def.���ʃe�[�u��.�m�[�c��(�}�[�N, �s) <> "" Then
                        MAX������ = MAX������ & �}�[�N����(�}�[�N)
                    End If
                Next �}�[�N
            
                MAX������ = MAX������ & " (" & Def.���ʃe�[�u��.�z�[���h�t���[����(�s)
                
            Else
                
                ���O�o�͗pMAX������ = ���O�o�͗pMAX������ & ">" & Def.��MAX�t���[���ő�l
                MAX������ = MAX������ & "(>" & Def.��MAX�t���[���ő�l
                
            End If
            
            ���O�o�͗pMAX������ = ���O�o�͗pMAX������ & "F)"
            MAX������ = MAX������ & "F)"
            
            ����MAX�\�� = False
            
        End If
        
        If MAX�\���t���O(�s) Then
        
            ����MAX�\�� = True
            
            If ���O�o�͗pMAX������ <> "" Then
                ���O�o�͗pMAX������ = ���O�o�͗pMAX������ & ", "
            End If
            
            ���O�o�͗pMAX������ = ���O�o�͗pMAX������ & �s & ���x�����܂��̓t���[���̎擾(�s) & "��"
            
            If MAX������ <> "" Then
                MAX������ = MAX������ & vbCrLf
            End If
            
            MAX������ = MAX������ & �m�[�c�ԍ�������̎擾(�s)
            
            For �}�[�N = 1 To �}�[�N��
                If Def.���ʃe�[�u��.�z�[���h�ʊJ�n�t���[����(�}�[�N, �s) > Def.���ʃe�[�u��.�z�[���h�ʊJ�n�t���[����(�}�[�N, �s - 1) Then
                    MAX������ = MAX������ & �}�[�N����(�}�[�N)
                End If
            Next �}�[�N
            
            MAX������ = MAX������ & ���x�����܂��̓t���[���̎擾(�s) & " ��"
                        
        End If
        
    Next �s
    
    Set �w��͈͂̑��x�������� = New OutputString
    �w��͈͂̑��x��������.�\�������� = MAX������
    �w��͈͂̑��x��������.���O�o�͗p������ = ���O�o�͗pMAX������
    
    Def.���ʃe�[�u��.�Čv�Z �J�n�s, �I���s, �����Čv�Z
    DoEvents
    
    ' �I������ -------------------------------------------------------------------------------------
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�w��͈͂̑��x��������" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�I��"
    #End If
    
'    For �s = �J�n�s To �I���s
'        MAX�\���t���O(�s) = False
'    Next �s
    
End Function

' ======================================================================================================================================================================================================
'
' MAX�\���t���O(�J�n�s to �I���s)���ݒ肳��܂��B
' �߂�l��MAX�\���̂���m�[�c�̏���\��������
' �� (�J�n�s - 1) �s�ڂ��]���ΏۂɂȂ�܂��B�� �J�n�s = 1 �̓_��
' �� �J�n�s���z�[���h���̍s������_��
' �܂��A�I���s�t�߂�MAX�����ނ��ƂŁAMAX�I���s�܂ŕ]���Ώۂ̖����̍s���ǉ������\��������܂��B
' �I���s���ύX���ꂽ�ꍇ�A�����̎Q�Ɠn���ɂ���Č��̕ϐ����ύX����܂��B
' ======================================================================================================================================================================================================

Private Function �u���b�N�����x��������( _
    ByVal �J�n�s As Long, _
    ByRef �I���s As Long, _
    ByRef ���݃z�[���h�J�n�s As Long, _
    ByRef ���݃z�[���h�{�^���� As Long, _
    ByRef ����MAX�\�� As Boolean, _
    Optional ByVal �����Čv�Z As Boolean = False)
    
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�u���b�N�����x��������" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�J�n"
    #End If
    
    Dim �s As Long
    Dim �}�[�N As Long
    
    Dim �{�^���� As Long
    'Dim �z�[���h�ω��t���O As Boolean
    
    �s = �J�n�s
    
    Do Until �s > �I���s
                
        If Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�s) <> Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�s - 1) Then
        
            ' 1. ���݂̑O��̃z�[���h�����瑁�x�܂��̓W���X�g(�^�C�~���O�s��)��ݒ�
            
            �{�^���� = 0
            
            For �}�[�N = 1 To �}�[�N��
                
                If Def.���ʃe�[�u��.�z�[���h�ʊJ�n�t���[����(�}�[�N, �s) > 0 Then
                    
                    �{�^���� = �{�^���� + 1
                    
                End If
                
            Next �}�[�N
            
            If Not Def.���x�蓮�w��t���O(�s) Then
                
                ' ���O�ɃX�^�[�g�����z�[���h��MAX������\��������ꍇ�͕ʈ���
                
                If ����MAX�\�� Then
                    
                    If Def.���ʃe�[�u��.�z�[���h�I���������������(�s) Then
                        
                        If �{�^���� > 0 Then
                            
                            Def.���ʃe�[�u��.���x�蓮�w���(�s) = ��COOL����
                            
                        Else
                            
                            Def.���ʃe�[�u��.���x�蓮�w���(�s) = �xCOOL����
                            
                        End If
                        
                    Else
                        
                        ' ���������Ȃ���(���݂̊�����ł�1���ȏ�]�T��������)MAX������ꍇ�͑��x�w��Ȃ�
                        
                    End If
                    
                Else
                    
                    If �{�^���� > ���݃z�[���h�{�^���� Then
                    
                        Def.���ʃe�[�u��.���x�蓮�w���(�s) = ��COOL����
                        
                    ElseIf �{�^���� = ���݃z�[���h�{�^���� Then
                        
                        Def.���ʃe�[�u��.���x�蓮�w���(�s) = �W���X�gCOOL����
                        
                    ElseIf �{�^���� < ���݃z�[���h�{�^���� Then
                        
                        Def.���ʃe�[�u��.���x�蓮�w���(�s) = �xCOOL����
                        
                    End If
                    
                End If
                
                Def.���ʃe�[�u��.�Čv�Z �s, �I���s, �����Čv�Z
                'DoEvents
                
            End If
            
            ' 2. MAX�\���̂�����̂ɂ���MAX������ꍇ�͂��悭����悤�ɑk���đ��x���C��
            
            If Def.���ʃe�[�u��.�z�[���h�{�[�i�X��(�s) > 0 Then
                
                If ����MAX�\�� Then
                    
                    ' a. MAX�\��������ꍇ�A�܂�����̃^�C�~���O��ύX=�x�����Ď����Ă݂�(�蓮�w�肳��Ă��Ȃ��ꍇ)
                    
                    Dim ���s�O���葁�x�w�� As String
                    Dim ���s�O�������x�w�� As String
                    
                    ���s�O���葁�x�w�� = Def.���ʃe�[�u��.���x�蓮�w���(���݃z�[���h�J�n�s)
                    ���s�O�������x�w�� = Def.���ʃe�[�u��.���x�蓮�w���(�s)
                    
                    If Not Def.���x�蓮�w��t���O(���݃z�[���h�J�n�s) Then
                        
                        If isMAX�^�C�~���O�s��(���݃z�[���h�J�n�s) And (Def.���ʃe�[�u��.���x�蓮�w���(���݃z�[���h�J�n�s) <> �xCOOL����) Then
                            
                            Def.���ʃe�[�u��.���x�蓮�w���(���݃z�[���h�J�n�s) = �W���X�gCOOL����
                            
                        Else
                            
                            Def.���ʃe�[�u��.���x�蓮�w���(���݃z�[���h�J�n�s) = �xCOOL����
                            
                        End If
                        
                        Def.���ʃe�[�u��.�Čv�Z ���݃z�[���h�J�n�s, �I���s, �����Čv�Z
                        'DoEvents
                        
                        ' ����̃^�C�~���O��x���������Ƃł��̍s��MAX�z�[���h�{�[�i�X������Ȃ��Ȃ����ꍇ�́A���݂̍s�̑��x�w����L�����Z�����Ă��̂܂܎��̍s��
                        ' �I���s�������ꍇ�͂���ɂ�����s�ǉ��ŕ]�������
                        
                        If Def.���ʃe�[�u��.�z�[���h�{�[�i�X��(�s) = 0 Then
                        
                            If Not Def.���x�蓮�w��t���O(�s) Then
                                
                                Def.���ʃe�[�u��.���x�蓮�w���(�s) = ""
                                                                
                            End If
                            
                            If �s = �I���s Then
                            
                                �I���s = �I���s + 1
                                Def.���ʃe�[�u��.OwnTable.ListRows(�I���s).Range.Rows.Hidden = False
                                
                            End If
                            
                            Def.���ʃe�[�u��.�Čv�Z �s, �I���s, �����Čv�Z
                            'DoEvents
                            
                        End If
                        
                        ' ���݂̑��x�̏�Ԃł��̍s�œ���{�[�i�X��MAX�łȂ��ꍇ�A����𑁂�����
                        
                        If Def.���ʃe�[�u��.�z�[���h�{�[�i�X��(�s) > 0 And Def.���ʃe�[�u��.�z�[���h�t���[����(�s) <= Def.��MAX�t���[���ő�l Then
                            
                            Def.���ʃe�[�u��.���x�蓮�w���(���݃z�[���h�J�n�s) = ��COOL����
                            
                            Def.���ʃe�[�u��.�Čv�Z ���݃z�[���h�J�n�s, �I���s, �����Čv�Z
                            'DoEvents
                            
                        End If
                        
                    End If
                    
                    ' b. ����̃^�C�~���O�̕ύX�����ł��̍s�œ���{�[�i�X��MAX�ɂȂ�Ȃ��ꍇ�͔����̃^�C�~���O���ύX���Ă݂�
                    
                    If Not Def.���x�蓮�w��t���O(�s) Then
                        
                        If Def.���ʃe�[�u��.�z�[���h�{�[�i�X��(�s) > 0 And Def.���ʃe�[�u��.�z�[���h�t���[����(�s) <= Def.��MAX�t���[���ő�l Then
                            
                            Def.���ʃe�[�u��.���x�蓮�w���(�s) = �xCOOL����
                            
                            Def.���ʃe�[�u��.�Čv�Z �s, �I���s, �����Čv�Z
                            'DoEvents
                            
                        End If
                        
                    End If
                    
                    ' c. ����ł����̍s�œ���{�[�i�X��MAX�ɂȂ�Ȃ��ꍇ��MAX�s�\�Ƃ��A����̑��x�̃^�C�~���O�����ɖ߂�
                    ' �܂��A���̒i�K��MAX�������Ă��A��������̃z�[���h���]�T�̂Ȃ�MAX�ł���ꍇ�ȂǂɁA����MAX���L�����Z�������ꍇ�����邩������Ȃ�
                    
                    If Def.���ʃe�[�u��.�z�[���h�{�[�i�X��(�s) > 0 And Def.���ʃe�[�u��.�z�[���h�t���[����(�s) <= Def.��MAX�t���[���ő�l Then
                        
                        If Not Def.���x�蓮�w��t���O(���݃z�[���h�J�n�s) Then
                        
                            Def.���ʃe�[�u��.���x�蓮�w���(���݃z�[���h�J�n�s) = ���s�O���葁�x�w��
                            
                        End If
                        
                        If Not Def.���x�蓮�w��t���O(�s) Then
                        
                            Def.���ʃe�[�u��.���x�蓮�w���(�s) = ���s�O�������x�w��
                            
                        End If
                        
                        Def.���ʃe�[�u��.�Čv�Z ���݃z�[���h�J�n�s, �I���s, �����Čv�Z
                        'DoEvents
                        
                    End If
                    
                    ' ���݃z�[���h�J�n�s = �J�n�s - 1 �s�ڂ̏ꍇ�A���̍s�̑��x��(����{�I��)�֌W�Ȃ��̂Ō��ɖ߂�
                    ' ��(�J�n�s - 1) �s�ڂ����O�̃z�[���h�J�n�s�������ꍇ�͍l�����Ȃ�?
                    ' �����x���ύX����Ă���̂Ŋm�F�s��
                    
'                    If ���݃z�[���h�J�n�s = �J�n�s - 1 Then
'
'                        ���x�蓮�w���(���݃z�[���h�J�n�s) = ���s�O���葁�x�w��

'                        Def.���ʃe�[�u��.�Čv�Z ���݃z�[���h�J�n�s, �I���s, �����Čv�Z
'                        'DoEvents
'
'                    End If
                    
                End If
                
            End If
            
            ' 3. ������x�z�[���h�ɕω������������`�F�b�N���A�ω����������ꍇ��
            '    ���݃z�[���h�J�n�s�ƃ{�^���������݂̍s�ƃ{�^�����ɍĐݒ�
            
            If Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�s) <> Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�s - 1) Then
            
                ���݃z�[���h�J�n�s = �s
                ���݃z�[���h�{�^���� = �{�^����
                
                ����MAX�\�� = MAX�\���t���O(���݃z�[���h�J�n�s)
                
            End If
            
        End If
        
        �s = �s + 1
        
    Loop
    
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�u���b�N�����x��������" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�I��"
    #End If
    
End Function

Private Function isMAX�^�C�~���O�s��(ByVal �z�[���h�J�n�s As Long) As Boolean
    isMAX�^�C�~���O�s�� = True
    Dim �}�[�N As Long
    For �}�[�N = 1 To Def.�}�[�N��
        If Def.���ʃe�[�u��.�z�[���h�ʊJ�n�t���[����(�}�[�N, �z�[���h�J�n�s) > 0 Then
            If Def.���ʃe�[�u��.�z�[���h�ʊJ�n�t���[����(�}�[�N, �z�[���h�J�n�s) < Def.���ʃe�[�u��.�z�[���h�J�n�t���[����(�z�[���h�J�n�s) Then
                isMAX�^�C�~���O�s�� = False
                Exit For
            End If
        End If
    Next �}�[�N
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Sub �w��͈͂̑��x�w��폜( _
    ByVal �J�n�s As Long, _
    ByVal �I���s As Long, _
    Optional ByVal �����Čv�Z As Boolean = False)
    
    ' ���x�w��폜 ---------------------------------------------------------------------------------
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�w��͈͂̑��x�w�����" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�J�n"
    #End If
    
    Dim �s As Long
    
    For �s = �J�n�s To �I���s
        
        If Not Def.���x�蓮�w��t���O(�s) Then
            
            If Def.���ʃe�[�u��.���x�蓮�w���(�s) <> "" Then
                
                Def.���ʃe�[�u��.���x�蓮�w���(�s) = ""
                
            End If
            
        End If
        
    Next �s
    
    For �s = �J�n�s To �I���s
        MAX�\���t���O(�s) = False
    Next �s
    
    Def.���ʃe�[�u��.�Čv�Z �J�n�s, �I���s, �����Čv�Z
    DoEvents
    
    #If �ؑ֏ڍ׃��O Then
        �ڍ׃��O�o�� "�w��͈͂̑��x�w�����" & vbTab & �J�n�s & vbTab & �I���s & vbTab & "�I��"
    #End If
    
End Sub

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Private Function �ؑ֕�����擾(ByVal �s As Long) As String
    
    Dim �}�[�N As Long
    
    �ؑ֕�����擾 = �m�[�c�ԍ�������̎擾(�s) & "["
    
    For �}�[�N = 1 To �}�[�N��
        If Def.���ʃe�[�u��.�z�[���h�ʊJ�n�t���[����(�}�[�N, �s - 1) > 0 Then
            �ؑ֕�����擾 = �ؑ֕�����擾 & �}�[�N����(�}�[�N)
        End If
    Next �}�[�N
    
    �ؑ֕�����擾 = �ؑ֕�����擾 & "��"
    
    For �}�[�N = 1 To �}�[�N��
        If Def.���ʃe�[�u��.�z�[���h�ʊJ�n�t���[����(�}�[�N, �s) > 0 Then
            �ؑ֕�����擾 = �ؑ֕�����擾 & �}�[�N����(�}�[�N)
        End If
    Next �}�[�N
    
    �ؑ֕�����擾 = �ؑ֕�����擾 & "]"
    
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Private Function ���x�����܂��̓t���[���̎擾(ByVal �s As Long) As String
    If Not Def.���ʃe�[�u��.���x�蓮�w���(�s) = "" Then
        ���x�����܂��̓t���[���̎擾 = "{" & Def.���ʃe�[�u��.���x�蓮�w���(�s) & "}"
    End If
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Private Function �m�[�c�ԍ�������̎擾( _
    ByVal �s As Long, _
    Optional ByVal �R���{0�o�� As Boolean = True, _
    Optional ByVal �R���{1�o�� As Boolean = True) _
    As String
    
    Dim �ؑփm�[�c�ԍ� As Long
    Dim �ؑփR���{�� As Long

    �ؑփm�[�c�ԍ� = Def.���ʃe�[�u��.�m�[�c�ԍ���(�s)
    �ؑփR���{�� = Def.���ʃe�[�u��.�R���{��(�s)
    
    �m�[�c�ԍ�������̎擾 = �ؑփm�[�c�ԍ�
    If �ؑփm�[�c�ԍ� <> �ؑփR���{�� Then
        If (�ؑփR���{�� = 0 And �R���{0�o��) Or (�ؑփR���{�� = 1 And �R���{1�o��) Or �ؑփR���{�� > 1 Then
            �m�[�c�ԍ�������̎擾 = �m�[�c�ԍ�������̎擾 & "<" & �ؑփR���{�� & ">"
        End If
    End If
    
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Sub Rescue()
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Private Function �ڍ׃��O������擾(ByVal �o�͕����� As String)
    Dim time As Double
    time = Timer
    �ڍ׃��O������擾 = "TIME:" & vbTab & Format(Now, "yyyy-mm-ddThh:nn:ss") & Format(time - Int(time), ".000") & vbTab & �o�͕�����
End Function

Private Function �ڍ׃��O�o��(ByVal �o�͕����� As String, Optional ByVal is���[�U�[�o�� As Boolean = False, Optional ByVal �o�͍s�ԍ� As Long = -1)
    If is���[�U�[�o�� Then
        Def.�������O.�o�� �ڍ׃��O������擾(�o�͕�����), True, �o�͍s�ԍ�
    Else
        Debug.Print �ڍ׃��O������擾(�o�͕�����)
    End If
End Function

