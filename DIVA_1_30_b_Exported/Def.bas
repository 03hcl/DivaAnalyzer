Attribute VB_Name = "Def"
Option Explicit
Option Base 1

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Const �}�[�N�� As Long = 4
Public Const �X���C�h�}�[�N�� As Long = 2

Public Const ��MAX�t���[���ő�l As Long = 300

Public �}�[�N����(�}�[�N��) As String
Public �X���C�h�}�[�N����(�X���C�h�}�[�N��) As String

Public ��COOL���� As String
Public �W���X�gCOOL���� As String
Public �xCOOL���� As String

Public HOLD���� As String

Public ��WRONG���� As String
Public WORST���� As String
Public MISSTAKE���� As String

Public �X�R�A�^�X�L�b�v�_ As Long
Public CC�ő�]�T�t���[���� As Long
Public �ő僉�C�t�� As Long

Public �]������ As Dictionary

Public ���ʃe�[�u�� As IChartTable
Public �������O As ProcessLog
Public �ؑ֌��ʃe�[�u�� As SwitchingTable

Public ���x�ؑֈꗗ�e�[�u�� As ElSwTable

Public �X�R�A�^�e�[�u�� As ScoreRouteTable

Public ���݃z�[���h�u���b�N As Long
Public ���x�蓮�w��t���O() As Boolean
Public MAX�\���t���O() As Boolean

Public Type �ʐؑփf�[�^
    �X�R�A As Long
    �ؑ֍s���X�g() As Long
End Type

Public Type �]���Z�b�g
    �J�n�t���[������ As Long
    �J�n�]�� As String
    �I���t���[������ As Long
    �I���]�� As String
    ���픻��g As Boolean
End Type

Public Type �s�X�R�A���
    �ő�e���s As Long
    �s�X�R�A() As Long
End Type

Public Type ���ʃZ�b�g
    �N���A�����N As String
    �B���� As Double
    �X�R�A As Long
End Type

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function �ؑ։�͏����ݒ���s(Optional ByVal ���ʃe�[�u���� As String = "") As Long
    
    Application.StatusBar = "�ؑ։�͂̏����ݒ���s���܂�......"
    
    �ؑ։�͏����ݒ���s = -1
    
    Def.�}�[�N�����ݒ�
    Def.�����萔�ݒ�
    
    If Def.���ʃe�[�u���ݒ�(���ʃe�[�u����) < 0 Then
        Exit Function
    End If
    
    If Def.�ؑ֌��ʃe�[�u���ݒ�(���ʃe�[�u��) < 0 Then
        Exit Function
    End If
    
    If Def.�������O�o�͐ݒ�() < 0 Then
        Exit Function
    End If
    
    Dim �f�[�^�s�� As Long
    �f�[�^�s�� = Def.���ʃe�[�u��.�f�[�^�s��
    
    ReDim ���x�蓮�w��t���O(�f�[�^�s��)
    ReDim MAX�\���t���O(�f�[�^�s��)
    
    Dim �s As Long
    For �s = 1 To �f�[�^�s��
        ���x�蓮�w��t���O(�s) = _
            Def.���ʃe�[�u��.���x�蓮�w���(�s) <> "" Or Def.���ʃe�[�u��.���x�t���[���蓮�w���(�s) <> ""
        MAX�\���t���O(�s) = False
    Next �s
    
    Def.�������O.�o�� "�ő�COOL: " & ���ʃe�[�u��.�ő�COOL�t���[�� & " / �ŒxCOOL: " & ���ʃe�[�u��.�ŒxCOOL�t���[��
    
    If MsgBox("���݂̃e�[�u���ŏ������J�n���܂��B��낵���ł����H" & vbCrLf & _
        "���ʃe�[�u����: " & Def.���ʃe�[�u��.OwnTable.name, vbOKCancel + vbInformation) = vbOK Then
        �ؑ։�͏����ݒ���s = 0
    Else
        MsgBox "�����𒆎~���܂��B", vbCritical
    End If
    
    Application.StatusBar = "�ؑ։�͂̏����ݒ肪�������܂����B"
    
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function �X�R�A�^��͏����ݒ���s(Optional ByVal ���ʃe�[�u���� As String = "") As Long
    
    Application.StatusBar = "�X�R�A�^��͂̏����ݒ���s���܂�......"
    
    �X�R�A�^��͏����ݒ���s = -1
    
    Def.�}�[�N�����ݒ�
    Def.�����萔�ݒ�
    
    If Def.�X�R�A�^��͗p�萔�ݒ� < 0 Then
        Exit Function
    End If
    
    If Def.���ʃe�[�u���ݒ�(���ʃe�[�u����) < 0 Then
        Exit Function
    End If
    
    Dim �ؑ֌��ʃe�[�u���� As String
    Set Def.�ؑ֌��ʃe�[�u�� = New SwitchingTable
    �ؑ֌��ʃe�[�u���� = ���ʃe�[�u��.OwnTable.name & "_�ؑ�"
    
    Do While Def.�ؑ֌��ʃe�[�u��.�I�u�W�F�N�g�ݒ�(�ؑ֌��ʃe�[�u����) < 0
        �ؑ֌��ʃe�[�u���� = InputBox("�ؑ֌��ʃe�[�u���������Ŕ����ł��܂���B" & vbCrLf & "�e�[�u��������͂��Ă��������B")
        If �ؑ֌��ʃe�[�u���� = "" Then
            If MsgBox("�e�[�u�����������͂ł��B�������I�����܂����H", vbOKCancel + vbQuestion) = vbOK Then
                GoTo �e�[�u���̐ݒ�𒆎~�����ꍇ
            End If
        End If
    Loop
    
    Dim ���x�ؑֈꗗ�e�[�u���� As String
    Set Def.���x�ؑֈꗗ�e�[�u�� = New ElSwTable
    ���x�ؑֈꗗ�e�[�u���� = ���ʃe�[�u��.OwnTable.name & "_���x�ؑփ��X�g"
    
    Do While Def.���x�ؑֈꗗ�e�[�u��.�I�u�W�F�N�g�ݒ�(���x�ؑֈꗗ�e�[�u����) < 0
        ���x�ؑֈꗗ�e�[�u���� = InputBox("���x�Ɛؑւ̈ꗗ�e�[�u���������Ŕ����ł��܂���B" & vbCrLf & "�e�[�u��������͂��Ă��������B")
        If ���x�ؑֈꗗ�e�[�u���� = "" Then
            If MsgBox("�e�[�u�����������͂ł��B�������I�����܂����H", vbOKCancel + vbQuestion) = vbOK Then
                GoTo �e�[�u���̐ݒ�𒆎~�����ꍇ
            End If
        End If
    Loop
    
    Def.���x�ؑֈꗗ�e�[�u��.���x�ؑփ��X�g�ݒ� Def.���ʃe�[�u��, True
    
    If Def.�X�R�A�^�e�[�u���ݒ�(Def.���ʃe�[�u��) < 0 Then
        Exit Function
    End If
    
    If Def.�������O�o�͐ݒ�() < 0 Then
        Exit Function
    End If
    
    Dim �f�[�^�s�� As Long
    �f�[�^�s�� = Def.���ʃe�[�u��.�f�[�^�s��
    
    ReDim ���x�蓮�w��t���O(�f�[�^�s��)
    ReDim MAX�\���t���O(�f�[�^�s��)
    
    Dim �s As Long
    For �s = 1 To �f�[�^�s��
        ���x�蓮�w��t���O(�s) = (Def.���ʃe�[�u��.���x�t���[���蓮�w���(�s) <> "")
        MAX�\���t���O(�s) = False
    Next �s
    
    Def.�������O.�o�� "�ő�COOL: " & ���ʃe�[�u��.�ő�COOL�t���[�� & " / �ŒxCOOL: " & ���ʃe�[�u��.�ŒxCOOL�t���[��
    
    If MsgBox("���L�̃e�[�u�����Ɉ�v����e�[�u�����g���ď������J�n���܂��B��낵���ł����H" & vbCrLf & _
        "���ʃe�[�u����: " & Def.���ʃe�[�u��.OwnTable.name & vbCrLf & _
        "�ؑ֌��ʃe�[�u����: " & Def.�ؑ֌��ʃe�[�u��.OwnTable.name & vbCrLf & _
        "���x�Ɛؑւ̈ꗗ�e�[�u����: " & Def.���x�ؑֈꗗ�e�[�u��.OwnTable.name, vbOKCancel + vbInformation) = vbOK Then
        �X�R�A�^��͏����ݒ���s = 0
    Else
        MsgBox "�����𒆎~���܂��B", vbCritical
    End If
    
    Application.StatusBar = "�X�R�A�^��͂̏����ݒ肪�������܂����B"
    Exit Function
    
�e�[�u���̐ݒ�𒆎~�����ꍇ:
    MsgBox "�e�[�u����ݒ�ł��Ȃ����߁A�����𒆎~���܂��B", vbCritical
    Exit Function
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function �ؑ։�͏I���ݒ���s( _
    ByVal �J�n���� As Long, _
    ByVal �I������ As Long, _
    ByVal ��ʍX�V As Boolean, _
    ByVal �����Čv�Z As Boolean) _
    As Long
    
    Application.StatusBar = "�ؑ։�͂̏I���������s���܂�......"
        
    Def.�ؑ֌��ʃe�[�u��.�I�u�W�F�N�g�ŏI���`
    
    Dim �őP�ؑ֏�� As Switching
    Set �őP�ؑ֏�� = Def.�ؑ֌��ʃe�[�u��.�őP�ؑ֏��擾(���ʃe�[�u��, �����Čv�Z)
    
    Def.�������O.�o�� "��͎���: " & CDbl(�I������ - �J�n����) / 1000 & " �b (��ʍX�V: " & ��ʍX�V & ", �����Čv�Z: " & �����Čv�Z & ")"
    Def.�������O.�o�� �őP�ؑ֏��.�ؑ֕�����, False
    
    Def.�������O.�t�@�C���o�͏I��
    
    Dim cb As New dataobject
    cb.SetText �őP�ؑ֏��.�ؑ֕�����
    cb.PutInClipboard
    Set cb = Nothing
    
    Dim MAX�\�� As OutputString
    If MsgBox("��͂ɂ��ؑ֌��ʂ͈ȉ��̒ʂ�ł��B" & vbCrLf & _
        "(���̕�����̓N���b�v�{�[�h�ɃR�s�[����Ă��܂��B)" & vbCrLf & vbCrLf & _
        �őP�ؑ֏��.�ؑ֕����� & vbCrLf & vbCrLf & _
        "��͎���: " & CDbl(�I������ - �J�n����) / 1000 & " �b" & vbCrLf & vbCrLf & _
        "���̐ؑ֏��𕈖ʃe�[�u���ɔ��f�����A���̏�����ɑ��x��ݒ肵�܂����H", _
        vbYesNo + vbInformation) = vbYes Then
        
        Set MAX�\�� = �őP�ؑ֏��.�֑ؑ��x��񔽉f(�����Čv�Z)
        Def.�������O.�o�� MAX�\��.�\��������
        
        If Def.���x�ؑֈꗗ�e�[�u���ݒ�(���ʃe�[�u��) < 0 Then
            Exit Function
        End If
        Def.���x�ؑֈꗗ�e�[�u��.���x�ؑ֏��ǂݍ��� Def.���ʃe�[�u��
        
    End If
    
    �ؑ։�͏I���ݒ���s = 0
    
    Application.StatusBar = "�ؑ։�͂̏I���������������܂����B"
    
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function �X�R�A�^��͏I���������s( _
    ByVal �J�n���� As Long, _
    ByVal �I������ As Long, _
    ByVal ��ʍX�V As Boolean, _
    ByVal �����Čv�Z As Boolean) _
    As Long
    
    Application.StatusBar = "�X�R�A�^��͂̏I���������s���܂�......"
    
    Def.�X�R�A�^�e�[�u��.�I�u�W�F�N�g�ŏI���` Def.���ʃe�[�u��
    
'    Def.���x�ؑֈꗗ�e�[�u��.���x�ؑ֏�񏑂��o�� Def.���ʃe�[�u��
    
    Def.�������O.�o�� "��͎���: " & CDbl(�I������ - �J�n����) / 1000 & " �b (��ʍX�V: " & ��ʍX�V & ", �����Čv�Z: " & �����Čv�Z & ")"
'    Def.�������O.�o�� �őP�ؑ֏��.�ؑ֕�����, False
    
    Def.�������O.�t�@�C���o�͏I��
    
'    Dim cb As New dataobject
'    cb.SetText "" '�N���b�v�{�[�h
'    cb.PutInClipboard
'    Set cb = Nothing
    
    �X�R�A�^��͏I���������s = 0
    
    Application.StatusBar = "�X�R�A�^��͂̏I���������������܂����B"
    
End Function


' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function ������A��(ByVal ������1 As String, ByVal ������2 As String, Optional ByVal �ڑ������� As String = vbCrLf) As String
    If ������1 <> "" Then
        If ������2 <> "" Then
            ������A�� = ������1 & �ڑ������� & ������2
        Else
            ������A�� = ������1
        End If
    Else
        If ������2 <> "" Then
            ������A�� = ������2
        Else
            ������A�� = ""
        End If
    End If
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function �e�[�u���ƃV�[�g�̌����ݒ�(ByRef �e�[�u�� As ListObject, ByRef �V�[�g As Worksheet, Optional ByVal �e�[�u���� As String = "") As Long

    On Error GoTo �e�[�u���𔭌��ł��Ȃ��ꍇ
    
    If �e�[�u���� = "" Then
        Set �e�[�u�� = ActiveSheet.ListObjects(1)
        Set �V�[�g = ActiveSheet
    Else
        Dim sheet As Worksheet
        Dim list As ListObject
        For Each sheet In ThisWorkbook.Worksheets
            For Each list In sheet.ListObjects
                If list.name = �e�[�u���� Then
                    Set �e�[�u�� = list
                    Set �V�[�g = sheet
                End If
            Next list
        Next sheet
        If �V�[�g Is Nothing Then
            GoTo �e�[�u���𔭌��ł��Ȃ��ꍇ
        End If
    End If
    
    On Error GoTo 0
    
    �e�[�u���ƃV�[�g�̌����ݒ� = 0
    Exit Function
    
�e�[�u���𔭌��ł��Ȃ��ꍇ:

    �e�[�u���ƃV�[�g�̌����ݒ� = -1
    Exit Function

End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function ���ʃe�[�u�������ݒ�(Optional ByVal ���ʃe�[�u���� As String = "") As Long
    
    Def.���ʃe�[�u���ݒ� ���ʃe�[�u����, "���ʃe�[�u���𔭌��ł��܂���B"
    
    Do While ���ʃe�[�u��.OwnTable Is Nothing
        ���ʃe�[�u���� = InputBox("���ʃe�[�u���������Ŕ����ł��܂���B" & vbCrLf & "���ʃe�[�u��������͂��Ă��������B")
        If ���ʃe�[�u���� = "" Then
            If MsgBox("���ʃe�[�u�����������͂ł��B�������I�����܂����H", vbOKCancel + vbQuestion) = vbOK Then
                GoTo �e�[�u���̐ݒ�𒆎~�����ꍇ
            End If
        End If
        Def.���ʃe�[�u���ݒ� ���ʃe�[�u����, "���ʃe�[�u���𔭌��ł��܂���B"
    Loop
    
    If MsgBox("���L�̃e�[�u�����̃e�[�u���ŏ������J�n���܂��B��낵���ł����H" & vbCrLf & _
        "�e�[�u����: " & Def.���ʃe�[�u��.OwnTable.name, vbOKCancel + vbInformation) <> vbOK Then
        GoTo �e�[�u���̐ݒ�𒆎~�����ꍇ
    End If
    
    ���ʃe�[�u�������ݒ� = 0
    Exit Function
    
�e�[�u���̐ݒ�𒆎~�����ꍇ:
    
    MsgBox "�����𒆎~���܂����B", vbCritical
    ���ʃe�[�u�������ݒ� = -1
    Exit Function
    
End Function

Public Function �}�[�N�����ݒ�()
    Def.�}�[�N����(1) = "��"
    Def.�}�[�N����(2) = "��"
    Def.�}�[�N����(3) = "�~"
    Def.�}�[�N����(4) = "��"
    Def.�X���C�h�}�[�N����(1) = "��"
    Def.�X���C�h�}�[�N����(2) = "��"
End Function

Public Function �����萔�ݒ�()
    Def.��COOL���� = ThisWorkbook.Names("EarlyCoolString").RefersToRange.value
    Def.�W���X�gCOOL���� = ThisWorkbook.Names("JustCoolString").RefersToRange.value
    Def.�xCOOL���� = ThisWorkbook.Names("LateCoolString").RefersToRange.value
    Def.HOLD���� = ThisWorkbook.Names("HoldMarker").RefersToRange.value
End Function

Public Function �X�R�A�^��͗p�萔�ݒ�() As Long
    
    Def.��WRONG���� = "��WRONG"
    Def.WORST���� = "WORST"
    
    Def.MISSTAKE���� = "MISS�~TAKE"
    
    Def.�X�R�A�^�X�L�b�v�_ = ThisWorkbook.Names("StoppingScoreAttackGap").RefersToRange.value
    If ThisWorkbook.Names("MaxDelayFrame").RefersToRange.value = "" Then
        Def.CC�ő�]�T�t���[���� = -Def.��MAX�t���[���ő�l
    Else
        Def.CC�ő�]�T�t���[���� = ThisWorkbook.Names("MaxDelayFrame").RefersToRange.value
    End If
    Def.�ő僉�C�t�� = ThisWorkbook.Names("MaximumLife").RefersToRange.value
    
    Dim �]���e�[�u�� As ListObject
    Set �]���e�[�u�� = ThisWorkbook.Worksheets("Rating").ListObjects("RatingTable")
    Dim rate�s As Long
    
    On Error GoTo �]�������̐ݒ�Ɏ��s�����ꍇ
    
    Set �]������ = New Dictionary
    For rate�s = 1 To �]���e�[�u��.ListRows.count
            �]������.Add �]���e�[�u��.ListColumns("Sign").DataBodyRange(rate�s).value, �]���e�[�u��.ListColumns("Small Sign").DataBodyRange(rate�s).value
    Next rate�s
    
    On Error GoTo 0
    
    �X�R�A�^��͗p�萔�ݒ� = 0
    Exit Function
    
�]�������̐ݒ�Ɏ��s�����ꍇ:
    
    MsgBox "ERR:�]�������̐ݒ�Ɏ��s���܂����B" & vbCrLf & "�������I�����܂��B", vbCritical
    
    �X�R�A�^��͗p�萔�ݒ� = -1
    Exit Function
    
End Function

Public Function ���ʃe�[�u���ݒ�(Optional ByVal ���ʃe�[�u���� As String = "", Optional ByVal �G���[������ As String = "") As Long
    If MsgBox("���ʃe�[�u���̉��z�����s���܂����H", vbYesNo + vbInformation) = vbYes Then
        Set Def.���ʃe�[�u�� = New ChartTable2
    Else
        Set Def.���ʃe�[�u�� = New ChartTable
    End If
    ���ʃe�[�u���ݒ� = Def.���ʃe�[�u��.�I�u�W�F�N�g�ݒ�(���ʃe�[�u����)
    If ���ʃe�[�u���ݒ� < 0 Then
        GoTo �e�[�u���ݒ�Ɏ��s�����ꍇ
    End If
    Exit Function
�e�[�u���ݒ�Ɏ��s�����ꍇ:
    If �G���[������ = "" Then
        �G���[������ = "ERR:�������s���e�[�u�����������܂���ł����B" & vbCrLf & "�������I�����܂��B"
    End If
    If ���ʃe�[�u���� <> "" Then
        �G���[������ = �G���[������ & vbCrLf & "�e�[�u����: " & ���ʃe�[�u����
    End If
    MsgBox �G���[������, vbCritical
    Exit Function
End Function

Public Function �������O�o�͐ݒ�() As Long
    Set Def.�������O = New ProcessLog
    �������O�o�͐ݒ� = Def.�������O.�t�@�C���o�͊J�n()
    If �������O�o�͐ݒ� < 0 Then
        GoTo ���O�o�͐ݒ�Ɏ��s�����ꍇ
    End If
    Exit Function
���O�o�͐ݒ�Ɏ��s�����ꍇ:
    MsgBox "ERR:���O�t�@�C�����쐬�ł��܂���B" & vbCrLf & "�������I�����܂��B", vbCritical
    Exit Function
End Function

Public Function �ؑ֌��ʃe�[�u���ݒ�(ByVal ���ʃe�[�u�� As IChartTable) As Long
    Set Def.�ؑ֌��ʃe�[�u�� = New SwitchingTable
    �ؑ֌��ʃe�[�u���ݒ� = Def.�ؑ֌��ʃe�[�u��.�I�u�W�F�N�g�V�K�쐬(���ʃe�[�u��)
    If �ؑ֌��ʃe�[�u���ݒ� < 0 Then
        GoTo �e�[�u���ݒ�Ɏ��s�����ꍇ
    End If
    Exit Function
�e�[�u���ݒ�Ɏ��s�����ꍇ:
    MsgBox "ERR:���ʃe�[�u���̍쐬�Ɏ��s���܂����B" & vbCrLf & "�������I�����܂��B", vbCritical
    Exit Function
End Function

Public Function ���x�ؑֈꗗ�e�[�u���ݒ�(ByVal ���ʃe�[�u�� As IChartTable) As Long
    Set Def.���x�ؑֈꗗ�e�[�u�� = New ElSwTable
    ���x�ؑֈꗗ�e�[�u���ݒ� = Def.���x�ؑֈꗗ�e�[�u��.�I�u�W�F�N�g�V�K�쐬(���ʃe�[�u��)
    If ���x�ؑֈꗗ�e�[�u���ݒ� < 0 Then
        GoTo �e�[�u���ݒ�Ɏ��s�����ꍇ
    End If
    Exit Function
�e�[�u���ݒ�Ɏ��s�����ꍇ:
    MsgBox "ERR:���x�Ɛؑւ̈ꗗ�e�[�u���̍쐬�Ɏ��s���܂����B" & vbCrLf & "�������I�����܂��B", vbCritical
    Exit Function
End Function

Public Function �X�R�A�^�e�[�u���ݒ�(ByVal ���ʃe�[�u�� As IChartTable) As Long
    Set Def.�X�R�A�^�e�[�u�� = New ScoreRouteTable
    �X�R�A�^�e�[�u���ݒ� = Def.�X�R�A�^�e�[�u��.�I�u�W�F�N�g�V�K�쐬(���ʃe�[�u��)
    If �X�R�A�^�e�[�u���ݒ� < 0 Then
        GoTo �e�[�u���ݒ�Ɏ��s�����ꍇ
    End If
    Exit Function
�e�[�u���ݒ�Ɏ��s�����ꍇ:
    MsgBox "ERR:�X�R�A�^���[�g�̌��ʃe�[�u���̍쐬�Ɏ��s���܂����B" & vbCrLf & "�������I�����܂��B", vbCritical
    Exit Function
End Function

