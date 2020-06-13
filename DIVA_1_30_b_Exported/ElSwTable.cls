VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ElSwTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private m_OwnTable As ListObject
Private m_Sheet As Worksheet

Private m_�ԍ��� As ListColumn
Private m_�s�ԍ��� As ListColumn
Private m_�]���� As ListColumn
Private m_���x�� As ListColumn
Private m_���x�t���[���� As ListColumn
Private m_�ؑ֗� As ListColumn

Private m_�]�����X�g() As String
Private m_���x���X�g() As String
Private m_���x�t���[�����X�g() As String
Private m_�ؑփ��X�g() As String

Public Function �I�u�W�F�N�g�V�K�쐬(Optional ByVal ���ʃe�[�u�� As IChartTable) As Long
    
    Me.�I�u�W�F�N�g���
    
    ' �V�[�g�쐬 -----------------------------------------------------------------------------------
    
    Application.ScreenUpdating = False
    
    Dim �A�N�e�B�u�V�[�g As Worksheet
    Set �A�N�e�B�u�V�[�g = ThisWorkbook.ActiveSheet
    
    On Error GoTo �V�[�g���쐬�ł��Ȃ��ꍇ
    
    Set m_Sheet = Worksheets.Add(after:=���ʃe�[�u��.�����V�[�g)
        
    On Error GoTo �V�[�g���ύX�G���[
    
    Dim ���ʃV�[�g�� As String
    ���ʃV�[�g�� = ���ʃe�[�u��.�����V�[�g.name & "_���x�ؑփ��X�g"
    Do Until Me.�����V�[�g.name = ���ʃV�[�g��
        Me.�����V�[�g.name = ���ʃV�[�g��
    Loop
    
    ' �e�[�u���쐬 ---------------------------------------------------------------------------------
    
    On Error GoTo �e�[�u�����쐬�ł��Ȃ��ꍇ
    
    Set m_OwnTable = Me.�����V�[�g.ListObjects.Add(xlSrcRange, Me.�����V�[�g.Range(Me.�����V�[�g.Cells(1, 1), Me.�����V�[�g.Cells(1, 6)), , xlYes)
    
    On Error GoTo �e�[�u�����ύX�G���[
    
    Dim ���ʃe�[�u���� As String
    ���ʃe�[�u���� = ���ʃe�[�u��.OwnTable.name & "_���x�ؑփ��X�g"
    Do Until Me.OwnTable.name = ���ʃe�[�u����
        Me.OwnTable.name = ���ʃe�[�u����
    Loop
    
    ' �e�[�u�����` ---------------------------------------------------------------------------------
    
    On Error GoTo 0
    
    Me.OwnTable.ShowTableStyleFirstColumn = True
    Me.OwnTable.TableStyle = ���ʃe�[�u��.OwnTable.TableStyle
    
    Me.OwnTable.ListColumns(1).name = "No"
    Me.OwnTable.ListColumns(2).name = "�s"
    Me.OwnTable.ListColumns(3).name = "�]��"
    Me.OwnTable.ListColumns(4).name = "���x"
    Me.OwnTable.ListColumns(5).name = "���xFr"
    Me.OwnTable.ListColumns(6).name = "�ؑ�"
    
    Me.OwnTable.ListColumns(1).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(2).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(3).Range.ColumnWidth = 8
    Me.OwnTable.ListColumns(4).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(5).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(6).Range.ColumnWidth = 6
    
    Set m_�ԍ��� = Me.OwnTable.ListColumns(1)
    Set m_�s�ԍ��� = Me.OwnTable.ListColumns(2)
    Set m_�]���� = Me.OwnTable.ListColumns(3)
    Set m_���x�� = Me.OwnTable.ListColumns(4)
    Set m_���x�t���[���� = Me.OwnTable.ListColumns(5)
    Set m_�ؑ֗� = Me.OwnTable.ListColumns(6)
    
    �A�N�e�B�u�V�[�g.Activate
    
    Application.ScreenUpdating = True
    
    �I�u�W�F�N�g�V�K�쐬 = 0
    Exit Function
    
�V�[�g���ύX�G���[:
    Dim ���ʃV�[�g���A�� As Long
    ���ʃV�[�g���A�� = ���ʃV�[�g���A�� + 1
    ���ʃV�[�g�� = ���ʃe�[�u��.�����V�[�g.name & "_���x�ؑփ��X�g_" & ���ʃV�[�g���A��
    Resume Next
    
�e�[�u�����ύX�G���[:
    Dim ���ʃe�[�u�����A�� As Long
    ���ʃe�[�u�����A�� = ���ʃe�[�u�����A�� + 1
    ���ʃe�[�u���� = ���ʃe�[�u��.OwnTable.name & "_���x�ؑփ��X�g_" & ���ʃe�[�u�����A��
    Resume Next
    
�V�[�g���쐬�ł��Ȃ��ꍇ:
    Application.ScreenUpdating = True
    �I�u�W�F�N�g�V�K�쐬 = -1
    GoTo �G���[����
    
�e�[�u�����쐬�ł��Ȃ��ꍇ:
    �I�u�W�F�N�g�V�K�쐬 = -2
    GoTo �G���[����
    
�G���[����:
    Me.�I�u�W�F�N�g���
    Exit Function
    
End Function

Public Function �I�u�W�F�N�g�ݒ�(Optional ByVal �e�[�u���� As String = "") As Long
    
    Me.�I�u�W�F�N�g���
    
    ' �e�[�u���ݒ� ---------------------------------------------------------------------------------
    
    On Error GoTo �e�[�u���𔭌��ł��Ȃ��ꍇ
    
    If Def.�e�[�u���ƃV�[�g�̌����ݒ�(m_OwnTable, m_Sheet, �e�[�u����) < 0 Then
        GoTo �e�[�u���𔭌��ł��Ȃ��ꍇ
    End If
    
    ' �����o�ݒ� -----------------------------------------------------------------------------------
    
    On Error GoTo �e�[�u���̃t�H�[�}�b�g�`�����قȂ�ꍇ
    
    Set m_�ԍ��� = Me.OwnTable.ListColumns("No")
    Set m_�s�ԍ��� = Me.OwnTable.ListColumns("�s")
    Set m_�]���� = Me.OwnTable.ListColumns("�]��")
    Set m_���x�� = Me.OwnTable.ListColumns("���x")
    Set m_���x�t���[���� = Me.OwnTable.ListColumns("���xFr")
    Set m_�ؑ֗� = Me.OwnTable.ListColumns("�ؑ�")
    
    On Error GoTo 0
    
    �I�u�W�F�N�g�ݒ� = 0
    Exit Function
    
    ' �G���[���x�� ---------------------------------------------------------------------------------
    
�e�[�u���𔭌��ł��Ȃ��ꍇ:
    �I�u�W�F�N�g�ݒ� = -1
    GoTo �G���[����
    
�e�[�u���̃t�H�[�}�b�g�`�����قȂ�ꍇ:
    �I�u�W�F�N�g�ݒ� = -2
    GoTo �G���[����
    
�G���[����:
    Me.�I�u�W�F�N�g���
    Exit Function
    
End Function

Public Sub �I�u�W�F�N�g���()
    
    Set m_OwnTable = Nothing
    Set m_Sheet = Nothing
    
    Set m_�ԍ��� = Nothing
    Set m_�s�ԍ��� = Nothing
    Set m_�]���� = Nothing
    Set m_���x�� = Nothing
    Set m_���x�t���[���� = Nothing
    Set m_�ؑ֗� = Nothing
    
    Erase m_�]�����X�g
    Erase m_���x���X�g
    Erase m_���x�t���[�����X�g
    Erase m_�ؑփ��X�g
    
End Sub

Public Function �I�u�W�F�N�g�ŏI���`(ByVal ���ʃe�[�u�� As IChartTable)
    
    Me.OwnTable.HeaderRowRange.HorizontalAlignment = xlCenter
    
    m_�]����.DataBodyRange.HorizontalAlignment = xlCenter
    m_���x��.DataBodyRange.HorizontalAlignment = xlCenter
    
    m_���x��.DataBodyRange.Validation.Add ���ʃe�[�u��.���x�蓮�w��e�[�u����.DataBodyRange.Validation.Type, Formula1:=���ʃe�[�u��.���x�蓮�w��e�[�u����.DataBodyRange.Validation.Formula1
    m_�ؑ֗�.DataBodyRange.Validation.Add ���ʃe�[�u��.�ؑ֔���e�[�u����.DataBodyRange.Validation.Type, Formula1:=���ʃe�[�u��.�ؑ֔���e�[�u����.DataBodyRange.Validation.Formula1
    
    Me.OwnTable.ShowAutoFilterDropDown = False
    
End Function

Public Function ���x�ؑ֏��ǂݍ���(ByVal ���ʃe�[�u�� As IChartTable, _
    Optional ByVal �J�n�s As Long = 1, _
    Optional ByVal �I���s As Long = -1, _
    Optional ByVal �X�e�[�^�X�o�[�i���o�� As Boolean = True)
    
    If �X�e�[�^�X�o�[�i���o�� Then
        Application.StatusBar = "���x�Ɛؑւ̈ꗗ�𕈖ʂ���V�[�g�ɏo�͂��Ă��܂�...... (������)"
    End If
    
    If �I���s = -1 Then
        �I���s = ���ʃe�[�u��.�f�[�^�s��
    End If
    
    Dim �s As Long
    Dim is���x�ؑ֎w�� As Boolean
    
    For �s = 1 To ���ʃe�[�u��.�f�[�^�s��
        
        If �X�e�[�^�X�o�[�i���o�� Then
            Application.StatusBar = "���x�Ɛؑւ̈ꗗ�𕈖ʂ���V�[�g�ɏo�͂��Ă��܂�...... ( " & �s & " / " & ���ʃe�[�u��.�f�[�^�s�� & " �s�� )"
        End If
        
        If ���ʃe�[�u��.�]����(�s) <> ���ʃe�[�u��.�t���[������ʕ]��(0) Then
            is���x�ؑ֎w�� = True
        ElseIf ���ʃe�[�u��.���x�蓮�w���(�s) <> "" Then
            is���x�ؑ֎w�� = True
        ElseIf ���ʃe�[�u��.���x�t���[���蓮�w���(�s) <> "" Then
            is���x�ؑ֎w�� = True
        ElseIf ���ʃe�[�u��.�ؑ֔��蕶����(�s) <> "" Then
            is���x�ؑ֎w�� = True
        Else
            is���x�ؑ֎w�� = False
        End If
        
        If is���x�ؑ֎w�� Then
            Me.�o�͍s�ǉ�
            �s�ԍ���(�f�[�^�s��) = �s
            If ���ʃe�[�u��.�]����(�s) <> ���ʃe�[�u��.�t���[������ʕ]��(0) Then
                �]����(�f�[�^�s��) = ���ʃe�[�u��.�]����(�s)
            End If
            ���x��(�f�[�^�s��) = ���ʃe�[�u��.���x�蓮�w���(�s)
            ���x�t���[����(�f�[�^�s��) = ���ʃe�[�u��.���x�t���[���蓮�w���(�s)
            �ؑ֗�(�f�[�^�s��) = ���ʃe�[�u��.�ؑ֔��蕶����(�s)
        End If
        
    Next �s
    
    DoEvents
    
    If �X�e�[�^�X�o�[�i���o�� Then
        Application.StatusBar = "���x�Ɛؑւ̈ꗗ�𕈖ʂ���V�[�g�ɏo�͂��Ă��܂�...... (�I����)"
    End If
    
    Me.�I�u�W�F�N�g�ŏI���` ���ʃe�[�u��
    Me.���x�ؑփ��X�g�ݒ� ���ʃe�[�u��
    
End Function

Public Function ���x�ؑ֏�񏑂��o��(ByVal ���ʃe�[�u�� As IChartTable, _
    Optional ByVal �J�n�s As Long = 1, _
    Optional ByVal �I���s As Long = -1, _
    Optional ByVal �X�e�[�^�X�o�[�i���o�� As Boolean = True)
    
    If �X�e�[�^�X�o�[�i���o�� Then
        Application.StatusBar = "���x�Ɛؑւ̈ꗗ���V�[�g���畈�ʂɔ��f���Ă��܂�...... (������)"
    End If
    
    Me.���x�ؑփ��X�g�ݒ� ���ʃe�[�u��
    
    If �X�e�[�^�X�o�[�i���o�� Then
        DoEvents
    End If
    
    If �I���s = -1 Then
        �I���s = ���ʃe�[�u��.�f�[�^�s��
    End If
    
    Dim �s As Long
    
    For �s = �J�n�s To �I���s
        
        If �X�e�[�^�X�o�[�i���o�� Then
            Application.StatusBar = "���x�Ɛؑւ̈ꗗ���V�[�g���畈�ʂɔ��f���Ă��܂�...... ( " & �s & " / " & ���ʃe�[�u��.�f�[�^�s�� & " �s�� )"
        End If
        
        If ���ʃe�[�u��.�]����(�s) <> m_�]�����X�g(�s) Then
            ���ʃe�[�u��.�]����(�s) = m_�]�����X�g(�s)
        End If
        If ���ʃe�[�u��.���x�蓮�w���(�s) <> m_���x���X�g(�s) Then
            ���ʃe�[�u��.���x�蓮�w���(�s) = m_���x���X�g(�s)
        End If
        If ���ʃe�[�u��.���x�t���[���蓮�w���(�s) <> m_���x�t���[�����X�g(�s) Then
            ���ʃe�[�u��.���x�t���[���蓮�w���(�s) = m_���x�t���[�����X�g(�s)
        End If
        If ���ʃe�[�u��.�ؑ֔��蕶����(�s) <> m_�ؑփ��X�g(�s) Then
            ���ʃe�[�u��.�ؑ֔��蕶����(�s) = m_�ؑփ��X�g(�s)
        End If
        
    Next �s
    
    If �X�e�[�^�X�o�[�i���o�� Then
        DoEvents
        Application.StatusBar = "���x�Ɛؑւ̈ꗗ���V�[�g���畈�ʂɔ��f���Ă��܂�...... (�I��)"
    End If
    
End Function

Public Function �o�͍s�ǉ�()
    Me.OwnTable.ListRows.Add
    �ԍ���(Me.�f�[�^�s��) = Me.�f�[�^�s��
End Function

Public Function ���x�ؑփ��X�g�ݒ�(ByVal ���ʃe�[�u�� As IChartTable, Optional ByVal �����Đݒ� As Boolean = False)
    
    '���Ƀ��X�g���ݒ肳��Ă���ꍇ�͂��Ȃ�
    If Not �����Đݒ� Then
        If Not Not m_�]�����X�g Then
            If Not Not m_���x���X�g Then
                If Not Not m_���x�t���[�����X�g Then
                    If Not Not m_�ؑփ��X�g Then
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    ReDim m_�]�����X�g(���ʃe�[�u��.�f�[�^�s��)
    ReDim m_���x���X�g(���ʃe�[�u��.�f�[�^�s��)
    ReDim m_���x�t���[�����X�g(���ʃe�[�u��.�f�[�^�s��)
    ReDim m_�ؑփ��X�g(���ʃe�[�u��.�f�[�^�s��)
    
    Dim �Ώە��ʍs As Long
    Dim �s As Long
    
    For �s = 1 To Me.�f�[�^�s��
        
        �Ώە��ʍs = �s�ԍ���(�s)
        
        If m_�]�����X�g(�Ώە��ʍs) <> �]����(�s) Then
            m_�]�����X�g(�Ώە��ʍs) = �]����(�s)
        End If
        If m_���x���X�g(�Ώە��ʍs) <> ���x��(�s) Then
            m_���x���X�g(�Ώە��ʍs) = ���x��(�s)
        End If
        If m_���x�t���[�����X�g(�Ώە��ʍs) <> ���x�t���[����(�s) Then
            m_���x�t���[�����X�g(�Ώە��ʍs) = ���x�t���[����(�s)
        End If
        If m_�ؑփ��X�g(�Ώە��ʍs) <> �ؑ֗�(�s) Then
            m_�ؑփ��X�g(�Ώە��ʍs) = �ؑ֗�(�s)
        End If
        
    Next �s
    
    For �s = 1 To ���ʃe�[�u��.�f�[�^�s��
        If m_�]�����X�g(�s) = "" Then
            m_�]�����X�g(�s) = ���ʃe�[�u��.�t���[������ʕ]��(0)
        End If
    Next
End Function

' **************************************************************************************************
' �e�[�u�����g
' **************************************************************************************************

Public Property Get OwnTable() As ListObject
    Set OwnTable = m_OwnTable
End Property

Public Property Get �����V�[�g() As Worksheet
    Set �����V�[�g = m_Sheet
End Property

' **************************************************************************************************
' �e�[�u���萔
' **************************************************************************************************

Public Property Get �f�[�^�s��() As Long
    If Me.OwnTable Is Nothing Then
        �f�[�^�s�� = -1
    Else
        �f�[�^�s�� = Me.OwnTable.ListRows.count
    End If
End Property

' **************************************************************************************************
' ��
' **************************************************************************************************

Private Property Get �ԍ���(ByVal �s As Long) As Long
    �ԍ��� = m_�ԍ���.DataBodyRange(�s).value
End Property

Private Property Let �ԍ���(ByVal �s As Long, ByVal �l As Long)
    If �ԍ���(�s) <> �l Then
        m_�ԍ���.DataBodyRange(�s).value = �l
    End If
End Property

Private Property Get �s�ԍ���(ByVal �s As Long) As Long
    �s�ԍ��� = m_�s�ԍ���.DataBodyRange(�s).value
End Property

Private Property Let �s�ԍ���(ByVal �s As Long, ByVal �l As Long)
    If �s�ԍ���(�s) <> �l Then
        m_�s�ԍ���.DataBodyRange(�s).value = �l
    End If
End Property

Private Property Get �]����(ByVal �s As Long) As String
    �]���� = m_�]����.DataBodyRange(�s).value
End Property

Private Property Let �]����(ByVal �s As Long, ByVal �l As String)
    If �]����(�s) <> �l Then
        m_�]����.DataBodyRange(�s).value = �l
    End If
End Property

Private Property Get ���x��(ByVal �s As Long) As String
    ���x�� = m_���x��.DataBodyRange(�s).value
End Property

Private Property Let ���x��(ByVal �s As Long, ByVal �l As String)
    If ���x��(�s) <> �l Then
        m_���x��.DataBodyRange(�s).value = �l
    End If
End Property

Private Property Get ���x�t���[����(ByVal �s As Long) As String
    ���x�t���[���� = m_���x�t���[����.DataBodyRange(�s).value
End Property

Private Property Let ���x�t���[����(ByVal �s As Long, ByVal �l As String)
    If ���x�t���[����(�s) <> �l Then
        m_���x�t���[����.DataBodyRange(�s).value = �l
    End If
End Property

Private Property Get �ؑ֗�(ByVal �s As Long) As String
    �ؑ֗� = m_�ؑ֗�.DataBodyRange(�s).value
End Property

Private Property Let �ؑ֗�(ByVal �s As Long, ByVal �l As String)
    If �ؑ֗�(�s) <> �l Then
        m_�ؑ֗�.DataBodyRange(�s).value = �l
    End If
End Property
