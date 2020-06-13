VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScoreRouteTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private m_OwnTable As ListObject
Private m_Sheet As Worksheet

Private m_�ԍ��� As ListColumn
Private m_�ő�X�R�A���� As ListColumn
Private m_�X�R�A�� As ListColumn
Private m_�B������ As ListColumn
Private m_�Œ჉�C�t�� As ListColumn
Private m_���[�g������ As ListColumn
Private m_���x�ؑփV�[�g�� As ListColumn
Private m_�e���J�n�� As ListColumn
Private m_�e���I���� As ListColumn

Private m_is�ȈՔ� As Boolean

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
    ���ʃV�[�g�� = ���ʃe�[�u��.�����V�[�g.name & "_�X�R�A�^���[�g"
    Do Until Me.�����V�[�g.name = ���ʃV�[�g��
        Me.�����V�[�g.name = ���ʃV�[�g��
    Loop
    
    ' �e�[�u���쐬 ---------------------------------------------------------------------------------
    
    On Error GoTo �e�[�u�����쐬�ł��Ȃ��ꍇ
    
    Set m_OwnTable = Me.�����V�[�g.ListObjects.Add(xlSrcRange, Me.�����V�[�g.Range(Me.�����V�[�g.Cells(1, 1), Me.�����V�[�g.Cells(1, 9)), , xlYes)
    
    On Error GoTo �e�[�u�����ύX�G���[
    
    Dim ���ʃe�[�u���� As String
    ���ʃe�[�u���� = ���ʃe�[�u��.OwnTable.name & "_�X�R�A�^���[�g"
    Do Until Me.OwnTable.name = ���ʃe�[�u����
        Me.OwnTable.name = ���ʃe�[�u����
    Loop
    
    ' �e�[�u�����` ---------------------------------------------------------------------------------
    
    On Error GoTo 0
    
    Me.OwnTable.ShowTableStyleFirstColumn = True
    Me.OwnTable.TableStyle = ���ʃe�[�u��.OwnTable.TableStyle
    
    Me.OwnTable.ListColumns(1).name = "No"
    Me.OwnTable.ListColumns(2).name = "�ő卷"
    Me.OwnTable.ListColumns(3).name = "�X�R�A"
    Me.OwnTable.ListColumns(4).name = "�B����"
    Me.OwnTable.ListColumns(5).name = "�Œ�LIFE"
    Me.OwnTable.ListColumns(6).name = "���[�g������"
    Me.OwnTable.ListColumns(7).name = "���x�ؑփV�[�g��"
    Me.OwnTable.ListColumns(8).name = "�e���J�n�s"
    Me.OwnTable.ListColumns(9).name = "�e���I���s"
    
    Me.OwnTable.ListColumns(1).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(2).Range.ColumnWidth = 10
    Me.OwnTable.ListColumns(3).Range.ColumnWidth = 10
    Me.OwnTable.ListColumns(4).Range.ColumnWidth = 8
    Me.OwnTable.ListColumns(5).Range.ColumnWidth = 8
    Me.OwnTable.ListColumns(6).Range.ColumnWidth = 72
    Me.OwnTable.ListColumns(7).Range.ColumnWidth = 30
    Me.OwnTable.ListColumns(8).Range.ColumnWidth = 10
    Me.OwnTable.ListColumns(9).Range.ColumnWidth = 10
    
    Set m_�ԍ��� = Me.OwnTable.ListColumns(1)
    Set m_�ő�X�R�A���� = Me.OwnTable.ListColumns(2)
    Set m_�X�R�A�� = Me.OwnTable.ListColumns(3)
    Set m_�B������ = Me.OwnTable.ListColumns(4)
    Set m_�Œ჉�C�t�� = Me.OwnTable.ListColumns(5)
    Set m_���[�g������ = Me.OwnTable.ListColumns(6)
    Set m_���x�ؑփV�[�g�� = Me.OwnTable.ListColumns(7)
    Set m_�e���J�n�� = Me.OwnTable.ListColumns(8)
    Set m_�e���I���� = Me.OwnTable.ListColumns(9)
    
    �A�N�e�B�u�V�[�g.Activate
    
    Application.ScreenUpdating = True
    
    m_is�ȈՔ� = False
    
    �I�u�W�F�N�g�V�K�쐬 = 0
    Exit Function
    
�V�[�g���ύX�G���[:
    Dim ���ʃV�[�g���A�� As Long
    ���ʃV�[�g���A�� = ���ʃV�[�g���A�� + 1
    ���ʃV�[�g�� = ���ʃe�[�u��.�����V�[�g.name & "_�X�R�A�^���[�g_" & ���ʃV�[�g���A��
    Resume Next
    
�e�[�u�����ύX�G���[:
    Dim ���ʃe�[�u�����A�� As Long
    ���ʃe�[�u�����A�� = ���ʃe�[�u�����A�� + 1
    ���ʃe�[�u���� = ���ʃe�[�u��.OwnTable.name & "_�X�R�A�^���[�g_" & ���ʃe�[�u�����A��
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
    Set m_�ő�X�R�A���� = Me.OwnTable.ListColumns("�ő卷")
    Set m_�X�R�A�� = Me.OwnTable.ListColumns("�X�R�A")
    Set m_�B������ = Me.OwnTable.ListColumns("�B����")
    Set m_�Œ჉�C�t�� = Me.OwnTable.ListColumns("�Œ�LIFE")
    Set m_���[�g������ = Me.OwnTable.ListColumns("���[�g������")
    Set m_���x�ؑփV�[�g�� = Me.OwnTable.ListColumns("���x�ؑփV�[�g��")
    Set m_�e���J�n�� = Me.OwnTable.ListColumns("�e���J�n�s")
    Set m_�e���I���� = Me.OwnTable.ListColumns("�e���I���s")
    
    On Error GoTo 0
    
    m_is�ȈՔ� = False
    
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
    Set m_�ő�X�R�A���� = Nothing
    Set m_�X�R�A�� = Nothing
    Set m_�B������ = Nothing
    Set m_�Œ჉�C�t�� = Nothing
    Set m_���[�g������ = Nothing
    Set m_���x�ؑփV�[�g�� = Nothing
    Set m_�e���J�n�� = Nothing
    Set m_�e���I���� = Nothing
    
    m_is�ȈՔ� = False
    
End Sub

Public Function �I�u�W�F�N�g�ŏI���`(ByVal ���ʃe�[�u�� As IChartTable)
    
    If Me.�f�[�^�s�� > 0 Then
        m_�ő�X�R�A����.DataBodyRange(�f�[�^�s��).Formula = "=[@�X�R�A]-MAX([�X�R�A])"
    End If
    
    Me.OwnTable.HeaderRowRange.HorizontalAlignment = xlCenter
    
    m_�B������.DataBodyRange.NumberFormatLocal = "0.00%"
    
'    Me.OwnTable.ShowAutoFilterDropDown = False
    
End Function

Public Function �o�͍s�ǉ�()
    Me.OwnTable.ListRows.Add
    �ԍ���(Me.�f�[�^�s��) = Me.�f�[�^�s��
End Function

Public Function ���݃��[�g�o��( _
    ByVal ���ʃe�[�u�� As IChartTable, _
    ByVal �X�R�A�^���[�g������ As String, _
    ByVal �e���J�n�s As Long, _
    ByVal �e���I���s As Long, _
    ByVal �B���� As Double, _
    ByVal �X�R�A As Long)
    
    Me.�o�͍s�ǉ�
    
    �X�R�A��(Me.�f�[�^�s��) = �X�R�A
    �B������(Me.�f�[�^�s��) = �B����
    �Œ჉�C�t��(Me.�f�[�^�s��) = ���ʃe�[�u��.�Œ჉�C�t
    ���[�g������(Me.�f�[�^�s��) = �X�R�A�^���[�g������
    �e���J�n��(Me.�f�[�^�s��) = �e���J�n�s
    �e���I����(Me.�f�[�^�s��) = �e���I���s
    
    If Not Me.Is�ȈՔ� Then
        Dim ���ݑ��x�ؑֈꗗ�e�[�u�� As ElSwTable
        Set ���ݑ��x�ؑֈꗗ�e�[�u�� = New ElSwTable
        If ���ݑ��x�ؑֈꗗ�e�[�u��.�I�u�W�F�N�g�V�K�쐬(���ʃe�[�u��) < 0 Then
            GoTo �e�[�u���ݒ�Ɏ��s�����ꍇ
        End If
        ���ݑ��x�ؑֈꗗ�e�[�u��.���x�ؑ֏��ǂݍ��� ���ʃe�[�u��
        
        ���x�ؑփV�[�g��(Me.�f�[�^�s��) = ���ݑ��x�ؑֈꗗ�e�[�u��.�����V�[�g.name
    End If
    
    Exit Function
    
�e�[�u���ݒ�Ɏ��s�����ꍇ:
    MsgBox "ERR:���x�Ɛؑւ̈ꗗ�e�[�u���̍쐬�Ɏ��s���܂����B" & vbCrLf & "�����͑��s����܂��B", vbCritical
    Exit Function
    
End Function

Public Function �g�ݍ��킹���[�g�o��( _
    ByVal ���ʃe�[�u�� As IChartTable, _
    ByVal ���x�ؑֈꗗ�e�[�u�� As ElSwTable, _
    Optional ByVal ���t���[�h As Boolean = False, _
    Optional ByVal ���O�e���I���s As Long = 0, _
    Optional ByVal �l�X�g�� As Long = 0)
    
    If Me.Is�ȈՔ� Then
        Exit Function
    End If
    
    Dim �s As Long
    
    Dim �e���J�n�s As Long
    Dim �e���I���s As Long
    
    Dim ���U���g As Def.���ʃZ�b�g
    
    For �s = 1 To Me.�f�[�^�s��
        �e���J�n�s = �e���J�n��(�s)
        If �e���J�n�s > ���O�e���I���s Then
            ���x�ؑֈꗗ�e�[�u��.���x�ؑ֏�񏑂��o�� ���ʃe�[�u��, ���O�e���I���s + 1, �e���J�n�s - 1
            Dim ���ݑ��x�ؑֈꗗ�e�[�u�� As ElSwTable
            Set ���ݑ��x�ؑֈꗗ�e�[�u�� = New ElSwTable
            ���ݑ��x�ؑֈꗗ�e�[�u��.�I�u�W�F�N�g�ݒ� ThisWorkbook.Worksheets(���x�ؑփV�[�g��(�s)).ListObjects(1).name
            �e���I���s = �e���I����(�s)
            ���ݑ��x�ؑֈꗗ�e�[�u��.���x�ؑ֏�񏑂��o�� ���ʃe�[�u��, �e���J�n�s, �e���I���s
            Me.�g�ݍ��킹���[�g�o�� ���ʃe�[�u��, ���x�ؑֈꗗ�e�[�u��, ���t���[�h, �e���I���s, �l�X�g�� + 1
        End If
    Next �s
    
    If �l�X�g�� > 1 Then
        ���x�ؑֈꗗ�e�[�u��.���x�ؑ֏�񏑂��o�� ���ʃe�[�u��, ���O�e���I���s + 1, ���ʃe�[�u��.�f�[�^�s��
        ���U���g = Def.���ʃe�[�u��.���U���g�Čv�Z()
        If ���t���[�h Or ���U���g.�N���A�����N <> Def.MISSTAKE���� Then
            Me.���݃��[�g�o�� ���ʃe�[�u��, Analyzing.�X�R�A�^���[�g������擾(), 1, ���ʃe�[�u��.�f�[�^�s��, ���U���g.�B����, ���U���g.�X�R�A
        End If
    End If
    
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

Private Property Get �X�R�A��(ByVal �s As Long) As Long
    �X�R�A�� = m_�X�R�A��.DataBodyRange(�s).value
End Property

Private Property Let �X�R�A��(ByVal �s As Long, ByVal �l As Long)
    If �X�R�A��(�s) <> �l Then
        m_�X�R�A��.DataBodyRange(�s).value = �l
    End If
End Property

Private Property Get �B������(ByVal �s As Long) As Double
    �B������ = m_�B������.DataBodyRange(�s).value
End Property

Private Property Let �B������(ByVal �s As Long, ByVal �l As Double)
    If �B������(�s) <> �l Then
        m_�B������.DataBodyRange(�s).value = �l
    End If
End Property

Private Property Get �Œ჉�C�t��(ByVal �s As Long) As Long
    �Œ჉�C�t�� = m_�Œ჉�C�t��.DataBodyRange(�s).value
End Property

Private Property Let �Œ჉�C�t��(ByVal �s As Long, ByVal �l As Long)
    If �Œ჉�C�t��(�s) <> �l Then
        m_�Œ჉�C�t��.DataBodyRange(�s).value = �l
    End If
End Property

Private Property Get ���[�g������(ByVal �s As Long) As String
    ���[�g������ = m_���[�g������.DataBodyRange(�s).value
End Property

Private Property Let ���[�g������(ByVal �s As Long, ByVal �l As String)
    If ���[�g������(�s) <> �l Then
        m_���[�g������.DataBodyRange(�s).value = �l
    End If
End Property

Private Property Get ���x�ؑփV�[�g��(ByVal �s As Long) As String
    ���x�ؑփV�[�g�� = m_���x�ؑփV�[�g��.DataBodyRange(�s).value
End Property

Private Property Let ���x�ؑփV�[�g��(ByVal �s As Long, ByVal �l As String)
    If ���x�ؑփV�[�g��(�s) <> �l Then
        m_���x�ؑփV�[�g��.DataBodyRange(�s).value = �l
    End If
End Property

Private Property Get �e���J�n��(ByVal �s As Long) As Long
    �e���J�n�� = m_�e���J�n��.DataBodyRange(�s).value
End Property

Private Property Let �e���J�n��(ByVal �s As Long, ByVal �l As Long)
    If �e���J�n��(�s) <> �l Then
        m_�e���J�n��.DataBodyRange(�s).value = �l
    End If
End Property

Private Property Get �e���I����(ByVal �s As Long) As Long
    �e���I���� = m_�e���I����.DataBodyRange(�s).value
End Property

Private Property Let �e���I����(ByVal �s As Long, ByVal �l As Long)
    If �e���I����(�s) <> �l Then
        m_�e���I����.DataBodyRange(�s).value = �l
    End If
End Property

Public Property Get Is�ȈՔ�() As Boolean
    Is�ȈՔ� = m_is�ȈՔ�
End Property

Public Property Let Is�ȈՔ�(ByVal �l As Boolean)
    m_is�ȈՔ� = �l
End Property

