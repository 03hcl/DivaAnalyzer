VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SwitchingTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private m_OwnTable As ListObject
Private m_Sheet As Worksheet

Private m_�ԍ��� As ListColumn
Private m_��̓u���b�N�� As ListColumn
Private m_�u���b�N�J�n�� As ListColumn
Private m_�u���b�N�I���� As ListColumn
Private m_�ő�X�R�A���� As ListColumn
Private m_�u���b�N�X�R�A�� As ListColumn
Private m_�ؑփX�R�A���� As ListColumn
Private m_�ؑ֕����� As ListColumn
Private m_MAX�\�������� As ListColumn
'Private m_�ؑ֗�() As listColumn

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
    ���ʃV�[�g�� = ���ʃe�[�u��.�����V�[�g.name & "_�ؑ�"
    Do Until Me.�����V�[�g.name = ���ʃV�[�g��
        Me.�����V�[�g.name = ���ʃV�[�g��
    Loop
    
    ' �e�[�u���쐬 ---------------------------------------------------------------------------------
    
    On Error GoTo �e�[�u�����쐬�ł��Ȃ��ꍇ
    
    Set m_OwnTable = Me.�����V�[�g.ListObjects.Add(xlSrcRange, Me.�����V�[�g.Range(Me.�����V�[�g.Cells(1, 1), Me.�����V�[�g.Cells(1, 9)), , xlYes)
    
    On Error GoTo �e�[�u�����ύX�G���[
    
    Dim ���ʃe�[�u���� As String
    ���ʃe�[�u���� = ���ʃe�[�u��.OwnTable.name & "_�ؑ�"
    Do Until Me.OwnTable.name = ���ʃe�[�u����
        Me.OwnTable.name = ���ʃe�[�u����
    Loop
    
    On Error GoTo 0
    
    ' �e�[�u�����` ---------------------------------------------------------------------------------
    
    Me.OwnTable.ShowTableStyleFirstColumn = True
    Me.OwnTable.TableStyle = ���ʃe�[�u��.OwnTable.TableStyle
    
    Me.OwnTable.ListColumns(1).name = "No"
    Me.OwnTable.ListColumns(2).name = "HBlock"
    Me.OwnTable.ListColumns(3).name = "�J�n�s"
    Me.OwnTable.ListColumns(4).name = "�I���s"
    Me.OwnTable.ListColumns(5).name = "�ő卷"
    Me.OwnTable.ListColumns(6).name = "�X�R�A"
    Me.OwnTable.ListColumns(7).name = "�ؑ֍�"
    Me.OwnTable.ListColumns(8).name = "�ؑ�"
    Me.OwnTable.ListColumns(9).name = "MAX�\��"
    
    Me.OwnTable.ListColumns(1).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(2).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(3).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(4).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(5).Range.ColumnWidth = 8
    Me.OwnTable.ListColumns(6).Range.ColumnWidth = 8
    Me.OwnTable.ListColumns(7).Range.ColumnWidth = 8
    Me.OwnTable.ListColumns(8).Range.ColumnWidth = 24
    Me.OwnTable.ListColumns(9).Range.ColumnWidth = 48
    
    Set m_�ԍ��� = Me.OwnTable.ListColumns(1)
    Set m_��̓u���b�N�� = Me.OwnTable.ListColumns(2)
    Set m_�u���b�N�J�n�� = Me.OwnTable.ListColumns(3)
    Set m_�u���b�N�I���� = Me.OwnTable.ListColumns(4)
    Set m_�ő�X�R�A���� = Me.OwnTable.ListColumns(5)
    Set m_�u���b�N�X�R�A�� = Me.OwnTable.ListColumns(6)
    Set m_�ؑփX�R�A���� = Me.OwnTable.ListColumns(7)
    Set m_�ؑ֕����� = Me.OwnTable.ListColumns(8)
    Set m_MAX�\�������� = Me.OwnTable.ListColumns(9)
    
    �A�N�e�B�u�V�[�g.Activate
    
    Application.ScreenUpdating = True
    
    �I�u�W�F�N�g�V�K�쐬 = 0
    Exit Function
    
�V�[�g���ύX�G���[:
    Dim ���ʃV�[�g���A�� As Long
    ���ʃV�[�g���A�� = ���ʃV�[�g���A�� + 1
    ���ʃV�[�g�� = ���ʃe�[�u��.�����V�[�g.name & "_�ؑ�_" & ���ʃV�[�g���A��
    Resume Next
    
�e�[�u�����ύX�G���[:
    Dim ���ʃe�[�u�����A�� As Long
    ���ʃe�[�u�����A�� = ���ʃe�[�u�����A�� + 1
    ���ʃe�[�u���� = ���ʃe�[�u��.OwnTable.name & "_�ؑ�_" & ���ʃe�[�u�����A��
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
    Set m_��̓u���b�N�� = Me.OwnTable.ListColumns("HBlock")
    Set m_�u���b�N�J�n�� = Me.OwnTable.ListColumns("�J�n�s")
    Set m_�u���b�N�I���� = Me.OwnTable.ListColumns("�I���s")
    Set m_�ő�X�R�A���� = Me.OwnTable.ListColumns("�ő卷")
    Set m_�u���b�N�X�R�A�� = Me.OwnTable.ListColumns("�X�R�A")
    Set m_�ؑփX�R�A���� = Me.OwnTable.ListColumns("�ؑ֍�")
    Set m_�ؑ֕����� = Me.OwnTable.ListColumns("�ؑ�")
    Set m_MAX�\�������� = Me.OwnTable.ListColumns("MAX�\��")
    
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

Public Function �I�u�W�F�N�g�ŏI���`() As Long
    
    If Me.�f�[�^�s�� > 0 Then
        m_�ő�X�R�A����.DataBodyRange(�f�[�^�s��).FormulaArray = "=[@�X�R�A]-MAX(([HBlock]=[@HBlock])*[�X�R�A])"
        m_�ؑփX�R�A����.DataBodyRange(�f�[�^�s��).FormulaArray = "=[@�X�R�A]-MAX(([HBlock]=[@HBlock])*([�ؑ�]="""")*[�X�R�A])"
    End If
    
    Dim �� As Long
    For �� = 9 + 1 To Me.OwnTable.ListColumns.count
        Me.OwnTable.ListColumns(��).name = "�s" & (�� - 9)
        Me.OwnTable.ListColumns(��).Range.ColumnWidth = 5
    Next ��
    
    Me.OwnTable.HeaderRowRange.HorizontalAlignment = xlCenter
    
End Function

Public Function �őP�ؑ֏��擾(ByVal ���ʃe�[�u�� As IChartTable, Optional ByVal �����Čv�Z As Boolean = False) As Switching
    
    Set �őP�ؑ֏��擾 = New Switching
    �őP�ؑ֏��擾.�I�u�W�F�N�g�ݒ� ���ʃe�[�u��
    Dim �s As Long
    Dim �� As Long
    Dim �֑ؑΏۍs As Long
    
    If Not �����Čv�Z Then
        Me.OwnTable.Range.Calculate
    End If
    
    For �s = 1 To Me.�f�[�^�s��
        If m_�ő�X�R�A����.DataBodyRange(�s).value = 0 Then
            �őP�ؑ֏��擾.�ؑ֕����� = Def.������A��(�őP�ؑ֏��擾.�ؑ֕�����, Me.�ؑ֕�����(�s), vbCrLf)
            For �� = 9 + 1 To Me.OwnTable.ListColumns.count
            �֑ؑΏۍs = Me.OwnTable.ListColumns(��).DataBodyRange(�s).value
                If �֑ؑΏۍs > 0 Then
                    �őP�ؑ֏��擾.�ؑփt���O(�֑ؑΏۍs) = True
                End If
            Next
        End If
    Next �s
    
    For �s = 1 To ���ʃe�[�u��.�f�[�^�s��
        If ���ʃe�[�u��.�ؑ֔����(�s) Then
            �őP�ؑ֏��擾.�ؑփt���O(�s) = True
        End If
    Next
    
End Function

Public Sub �I�u�W�F�N�g���()
    
    Set m_OwnTable = Nothing
    Set m_Sheet = Nothing
    
    Set m_�ԍ��� = Nothing
    Set m_��̓u���b�N�� = Nothing
    Set m_�u���b�N�J�n�� = Nothing
    Set m_�u���b�N�I���� = Nothing
    Set m_�ő�X�R�A���� = Nothing
    Set m_�u���b�N�X�R�A�� = Nothing
    Set m_�ؑփX�R�A���� = Nothing
    Set m_�ؑ֕����� = Nothing
    Set m_MAX�\�������� = Nothing
    
End Sub

Public Function �ؑ֍s����(ByVal ���͍s�� As Long, ByVal �ؑ֐� As Long)
    Me.�����V�[�g.Cells(Me.�f�[�^�s�� + 1, �ؑ֐� + 9).value = ���͍s��
End Function

Public Function �o�͍s�ǉ�()
    Me.OwnTable.ListRows.Add
    �ԍ���(Me.�f�[�^�s��) = Me.�f�[�^�s��
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

Public Property Get �ő�u���b�N�J�n�s() As Long
    �ő�u���b�N�J�n�s = Application.WorksheetFunction.Max(m_�u���b�N�J�n��.DataBodyRange)
End Property

Public Property Get �ő�u���b�N�I���s() As Long
    �ő�u���b�N�I���s = Application.WorksheetFunction.Max(m_�u���b�N�I����.DataBodyRange)
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

Public Property Get ��̓u���b�N��(ByVal �s As Long) As Long
    ��̓u���b�N�� = m_��̓u���b�N��.DataBodyRange(�s).value
End Property

Public Property Let ��̓u���b�N��(ByVal �s As Long, ByVal �l As Long)
    If ��̓u���b�N��(�s) <> �l Then
        m_��̓u���b�N��.DataBodyRange(�s).value = �l
    End If
End Property

Public Property Get �u���b�N�J�n��(ByVal �s As Long) As Long
    �u���b�N�J�n�� = m_�u���b�N�J�n��.DataBodyRange(�s).value
End Property

Public Property Let �u���b�N�J�n��(ByVal �s As Long, ByVal �l As Long)
    If �u���b�N�J�n��(�s) <> �l Then
        m_�u���b�N�J�n��.DataBodyRange(�s).value = �l
    End If
End Property

Public Property Get �u���b�N�I����(ByVal �s As Long) As Long
    �u���b�N�I���� = m_�u���b�N�I����.DataBodyRange(�s).value
End Property

Public Property Let �u���b�N�I����(ByVal �s As Long, ByVal �l As Long)
    If �u���b�N�I����(�s) <> �l Then
        m_�u���b�N�I����.DataBodyRange(�s).value = �l
    End If
End Property

Public Property Get �u���b�N�X�R�A��(ByVal �s As Long) As Long
    �u���b�N�X�R�A�� = m_�u���b�N�X�R�A��.DataBodyRange(�s).value
End Property

Public Property Let �u���b�N�X�R�A��(ByVal �s As Long, ByVal �l As Long)
    If �u���b�N�X�R�A��(�s) <> �l Then
        m_�u���b�N�X�R�A��.DataBodyRange(�s).value = �l
    End If
End Property

Public Property Get �ؑ֕�����(ByVal �s As Long) As String
    �ؑ֕����� = m_�ؑ֕�����.DataBodyRange(�s).value
End Property

Public Property Let �ؑ֕�����(ByVal �s As Long, ByVal �l As String)
    If �ؑ֕�����(�s) <> �l Then
        m_�ؑ֕�����.DataBodyRange(�s).value = �l
    End If
End Property

Public Property Get MAX�\��������(ByVal �s As Long) As String
    MAX�\�������� = m_MAX�\��������.DataBodyRange(�s).value
End Property

Public Property Let MAX�\��������(ByVal �s As Long, ByVal �l As String)
    If MAX�\��������(�s) <> �l Then
        m_MAX�\��������.DataBodyRange(�s).value = �l
    End If
End Property
