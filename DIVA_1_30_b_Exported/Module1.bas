Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Private Sub ��\���̖��O��\��()
    Dim name As Object
    For Each name In Names
        If name.Visible = False Then
            name.Visible = True
        End If
    Next
End Sub

Private Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Dim test1 As Def.�ʐؑփf�[�^
    Dim test2 As Def.�ʐؑփf�[�^
    If Not test1.�ؑ֍s���X�g Then
        ReDim test1.�ؑ֍s���X�g(3)
    End If
    test1.�ؑ֍s���X�g(1) = 1
    test2 = test1
    test2.�ؑ֍s���X�g(2) = 2
    Erase test1.�ؑ֍s���X�g
    'ReDim test1.�ؑ֍s���X�g(0) '�G���[
    
End Sub

Private Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Dim test1(3 To 5) As Long
    Dim test2(-1 To 3) As Long
    test1(4) = 3
    test1(5) = -2
    Dim t As Long
    t = Application.WorksheetFunction.Sum(test1)
    test2(-1) = 2
    
    Dim str As String
    str = ThisWorkbook.Worksheets("Rating").ListObjects("RatingTable").ListColumns(11).name
    Def.�X�R�A�^��͗p�萔�ݒ�
    str = Def.�]������("��WRONG")
    
End Sub

Private Sub Macro3()
'
' Macro3 Macro
'

'
    Def.�}�[�N�����ݒ�
    Def.�����萔�ݒ�
    
    If Def.�X�R�A�^��͗p�萔�ݒ� < 0 Then
        Exit Sub
    End If
    
    If Def.���ʃe�[�u���ݒ�() < 0 Then
        Exit Sub
    End If
    
'    Dim t As Long
'    t = Def.���ʃe�[�u��.�z�[���h�X�R�A��(100)
'    Dim b As Boolean
'    b = Def.���ʃe�[�u��.�ؑ֔����(100)
'    Def.���ʃe�[�u��.�ؑ֔����(100) = False
'    Def.���ʃe�[�u��.�ؑ֔����(100) = True
    
    'Range(Def.���ʃe�[�u��.�]����(1), Def.���ʃe�[�u��.OwnTable.ListColumns(Def.���ʃe�[�u��.OwnTable.ListColumns.count).DataBodyRange(3)).Select
    
    Def.���ʃe�[�u��.���U���g�Čv�Z
    
    Dim fr As Long
    'fr = Def.���ʃe�[�u��.�]���ʍő��t���[��("SAFE")
    
    Dim �]��() As Def.�]���Z�b�g
    �]�� = Def.���ʃe�[�u��.�]�����X�g�擾(7, -6, -6)
    
End Sub

Private Sub Macro4()
'
' Macro4 Macro
'

'
'    Range("DifficultyTable[[MISS�~TAKE]:[EXCELLENT]]").Select
    Dim m_��Փx�e�[�u�� As ListObject
    Set m_��Փx�e�[�u�� = ThisWorkbook.Worksheets("Difficulty").ListObjects("DifficultyTable")
    Dim a
    a = Application.WorksheetFunction.index( _
        Range(m_��Փx�e�[�u��.ListColumns("MISS�~TAKE").DataBodyRange, m_��Փx�e�[�u��.ListColumns("EXCELLENT").DataBodyRange), _
        Application.WorksheetFunction.Match( _
            "EXTREME", m_��Փx�e�[�u��.ListColumns("Difficulty").DataBodyRange, _
            0), _
        0)
    Dim b As Long
    b = Application.WorksheetFunction.Match( _
                    0.6, a, 1)
    Dim c As String
    c = m_��Փx�e�[�u��.ListColumns(m_��Փx�e�[�u��.ListColumns("MISS�~TAKE").index - 1 + b).name
End Sub


Private Sub test()
    Dim wShell As IWshRuntimeLibrary.WshShell
    wShell = New WshShell
    Dim wExec As WshExec
    Set wExec = shell.exec
    Do While wExec.Status = 0
        DoEvents
    Loop
    
    wExec.Terminate
    Set wExec = Nothing
    Set wShell = Nothing
End Sub
