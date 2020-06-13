Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Private Sub 非表示の名前を表示()
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
    Dim test1 As Def.個別切替データ
    Dim test2 As Def.個別切替データ
    If Not test1.切替行リスト Then
        ReDim test1.切替行リスト(3)
    End If
    test1.切替行リスト(1) = 1
    test2 = test1
    test2.切替行リスト(2) = 2
    Erase test1.切替行リスト
    'ReDim test1.切替行リスト(0) 'エラー
    
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
    Def.スコアタ解析用定数設定
    str = Def.評価略号("赤WRONG")
    
End Sub

Private Sub Macro3()
'
' Macro3 Macro
'

'
    Def.マーク文字設定
    Def.文字定数設定
    
    If Def.スコアタ解析用定数設定 < 0 Then
        Exit Sub
    End If
    
    If Def.譜面テーブル設定() < 0 Then
        Exit Sub
    End If
    
'    Dim t As Long
'    t = Def.譜面テーブル.ホールドスコア列(100)
'    Dim b As Boolean
'    b = Def.譜面テーブル.切替判定列(100)
'    Def.譜面テーブル.切替判定列(100) = False
'    Def.譜面テーブル.切替判定列(100) = True
    
    'Range(Def.譜面テーブル.評価列(1), Def.譜面テーブル.OwnTable.ListColumns(Def.譜面テーブル.OwnTable.ListColumns.count).DataBodyRange(3)).Select
    
    Def.譜面テーブル.リザルト再計算
    
    Dim fr As Long
    'fr = Def.譜面テーブル.評価別最早フレーム("SAFE")
    
    Dim 評価() As Def.評価セット
    評価 = Def.譜面テーブル.評価リスト取得(7, -6, -6)
    
End Sub

Private Sub Macro4()
'
' Macro4 Macro
'

'
'    Range("DifficultyTable[[MISS×TAKE]:[EXCELLENT]]").Select
    Dim m_難易度テーブル As ListObject
    Set m_難易度テーブル = ThisWorkbook.Worksheets("Difficulty").ListObjects("DifficultyTable")
    Dim a
    a = Application.WorksheetFunction.index( _
        Range(m_難易度テーブル.ListColumns("MISS×TAKE").DataBodyRange, m_難易度テーブル.ListColumns("EXCELLENT").DataBodyRange), _
        Application.WorksheetFunction.Match( _
            "EXTREME", m_難易度テーブル.ListColumns("Difficulty").DataBodyRange, _
            0), _
        0)
    Dim b As Long
    b = Application.WorksheetFunction.Match( _
                    0.6, a, 1)
    Dim c As String
    c = m_難易度テーブル.ListColumns(m_難易度テーブル.ListColumns("MISS×TAKE").index - 1 + b).name
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
