Attribute VB_Name = "Arranging"
Option Explicit
Option Base 1

Public Sub 切替結果から早遅と切替を譜面に反映()
    
    Def.マーク文字設定
    Def.文字定数設定
    
    Set Def.切替結果テーブル = New SwitchingTable
    If Def.切替結果テーブル.オブジェクト設定() < 0 Then
        GoTo 切替結果テーブルの設定に失敗した場合
    End If
    
    Dim 譜面テーブル名 As String
    譜面テーブル名 = Def.切替結果テーブル.OwnTable.name
    譜面テーブル名 = Left(譜面テーブル名, Application.WorksheetFunction.Max(0, InStr(譜面テーブル名, "_切替") - 1))
    If 譜面テーブル検索設定(譜面テーブル名) < 0 Then
        Exit Sub
    End If
    
    ReDim 早遅手動指定フラグ(Def.譜面テーブル.データ行数)
    ReDim MAX可能性フラグ(Def.譜面テーブル.データ行数)
    
    Application.StatusBar = "切替と早遅の情報を譜面テーブルに反映しています......"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim 最善切替情報 As Switching
    Set 最善切替情報 = Def.切替結果テーブル.最善切替情報取得(譜面テーブル)
    Debug.Print 最善切替情報.切替文字列
    Dim MAX可能性 As OutputString
    Set MAX可能性 = 最善切替情報.切替早遅情報反映()
    Debug.Print MAX可能性.表示文字列
    
    MsgBox "正常に終了しました。", vbInformation
    
    Rescue
    
    Exit Sub
    
切替結果テーブルの設定に失敗した場合:
    
    MsgBox "ERR:切替結果テーブルの設定に失敗しました。" & vbCrLf & "処理を終了します。", vbCritical
    Exit Sub
    
End Sub

Public Sub 早遅と切替を別シートに出力()
    
    Def.マーク文字設定
    Def.文字定数設定
    
    If Def.譜面テーブル設定() < 0 Then
        Exit Sub
    End If
    
    If Def.早遅切替一覧テーブル設定(Def.譜面テーブル) < 0 Then
        Exit Sub
    End If
    
    If MsgBox("現在のテーブルで処理を開始します。よろしいですか？" & vbCrLf & _
        "テーブル名: " & Def.譜面テーブル.OwnTable.name, vbOKCancel + vbInformation) <> vbOK Then
        MsgBox "処理を中止しました。", vbCritical
        Exit Sub
    End If
    
    Application.StatusBar = "早遅と切替の一覧を譜面からシートに出力しています......"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Def.早遅切替一覧テーブル.早遅切替情報読み込み Def.譜面テーブル
    
    MsgBox "正常に終了しました。", vbInformation
    
    Rescue
    
End Sub

Public Sub 早遅と切替を譜面に反映()
    
    Def.マーク文字設定
    Def.文字定数設定
    
    Set Def.早遅切替一覧テーブル = New ElSwTable
    If Def.早遅切替一覧テーブル.オブジェクト設定() < 0 Then
        GoTo 早遅切替一覧テーブルの設定に失敗した場合
    End If
    
    Dim 譜面テーブル名 As String
    譜面テーブル名 = Def.早遅切替一覧テーブル.OwnTable.name
    譜面テーブル名 = Left(譜面テーブル名, Application.WorksheetFunction.Max(0, InStr(譜面テーブル名, "_早遅切替リスト") - 1))
    If 譜面テーブル検索設定(譜面テーブル名) < 0 Then
        Exit Sub
    End If
    
    Application.StatusBar = "早遅と切替の一覧をシートから譜面に反映しています......"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Def.早遅切替一覧テーブル.早遅切替情報書き出し Def.譜面テーブル
    
    MsgBox "正常に終了しました。", vbInformation
    
    Rescue
    Exit Sub
    
早遅切替一覧テーブルの設定に失敗した場合:
    
    MsgBox "ERR:早遅と切替の一覧テーブルの設定に失敗しました。" & vbCrLf & "処理を終了します。", vbCritical
    
    Rescue
    Exit Sub
    
End Sub

Public Sub 現在の状態でのスコアタルート文字列取得()
    
    Def.マーク文字設定
    Def.文字定数設定
    
    If Def.スコアタ解析用定数設定() < 0 Then
        Exit Sub
    End If
    
    If Def.譜面テーブル設定() < 0 Then
        Exit Sub
    End If
    
    Dim スコアタルート文字列 As String
    スコアタルート文字列 = Analyzing.スコアタルート文字列取得()
    
    Dim cb As New dataobject
    cb.SetText スコアタルート文字列
    cb.PutInClipboard
    Set cb = Nothing
    
    MsgBox "現在の状態でのスコアタルート文字列は以下の通りです。" & vbCrLf & _
        "(この文字列はクリップボードにコピーされています。)" & vbCrLf & vbCrLf & _
        スコアタルート文字列, _
        vbInformation
    
    Rescue
    
    Exit Sub
End Sub
