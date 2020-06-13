Attribute VB_Name = "Def"
Option Explicit
Option Base 1

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Const マーク数 As Long = 4
Public Const スライドマーク数 As Long = 2

Public Const 未MAXフレーム最大値 As Long = 300

Public マーク文字(マーク数) As String
Public スライドマーク文字(スライドマーク数) As String

Public 早COOL文字 As String
Public ジャストCOOL文字 As String
Public 遅COOL文字 As String

Public HOLD文字 As String

Public 赤WRONG文字 As String
Public WORST文字 As String
Public MISSTAKE文字 As String

Public スコアタスキップ点 As Long
Public CC最大余裕フレーム量 As Long
Public 最大ライフ量 As Long

Public 評価略号 As Dictionary

Public 譜面テーブル As IChartTable
Public 処理ログ As ProcessLog
Public 切替結果テーブル As SwitchingTable

Public 早遅切替一覧テーブル As ElSwTable

Public スコアタテーブル As ScoreRouteTable

Public 現在ホールドブロック As Long
Public 早遅手動指定フラグ() As Boolean
Public MAX可能性フラグ() As Boolean

Public Type 個別切替データ
    スコア As Long
    切替行リスト() As Long
End Type

Public Type 評価セット
    開始フレームずれ As Long
    開始評価 As String
    終了フレームずれ As Long
    終了評価 As String
    正常判定枠 As Boolean
End Type

Public Type 行スコア情報
    最大影響行 As Long
    行スコア() As Long
End Type

Public Type 結果セット
    クリアランク As String
    達成率 As Double
    スコア As Long
End Type

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function 切替解析初期設定実行(Optional ByVal 譜面テーブル名 As String = "") As Long
    
    Application.StatusBar = "切替解析の初期設定を行います......"
    
    切替解析初期設定実行 = -1
    
    Def.マーク文字設定
    Def.文字定数設定
    
    If Def.譜面テーブル設定(譜面テーブル名) < 0 Then
        Exit Function
    End If
    
    If Def.切替結果テーブル設定(譜面テーブル) < 0 Then
        Exit Function
    End If
    
    If Def.処理ログ出力設定() < 0 Then
        Exit Function
    End If
    
    Dim データ行数 As Long
    データ行数 = Def.譜面テーブル.データ行数
    
    ReDim 早遅手動指定フラグ(データ行数)
    ReDim MAX可能性フラグ(データ行数)
    
    Dim 行 As Long
    For 行 = 1 To データ行数
        早遅手動指定フラグ(行) = _
            Def.譜面テーブル.早遅手動指定列(行) <> "" Or Def.譜面テーブル.早遅フレーム手動指定列(行) <> ""
        MAX可能性フラグ(行) = False
    Next 行
    
    Def.処理ログ.出力 "最早COOL: " & 譜面テーブル.最早COOLフレーム & " / 最遅COOL: " & 譜面テーブル.最遅COOLフレーム
    
    If MsgBox("現在のテーブルで処理を開始します。よろしいですか？" & vbCrLf & _
        "譜面テーブル名: " & Def.譜面テーブル.OwnTable.name, vbOKCancel + vbInformation) = vbOK Then
        切替解析初期設定実行 = 0
    Else
        MsgBox "処理を中止します。", vbCritical
    End If
    
    Application.StatusBar = "切替解析の初期設定が完了しました。"
    
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function スコアタ解析初期設定実行(Optional ByVal 譜面テーブル名 As String = "") As Long
    
    Application.StatusBar = "スコアタ解析の初期設定を行います......"
    
    スコアタ解析初期設定実行 = -1
    
    Def.マーク文字設定
    Def.文字定数設定
    
    If Def.スコアタ解析用定数設定 < 0 Then
        Exit Function
    End If
    
    If Def.譜面テーブル設定(譜面テーブル名) < 0 Then
        Exit Function
    End If
    
    Dim 切替結果テーブル名 As String
    Set Def.切替結果テーブル = New SwitchingTable
    切替結果テーブル名 = 譜面テーブル.OwnTable.name & "_切替"
    
    Do While Def.切替結果テーブル.オブジェクト設定(切替結果テーブル名) < 0
        切替結果テーブル名 = InputBox("切替結果テーブルを自動で発見できません。" & vbCrLf & "テーブル名を入力してください。")
        If 切替結果テーブル名 = "" Then
            If MsgBox("テーブル名が未入力です。処理を終了しますか？", vbOKCancel + vbQuestion) = vbOK Then
                GoTo テーブルの設定を中止した場合
            End If
        End If
    Loop
    
    Dim 早遅切替一覧テーブル名 As String
    Set Def.早遅切替一覧テーブル = New ElSwTable
    早遅切替一覧テーブル名 = 譜面テーブル.OwnTable.name & "_早遅切替リスト"
    
    Do While Def.早遅切替一覧テーブル.オブジェクト設定(早遅切替一覧テーブル名) < 0
        早遅切替一覧テーブル名 = InputBox("早遅と切替の一覧テーブルを自動で発見できません。" & vbCrLf & "テーブル名を入力してください。")
        If 早遅切替一覧テーブル名 = "" Then
            If MsgBox("テーブル名が未入力です。処理を終了しますか？", vbOKCancel + vbQuestion) = vbOK Then
                GoTo テーブルの設定を中止した場合
            End If
        End If
    Loop
    
    Def.早遅切替一覧テーブル.早遅切替リスト設定 Def.譜面テーブル, True
    
    If Def.スコアタテーブル設定(Def.譜面テーブル) < 0 Then
        Exit Function
    End If
    
    If Def.処理ログ出力設定() < 0 Then
        Exit Function
    End If
    
    Dim データ行数 As Long
    データ行数 = Def.譜面テーブル.データ行数
    
    ReDim 早遅手動指定フラグ(データ行数)
    ReDim MAX可能性フラグ(データ行数)
    
    Dim 行 As Long
    For 行 = 1 To データ行数
        早遅手動指定フラグ(行) = (Def.譜面テーブル.早遅フレーム手動指定列(行) <> "")
        MAX可能性フラグ(行) = False
    Next 行
    
    Def.処理ログ.出力 "最早COOL: " & 譜面テーブル.最早COOLフレーム & " / 最遅COOL: " & 譜面テーブル.最遅COOLフレーム
    
    If MsgBox("下記のテーブル名に一致するテーブルを使って処理を開始します。よろしいですか？" & vbCrLf & _
        "譜面テーブル名: " & Def.譜面テーブル.OwnTable.name & vbCrLf & _
        "切替結果テーブル名: " & Def.切替結果テーブル.OwnTable.name & vbCrLf & _
        "早遅と切替の一覧テーブル名: " & Def.早遅切替一覧テーブル.OwnTable.name, vbOKCancel + vbInformation) = vbOK Then
        スコアタ解析初期設定実行 = 0
    Else
        MsgBox "処理を中止します。", vbCritical
    End If
    
    Application.StatusBar = "スコアタ解析の初期設定が完了しました。"
    Exit Function
    
テーブルの設定を中止した場合:
    MsgBox "テーブルを設定できないため、処理を中止します。", vbCritical
    Exit Function
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function 切替解析終了設定実行( _
    ByVal 開始時間 As Long, _
    ByVal 終了時間 As Long, _
    ByVal 画面更新 As Boolean, _
    ByVal 自動再計算 As Boolean) _
    As Long
    
    Application.StatusBar = "切替解析の終了処理を行います......"
        
    Def.切替結果テーブル.オブジェクト最終整形
    
    Dim 最善切替情報 As Switching
    Set 最善切替情報 = Def.切替結果テーブル.最善切替情報取得(譜面テーブル, 自動再計算)
    
    Def.処理ログ.出力 "解析時間: " & CDbl(終了時間 - 開始時間) / 1000 & " 秒 (画面更新: " & 画面更新 & ", 自動再計算: " & 自動再計算 & ")"
    Def.処理ログ.出力 最善切替情報.切替文字列, False
    
    Def.処理ログ.ファイル出力終了
    
    Dim cb As New dataobject
    cb.SetText 最善切替情報.切替文字列
    cb.PutInClipboard
    Set cb = Nothing
    
    Dim MAX可能性 As OutputString
    If MsgBox("解析による切替結果は以下の通りです。" & vbCrLf & _
        "(この文字列はクリップボードにコピーされています。)" & vbCrLf & vbCrLf & _
        最善切替情報.切替文字列 & vbCrLf & vbCrLf & _
        "解析時間: " & CDbl(終了時間 - 開始時間) / 1000 & " 秒" & vbCrLf & vbCrLf & _
        "この切替情報を譜面テーブルに反映させ、その情報を基に早遅を設定しますか？", _
        vbYesNo + vbInformation) = vbYes Then
        
        Set MAX可能性 = 最善切替情報.切替早遅情報反映(自動再計算)
        Def.処理ログ.出力 MAX可能性.表示文字列
        
        If Def.早遅切替一覧テーブル設定(譜面テーブル) < 0 Then
            Exit Function
        End If
        Def.早遅切替一覧テーブル.早遅切替情報読み込み Def.譜面テーブル
        
    End If
    
    切替解析終了設定実行 = 0
    
    Application.StatusBar = "切替解析の終了処理が完了しました。"
    
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function スコアタ解析終了処理実行( _
    ByVal 開始時間 As Long, _
    ByVal 終了時間 As Long, _
    ByVal 画面更新 As Boolean, _
    ByVal 自動再計算 As Boolean) _
    As Long
    
    Application.StatusBar = "スコアタ解析の終了処理を行います......"
    
    Def.スコアタテーブル.オブジェクト最終整形 Def.譜面テーブル
    
'    Def.早遅切替一覧テーブル.早遅切替情報書き出し Def.譜面テーブル
    
    Def.処理ログ.出力 "解析時間: " & CDbl(終了時間 - 開始時間) / 1000 & " 秒 (画面更新: " & 画面更新 & ", 自動再計算: " & 自動再計算 & ")"
'    Def.処理ログ.出力 最善切替情報.切替文字列, False
    
    Def.処理ログ.ファイル出力終了
    
'    Dim cb As New dataobject
'    cb.SetText "" 'クリップボード
'    cb.PutInClipboard
'    Set cb = Nothing
    
    スコアタ解析終了処理実行 = 0
    
    Application.StatusBar = "スコアタ解析の終了処理が完了しました。"
    
End Function


' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function 文字列連結(ByVal 文字列1 As String, ByVal 文字列2 As String, Optional ByVal 接続文字列 As String = vbCrLf) As String
    If 文字列1 <> "" Then
        If 文字列2 <> "" Then
            文字列連結 = 文字列1 & 接続文字列 & 文字列2
        Else
            文字列連結 = 文字列1
        End If
    Else
        If 文字列2 <> "" Then
            文字列連結 = 文字列2
        Else
            文字列連結 = ""
        End If
    End If
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function テーブルとシートの検索設定(ByRef テーブル As ListObject, ByRef シート As Worksheet, Optional ByVal テーブル名 As String = "") As Long

    On Error GoTo テーブルを発見できない場合
    
    If テーブル名 = "" Then
        Set テーブル = ActiveSheet.ListObjects(1)
        Set シート = ActiveSheet
    Else
        Dim sheet As Worksheet
        Dim list As ListObject
        For Each sheet In ThisWorkbook.Worksheets
            For Each list In sheet.ListObjects
                If list.name = テーブル名 Then
                    Set テーブル = list
                    Set シート = sheet
                End If
            Next list
        Next sheet
        If シート Is Nothing Then
            GoTo テーブルを発見できない場合
        End If
    End If
    
    On Error GoTo 0
    
    テーブルとシートの検索設定 = 0
    Exit Function
    
テーブルを発見できない場合:

    テーブルとシートの検索設定 = -1
    Exit Function

End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function 譜面テーブル検索設定(Optional ByVal 譜面テーブル名 As String = "") As Long
    
    Def.譜面テーブル設定 譜面テーブル名, "譜面テーブルを発見できません。"
    
    Do While 譜面テーブル.OwnTable Is Nothing
        譜面テーブル名 = InputBox("譜面テーブルを自動で発見できません。" & vbCrLf & "譜面テーブル名を入力してください。")
        If 譜面テーブル名 = "" Then
            If MsgBox("譜面テーブル名が未入力です。処理を終了しますか？", vbOKCancel + vbQuestion) = vbOK Then
                GoTo テーブルの設定を中止した場合
            End If
        End If
        Def.譜面テーブル設定 譜面テーブル名, "譜面テーブルを発見できません。"
    Loop
    
    If MsgBox("下記のテーブル名のテーブルで処理を開始します。よろしいですか？" & vbCrLf & _
        "テーブル名: " & Def.譜面テーブル.OwnTable.name, vbOKCancel + vbInformation) <> vbOK Then
        GoTo テーブルの設定を中止した場合
    End If
    
    譜面テーブル検索設定 = 0
    Exit Function
    
テーブルの設定を中止した場合:
    
    MsgBox "処理を中止しました。", vbCritical
    譜面テーブル検索設定 = -1
    Exit Function
    
End Function

Public Function マーク文字設定()
    Def.マーク文字(1) = "△"
    Def.マーク文字(2) = "□"
    Def.マーク文字(3) = "×"
    Def.マーク文字(4) = "○"
    Def.スライドマーク文字(1) = "←"
    Def.スライドマーク文字(2) = "→"
End Function

Public Function 文字定数設定()
    Def.早COOL文字 = ThisWorkbook.Names("EarlyCoolString").RefersToRange.value
    Def.ジャストCOOL文字 = ThisWorkbook.Names("JustCoolString").RefersToRange.value
    Def.遅COOL文字 = ThisWorkbook.Names("LateCoolString").RefersToRange.value
    Def.HOLD文字 = ThisWorkbook.Names("HoldMarker").RefersToRange.value
End Function

Public Function スコアタ解析用定数設定() As Long
    
    Def.赤WRONG文字 = "赤WRONG"
    Def.WORST文字 = "WORST"
    
    Def.MISSTAKE文字 = "MISS×TAKE"
    
    Def.スコアタスキップ点 = ThisWorkbook.Names("StoppingScoreAttackGap").RefersToRange.value
    If ThisWorkbook.Names("MaxDelayFrame").RefersToRange.value = "" Then
        Def.CC最大余裕フレーム量 = -Def.未MAXフレーム最大値
    Else
        Def.CC最大余裕フレーム量 = ThisWorkbook.Names("MaxDelayFrame").RefersToRange.value
    End If
    Def.最大ライフ量 = ThisWorkbook.Names("MaximumLife").RefersToRange.value
    
    Dim 評価テーブル As ListObject
    Set 評価テーブル = ThisWorkbook.Worksheets("Rating").ListObjects("RatingTable")
    Dim rate行 As Long
    
    On Error GoTo 評価略号の設定に失敗した場合
    
    Set 評価略号 = New Dictionary
    For rate行 = 1 To 評価テーブル.ListRows.count
            評価略号.Add 評価テーブル.ListColumns("Sign").DataBodyRange(rate行).value, 評価テーブル.ListColumns("Small Sign").DataBodyRange(rate行).value
    Next rate行
    
    On Error GoTo 0
    
    スコアタ解析用定数設定 = 0
    Exit Function
    
評価略号の設定に失敗した場合:
    
    MsgBox "ERR:評価略号の設定に失敗しました。" & vbCrLf & "処理を終了します。", vbCritical
    
    スコアタ解析用定数設定 = -1
    Exit Function
    
End Function

Public Function 譜面テーブル設定(Optional ByVal 譜面テーブル名 As String = "", Optional ByVal エラー文字列 As String = "") As Long
    If MsgBox("譜面テーブルの仮想化を行いますか？", vbYesNo + vbInformation) = vbYes Then
        Set Def.譜面テーブル = New ChartTable2
    Else
        Set Def.譜面テーブル = New ChartTable
    End If
    譜面テーブル設定 = Def.譜面テーブル.オブジェクト設定(譜面テーブル名)
    If 譜面テーブル設定 < 0 Then
        GoTo テーブル設定に失敗した場合
    End If
    Exit Function
テーブル設定に失敗した場合:
    If エラー文字列 = "" Then
        エラー文字列 = "ERR:処理を行うテーブルを見つけられませんでした。" & vbCrLf & "処理を終了します。"
    End If
    If 譜面テーブル名 <> "" Then
        エラー文字列 = エラー文字列 & vbCrLf & "テーブル名: " & 譜面テーブル名
    End If
    MsgBox エラー文字列, vbCritical
    Exit Function
End Function

Public Function 処理ログ出力設定() As Long
    Set Def.処理ログ = New ProcessLog
    処理ログ出力設定 = Def.処理ログ.ファイル出力開始()
    If 処理ログ出力設定 < 0 Then
        GoTo ログ出力設定に失敗した場合
    End If
    Exit Function
ログ出力設定に失敗した場合:
    MsgBox "ERR:ログファイルを作成できません。" & vbCrLf & "処理を終了します。", vbCritical
    Exit Function
End Function

Public Function 切替結果テーブル設定(ByVal 譜面テーブル As IChartTable) As Long
    Set Def.切替結果テーブル = New SwitchingTable
    切替結果テーブル設定 = Def.切替結果テーブル.オブジェクト新規作成(譜面テーブル)
    If 切替結果テーブル設定 < 0 Then
        GoTo テーブル設定に失敗した場合
    End If
    Exit Function
テーブル設定に失敗した場合:
    MsgBox "ERR:結果テーブルの作成に失敗しました。" & vbCrLf & "処理を終了します。", vbCritical
    Exit Function
End Function

Public Function 早遅切替一覧テーブル設定(ByVal 譜面テーブル As IChartTable) As Long
    Set Def.早遅切替一覧テーブル = New ElSwTable
    早遅切替一覧テーブル設定 = Def.早遅切替一覧テーブル.オブジェクト新規作成(譜面テーブル)
    If 早遅切替一覧テーブル設定 < 0 Then
        GoTo テーブル設定に失敗した場合
    End If
    Exit Function
テーブル設定に失敗した場合:
    MsgBox "ERR:早遅と切替の一覧テーブルの作成に失敗しました。" & vbCrLf & "処理を終了します。", vbCritical
    Exit Function
End Function

Public Function スコアタテーブル設定(ByVal 譜面テーブル As IChartTable) As Long
    Set Def.スコアタテーブル = New ScoreRouteTable
    スコアタテーブル設定 = Def.スコアタテーブル.オブジェクト新規作成(譜面テーブル)
    If スコアタテーブル設定 < 0 Then
        GoTo テーブル設定に失敗した場合
    End If
    Exit Function
テーブル設定に失敗した場合:
    MsgBox "ERR:スコアタルートの結果テーブルの作成に失敗しました。" & vbCrLf & "処理を終了します。", vbCritical
    Exit Function
End Function

