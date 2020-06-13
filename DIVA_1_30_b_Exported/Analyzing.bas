Attribute VB_Name = "Analyzing"
Option Explicit
Option Base 1

#Const 切替詳細ログ = False
#Const スコアタ詳細ログ = False
#Const スコアタ通常ログ = True
Private スコアタネストカウンタ As Long

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Sub スコアタ解析()
    
    Application.StatusBar = "スコアタ解析を開始します......"
    
    Application.Calculation = xlCalculationAutomatic
    DoEvents
    
    #If スコアタ詳細ログ Then
        詳細ログ出力 "スコアタ解析初期設定実行" & vbTab & "開始"
    #End If
    
    If Def.スコアタ解析初期設定実行() <> 0 Then
        Rescue
        Exit Sub
    End If
    
    #If スコアタ詳細ログ Then
        詳細ログ出力 "スコアタ解析初期設定実行" & vbTab & "終了"
    #End If
    
    Def.スコアタテーブル.Is簡易版 = (MsgBox("簡易版にしますか？" & vbCrLf & _
        "(十分に離れた部分のルートを分離して検討しませんが、早遅切替一覧テーブルの出力に時間を使いません。)", _
        vbYesNo + vbInformation) = vbYes)
'    Def.スコアタテーブル.Is簡易版 = True
    
    Dim 画面更新 As Boolean
    Dim 自動再計算 As Boolean
    
    Dim 開始時間 As Long
    Dim 終了時間 As Long
    
    画面更新 = MsgBox("試行ルートごとの画面更新を行いますか？", vbYesNo + vbInformation) = vbYes
    自動再計算 = MsgBox("ブックの自動再計算を行いますか？", vbYesNo + vbInformation) = vbYes
    
    開始時間 = Def.GetTickCount
    
    #If スコアタ通常ログ Then
        Def.処理ログ.フォーム出力開始
    #End If
    
    自動スコアタルート判定 画面更新, 自動再計算 ', 670
    
    #If スコアタ通常ログ Then
        Def.処理ログ.フォーム出力終了
    #End If
    
    終了時間 = Def.GetTickCount
    
    #If スコアタ詳細ログ Then
        詳細ログ出力 "スコアタ解析終了設定実行" & vbTab & "開始"
    #End If
    
    Def.スコアタ解析終了処理実行 開始時間, 終了時間, 画面更新, 自動再計算
    
    #If スコアタ詳細ログ Then
        詳細ログ出力 "スコアタ解析終了設定実行" & vbTab & "終了"
    #End If
    
    MsgBox "正常に終了しました。ありがほー。" & vbCrLf & "解析時間: " & CDbl(終了時間 - 開始時間) / 1000 & " 秒"
    
    Rescue
    
End Sub

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Sub 切替解析()
    
    Application.StatusBar = "切替解析を開始します......"
    
    Application.Calculation = xlCalculationAutomatic
    DoEvents
    
    #If 切替詳細ログ Then
        詳細ログ出力 "切替解析初期設定実行" & vbTab & "開始"
    #End If
    
    If Def.切替解析初期設定実行() <> 0 Then
        Rescue
        Exit Sub
    End If
    
    #If 切替詳細ログ Then
        詳細ログ出力 "切替解析初期設定実行" & vbTab & "終了"
    #End If
    
    Dim 画面更新 As Boolean
    Dim 自動再計算 As Boolean
    
    Dim 開始時間 As Long
    Dim 終了時間 As Long
    
    画面更新 = MsgBox("試行ルートごとの画面更新を行いますか？", vbYesNo + vbInformation) = vbYes
    自動再計算 = MsgBox("ブックの自動再計算を行いますか？", vbYesNo + vbInformation) = vbYes
    
    開始時間 = Def.GetTickCount
    
    自動切替判定 画面更新, 自動再計算
    
    終了時間 = Def.GetTickCount
    
    #If 切替詳細ログ Then
        詳細ログ出力 "切替解析終了設定実行" & vbTab & "開始"
    #End If
    
    Def.切替解析終了設定実行 開始時間, 終了時間, 画面更新, 自動再計算
    
    #If 切替詳細ログ Then
        詳細ログ出力 "切替解析終了設定実行" & vbTab & "終了"
    #End If
    
    MsgBox "正常に終了しました。ありがほー。"
    
    Rescue
    
End Sub

' ======================================================================================================================================================================================================
'
' ※ (開始行 - 1) 行目も評価対象になります。→ 開始行 = 1 はダメ
' ======================================================================================================================================================================================================

Public Function 自動スコアタルート判定( _
    Optional ByVal 画面更新 As Boolean = False, _
    Optional ByVal 自動再計算 As Boolean = False, _
    Optional ByVal 判定開始行 As Long = 1, _
    Optional ByVal 判定終了行 As Long = -1, _
    Optional ByVal 完奏モード As Boolean = False)
    
    ' 開始処理 -------------------------------------------------------------------------------------
    
    #If スコアタ詳細ログ Then
        詳細ログ出力 "自動スコアタルート判定" & vbTab & "初期設定開始", True
    #End If
    
    Application.ScreenUpdating = False
    
    If 自動再計算 Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    
    If 判定終了行 = -1 Then
        判定終了行 = Def.譜面テーブル.データ行数
    End If
    
    ' 早遅を反映してその確定スコアを配列に格納
    
    Def.早遅切替一覧テーブル.早遅切替情報書き出し Def.譜面テーブル
    Def.譜面テーブル.再計算 判定開始行, 判定終了行, 自動再計算
    Dim 行スコア() As Long
    行スコア = 行スコア設定(判定開始行, 判定終了行)
    
    ' メイン処理 -----------------------------------------------------------------------------------
    
    #If スコアタ詳細ログ Then
        詳細ログ出力 "自動スコアタルート判定" & vbTab & "開始", True
    #End If
    
    Application.StatusBar = "スコアタ解析を開始します......"
    
    スコアタネストカウンタ = -1
    指定範囲のルート判定 判定開始行, 判定終了行, 行スコア, 完奏モード, 画面更新, 自動再計算
    Def.スコアタテーブル.組み合わせルート出力 Def.譜面テーブル, Def.早遅切替一覧テーブル, 完奏モード
    
    ' 終了処理 ------------------------------------------------------------------------------------
    
    Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True
    
    #If スコアタ詳細ログ Then
        詳細ログ出力 "自動スコアタルート判定" & vbTab & "終了", True
    #End If
    
End Function

Private Function 行スコア設定(ByVal 開始行 As Long, ByVal 終了行 As Long) As Long()
    Dim 行スコア() As Long
    ReDim 行スコア(Def.譜面テーブル.データ行数)
    Dim 行 As Long
    For 行 = 開始行 To 終了行
        行スコア(行) = 現在行スコア取得(行)
    Next 行
    行スコア設定 = 行スコア
End Function

Private Function 現在行スコア取得(ByVal 行 As Long) As Long
    If 行 = 1 Then
        現在行スコア取得 = Def.譜面テーブル.スコア列(行)
    ElseIf Def.譜面テーブル.ホールド開始フレーム列(行 - 1) = 0 Then
        現在行スコア取得 = Def.譜面テーブル.スコア列(行)
    ElseIf Def.譜面テーブル.ホールドボーナス列(行) > 0 Then
        現在行スコア取得 = Def.譜面テーブル.スコア列(行)
    Else
        現在行スコア取得 = 0
    End If
End Function

' ======================================================================================================================================================================================================
'
' 戻り値はルートを完走した場合のみ、そのルートの行スコアとなる
' ======================================================================================================================================================================================================

Private Function 指定範囲のルート判定( _
    ByVal 開始行 As Long, _
    ByVal 終了行 As Long, _
    ByRef 現在ルート行スコア() As Long, _
    Optional ByVal 完奏モード As Boolean = False, _
    Optional ByVal 画面更新 As Boolean = False, _
    Optional ByVal 自動再計算 As Boolean = False, _
    Optional ByVal 影響開始行 As Long = 0) _
    As Def.行スコア情報
'    Optional ByRef 比較ルート行スコア As Def.行スコア情報, _

    #If スコアタ詳細ログ Then
        詳細ログ出力 "指定範囲のルート判定" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "開始", True
    #End If
    
    #If スコアタ通常ログ Then
        スコアタネストカウンタ = スコアタネストカウンタ + 1
        If スコアタネストカウンタ > 0 Then
            詳細ログ出力 String(スコアタネストカウンタ - 1, "│") & "├" & 開始行 & ",", True, スコアタネストカウンタ + 1
        Else
            詳細ログ出力 開始行 & ",", True, スコアタネストカウンタ + 1
        End If
    #End If
    
    Dim 対象ルート行スコア As Def.行スコア情報
    対象ルート行スコア = 現ルート有効判定と出力(開始行, 終了行, 現在ルート行スコア, 完奏モード, 画面更新, 自動再計算, 影響開始行)
    
    If 対象ルート行スコア.最大影響行 = 0 Then
        対象ルート行スコア.行スコア = 現在ルート行スコア
    End If
    
    Dim isホールド開始(Def.マーク数) As Boolean
    
    Dim 直前ホールド開始フレーム As Long
    Dim 切替再計算開始行 As Long
    Dim 開始早ずれ許容フレーム As Long
    
    Dim isルート開始行 As Boolean
    
    Dim 切替可能性 As Boolean
    
    Dim 行 As Long
    Dim マーク As Long
    
    For 行 = 開始行 To 終了行
        
        Def.譜面テーブル.再計算 行, , 自動再計算
        
        ' 現在の比較対象の行スコアより一定以上スコアが落ちた場合は終了
        Dim 現在スコア As Long
        現在スコア = 現在行スコア取得(行)
        If 現在ルート行スコア(行) > 0 And 現在スコア > 0 Then
            If 現在スコア < 現在ルート行スコア(行) + Def.スコアタスキップ点 Then
                Exit For
            End If
        End If
        
        '完奏モードでない場合、ライフが0なら終了
        If (Not 完奏モード) And Def.譜面テーブル.ライフ列(行) = 0 Then
            Exit For
        End If
        
        '50コンボ以上でライフが最大なら、以前に設定された現在ルート行スコアと同じスコア上昇ペースになるので終了
        If Not (Def.スコアタテーブル.Is簡易版) And 影響開始行 > 0 Then
            If Def.譜面テーブル.コンボ列(行) >= 50 And Def.譜面テーブル.ライフ列(行) = Def.最大ライフ量 Then
                Exit For
            End If
        End If
        
        ' ルート開始位置になれるかどうかを判定
        isルート開始行 = False
        For マーク = 1 To Def.マーク数
            If Def.譜面テーブル.ノーツ列(マーク, 行) = Def.HOLD文字 Then
                isホールド開始(マーク) = True
                isルート開始行 = True
            Else
                isホールド開始(マーク) = False
            End If
        Next マーク
        
        If isルート開始行 Then
            ' ルート開始位置になれる場合の処理
            
            ' 開始行の切替設定(切替は強制変更)
            切替可能性 = False
            For マーク = 1 To Def.マーク数
                If Def.譜面テーブル.ホールド個別開始フレーム列(マーク, 行 - 1) > 0 Then
                    切替可能性 = True
                End If
            Next
            If 切替可能性 And Not (Def.譜面テーブル.ホールドボーナス列(行) > 0) Then
                Def.譜面テーブル.切替判定列(行) = True
            End If
            
            ' 直前ブロックの切替と早遅の再設定
            切替再計算開始行 = 指定行直前までの切替再計算(開始行, 行, 画面更新, 自動再計算)
            
            ' 開始行となる現在の行のフレーム前ずれ許容フレームを設定して、早遅フレームと評価をリセット
            ' ※直前がC-SdでMAXのルートなど、最早COOLでMAXにならない場合はリセットしない
            開始早ずれ許容フレーム = Def.譜面テーブル.デフォルト早ずれ許容フレーム
            If Def.譜面テーブル.ホールドボーナス列(行) > 0 Then
                ' MAXが入る場合のみデフォルトから変更
                If Def.譜面テーブル.ホールドフレーム列(行) > Def.未MAXフレーム最大値 Then
                    開始早ずれ許容フレーム = Application.WorksheetFunction.Max(開始早ずれ許容フレーム, _
                        Def.未MAXフレーム最大値 - Def.譜面テーブル.ホールドフレーム列(行) + Def.譜面テーブル.早遅フレーム列(行) + 1)
                End If
            End If
            
            早遅フレームと評価の解除 行, , Not (開始早ずれ許容フレーム > Def.譜面テーブル.最早COOLフレーム)
            Def.譜面テーブル.再計算 行, , 自動再計算
            
            '現在行の次の行からのルート探索
            If 影響開始行 = 0 Then
                指定行からのルート探索 行, 終了行, isホールド開始, 対象ルート行スコア.行スコア, 開始早ずれ許容フレーム, , 完奏モード, 画面更新, 自動再計算, 切替再計算開始行
            Else
                指定行からのルート探索 行, 終了行, isホールド開始, 対象ルート行スコア.行スコア, 開始早ずれ許容フレーム, , 完奏モード, 画面更新, 自動再計算, 影響開始行
            End If
            
            '切替と早遅を元に戻す
            Def.早遅切替一覧テーブル.早遅切替情報書き出し Def.譜面テーブル, 切替再計算開始行, 行, False
            Def.譜面テーブル.再計算 切替再計算開始行, 行, 自動再計算
            
        End If
        
    Next 行
    
    ' 戻り値を設定
    指定範囲のルート判定 = 対象ルート行スコア
    
    ' 早遅リセットはなくても大丈夫?
    
    #If スコアタ詳細ログ Then
        詳細ログ出力 "指定範囲のルート判定" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "終了", True
    #End If
    
    #If スコアタ通常ログ Then
        Def.処理ログ.フォーム文字列削除 スコアタネストカウンタ + 1
        スコアタネストカウンタ = スコアタネストカウンタ - 1
    #End If
    
End Function

Private Function 現ルート有効判定と出力( _
    ByVal 開始行 As Long, _
    ByVal 終了行 As Long, _
    ByRef 現在ルート行スコア() As Long, _
    Optional ByVal 完奏モード As Boolean = False, _
    Optional ByVal 画面更新 As Boolean = False, _
    Optional ByVal 自動再計算 As Boolean = False, _
    Optional ByVal 影響開始行 As Long = 0) _
    As Def.行スコア情報
'    Optional ByRef 比較ルート行スコア As Def.行スコア情報, _

    #If スコアタ詳細ログ Then
        詳細ログ出力 "現ルート有効判定と出力" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "開始", True
    #End If
    
    Dim is有効ルート As Boolean
    is有効ルート = True
    
    Dim 行 As Long
    
    For 行 = 開始行 To 終了行
        
        Def.譜面テーブル.再計算 行, , 自動再計算
        
        If is有効ルート Then
            
            ' 現在の比較対象の行スコアより一定以上スコアが落ちた場合は最後まで計算せずに終了
            Dim 現在スコア As Long
            現在スコア = 現在行スコア取得(行)
            If 現在ルート行スコア(行) > 0 And 現在スコア > 0 Then
                If 現在スコア < 現在ルート行スコア(行) + Def.スコアタスキップ点 Then
                    is有効ルート = False
                    Exit For
                End If
            End If
        
            '完奏モードでない場合、ライフが0なら終了
            If (Not 完奏モード) And Def.譜面テーブル.ライフ列(行) = 0 Then
                is有効ルート = False
                Exit For
            End If
            
            '50コンボ以上でライフが最大なら、以前に設定された現在ルート行スコアを下回ることはないので確定
'            If 行 > 比較ルート行スコア.最大影響行 Then
            If Not (Def.スコアタテーブル.Is簡易版) And 影響開始行 > 0 Then
                If Def.譜面テーブル.コンボ列(行) >= 50 And Def.譜面テーブル.ライフ列(行) = Def.最大ライフ量 Then
                    Exit For
                End If
            End If
            
        End If
        
    Next 行
    
    Dim リザルト As Def.結果セット
    
    ' 完奏モードでない場合、MISS×TAKE(クリアゲージが満たない)なら終了
    If is有効ルート Then
        'Def.譜面テーブル.再計算 行 + 1, 終了行, 自動再計算
        Def.譜面テーブル.再計算 行 + 1, Def.譜面テーブル.データ行数, 自動再計算
        リザルト = Def.譜面テーブル.リザルト再計算()
        If (Not 完奏モード) And リザルト.クリアランク = Def.MISSTAKE文字 Then
            is有効ルート = False
        End If
    End If
    
    Dim 結果 As Def.行スコア情報
    
    ' 戻り値となるルート行スコアを設定
    If is有効ルート Then
        
        結果.最大影響行 = 切替影響終了行取得(行)
        
        結果.行スコア = 行スコア設定(開始行, 終了行)
        
        現ルート有効判定と出力 = 結果
    Else
'        現ルート有効判定と出力 = 現在ルート行スコア
'        現ルート有効判定と出力 = 結果
        #If スコアタ詳細ログ Then
            詳細ログ出力 "現ルート有効判定と出力" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "終了(無効)", True
        #End If
        Exit Function
    End If
    
    ' 有効ルートでない場合はここで終了
    ' 有効ルートである場合のみ出力処理
    
    #If スコアタ詳細ログ Then
        詳細ログ出力 "現ルート有効判定と出力" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "出力処理開始", True
    #End If
    
    Dim スコアタルート文字列 As String
    スコアタルート文字列 = スコアタルート文字列取得()
    
    ' スコアタテーブル(と新しい早遅切替一覧テーブル)に結果出力
    If 影響開始行 = 0 Then
        影響開始行 = 開始行
    End If
    
    Def.スコアタテーブル.現在ルート出力 Def.譜面テーブル, スコアタルート文字列, 影響開始行, 結果.最大影響行, リザルト.達成率, リザルト.スコア
    
    #If スコアタ詳細ログ Then
        詳細ログ出力 "現ルート有効判定と出力" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "終了(有効)", True
    #End If
    
End Function

Public Function スコアタルート文字列取得() As String
    
    Dim スコアタ文字列 As String
    Dim 現在ルート中文字列 As String
    スコアタ文字列 = ""
    
    Dim isホールド開始(Def.マーク数) As Boolean
    
    Dim isルート開始行 As Boolean
    
    Dim 行 As Long
    Dim マーク As Long
    
    Dim isルート中 As Boolean
    Dim isルート終了 As Boolean
    isルート中 = False
    isルート終了 = False
    
    Dim is仮ルート中 As Boolean
    is仮ルート中 = False
        
    Dim 直前ホールド開始行 As Long
    Dim w数 As Long
    Dim isホールド中(Def.マーク数) As Boolean
    
    Dim isホールド開始行 As Boolean
    
    Dim 開始フレームずれ As Long
    Dim 終了フレームずれ As Long
    Dim 開始フレーム余裕 As Long
    Dim 終了フレーム余裕 As Long
    
    For 行 = 1 To Def.譜面テーブル.データ行数
        
        ' ルート終了かどうかの判定
        If Def.譜面テーブル.ホールドボーナス列(行) > 0 Then
            If isルート中 Then
                isルート終了 = True
            ElseIf Def.譜面テーブル.ホールドフレーム列(行) > Def.未MAXフレーム最大値 Then
                ' ルート中でなくても余裕のないMAXの場合は遡ってルートとして文字列追加
                If Def.譜面テーブル.ホールド終了押し直し判定列(行) And _
                    Def.譜面テーブル.ホールドフレーム列(行) < Def.未MAXフレーム最大値 + 1 + Def.CC最大余裕フレーム量 Then
                    ルート開始時変数設定 行 - 1, True, 直前ホールド開始行, 現在ルート中文字列, w数, isホールド中
                    isルート終了 = True
                End If
            End If
        End If
        
        'ルート終了の場合の文字列処理
        If isルート終了 Then
            
            現在ルート中文字列 = 現在ルート中文字列 & " →【"
            
            Dim ボタン数 As Long
            ボタン数 = 0
            For マーク = 1 To マーク数
                If isホールド中(マーク) Then
                    ボタン数 = ボタン数 + 1
                End If
            Next マーク
            If ボタン数 > 1 Then
                現在ルート中文字列 = 現在ルート中文字列 & ボタン数
            End If
            
            If Def.譜面テーブル.ホールドフレーム列(行) > Def.未MAXフレーム最大値 Then
                現在ルート中文字列 = 現在ルート中文字列 & "MAX】"
            Else
                現在ルート中文字列 = 現在ルート中文字列 & "HOLD】"
            End If
            
            Dim 次ノーツ行 As Long
            Dim 抜けノーツ文字列 As String
            抜けノーツ文字列 = ""
            For 次ノーツ行 = 行 To Def.譜面テーブル.データ行数
                For マーク = 1 To Def.マーク数
                    If Def.譜面テーブル.ノーツ列(マーク, 次ノーツ行) <> "" Then
                        抜けノーツ文字列 = 抜けノーツ文字列 & Def.マーク文字(マーク)
                    End If
                Next マーク
                If 抜けノーツ文字列 <> "" Then
                    Exit For
                End If
            Next
            If 抜けノーツ文字列 <> "" Then
                現在ルート中文字列 = 現在ルート中文字列 & "→ " & ノーツ番号文字列の取得(行) & 抜けノーツ文字列
            End If
            
            If Def.譜面テーブル.ホールドフレーム列(行) > Def.未MAXフレーム最大値 And _
                Def.譜面テーブル.ホールドフレーム列(行) < Def.未MAXフレーム最大値 + 1 + Def.CC最大余裕フレーム量 _
                And Def.譜面テーブル.ホールド終了押し直し判定列(行) = True Then
                
                '余裕のないMAXである場合は詳細を現在ルート中文字列に追記
                現在ルート中文字列 = 現在ルート中文字列 & "《"
                
                開始フレームずれ = Def.譜面テーブル.早遅フレーム列(直前ホールド開始行)
                終了フレームずれ = Def.譜面テーブル.早遅フレーム列(行)
                
                If 開始フレームずれ < 0 Then
                    開始フレーム余裕 = -1
                ElseIf 開始フレームずれ > 0 Then
                    If Def.譜面テーブル.評価列(直前ホールド開始行) <> Def.譜面テーブル.フレームずれ別評価(0) Then
                        現在ルート中文字列 = 現在ルート中文字列 & "遅"
                    End If
                    開始フレーム余裕 = 0
                Else
                    開始フレーム余裕 = 0
                End If
                現在ルート中文字列 = 現在ルート中文字列 & Def.評価略号(Def.譜面テーブル.評価列(直前ホールド開始行)) & "-"
                Do Until Def.譜面テーブル.フレームずれ別評価(開始フレームずれ - 開始フレーム余裕) = Def.譜面テーブル.フレームずれ別評価(開始フレームずれ)
                    開始フレーム余裕 = 開始フレーム余裕 + 1
                Loop
                
                If 終了フレームずれ < 0 Then
                    If Def.譜面テーブル.評価列(行) <> Def.譜面テーブル.フレームずれ別評価(0) Then
                        現在ルート中文字列 = 現在ルート中文字列 & "早"
                    End If
                    終了フレーム余裕 = 0
                ElseIf 終了フレームずれ > 0 Then
                    終了フレーム余裕 = -1
                Else
                    終了フレーム余裕 = 0
                End If
                現在ルート中文字列 = 現在ルート中文字列 & Def.評価略号(Def.譜面テーブル.評価列(行)) & ",猶予:"
                Do Until Def.譜面テーブル.フレームずれ別評価(終了フレームずれ + 終了フレーム余裕) = Def.譜面テーブル.フレームずれ別評価(終了フレームずれ)
                    終了フレーム余裕 = 終了フレーム余裕 + 1
                Loop
                
                現在ルート中文字列 = 現在ルート中文字列 & (開始フレーム余裕 + 終了フレーム余裕) & "》"
                
            End If
            
            If w数 > 0 Then
                現在ルート中文字列 = 現在ルート中文字列 & " (W" & w数 & ")"
            End If
            
            スコアタ文字列 = Def.文字列連結(スコアタ文字列, 現在ルート中文字列, vbCrLf)
            
            isルート中 = False
            isルート終了 = False
            
            is仮ルート中 = False
            
        End If
        
        ' 切替がある場合の文字列処理
        If Def.譜面テーブル.切替判定列(行) And (Not isルート中) Then
            スコアタ文字列 = Def.文字列連結(スコアタ文字列, 切替文字列取得(行), vbCrLf)
        End If
        
        '仮ルート中の場合
        If is仮ルート中 Then
            
            isホールド開始行 = False
            For マーク = 1 To Def.マーク数
                If Def.譜面テーブル.ノーツ列(マーク, 行) = Def.HOLD文字 Then
                    isホールド開始行 = True
                    Exit For
                End If
            Next マーク
            
            If isホールド開始行 Then
                If Def.譜面テーブル.ホールドフレーム列(行) > Def.未MAXフレーム最大値 - 1 - Def.CC最大余裕フレーム量 Then
                    isルート中 = True
                    ルート開始時変数設定 行 - 1, True, 直前ホールド開始行, 現在ルート中文字列, w数, isホールド中
                End If
                is仮ルート中 = False
            End If
            
        End If
        
        'ルート開始かどうかの判定
        If Not isルート中 Then
            
            If Def.譜面テーブル.評価列(行) = Def.赤WRONG文字 Or Def.譜面テーブル.評価列(行) = Def.WORST文字 Then
                
                '現在の行が赤WRONGやWORSTだった場合は(開始行を遡って検索して)ルート開始
                isルート中 = True
                ルート開始時変数設定 行, True, 直前ホールド開始行, 現在ルート中文字列, w数, isホールド中
                is仮ルート中 = False
                
            ElseIf Def.譜面テーブル.評価列(行) <> Def.譜面テーブル.フレームずれ別評価(0) Then
'                Def.譜面テーブル.ホールドフレーム列(行) < Def.未MAXフレーム最大値 + 1 + Def.CC最大余裕フレーム量
                
                If Def.譜面テーブル.ホールドボーナス列(行) = 0 Then
                    '現在の行のノーツを捨てておらずCOOL以外で、かつホールド終了行でない場合はルート開始
                    isルート中 = True
                    ルート開始時変数設定 行, False, 直前ホールド開始行, 現在ルート中文字列, w数, isホールド中
                    is仮ルート中 = False
                Else
                    'ホールド終了行かつホールド開始行だった場合は仮ルート開始
                    For マーク = 1 To Def.マーク数
                        If Def.譜面テーブル.ノーツ列(マーク, 行) = Def.HOLD文字 Then
                            is仮ルート中 = True
                            Exit For
                        End If
                    Next マーク
                End If
                
            End If
            
        End If
        
        'ルート中の場合の文字列処理
        If isルート中 Then
            
            '捨ててたらw数カウント増加
            If Def.譜面テーブル.評価列(行) = Def.赤WRONG文字 Or Def.譜面テーブル.評価列(行) = Def.WORST文字 Then
                w数 = w数 + 1
            End If
            
            Dim cool可能性 As Boolean
            cool可能性 = COOL可能性取得(行, isホールド中)
            
            If cool可能性 Then
                
                isホールド開始行 = False
                For マーク = 1 To Def.マーク数
                    If Def.譜面テーブル.ノーツ列(マーク, 行) = Def.HOLD文字 Then
                        isホールド開始(マーク) = True
                        isホールド開始行 = True
                    Else
                        isホールド開始(マーク) = False
                    End If
                Next マーク
                
                If isホールド開始行 Then
                    
                    If Def.譜面テーブル.評価列(行) = Def.赤WRONG文字 Or Def.譜面テーブル.評価列(行) = Def.WORST文字 Then
                        'COOLの可能性があるホールドを捨てている場合は現在ルート中文字列に追記
                        
                        現在ルート中文字列 = 現在ルート中文字列 & " (" & ノーツ番号文字列の取得(行, False)
                        For マーク = 1 To Def.マーク数
                            If isホールド開始(マーク) Then
                                現在ルート中文字列 = 現在ルート中文字列 & マーク文字(マーク)
                            End If
                        Next マーク
                        現在ルート中文字列 = 現在ルート中文字列 & Def.評価略号(Def.譜面テーブル.評価列(行)) & ")"
                    Else
                        'ホールドをCOOLでとっている場合はホールド中をTrueにして現在ルート中文字列に追記
                        
                        現在ルート中文字列 = 現在ルート中文字列 & " " & ノーツ番号文字列の取得(行)
                        For マーク = 1 To Def.マーク数
                            If isホールド開始(マーク) Then
                                isホールド中(マーク) = True
                                現在ルート中文字列 = 現在ルート中文字列 & マーク文字(マーク)
                            End If
                        Next
                        
                        If Def.譜面テーブル.ホールドフレーム列(行) > Def.未MAXフレーム最大値 - 1 - Def.CC最大余裕フレーム量 Then
                            
                            '余裕のない接続である場合は詳細を現在ルート中文字列に追記
                            現在ルート中文字列 = 現在ルート中文字列 & "《"
                            
                            開始フレームずれ = Def.譜面テーブル.早遅フレーム列(直前ホールド開始行)
                            終了フレームずれ = Def.譜面テーブル.早遅フレーム列(行)
                            
                            If 開始フレームずれ < 0 Then
                                開始フレーム余裕 = 0
                            ElseIf 開始フレームずれ > 0 Then
                                If Def.譜面テーブル.評価列(直前ホールド開始行) <> Def.譜面テーブル.フレームずれ別評価(0) Then
                                    現在ルート中文字列 = 現在ルート中文字列 & "遅"
                                End If
                                開始フレーム余裕 = -1
                            Else
                                開始フレーム余裕 = 0
                            End If
                            現在ルート中文字列 = 現在ルート中文字列 & Def.評価略号(Def.譜面テーブル.評価列(直前ホールド開始行)) & "-"
                            Do Until Def.譜面テーブル.フレームずれ別評価(開始フレームずれ + 開始フレーム余裕) = Def.譜面テーブル.フレームずれ別評価(開始フレームずれ)
                                開始フレーム余裕 = 開始フレーム余裕 + 1
                            Loop
                            
                            If 終了フレームずれ < 0 Then
                                終了フレーム余裕 = -1
                            ElseIf 終了フレームずれ > 0 Then
                                If Def.譜面テーブル.評価列(行) <> Def.譜面テーブル.フレームずれ別評価(0) Then
                                    現在ルート中文字列 = 現在ルート中文字列 & "遅"
                                End If
                                終了フレーム余裕 = 0
                            Else
                                終了フレーム余裕 = 0
                            End If
                            現在ルート中文字列 = 現在ルート中文字列 & Def.評価略号(Def.譜面テーブル.評価列(行)) & ",猶予:"
                            Do Until Def.譜面テーブル.フレームずれ別評価(終了フレームずれ - 終了フレーム余裕) = Def.譜面テーブル.フレームずれ別評価(終了フレームずれ)
                                終了フレーム余裕 = 終了フレーム余裕 + 1
                            Loop
                            
                            現在ルート中文字列 = 現在ルート中文字列 & (開始フレーム余裕 + 終了フレーム余裕) & "》"
                            
                            直前ホールド開始行 = 行
                            
                        End If
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
    Next 行
    
    スコアタルート文字列取得 = スコアタ文字列
    
End Function

Private Function ルート開始時変数設定( _
    ByVal 現在行 As Long, _
    ByVal 直前ルート開始行探索 As Boolean, _
    ByRef 直前ホールド開始行 As Long, _
    ByRef 現在ルート中文字列 As String, _
    ByRef w数 As Long, _
    ByRef isホールド中() As Boolean)
    
    Dim マーク As Long
    
    ' ルート開始行
    直前ホールド開始行 = 現在行
    If 直前ルート開始行探索 Then
        Do Until Def.譜面テーブル.ホールド開始フレーム列(直前ホールド開始行 - 1) < Def.譜面テーブル.ホールド開始フレーム列(直前ホールド開始行)
            直前ホールド開始行 = 直前ホールド開始行 - 1
        Loop
    End If
    
    ' 現在ルート中文字列
    現在ルート中文字列 = ノーツ番号文字列の取得(直前ホールド開始行)
    For マーク = 1 To Def.マーク数
        If Def.譜面テーブル.ホールド個別開始フレーム列(マーク, 直前ホールド開始行) > Def.譜面テーブル.ホールド個別開始フレーム列(マーク, 直前ホールド開始行 - 1) Then
            現在ルート中文字列 = 現在ルート中文字列 & マーク文字(マーク)
        End If
    Next マーク
    
    ' w数
    w数 = 0
    
    ' isホールド中
    For マーク = 1 To Def.マーク数
        isホールド中(マーク) = (Def.譜面テーブル.ホールド個別開始フレーム列(マーク, 直前ホールド開始行) > 0)
    Next マーク
    
End Function

Private Function 早遅フレームと評価の解除(ByVal 開始行 As Long, Optional ByVal 終了行 As Long = -1, Optional ByVal 評価の解除 As Boolean = True, Optional ByVal 切替の解除 As Boolean = False)
    
    If 終了行 = -1 Then
        終了行 = 開始行
    End If
    
    Dim 行 As Long
    For 行 = 開始行 To 終了行
        If 評価の解除 And (Not Def.早遅手動指定フラグ(行)) Then
            If Def.譜面テーブル.評価列(行) <> Def.譜面テーブル.フレームずれ別評価(0) Then
                Def.譜面テーブル.評価列(行) = Def.譜面テーブル.フレームずれ別評価(0)
            End If
            If 切替の解除 And (Not Def.譜面テーブル.切替判定列(行)) Then
                Def.譜面テーブル.切替判定列(行) = False
            End If
            If Def.譜面テーブル.早遅手動指定列(行) <> "" Then
                Def.譜面テーブル.早遅手動指定列(行) = ""
            End If
            If Def.譜面テーブル.早遅フレーム手動指定列(行) <> "" Then
                Def.譜面テーブル.早遅フレーム手動指定列(行) = ""
            End If
        End If
    Next
    
End Function

Private Function 切替影響開始行取得(ByVal 開始行 As Long, ByVal 指定行 As Long)

    Dim 切替結果行 As Long
    Dim 切替結果ブロック開始行 As Long
    Dim 切替再計算開始行 As Long
    
    For 切替結果行 = 1 To Def.切替結果テーブル.データ行数
        
        切替結果ブロック開始行 = Def.切替結果テーブル.ブロック開始列(切替結果行)
        If 切替結果ブロック開始行 = 指定行 Then
            切替再計算開始行 = 指定行
            Exit For
        ElseIf 切替結果ブロック開始行 < 指定行 Then
            切替再計算開始行 = Application.WorksheetFunction.Max(切替再計算開始行, 切替結果ブロック開始行)
        End If
        
    Next 切替結果行
    
    切替影響開始行取得 = Application.WorksheetFunction.Max(切替再計算開始行, 開始行)
    
End Function

Private Function 指定行直前までの切替再計算(ByVal 開始行 As Long, ByVal 指定行 As Long, Optional ByVal 画面更新 As Boolean = False, Optional ByVal 自動再計算 As Boolean = False) As Long
    
    Dim 切替再計算開始行 As Long
    切替再計算開始行 = 切替影響開始行取得(開始行, 指定行)
    
    If 切替再計算開始行 < 指定行 Then
        
        早遅フレームと評価の解除 切替再計算開始行, 指定行 - 1, 切替の解除:=True
        Def.譜面テーブル.再計算 切替再計算開始行, 指定行 - 1, 自動再計算
        
        Dim ホールド計算開始行 As Long
        ホールド計算開始行 = 切替再計算開始行
        Do Until Def.譜面テーブル.ホールド開始フレーム列(ホールド計算開始行) <> Def.譜面テーブル.ホールド開始フレーム列(ホールド計算開始行 - 1) Or _
            Def.譜面テーブル.ホールド開始フレーム列(ホールド計算開始行) = 0
            ホールド計算開始行 = ホールド計算開始行 + 1
        Loop
        
        Dim 切替データ As Def.個別切替データ
        切替データ = 指定範囲のホールド計算(ホールド計算開始行, 指定行, 画面更新, 自動再計算, 切替結果出力:=False)
        If Not Not 切替データ.切替行リスト Then
            Dim index As Long
            For index = 1 To UBound(切替データ.切替行リスト)
                Def.譜面テーブル.切替判定列(切替データ.切替行リスト(index)) = True
            Next index
            Def.譜面テーブル.再計算 切替再計算開始行, 指定行 - 1, 自動再計算
        End If
        
        指定範囲の早遅自動判定 ホールド計算開始行, 指定行, 自動再計算
        'Set こうすればMAX文字列は得られる = 指定範囲の早遅自動判定(切替再計算開始行, 指定行, 自動再計算)
    End If
    
    指定行直前までの切替再計算 = 切替再計算開始行
    
End Function

Private Function 切替影響終了行取得(ByVal 指定行 As Long)
    
    Dim 切替結果行 As Long
    Dim 切替結果ブロック終了行 As Long
    Dim 切替再計算終了行 As Long
    切替再計算終了行 = Def.切替結果テーブル.最大ブロック終了行
    
    Dim 切替再計算開始行 As Long
    切替再計算開始行 = Def.切替結果テーブル.最大ブロック開始行
    
    For 切替結果行 = 1 To Def.切替結果テーブル.データ行数
        
        切替結果ブロック終了行 = Def.切替結果テーブル.ブロック終了列(切替結果行)
        If 切替結果ブロック終了行 = 指定行 Then
            切替再計算終了行 = 指定行
            Exit For
        ElseIf 切替結果ブロック終了行 > 指定行 Then
            切替再計算終了行 = Application.WorksheetFunction.Min(切替再計算終了行, 切替結果ブロック終了行)
            切替再計算開始行 = Application.WorksheetFunction.Min(切替再計算開始行, Def.切替結果テーブル.ブロック開始列(切替結果行))
        End If
        
    Next 切替結果行
    
    If 切替再計算開始行 > 指定行 Then
        切替影響終了行取得 = 指定行
    Else
        切替影響終了行取得 = 切替再計算終了行
    End If
    
End Function

Private Function 指定行直後からの切替再計算(ByVal 指定行 As Long, Optional ByVal 画面更新 As Boolean = False, Optional ByVal 自動再計算 As Boolean = False) As Long
    
    Dim 切替再計算終了行 As Long
    切替再計算終了行 = 切替影響終了行取得(指定行)
    
    If 切替再計算終了行 > 指定行 Then
        
        早遅フレームと評価の解除 指定行 + 1, 切替再計算終了行, 切替の解除:=True
        Def.譜面テーブル.再計算 指定行 + 1, 切替再計算終了行, 自動再計算
        
        Dim ホールド計算開始行 As Long
        ホールド計算開始行 = 指定行
        Do Until Def.譜面テーブル.ホールド開始フレーム列(ホールド計算開始行) <> Def.譜面テーブル.ホールド開始フレーム列(ホールド計算開始行 - 1) Or _
            Def.譜面テーブル.ホールド開始フレーム列(ホールド計算開始行) = 0
            ホールド計算開始行 = ホールド計算開始行 + 1
        Loop
        
        Dim 切替データ As Def.個別切替データ
        切替データ = 指定範囲のホールド計算(ホールド計算開始行, 切替再計算終了行, 画面更新, 自動再計算, 切替結果出力:=False)
        If Not Not 切替データ.切替行リスト Then
            Dim index As Long
            For index = 1 To UBound(切替データ.切替行リスト)
                Def.譜面テーブル.切替判定列(切替データ.切替行リスト(index)) = True
            Next index
            Def.譜面テーブル.再計算 指定行 + 1, 切替再計算終了行, 自動再計算
        End If
        
        指定範囲の早遅自動判定 ホールド計算開始行, 切替再計算終了行 + 1, 自動再計算, True
        
    Else
        
        切替再計算終了行 = 指定行
        
    End If
    
    指定行直後からの切替再計算 = 切替再計算終了行
    
End Function

Private Function COOL可能性取得(ByVal 行 As Long, isホールド中() As Boolean) As Boolean
    COOL可能性取得 = True
    Dim マーク As Long
    For マーク = 1 To Def.マーク数
        If Def.譜面テーブル.ノーツ列(マーク, 行) <> "" And isホールド中(マーク) Then
            COOL可能性取得 = False
            Exit For
        End If
    Next マーク
End Function

' ======================================================================================================================================================================================================
'
' ※開始行がホールド開始行と一致するはず
' ※開始行の評価だけは予め設定されている可能性がある(それ以外=早遅などは設定されていないはず)
' 戻り値は最大影響行
' ======================================================================================================================================================================================================

Public Function 指定行からのルート探索( _
    ByVal 開始行 As Long, _
    ByVal 終了行 As Long, _
    ByRef isホールド中() As Boolean, _
    ByRef 現在ルート行スコア() As Long, _
    Optional ByVal 開始早ずれ許容フレーム As Long = -Def.未MAXフレーム最大値, _
    Optional ByVal 開始遅ずれ許容フレーム As Long = Def.未MAXフレーム最大値, _
    Optional ByVal 完奏モード As Boolean = False, _
    Optional ByVal 画面更新 As Boolean = False, _
    Optional ByVal 自動再計算 As Boolean = False, _
    Optional ByVal 影響開始行 As Long = 0) ', _
    Optional ByVal 事前w数 As Long = 0)
    
    #If スコアタ詳細ログ Then
        詳細ログ出力 "指定行からのルート探索" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "開始", True
    #End If
    
    #If スコアタ通常ログ Then
        スコアタネストカウンタ = スコアタネストカウンタ + 1
        詳細ログ出力 String(スコアタネストカウンタ - 1, "│") & "├" & 開始行 & "～", True, スコアタネストカウンタ + 1
    #End If
    
    Dim フレーム差 As Long
    
    Dim isホールド開始(Def.マーク数) As Boolean
    Dim ルート開始可能性 As Boolean
    
    Dim cool可能性 As Boolean
    Dim wrong可能性 As Boolean
    Dim w数 As Long
    w数 = 0 '事前w数
    
    Dim ルート別評価() As Def.評価セット
    Dim 暫定評価 As String
    
    Dim 行 As Long
    Dim マーク As Long
    Dim 評価index As Long
    
    For 行 = 開始行 + 1 To 終了行
        
        早遅フレームと評価の解除 行
        Def.譜面テーブル.再計算 行, , 自動再計算
        フレーム差 = Def.譜面テーブル.フレーム列(行) - Def.譜面テーブル.フレーム列(開始行)
        
        ルート開始可能性 = False
        'どう頑張ってもホールドを接続できない行になれば終了
        If フレーム差 > Def.未MAXフレーム最大値 + Def.譜面テーブル.最大ホールド接続フレームずれ + 1 _
            - Application.WorksheetFunction.Max(Def.譜面テーブル.デフォルト遅ずれ許容フレーム - 開始遅ずれ許容フレーム, 0) Then
            Exit For
        End If
        
        '現在行の早遅フレームと評価をリセット
        早遅フレームと評価の解除 行
        Def.譜面テーブル.再計算 行, , 自動再計算
        
        'COOLになれる=ボタンを押せる可能性を判定
        cool可能性 = COOL可能性取得(行, isホールド中)
        
        If cool可能性 Then
            
            ' 新たにルート開始位置となることが可能かどうかを判定
            ルート開始可能性 = False
            For マーク = 1 To Def.マーク数
                If Def.譜面テーブル.ノーツ列(マーク, 行) = Def.HOLD文字 Then
                    isホールド開始(マーク) = True
                    ルート開始可能性 = True
                ElseIf isホールド中(マーク) Then
                    isホールド開始(マーク) = True
                Else
                    isホールド開始(マーク) = False
                End If
            Next マーク
            
            'ルート開始位置になれる場合は、ホールドを開始して次の行から新たにルート探索
            If ルート開始可能性 Then
                
                Erase ルート別評価
                暫定評価 = Def.譜面テーブル.評価列(開始行)
                
                ' 直前のホールド開始行と現在行の評価とフレームずれを設定
                If フレーム差 - Application.WorksheetFunction.Min(開始遅ずれ許容フレーム - Def.譜面テーブル.最早COOLフレーム, 0) > Def.未MAXフレーム最大値 Then
                    '開始行を最早COOL(遅ずれ許容フレームが最早COOL以前の場合はそのフレーム)にした場合に、現在行の早遅フレームが最早COOLで接続できない場合
                    ルート別評価 = Def.譜面テーブル.評価リスト取得(Def.未MAXフレーム最大値 - フレーム差, 開始早ずれ許容フレーム, 開始遅ずれ許容フレーム)
                Else
                    '開始行を最早COOL(遅ずれ許容フレームが最早COOL以前の場合はそのフレーム)にした場合に、現在行の早遅フレームが最早COOLで接続できる場合
                    ReDim ルート別評価(1)
                    ルート別評価(1).開始評価 = Def.譜面テーブル.評価列(行)
                    ルート別評価(1).開始フレームずれ = Application.WorksheetFunction.Max(開始早ずれ許容フレーム, _
                        Application.WorksheetFunction.Min(開始遅ずれ許容フレーム, Def.譜面テーブル.最早COOLフレーム))
                End If
                
                Dim 現在早ずれ許容フレーム As Long
                Dim 現在遅ずれ許容フレーム As Long
                
                ' 評価ごとに開始行(と現在行)を設定して、新たにルート探索
                If Not Not ルート別評価 Then
                    
                    For 評価index = 1 To UBound(ルート別評価)
                    
                        Def.譜面テーブル.評価列(開始行) = ルート別評価(評価index).開始評価
                        Def.譜面テーブル.早遅フレーム手動指定列(開始行) = ルート別評価(評価index).開始フレームずれ
                        
                        If ルート別評価(評価index).終了評価 = "" Then
                            Def.譜面テーブル.評価列(行) = Def.譜面テーブル.フレームずれ別評価(0)
                            現在早ずれ許容フレーム = Def.譜面テーブル.デフォルト早ずれ許容フレーム
                            現在遅ずれ許容フレーム = Application.WorksheetFunction.Min(Def.譜面テーブル.デフォルト遅ずれ許容フレーム, _
                            Def.未MAXフレーム最大値 + Def.譜面テーブル.ホールド開始フレーム列(開始行) - Def.譜面テーブル.フレーム列(行) + 1)
                        Else
                            Def.譜面テーブル.評価列(行) = ルート別評価(評価index).終了評価
                            現在早ずれ許容フレーム = ルート別評価(評価index).終了フレームずれ
                            現在遅ずれ許容フレーム = ルート別評価(評価index).終了フレームずれ
                        End If
                        
                        Dim 現在行切替 As String
                        現在行切替 = Def.譜面テーブル.切替判定文字列(行)
                        If 現在行切替 <> "" Then
                            Def.譜面テーブル.切替判定列(行) = False
                        End If
                        Def.譜面テーブル.再計算 開始行, 行, 自動再計算
                        
                        ' 新たにルート探索
                        指定行からのルート探索 行, 終了行, isホールド開始, 現在ルート行スコア, 現在早ずれ許容フレーム, 現在遅ずれ許容フレーム, 完奏モード, 画面更新, 自動再計算, 影響開始行 ', w数
                        
                        If 現在行切替 <> "" Then
                            Def.譜面テーブル.切替判定文字列(行) = 現在行切替
                        End If
                        
                    Next 評価index
                    
                End If
                
                ' 直前のホールド開始行のリセット
                早遅フレームと評価の解除 開始行
                Def.譜面テーブル.評価列(開始行) = 暫定評価
                早遅フレームと評価の解除 行
                Def.譜面テーブル.再計算 開始行, 行, 自動再計算
                
                ' 以後、このホールドを潰す可能性を検討させる
                cool可能性 = False
                
            End If
            
        End If
        
        ' COOLにならない可能性がある場合(上記のCOOLになれる可能性のあるホールドを潰す可能性も含める)
        If Not cool可能性 Then
            
            ' WRONGで潰すかWORSTかを判定させる
            wrong可能性 = False
            For マーク = 1 To Def.マーク数
                If Def.譜面テーブル.ノーツ列(マーク, 行) = "" And (Not isホールド中(マーク)) Then
                    wrong可能性 = True
                    Exit For
                End If
            Next マーク
            
            ' 赤WRONGまたはWORSTを代入・再計算
            If wrong可能性 Then
                Def.譜面テーブル.評価列(行) = Def.赤WRONG文字
            Else
                Def.譜面テーブル.評価列(行) = Def.WORST文字
            End If
            Def.譜面テーブル.再計算 行, , 自動再計算
            
            w数 = w数 + 1
            
        End If
        
    Next
    
    If 行 > 終了行 Then
        行 = 終了行
    End If
    
    ' 現在の行では既にMAXが入っている(はずの)ため、MAXが入る直後の行まで遡る
    Do Until Def.譜面テーブル.ホールド開始フレーム列(行 - 1) = Def.譜面テーブル.ホールド開始フレーム列(開始行)
        If Def.譜面テーブル.評価列(行) <> Def.譜面テーブル.フレームずれ別評価(0) Then
            Def.譜面テーブル.評価列(行) = Def.譜面テーブル.フレームずれ別評価(0)
            w数 = w数 - 1
        End If
        Def.早遅切替一覧テーブル.早遅切替情報書き出し Def.譜面テーブル, 行, 行, False
        行 = 行 - 1
    Loop
    
    Dim 開始行早遅フレーム As Long
    ' 開始行の早遅フレームを接続がある場合は最遅COOL、ない場合は最早COOLに設定
    If isMAXタイミング不問(開始行) Then
        開始行早遅フレーム = Application.WorksheetFunction.Max(開始早ずれ許容フレーム, _
            Application.WorksheetFunction.Min(開始遅ずれ許容フレーム, Def.譜面テーブル.最早COOLフレーム))
    Else
        開始行早遅フレーム = Application.WorksheetFunction.Max(開始早ずれ許容フレーム, _
            Application.WorksheetFunction.Min(開始遅ずれ許容フレーム, Def.譜面テーブル.最遅COOLフレーム))
    End If
    Def.譜面テーブル.早遅フレーム手動指定列(開始行) = 開始行早遅フレーム
    Def.譜面テーブル.再計算 開始行, 行, 自動再計算
    
'    Dim 最大影響行 As Long
'    最大影響行 = 行
'    指定行からのルート探索 = 行
    
    ' ※以後再計算見直し
    
    Dim 切替再計算終了行 As Long
'    Dim 対象ルート行スコア As Def.行スコア情報
    
    ' 捨て量が最大の時の、その行以降の行スコアを一度検証
    暫定評価 = Def.譜面テーブル.評価列(行)
    If 暫定評価 <> Def.譜面テーブル.フレームずれ別評価(0) Then
        Def.譜面テーブル.評価列(行) = Def.譜面テーブル.フレームずれ別評価(0)
        w数 = w数 - 1
        If Def.譜面テーブル.ホールドフレーム列(行) + Def.譜面テーブル.最遅COOLフレーム <= Def.未MAXフレーム最大値 Then
            ' COOLにした場合に最遅でもMAXが入らなくなる場合(ある? 赤WRONGで301F,COOLで押し直すと300Fになり、かつ最遅COOLが0Fの場合はあり得る?)
            ' このときは赤WRONG(またはWORST?)のままルートを確定させる
            Def.譜面テーブル.評価列(行) = 暫定評価
            w数 = w数 + 1
        End If
    End If
    
    If w数 > 0 Then
'        対象ルート行スコア = ルート確定後処理(開始行, 行, 終了行, 現在ルート行スコア, 完奏モード, 画面更新, 自動再計算, 影響開始行)
        ルート確定後処理 開始行, 行, 終了行, 現在ルート行スコア, 完奏モード, 画面更新, 自動再計算, 影響開始行
    End If
    
    ' ここから行を遡って、一行ずつ赤WRONGやWORSTを外した時の行スコアを検証
    ' ※スタート行の評価から赤WRONGやWORSTの可能性あり
    暫定評価 = Def.譜面テーブル.評価列(開始行)
    
    For 行 = 行 To 開始行 + 1 Step -1
        
        If Def.譜面テーブル.評価列(行) <> Def.譜面テーブル.フレームずれ別評価(0) Then
            
            Def.譜面テーブル.評価列(行) = Def.譜面テーブル.フレームずれ別評価(0)
            
            Erase ルート別評価
            ルート別評価 = Def.譜面テーブル.評価リスト取得( _
                Def.未MAXフレーム最大値 + 3 + Def.譜面テーブル.フレーム列(開始行) - Def.譜面テーブル.フレーム列(行), _
                開始早ずれ許容フレーム, 開始遅ずれ許容フレーム)
            
            ' MAXを入れることが出来る場合、その状態での行スコアを検証
            If Not Not ルート別評価 Then
                For 評価index = 1 To UBound(ルート別評価)
                    Def.譜面テーブル.評価列(開始行) = ルート別評価(評価index).開始評価
                    Def.譜面テーブル.早遅フレーム手動指定列(開始行) = ルート別評価(評価index).開始フレームずれ
                    Def.譜面テーブル.評価列(行) = ルート別評価(評価index).終了評価
                    Def.譜面テーブル.早遅フレーム手動指定列(行) = ルート別評価(評価index).終了フレームずれ
                    
'                    ルート確定後処理 開始行, 行, 終了行, 対象ルート行スコア, 完奏モード, 画面更新, 自動再計算, 影響開始行
                    ルート確定後処理 開始行, 行, 終了行, 現在ルート行スコア, 完奏モード, 画面更新, 自動再計算, 影響開始行
                Next
            End If
            
            ' MAXを入れない場合の行スコアを検証
            Def.譜面テーブル.評価列(開始行) = 暫定評価
            Def.譜面テーブル.早遅フレーム手動指定列(開始行) = 開始行早遅フレーム
            Def.譜面テーブル.評価列(行) = Def.譜面テーブル.フレームずれ別評価(0)
            Def.譜面テーブル.早遅フレーム手動指定列(行) = Def.譜面テーブル.最遅COOLフレーム
            
            w数 = w数 - 1
            If w数 > 0 Then
'                ルート確定後処理 開始行, 行, 終了行, 対象ルート行スコア, 完奏モード, 画面更新, 自動再計算, 影響開始行
                ルート確定後処理 開始行, 行, 終了行, 現在ルート行スコア, 完奏モード, 画面更新, 自動再計算, 影響開始行
            End If
            
        End If
        
        Def.早遅切替一覧テーブル.早遅切替情報書き出し Def.譜面テーブル, 行, 行, False
        
    Next 行
    
    '切替と早遅を元に戻す ⇒ 早遅切替は1行毎に戻す、再計算はいらない?
'    Def.早遅切替一覧テーブル.早遅切替情報書き出し Def.譜面テーブル, 開始行 + 1, 切替再計算終了行, False
'    Def.譜面テーブル.再計算 開始行 + 1, 切替再計算終了行, 自動再計算
    
    #If スコアタ詳細ログ Then
        詳細ログ出力 "指定行からのルート探索" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "終了", True
    #End If
    
    #If スコアタ通常ログ Then
        Def.処理ログ.フォーム文字列削除 スコアタネストカウンタ + 1
        スコアタネストカウンタ = スコアタネストカウンタ - 1
    #End If
    
End Function

Private Function ルート確定後処理( _
    ByVal 開始行 As Long, _
    ByVal 現在行 As Long, _
    ByVal 終了行 As Long, _
    ByRef 現在ルート行スコア() As Long, _
    Optional ByVal 完奏モード As Boolean = False, _
    Optional ByVal 画面更新 As Boolean = False, _
    Optional ByVal 自動再計算 As Boolean = False, _
    Optional ByVal 影響開始行 As Long = 0) _
    As Def.行スコア情報
'    Optional ByRef 比較ルート行スコア As Def.行スコア情報, _

    Dim 切替再計算終了行 As Long
    
    Def.譜面テーブル.再計算 開始行, 現在行, 自動再計算
    切替再計算終了行 = 指定行直後からの切替再計算(現在行, 画面更新, 自動再計算)
    DoEvents
    ルート確定後処理 = 指定範囲のルート判定(現在行, 終了行, 現在ルート行スコア, 完奏モード, 画面更新, 自動再計算, 影響開始行)
    Def.早遅切替一覧テーブル.早遅切替情報書き出し Def.譜面テーブル, 現在行 + 1, 切替再計算終了行, False
'    Def.譜面テーブル.再計算 現在行 + 1, 切替再計算終了行, 自動再計算
    
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Function 自動切替判定( _
    Optional ByVal 画面更新 As Boolean = False, _
    Optional ByVal 自動再計算 As Boolean = False, _
    Optional ByVal 判定開始行 As Long = 1, _
    Optional ByVal 判定終了行 As Long = -1)
    
    #If 切替詳細ログ Then
        詳細ログ出力 "自動切替判定" & vbTab & "開始"
    #End If
    
    Application.ScreenUpdating = False
    
    If 自動再計算 Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    
    If 判定終了行 = -1 Then
        判定終了行 = Def.譜面テーブル.データ行数
    End If
    
    ' 画面更新ON時のフィルタ更新用初期設定 ---------------------------------------------------------
    
    If 画面更新 Then
        
        Dim block As Long
        Dim hBlock As Long
        Dim hBlockColor As Long
        
        block = Def.譜面テーブル.OwnTable.ListColumns("block").index
        hBlock = Def.譜面テーブル.OwnTable.ListColumns("HBlock").index
        hBlockColor = RGB(146, 208, 80)
        
        Def.譜面テーブル.OwnTable.Range.AutoFilter hBlock, hBlockColor, Operator:=xlFilterCellColor
        
    End If
    
    ' 解析ブロック分割 -----------------------------------------------------------------------------
    
    Dim 行 As Long
    Dim 開始行 As Long
    Dim 終了行 As Long
    
    Dim 現在切替検討数 As Long
    Dim 結果行 As Long
    現在切替検討数 = Def.切替結果テーブル.データ行数
    現在ホールドブロック = 0
    
    For 行 = 判定開始行 To 判定終了行
        
        If Def.譜面テーブル.ホールドブロック列(行) <> 現在ホールドブロック Then
            
            If 現在ホールドブロック > 0 Then
                
                終了行 = 行 - 1
                
                If 画面更新 Then
                    Def.譜面テーブル.OwnTable.Range.AutoFilter block, 現在ホールドブロック
                    Def.譜面テーブル.OwnTable.ListRows(行).Range.Rows.Hidden = False
                End If
                
                Def.処理ログ.出力 "【" & 現在ホールドブロック & "ブロック目解析スタート】(" & 開始行 & "～" & 終了行 & "行目)"
                
                Application.ScreenUpdating = 画面更新
                DoEvents
                Application.ScreenUpdating = False
                
                現在切替検討数 = Def.切替結果テーブル.データ行数
                
                指定範囲のホールド計算 開始行, 終了行, 画面更新, 自動再計算
                
                ' ブロック結果出力
                For 結果行 = 現在切替検討数 + 1 To Def.切替結果テーブル.データ行数
                    Def.切替結果テーブル.ブロック開始列(結果行) = 開始行
                    Def.切替結果テーブル.ブロック終了列(結果行) = 終了行
                Next 結果行
                
                'Application.ScreenUpdating = False
                
                If 画面更新 Then
                    Def.譜面テーブル.OwnTable.Range.AutoFilter block
                End If
                
            End If
            
            現在ホールドブロック = Def.譜面テーブル.ホールドブロック列(行)
            開始行 = 行
            
        End If
        
    Next 行
    
    If 画面更新 Then
        Def.譜面テーブル.OwnTable.Range.AutoFilter hBlock
    End If
    
    Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True
    
    #If 切替詳細ログ Then
        詳細ログ出力 "自動切替判定" & vbTab & "終了"
    #End If
    
End Function

' ======================================================================================================================================================================================================
'
' 切替結果出力 が False の時のみ最大スコアになる切替データを格納したオブジェクトが返されます。
' ※ (開始行 - 1) 行目も評価対象になります。→ 開始行 = 1 はダメ
' ※ 早遅判定のために (終了行 + 1) 行目まで評価対象になります。→ 終了行 = テーブル最終行 はダメ
' ※ また早遅判定の影響によって、評価対象の行の末尾がさらに追加される可能性もあります。
' ======================================================================================================================================================================================================

Public Function 指定範囲のホールド計算( _
    ByVal 開始行 As Long, _
    ByRef 終了行 As Long, _
    Optional ByVal 画面更新 As Boolean = False, _
    Optional ByVal 自動再計算 As Boolean = False, _
    Optional ByVal ブロック開始行 As Long = 0, _
    Optional ByRef MAX文字列 As OutputString = Nothing, _
    Optional ByVal 切替結果出力 As Boolean = True) _
    As Def.個別切替データ
    
    ' 初期設定 -------------------------------------------------------------------------------------
    #If 切替詳細ログ Then
        詳細ログ出力 "指定範囲のホールド計算" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "初期設定開始"
    #End If
    
    If ブロック開始行 = 0 Then
        ブロック開始行 = 開始行
    End If
    
    If MAX文字列 Is Nothing Then
        Set MAX文字列 = New OutputString
        MAX文字列.表示文字列 = ""
        MAX文字列.ログ出力用文字列 = ""
    End If
    
    Dim ホールド開始フレーム(マーク数) As Double
    Dim ホールド終了行(マーク数) As Long
    
    Dim マーク As Long
    For マーク = 1 To マーク数
        ホールド開始フレーム(マーク) = 0
    Next
    
    Dim 切替可能性 As Boolean
    
    If Not 切替結果出力 Then
        Dim 最大スコア切替データ As Def.個別切替データ
        最大スコア切替データ.スコア = 0
        Dim 現在切替データ As Def.個別切替データ
    End If
    
    ' 自動切替判定 ---------------------------------------------------------------------------------
    #If 切替詳細ログ Then
        詳細ログ出力 "指定範囲のホールド計算" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "切替判定開始"
    #End If
    
    Dim 行 As Long
    'Dim マーク As Long
    Dim フレーム As Double
    
    Dim 現在ホールド終了行 As Long
    Dim 早遅判定終了行 As Long
    Dim 現在MAX文字列 As OutputString
    
'    Dim 早遅判定開始行 As Long
'    早遅判定開始行 = 開始行
    
    For 行 = 開始行 To 終了行
        
        切替可能性 = False
        
        ' 1. マークごとにホールド開始フレームとホールド終了行をその行のものに設定
        
        For マーク = 1 To マーク数
            
            フレーム = Def.譜面テーブル.ホールド可能性判定列(マーク, 行)
            
            If フレーム > ホールド開始フレーム(マーク) Then
                
                ホールド開始フレーム(マーク) = フレーム
                
                現在ホールド終了行 = 行
                
                Do While ホールド開始フレーム(マーク) = Def.譜面テーブル.ホールド可能性判定列(マーク, 現在ホールド終了行)
                    現在ホールド終了行 = 現在ホールド終了行 + 1
                Loop
                
                ' ホールド終了行は(適切な終了行設定でも) 終了行 + 1 になる可能性がある
                ホールド終了行(マーク) = 現在ホールド終了行
                
                切替可能性 = True
                                
            ElseIf フレーム = 0 And ホールド開始フレーム(マーク) > 0 Then
                
                ホールド開始フレーム(マーク) = 0
                ホールド終了行(マーク) = 0
                
            End If
            
        Next マーク
        
        ' 2. 切替可能性があり、切替が手動指定されていない行である場合、切替を指定し、同じ解析ブロック内の続きを再帰で探索
        
        If 切替可能性 And Def.譜面テーブル.切替判定文字列(行) = "" Then
            
            For マーク = 1 To マーク数
            
                If ホールド終了行(マーク) < 現在ホールド終了行 _
                    And ホールド開始フレーム(マーク) = Def.譜面テーブル.ホールド個別開始フレーム列(マーク, 行) _
                    And ホールド開始フレーム(マーク) > 0 Then
                    
                    Def.譜面テーブル.切替判定列(行) = True
                    
                    'Def.譜面テーブル.再計算 行, 終了行 + 1, 自動再計算
                    'DoEvents
                    
                    早遅判定終了行 = 行
                    
                    Set 現在MAX文字列 = 指定範囲の早遅自動判定(開始行, 早遅判定終了行, 自動再計算)
                    
                    If 早遅判定終了行 <> 行 Then
                        MsgBox "予想外のことが発生しました。なぜですか。教えて下さい。でもまあ処理は続けますよ。"
                    End If
                    
                    Def.譜面テーブル.再計算 行, 終了行 + 1, 自動再計算
                    DoEvents
                    
                    現在MAX文字列.表示文字列 = Def.文字列連結(MAX文字列.表示文字列, 現在MAX文字列.表示文字列, vbCrLf)
                    現在MAX文字列.ログ出力用文字列 = Def.文字列連結(MAX文字列.ログ出力用文字列, 現在MAX文字列.ログ出力用文字列, ", ")
                    現在切替データ = 指定範囲のホールド計算(行, 終了行, 画面更新, 自動再計算, ブロック開始行, 現在MAX文字列, 切替結果出力)
                    If Not 切替結果出力 Then
                        If 現在切替データ.スコア > 最大スコア切替データ.スコア Then
                            最大スコア切替データ.スコア = 現在切替データ.スコア
                            最大スコア切替データ.切替行リスト = 現在切替データ.切替行リスト
                        End If
                    End If
                    ' 画面再ロック
                    Application.ScreenUpdating = False
                    
                    指定範囲の早遅指定削除 開始行, 早遅判定終了行, 自動再計算
                    
                    'Def.譜面テーブル.再計算 行, 終了行 + 1, 自動再計算
                    
                    Def.譜面テーブル.切替判定列(行) = False
                    
                    Def.譜面テーブル.再計算 行, 終了行 + 1, 自動再計算
                    'DoEvents
                    
                    Exit For
                    
                End If
                
            Next
            
        End If
        
    Next 行
    
    ' 現ルートによる結果の表示 ---------------------------------------------------------------------
    #If 切替詳細ログ Then
        詳細ログ出力 "指定範囲のホールド計算" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "結果表示開始"
    #End If
    
    'Dim 現在MAX文字列 As String
    
    早遅判定終了行 = 終了行 + 1
    Set 現在MAX文字列 = 指定範囲の早遅自動判定(開始行, 早遅判定終了行, 自動再計算)
    DoEvents
    終了行 = 早遅判定終了行 - 1
    
    現在MAX文字列.表示文字列 = Def.文字列連結(MAX文字列.表示文字列, 現在MAX文字列.表示文字列, vbCrLf)
    現在MAX文字列.ログ出力用文字列 = Def.文字列連結(MAX文字列.ログ出力用文字列, 現在MAX文字列.ログ出力用文字列, ", ")
    
    ' 切替結果出力
    If 切替結果出力 Then
        切替結果テーブル.出力行追加
    Else
'        現在切替データ.スコア = 0
        Erase 現在切替データ.切替行リスト
    End If
    
    Dim ログ出力用切替文字列 As String
    Dim 切替文字列 As String
    Dim 切替ノーツ番号 As Long
    Dim 切替コンボ数 As Long
    ログ出力用切替文字列 = ""
    切替文字列 = ""
    
    Dim 切替数 As Long
    切替数 = 0
    
    For 行 = ブロック開始行 To 終了行
        
        If Def.譜面テーブル.切替判定列(行) Then
            
            切替数 = 切替数 + 1
            
            切替文字列 = Def.文字列連結(切替文字列, 切替文字列取得(行), vbCrLf)
            
            If 切替結果出力 Then
                Def.切替結果テーブル.切替行入力 行, 切替数
            Else
                ReDim Preserve 現在切替データ.切替行リスト(切替数)
                現在切替データ.切替行リスト(切替数) = 行
            End If
            
            ログ出力用切替文字列 = Def.文字列連結(ログ出力用切替文字列, 行 & "行目", ", ")
            
        End If
        
    Next 行
        
    Dim スコア As Long
    スコア = Def.譜面テーブル.ホールドスコア列(早遅判定終了行) - Def.譜面テーブル.ホールドスコア列(ブロック開始行)
    
    If 切替結果出力 Then
        Def.切替結果テーブル.解析ブロック列(Def.切替結果テーブル.データ行数) = 現在ホールドブロック
        Def.切替結果テーブル.ブロックスコア列(Def.切替結果テーブル.データ行数) = スコア
        Def.切替結果テーブル.切替文字列(Def.切替結果テーブル.データ行数) = 切替文字列
        Def.切替結果テーブル.MAX可能性文字列(Def.切替結果テーブル.データ行数) = 現在MAX文字列.表示文字列
    Else
'        現在切替データ.スコア = スコア
        If スコア > 最大スコア切替データ.スコア Then
            最大スコア切替データ.スコア = スコア
            最大スコア切替データ.切替行リスト = 現在切替データ.切替行リスト
        End If
        指定範囲のホールド計算 = 最大スコア切替データ
    End If
    
    If 切替結果出力 Then
        
        If ログ出力用切替文字列 = "" Then
            ログ出力用切替文字列 = "(なし)"
        End If
        
        Dim ログ出力文字列 As String
        ログ出力文字列 = 現在ホールドブロック & "ブロック目 / スコア: " & スコア & " / 切替: " & ログ出力用切替文字列
        
        If 現在MAX文字列.ログ出力用文字列 <> "" Then
            ログ出力文字列 = ログ出力文字列 & " / MAX可能性: " & 現在MAX文字列.ログ出力用文字列
        End If
        Def.処理ログ.出力 ログ出力文字列
        
    End If
    
    ' シートに結果表示
    Application.ScreenUpdating = 画面更新
    DoEvents
    Application.ScreenUpdating = False
    
    指定範囲の早遅指定削除 開始行, 早遅判定終了行, 自動再計算
    
    Def.譜面テーブル.再計算 開始行, 早遅判定終了行, 自動再計算
    
    'DoEvents
    
    #If 切替詳細ログ Then
        詳細ログ出力 "指定範囲のホールド計算" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "終了"
    #End If
    
End Function

' ======================================================================================================================================================================================================
'
' MAX可能性フラグ(開始行 to 終了行)も設定されます。
' 戻り値はMAX可能性のあるノーツの情報を表す文字列
' ※ (開始行 - 1) 行目も評価対象になります。→ 開始行 = 1 はダメ
' ※ 開始行がホールド中の行からもダメ(ホールド中でなければホールドブロック途中の行からでもOK??)
' また、終了行付近でMAXが絡むことで、MAX終了行まで評価対象の末尾の行が追加される可能性もあります。
' 終了行が変更された場合、引数の参照渡しによって元の変数も変更されます。
' ======================================================================================================================================================================================================

Public Function 指定範囲の早遅自動判定( _
    ByVal 開始行 As Long, _
    ByRef 終了行 As Long, _
    Optional ByVal 自動再計算 As Boolean = False, _
    Optional ByVal 直前MAX確認 As Boolean = False) _
    As OutputString
    
    ' 初期設定 -------------------------------------------------------------------------------------
    #If 切替詳細ログ Then
        詳細ログ出力 "指定範囲の早遅自動判定" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "初期設定開始"
    #End If
    
    Def.譜面テーブル.再計算 開始行, 終了行, 自動再計算
    
    ' MAXの可能性が存在するホールド箇所を検索 ------------------------------------------------------
    #If 切替詳細ログ Then
        詳細ログ出力 "指定範囲の早遅自動判定" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "MAX検索開始"
    #End If
    
    
    Dim 行 As Long
    
    Dim 現在ホールド開始行 As Long
    Dim 現在ホールド開始フレームずれ量 As Double
    
    'Dim 行 As Long
    
    For 行 = 開始行 To 終了行
        
        If Def.譜面テーブル.ホールドボーナス列(行) > 0 And 行 > 開始行 Then
            
            If Def.譜面テーブル.ホールドフレーム列(行) - (Def.譜面テーブル.早遅フレーム列(行) - 現在ホールド開始フレームずれ量) > _
                Def.未MAXフレーム最大値 - (Def.譜面テーブル.最遅COOLフレーム - Def.譜面テーブル.最早COOLフレーム) Then
                MAX可能性フラグ(現在ホールド開始行) = True
            End If
            
            現在ホールド開始行 = 0
            現在ホールド開始フレームずれ量 = 0
            
        End If
        
        If Def.譜面テーブル.ホールド開始フレーム列(行) > Def.譜面テーブル.ホールド開始フレーム列(行 - 1) Then
            
            現在ホールド開始行 = 行
            現在ホールド開始フレームずれ量 = Def.譜面テーブル.早遅フレーム列(行)
            
        End If
    
    Next
    
    ' 上から順にブロックごとに早遅自動判定(MAXが絡む場合のみ遡って変更の可能性あり) ------------------------------
    #If 切替詳細ログ Then
        詳細ログ出力 "指定範囲の早遅自動判定" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "早遅設定開始"
    #End If
    
    Dim マーク As Long
    
    'Dim 現在ホールド開始行 As Long
    Dim 現在ホールドボタン数 As Long
    Dim 現在MAX可能性 As Boolean
    
    現在ホールド開始行 = 0
    現在ホールドボタン数 = 0
    For マーク = 1 To マーク数
        If Def.譜面テーブル.ホールド個別開始フレーム列(マーク, 開始行 - 1) > 0 Then
            現在ホールドボタン数 = 現在ホールドボタン数 + 1
        End If
    Next マーク
    
    If 直前MAX確認 Then
        現在MAX可能性 = Def.譜面テーブル.ホールドフレーム列(開始行) > Def.未MAXフレーム最大値
        If 現在MAX可能性 Then
            For 現在ホールド開始行 = 開始行 - 1 To 2 Step -1
                If Def.譜面テーブル.ホールド開始フレーム列(現在ホールド開始行) > Def.譜面テーブル.ホールド開始フレーム列(現在ホールド開始行 - 1) Then
                    Exit For
                End If
            Next
            If 現在ホールド開始行 = 1 Then
                現在ホールド開始行 = 0
            End If
        End If
    Else
        現在MAX可能性 = False
    End If
    
    Dim 仮ホールドブロック As Long
    Dim ブロック開始行 As Long
    
    仮ホールドブロック = Def.譜面テーブル.ホールドブロック列(開始行)
    ブロック開始行 = 開始行
    
    For 行 = 開始行 To 終了行 - 1
        If Def.譜面テーブル.ホールドブロック列(行 + 1) <> 仮ホールドブロック Then
            ブロック内早遅自動判定 ブロック開始行, 行, 現在ホールド開始行, 現在ホールドボタン数, 現在MAX可能性, 自動再計算
            DoEvents
            仮ホールドブロック = Def.譜面テーブル.ホールドブロック列(行 + 1)
            ブロック開始行 = 行 + 1
        End If
    Next 行
    
    ' (ラストまで早遅計算)
    ブロック内早遅自動判定 ブロック開始行, 終了行, 現在ホールド開始行, 現在ホールドボタン数, 現在MAX可能性, 自動再計算
    
    ' MAX可能性のある箇所のデバッグ出力 ------------------------------------------------------------
    #If 切替詳細ログ Then
        詳細ログ出力 "指定範囲の早遅自動判定" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "MAX出力開始"
    #End If
    
    Dim ログ出力用MAX文字列 As String
    Dim MAX文字列 As String
    'Dim 現在MAX可能性 As Boolean
    
    ログ出力用MAX文字列 = ""
    MAX文字列 = ""
    現在MAX可能性 = False
    
    'Dim 行 As Long
    'Dim マーク As Long
    Dim ボタン数 As Long
    
    For 行 = 開始行 To 終了行
        
        If 現在MAX可能性 And Def.譜面テーブル.ホールドボーナス列(行) > 0 Then
            
            ボタン数 = 0
            For マーク = 1 To マーク数
                If Def.譜面テーブル.ホールド個別開始フレーム列(マーク, 行 - 1) > 0 Then
                    ボタン数 = ボタン数 + 1
                End If
            Next マーク
            
            MAX文字列 = MAX文字列 & "【"
            If ボタン数 > 1 Then
                MAX文字列 = MAX文字列 & ボタン数
            End If
            MAX文字列 = MAX文字列 & "MAX】"
            
            ログ出力用MAX文字列 = ログ出力用MAX文字列 & 行 & 早遅文字またはフレームの取得(行) & "行目 ("
            
            If Def.譜面テーブル.ホールド終了押し直し判定列(行) Then
                
                ログ出力用MAX文字列 = ログ出力用MAX文字列 & Def.譜面テーブル.ホールドフレーム列(行)
                
                MAX文字列 = MAX文字列 & "→ " & ノーツ番号文字列の取得(行)
                
                For マーク = 1 To マーク数
                    If Def.譜面テーブル.ノーツ列(マーク, 行) <> "" Then
                        MAX文字列 = MAX文字列 & マーク文字(マーク)
                    End If
                Next マーク
            
                MAX文字列 = MAX文字列 & " (" & Def.譜面テーブル.ホールドフレーム列(行)
                
            Else
                
                ログ出力用MAX文字列 = ログ出力用MAX文字列 & ">" & Def.未MAXフレーム最大値
                MAX文字列 = MAX文字列 & "(>" & Def.未MAXフレーム最大値
                
            End If
            
            ログ出力用MAX文字列 = ログ出力用MAX文字列 & "F)"
            MAX文字列 = MAX文字列 & "F)"
            
            現在MAX可能性 = False
            
        End If
        
        If MAX可能性フラグ(行) Then
        
            現在MAX可能性 = True
            
            If ログ出力用MAX文字列 <> "" Then
                ログ出力用MAX文字列 = ログ出力用MAX文字列 & ", "
            End If
            
            ログ出力用MAX文字列 = ログ出力用MAX文字列 & 行 & 早遅文字またはフレームの取得(行) & "→"
            
            If MAX文字列 <> "" Then
                MAX文字列 = MAX文字列 & vbCrLf
            End If
            
            MAX文字列 = MAX文字列 & ノーツ番号文字列の取得(行)
            
            For マーク = 1 To マーク数
                If Def.譜面テーブル.ホールド個別開始フレーム列(マーク, 行) > Def.譜面テーブル.ホールド個別開始フレーム列(マーク, 行 - 1) Then
                    MAX文字列 = MAX文字列 & マーク文字(マーク)
                End If
            Next マーク
            
            MAX文字列 = MAX文字列 & 早遅文字またはフレームの取得(行) & " →"
                        
        End If
        
    Next 行
    
    Set 指定範囲の早遅自動判定 = New OutputString
    指定範囲の早遅自動判定.表示文字列 = MAX文字列
    指定範囲の早遅自動判定.ログ出力用文字列 = ログ出力用MAX文字列
    
    Def.譜面テーブル.再計算 開始行, 終了行, 自動再計算
    DoEvents
    
    ' 終了処理 -------------------------------------------------------------------------------------
    #If 切替詳細ログ Then
        詳細ログ出力 "指定範囲の早遅自動判定" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "終了"
    #End If
    
'    For 行 = 開始行 To 終了行
'        MAX可能性フラグ(行) = False
'    Next 行
    
End Function

' ======================================================================================================================================================================================================
'
' MAX可能性フラグ(開始行 to 終了行)も設定されます。
' 戻り値はMAX可能性のあるノーツの情報を表す文字列
' ※ (開始行 - 1) 行目も評価対象になります。→ 開始行 = 1 はダメ
' ※ 開始行がホールド中の行からもダメ
' また、終了行付近でMAXが絡むことで、MAX終了行まで評価対象の末尾の行が追加される可能性もあります。
' 終了行が変更された場合、引数の参照渡しによって元の変数も変更されます。
' ======================================================================================================================================================================================================

Private Function ブロック内早遅自動判定( _
    ByVal 開始行 As Long, _
    ByRef 終了行 As Long, _
    ByRef 現在ホールド開始行 As Long, _
    ByRef 現在ホールドボタン数 As Long, _
    ByRef 現在MAX可能性 As Boolean, _
    Optional ByVal 自動再計算 As Boolean = False)
    
    #If 切替詳細ログ Then
        詳細ログ出力 "ブロック内早遅自動判定" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "開始"
    #End If
    
    Dim 行 As Long
    Dim マーク As Long
    
    Dim ボタン数 As Long
    'Dim ホールド変化フラグ As Boolean
    
    行 = 開始行
    
    Do Until 行 > 終了行
                
        If Def.譜面テーブル.ホールド開始フレーム列(行) <> Def.譜面テーブル.ホールド開始フレーム列(行 - 1) Then
        
            ' 1. 現在の前後のホールド数から早遅またはジャスト(タイミング不問)を設定
            
            ボタン数 = 0
            
            For マーク = 1 To マーク数
                
                If Def.譜面テーブル.ホールド個別開始フレーム列(マーク, 行) > 0 Then
                    
                    ボタン数 = ボタン数 + 1
                    
                End If
                
            Next マーク
            
            If Not Def.早遅手動指定フラグ(行) Then
                
                ' 直前にスタートしたホールドがMAXが入る可能性がある場合は別扱い
                
                If 現在MAX可能性 Then
                    
                    If Def.譜面テーブル.ホールド終了押し直し判定列(行) Then
                        
                        If ボタン数 > 0 Then
                            
                            Def.譜面テーブル.早遅手動指定列(行) = 早COOL文字
                            
                        Else
                            
                            Def.譜面テーブル.早遅手動指定列(行) = 遅COOL文字
                            
                        End If
                        
                    Else
                        
                        ' 押し直しなしで(現在の基準拍数での1拍以上余裕を持って)MAXが入る場合は早遅指定なし
                        
                    End If
                    
                Else
                    
                    If ボタン数 > 現在ホールドボタン数 Then
                    
                        Def.譜面テーブル.早遅手動指定列(行) = 早COOL文字
                        
                    ElseIf ボタン数 = 現在ホールドボタン数 Then
                        
                        Def.譜面テーブル.早遅手動指定列(行) = ジャストCOOL文字
                        
                    ElseIf ボタン数 < 現在ホールドボタン数 Then
                        
                        Def.譜面テーブル.早遅手動指定列(行) = 遅COOL文字
                        
                    End If
                    
                End If
                
                Def.譜面テーブル.再計算 行, 終了行, 自動再計算
                'DoEvents
                
            End If
            
            ' 2. MAX可能性のあるものについてMAXが入る場合はよりよく入るように遡って早遅を修正
            
            If Def.譜面テーブル.ホールドボーナス列(行) > 0 Then
                
                If 現在MAX可能性 Then
                    
                    ' a. MAX可能性がある場合、まず入りのタイミングを変更=遅くして試してみる(手動指定されていない場合)
                    
                    Dim 試行前入り早遅指定 As String
                    Dim 試行前抜け早遅指定 As String
                    
                    試行前入り早遅指定 = Def.譜面テーブル.早遅手動指定列(現在ホールド開始行)
                    試行前抜け早遅指定 = Def.譜面テーブル.早遅手動指定列(行)
                    
                    If Not Def.早遅手動指定フラグ(現在ホールド開始行) Then
                        
                        If isMAXタイミング不問(現在ホールド開始行) And (Def.譜面テーブル.早遅手動指定列(現在ホールド開始行) <> 遅COOL文字) Then
                            
                            Def.譜面テーブル.早遅手動指定列(現在ホールド開始行) = ジャストCOOL文字
                            
                        Else
                            
                            Def.譜面テーブル.早遅手動指定列(現在ホールド開始行) = 遅COOL文字
                            
                        End If
                        
                        Def.譜面テーブル.再計算 現在ホールド開始行, 終了行, 自動再計算
                        'DoEvents
                        
                        ' 入りのタイミングを遅くしたことでこの行でMAXホールドボーナスが入らなくなった場合は、現在の行の早遅指定をキャンセルしてそのまま次の行へ
                        ' 終了行だった場合はさらにもう一行追加で評価される
                        
                        If Def.譜面テーブル.ホールドボーナス列(行) = 0 Then
                        
                            If Not Def.早遅手動指定フラグ(行) Then
                                
                                Def.譜面テーブル.早遅手動指定列(行) = ""
                                                                
                            End If
                            
                            If 行 = 終了行 Then
                            
                                終了行 = 終了行 + 1
                                Def.譜面テーブル.OwnTable.ListRows(終了行).Range.Rows.Hidden = False
                                
                            End If
                            
                            Def.譜面テーブル.再計算 行, 終了行, 自動再計算
                            'DoEvents
                            
                        End If
                        
                        ' 現在の早遅の状態でこの行で入るボーナスがMAXでない場合、入りを早くする
                        
                        If Def.譜面テーブル.ホールドボーナス列(行) > 0 And Def.譜面テーブル.ホールドフレーム列(行) <= Def.未MAXフレーム最大値 Then
                            
                            Def.譜面テーブル.早遅手動指定列(現在ホールド開始行) = 早COOL文字
                            
                            Def.譜面テーブル.再計算 現在ホールド開始行, 終了行, 自動再計算
                            'DoEvents
                            
                        End If
                        
                    End If
                    
                    ' b. 入りのタイミングの変更だけでこの行で入るボーナスがMAXにならない場合は抜けのタイミングも変更してみる
                    
                    If Not Def.早遅手動指定フラグ(行) Then
                        
                        If Def.譜面テーブル.ホールドボーナス列(行) > 0 And Def.譜面テーブル.ホールドフレーム列(行) <= Def.未MAXフレーム最大値 Then
                            
                            Def.譜面テーブル.早遅手動指定列(行) = 遅COOL文字
                            
                            Def.譜面テーブル.再計算 行, 終了行, 自動再計算
                            'DoEvents
                            
                        End If
                        
                    End If
                    
                    ' c. それでもこの行で入るボーナスがMAXにならない場合はMAX不可能とし、入りの早遅のタイミングを元に戻す
                    ' また、この段階でMAXが入っても、もし直後のホールドが余裕のないMAXである場合などに、このMAXがキャンセルされる場合があるかもしれない
                    
                    If Def.譜面テーブル.ホールドボーナス列(行) > 0 And Def.譜面テーブル.ホールドフレーム列(行) <= Def.未MAXフレーム最大値 Then
                        
                        If Not Def.早遅手動指定フラグ(現在ホールド開始行) Then
                        
                            Def.譜面テーブル.早遅手動指定列(現在ホールド開始行) = 試行前入り早遅指定
                            
                        End If
                        
                        If Not Def.早遅手動指定フラグ(行) Then
                        
                            Def.譜面テーブル.早遅手動指定列(行) = 試行前抜け早遅指定
                            
                        End If
                        
                        Def.譜面テーブル.再計算 現在ホールド開始行, 終了行, 自動再計算
                        'DoEvents
                        
                    End If
                    
                    ' 現在ホールド開始行 = 開始行 - 1 行目の場合、その行の早遅は(※基本的に)関係ないので元に戻す
                    ' ※(開始行 - 1) 行目が直前のホールド開始行だった場合は考慮しない?
                    ' →早遅が変更されているので確認不可
                    
'                    If 現在ホールド開始行 = 開始行 - 1 Then
'
'                        早遅手動指定列(現在ホールド開始行) = 試行前入り早遅指定

'                        Def.譜面テーブル.再計算 現在ホールド開始行, 終了行, 自動再計算
'                        'DoEvents
'
'                    End If
                    
                End If
                
            End If
            
            ' 3. もう一度ホールドに変化があったかチェックし、変化があった場合は
            '    現在ホールド開始行とボタン数を現在の行とボタン数に再設定
            
            If Def.譜面テーブル.ホールド開始フレーム列(行) <> Def.譜面テーブル.ホールド開始フレーム列(行 - 1) Then
            
                現在ホールド開始行 = 行
                現在ホールドボタン数 = ボタン数
                
                現在MAX可能性 = MAX可能性フラグ(現在ホールド開始行)
                
            End If
            
        End If
        
        行 = 行 + 1
        
    Loop
    
    #If 切替詳細ログ Then
        詳細ログ出力 "ブロック内早遅自動判定" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "終了"
    #End If
    
End Function

Private Function isMAXタイミング不問(ByVal ホールド開始行 As Long) As Boolean
    isMAXタイミング不問 = True
    Dim マーク As Long
    For マーク = 1 To Def.マーク数
        If Def.譜面テーブル.ホールド個別開始フレーム列(マーク, ホールド開始行) > 0 Then
            If Def.譜面テーブル.ホールド個別開始フレーム列(マーク, ホールド開始行) < Def.譜面テーブル.ホールド開始フレーム列(ホールド開始行) Then
                isMAXタイミング不問 = False
                Exit For
            End If
        End If
    Next マーク
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Public Sub 指定範囲の早遅指定削除( _
    ByVal 開始行 As Long, _
    ByVal 終了行 As Long, _
    Optional ByVal 自動再計算 As Boolean = False)
    
    ' 早遅指定削除 ---------------------------------------------------------------------------------
    #If 切替詳細ログ Then
        詳細ログ出力 "指定範囲の早遅指定解除" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "開始"
    #End If
    
    Dim 行 As Long
    
    For 行 = 開始行 To 終了行
        
        If Not Def.早遅手動指定フラグ(行) Then
            
            If Def.譜面テーブル.早遅手動指定列(行) <> "" Then
                
                Def.譜面テーブル.早遅手動指定列(行) = ""
                
            End If
            
        End If
        
    Next 行
    
    For 行 = 開始行 To 終了行
        MAX可能性フラグ(行) = False
    Next 行
    
    Def.譜面テーブル.再計算 開始行, 終了行, 自動再計算
    DoEvents
    
    #If 切替詳細ログ Then
        詳細ログ出力 "指定範囲の早遅指定解除" & vbTab & 開始行 & vbTab & 終了行 & vbTab & "終了"
    #End If
    
End Sub

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Private Function 切替文字列取得(ByVal 行 As Long) As String
    
    Dim マーク As Long
    
    切替文字列取得 = ノーツ番号文字列の取得(行) & "["
    
    For マーク = 1 To マーク数
        If Def.譜面テーブル.ホールド個別開始フレーム列(マーク, 行 - 1) > 0 Then
            切替文字列取得 = 切替文字列取得 & マーク文字(マーク)
        End If
    Next マーク
    
    切替文字列取得 = 切替文字列取得 & "→"
    
    For マーク = 1 To マーク数
        If Def.譜面テーブル.ホールド個別開始フレーム列(マーク, 行) > 0 Then
            切替文字列取得 = 切替文字列取得 & マーク文字(マーク)
        End If
    Next マーク
    
    切替文字列取得 = 切替文字列取得 & "]"
    
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Private Function 早遅文字またはフレームの取得(ByVal 行 As Long) As String
    If Not Def.譜面テーブル.早遅手動指定列(行) = "" Then
        早遅文字またはフレームの取得 = "{" & Def.譜面テーブル.早遅手動指定列(行) & "}"
    End If
End Function

' ======================================================================================================================================================================================================
'
' ======================================================================================================================================================================================================

Private Function ノーツ番号文字列の取得( _
    ByVal 行 As Long, _
    Optional ByVal コンボ0出力 As Boolean = True, _
    Optional ByVal コンボ1出力 As Boolean = True) _
    As String
    
    Dim 切替ノーツ番号 As Long
    Dim 切替コンボ数 As Long

    切替ノーツ番号 = Def.譜面テーブル.ノーツ番号列(行)
    切替コンボ数 = Def.譜面テーブル.コンボ列(行)
    
    ノーツ番号文字列の取得 = 切替ノーツ番号
    If 切替ノーツ番号 <> 切替コンボ数 Then
        If (切替コンボ数 = 0 And コンボ0出力) Or (切替コンボ数 = 1 And コンボ1出力) Or 切替コンボ数 > 1 Then
            ノーツ番号文字列の取得 = ノーツ番号文字列の取得 & "<" & 切替コンボ数 & ">"
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

Private Function 詳細ログ文字列取得(ByVal 出力文字列 As String)
    Dim time As Double
    time = Timer
    詳細ログ文字列取得 = "TIME:" & vbTab & Format(Now, "yyyy-mm-ddThh:nn:ss") & Format(time - Int(time), ".000") & vbTab & 出力文字列
End Function

Private Function 詳細ログ出力(ByVal 出力文字列 As String, Optional ByVal isユーザー出力 As Boolean = False, Optional ByVal 出力行番号 As Long = -1)
    If isユーザー出力 Then
        Def.処理ログ.出力 詳細ログ文字列取得(出力文字列), True, 出力行番号
    Else
        Debug.Print 詳細ログ文字列取得(出力文字列)
    End If
End Function

