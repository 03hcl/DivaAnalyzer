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

Private m_番号列 As ListColumn
Private m_最大スコア差列 As ListColumn
Private m_スコア列 As ListColumn
Private m_達成率列 As ListColumn
Private m_最低ライフ列 As ListColumn
Private m_ルート文字列 As ListColumn
Private m_早遅切替シート列 As ListColumn
Private m_影響開始列 As ListColumn
Private m_影響終了列 As ListColumn

Private m_is簡易版 As Boolean

Public Function オブジェクト新規作成(Optional ByVal 譜面テーブル As IChartTable) As Long
    
    Me.オブジェクト解放
    
    ' シート作成 -----------------------------------------------------------------------------------
    
    Application.ScreenUpdating = False
    
    Dim アクティブシート As Worksheet
    Set アクティブシート = ThisWorkbook.ActiveSheet
    
    On Error GoTo シートを作成できない場合
    
    Set m_Sheet = Worksheets.Add(after:=譜面テーブル.所属シート)
        
    On Error GoTo シート名変更エラー
    
    Dim 結果シート名 As String
    結果シート名 = 譜面テーブル.所属シート.name & "_スコアタルート"
    Do Until Me.所属シート.name = 結果シート名
        Me.所属シート.name = 結果シート名
    Loop
    
    ' テーブル作成 ---------------------------------------------------------------------------------
    
    On Error GoTo テーブルを作成できない場合
    
    Set m_OwnTable = Me.所属シート.ListObjects.Add(xlSrcRange, Me.所属シート.Range(Me.所属シート.Cells(1, 1), Me.所属シート.Cells(1, 9)), , xlYes)
    
    On Error GoTo テーブル名変更エラー
    
    Dim 結果テーブル名 As String
    結果テーブル名 = 譜面テーブル.OwnTable.name & "_スコアタルート"
    Do Until Me.OwnTable.name = 結果テーブル名
        Me.OwnTable.name = 結果テーブル名
    Loop
    
    ' テーブル整形 ---------------------------------------------------------------------------------
    
    On Error GoTo 0
    
    Me.OwnTable.ShowTableStyleFirstColumn = True
    Me.OwnTable.TableStyle = 譜面テーブル.OwnTable.TableStyle
    
    Me.OwnTable.ListColumns(1).name = "No"
    Me.OwnTable.ListColumns(2).name = "最大差"
    Me.OwnTable.ListColumns(3).name = "スコア"
    Me.OwnTable.ListColumns(4).name = "達成率"
    Me.OwnTable.ListColumns(5).name = "最低LIFE"
    Me.OwnTable.ListColumns(6).name = "ルート文字列"
    Me.OwnTable.ListColumns(7).name = "早遅切替シート名"
    Me.OwnTable.ListColumns(8).name = "影響開始行"
    Me.OwnTable.ListColumns(9).name = "影響終了行"
    
    Me.OwnTable.ListColumns(1).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(2).Range.ColumnWidth = 10
    Me.OwnTable.ListColumns(3).Range.ColumnWidth = 10
    Me.OwnTable.ListColumns(4).Range.ColumnWidth = 8
    Me.OwnTable.ListColumns(5).Range.ColumnWidth = 8
    Me.OwnTable.ListColumns(6).Range.ColumnWidth = 72
    Me.OwnTable.ListColumns(7).Range.ColumnWidth = 30
    Me.OwnTable.ListColumns(8).Range.ColumnWidth = 10
    Me.OwnTable.ListColumns(9).Range.ColumnWidth = 10
    
    Set m_番号列 = Me.OwnTable.ListColumns(1)
    Set m_最大スコア差列 = Me.OwnTable.ListColumns(2)
    Set m_スコア列 = Me.OwnTable.ListColumns(3)
    Set m_達成率列 = Me.OwnTable.ListColumns(4)
    Set m_最低ライフ列 = Me.OwnTable.ListColumns(5)
    Set m_ルート文字列 = Me.OwnTable.ListColumns(6)
    Set m_早遅切替シート列 = Me.OwnTable.ListColumns(7)
    Set m_影響開始列 = Me.OwnTable.ListColumns(8)
    Set m_影響終了列 = Me.OwnTable.ListColumns(9)
    
    アクティブシート.Activate
    
    Application.ScreenUpdating = True
    
    m_is簡易版 = False
    
    オブジェクト新規作成 = 0
    Exit Function
    
シート名変更エラー:
    Dim 結果シート名連番 As Long
    結果シート名連番 = 結果シート名連番 + 1
    結果シート名 = 譜面テーブル.所属シート.name & "_スコアタルート_" & 結果シート名連番
    Resume Next
    
テーブル名変更エラー:
    Dim 結果テーブル名連番 As Long
    結果テーブル名連番 = 結果テーブル名連番 + 1
    結果テーブル名 = 譜面テーブル.OwnTable.name & "_スコアタルート_" & 結果テーブル名連番
    Resume Next
    
シートを作成できない場合:
    Application.ScreenUpdating = True
    オブジェクト新規作成 = -1
    GoTo エラー処理
    
テーブルを作成できない場合:
    オブジェクト新規作成 = -2
    GoTo エラー処理
    
エラー処理:
    Me.オブジェクト解放
    Exit Function
    
End Function

Public Function オブジェクト設定(Optional ByVal テーブル名 As String = "") As Long
    
    Me.オブジェクト解放
    
    ' テーブル設定 ---------------------------------------------------------------------------------
    
    On Error GoTo テーブルを発見できない場合
    
    If Def.テーブルとシートの検索設定(m_OwnTable, m_Sheet, テーブル名) < 0 Then
        GoTo テーブルを発見できない場合
    End If
    
    ' メンバ設定 -----------------------------------------------------------------------------------
    
    On Error GoTo テーブルのフォーマット形式が異なる場合
    
    Set m_番号列 = Me.OwnTable.ListColumns("No")
    Set m_最大スコア差列 = Me.OwnTable.ListColumns("最大差")
    Set m_スコア列 = Me.OwnTable.ListColumns("スコア")
    Set m_達成率列 = Me.OwnTable.ListColumns("達成率")
    Set m_最低ライフ列 = Me.OwnTable.ListColumns("最低LIFE")
    Set m_ルート文字列 = Me.OwnTable.ListColumns("ルート文字列")
    Set m_早遅切替シート列 = Me.OwnTable.ListColumns("早遅切替シート名")
    Set m_影響開始列 = Me.OwnTable.ListColumns("影響開始行")
    Set m_影響終了列 = Me.OwnTable.ListColumns("影響終了行")
    
    On Error GoTo 0
    
    m_is簡易版 = False
    
    オブジェクト設定 = 0
    Exit Function
    
    ' エラーラベル ---------------------------------------------------------------------------------
    
テーブルを発見できない場合:
    オブジェクト設定 = -1
    GoTo エラー処理
    
テーブルのフォーマット形式が異なる場合:
    オブジェクト設定 = -2
    GoTo エラー処理
    
エラー処理:
    Me.オブジェクト解放
    Exit Function
    
End Function

Public Sub オブジェクト解放()
    
    Set m_OwnTable = Nothing
    Set m_Sheet = Nothing
    
    Set m_番号列 = Nothing
    Set m_最大スコア差列 = Nothing
    Set m_スコア列 = Nothing
    Set m_達成率列 = Nothing
    Set m_最低ライフ列 = Nothing
    Set m_ルート文字列 = Nothing
    Set m_早遅切替シート列 = Nothing
    Set m_影響開始列 = Nothing
    Set m_影響終了列 = Nothing
    
    m_is簡易版 = False
    
End Sub

Public Function オブジェクト最終整形(ByVal 譜面テーブル As IChartTable)
    
    If Me.データ行数 > 0 Then
        m_最大スコア差列.DataBodyRange(データ行数).Formula = "=[@スコア]-MAX([スコア])"
    End If
    
    Me.OwnTable.HeaderRowRange.HorizontalAlignment = xlCenter
    
    m_達成率列.DataBodyRange.NumberFormatLocal = "0.00%"
    
'    Me.OwnTable.ShowAutoFilterDropDown = False
    
End Function

Public Function 出力行追加()
    Me.OwnTable.ListRows.Add
    番号列(Me.データ行数) = Me.データ行数
End Function

Public Function 現在ルート出力( _
    ByVal 譜面テーブル As IChartTable, _
    ByVal スコアタルート文字列 As String, _
    ByVal 影響開始行 As Long, _
    ByVal 影響終了行 As Long, _
    ByVal 達成率 As Double, _
    ByVal スコア As Long)
    
    Me.出力行追加
    
    スコア列(Me.データ行数) = スコア
    達成率列(Me.データ行数) = 達成率
    最低ライフ列(Me.データ行数) = 譜面テーブル.最低ライフ
    ルート文字列(Me.データ行数) = スコアタルート文字列
    影響開始列(Me.データ行数) = 影響開始行
    影響終了列(Me.データ行数) = 影響終了行
    
    If Not Me.Is簡易版 Then
        Dim 現在早遅切替一覧テーブル As ElSwTable
        Set 現在早遅切替一覧テーブル = New ElSwTable
        If 現在早遅切替一覧テーブル.オブジェクト新規作成(譜面テーブル) < 0 Then
            GoTo テーブル設定に失敗した場合
        End If
        現在早遅切替一覧テーブル.早遅切替情報読み込み 譜面テーブル
        
        早遅切替シート列(Me.データ行数) = 現在早遅切替一覧テーブル.所属シート.name
    End If
    
    Exit Function
    
テーブル設定に失敗した場合:
    MsgBox "ERR:早遅と切替の一覧テーブルの作成に失敗しました。" & vbCrLf & "処理は続行されます。", vbCritical
    Exit Function
    
End Function

Public Function 組み合わせルート出力( _
    ByVal 譜面テーブル As IChartTable, _
    ByVal 早遅切替一覧テーブル As ElSwTable, _
    Optional ByVal 完奏モード As Boolean = False, _
    Optional ByVal 直前影響終了行 As Long = 0, _
    Optional ByVal ネスト数 As Long = 0)
    
    If Me.Is簡易版 Then
        Exit Function
    End If
    
    Dim 行 As Long
    
    Dim 影響開始行 As Long
    Dim 影響終了行 As Long
    
    Dim リザルト As Def.結果セット
    
    For 行 = 1 To Me.データ行数
        影響開始行 = 影響開始列(行)
        If 影響開始行 > 直前影響終了行 Then
            早遅切替一覧テーブル.早遅切替情報書き出し 譜面テーブル, 直前影響終了行 + 1, 影響開始行 - 1
            Dim 現在早遅切替一覧テーブル As ElSwTable
            Set 現在早遅切替一覧テーブル = New ElSwTable
            現在早遅切替一覧テーブル.オブジェクト設定 ThisWorkbook.Worksheets(早遅切替シート列(行)).ListObjects(1).name
            影響終了行 = 影響終了列(行)
            現在早遅切替一覧テーブル.早遅切替情報書き出し 譜面テーブル, 影響開始行, 影響終了行
            Me.組み合わせルート出力 譜面テーブル, 早遅切替一覧テーブル, 完奏モード, 影響終了行, ネスト数 + 1
        End If
    Next 行
    
    If ネスト数 > 1 Then
        早遅切替一覧テーブル.早遅切替情報書き出し 譜面テーブル, 直前影響終了行 + 1, 譜面テーブル.データ行数
        リザルト = Def.譜面テーブル.リザルト再計算()
        If 完奏モード Or リザルト.クリアランク <> Def.MISSTAKE文字 Then
            Me.現在ルート出力 譜面テーブル, Analyzing.スコアタルート文字列取得(), 1, 譜面テーブル.データ行数, リザルト.達成率, リザルト.スコア
        End If
    End If
    
End Function

' **************************************************************************************************
' テーブル自身
' **************************************************************************************************

Public Property Get OwnTable() As ListObject
    Set OwnTable = m_OwnTable
End Property

Public Property Get 所属シート() As Worksheet
    Set 所属シート = m_Sheet
End Property

' **************************************************************************************************
' テーブル定数
' **************************************************************************************************

Public Property Get データ行数() As Long
    If Me.OwnTable Is Nothing Then
        データ行数 = -1
    Else
        データ行数 = Me.OwnTable.ListRows.count
    End If
End Property

' **************************************************************************************************
' 列
' **************************************************************************************************

Private Property Get 番号列(ByVal 行 As Long) As Long
    番号列 = m_番号列.DataBodyRange(行).value
End Property

Private Property Let 番号列(ByVal 行 As Long, ByVal 値 As Long)
    If 番号列(行) <> 値 Then
        m_番号列.DataBodyRange(行).value = 値
    End If
End Property

Private Property Get スコア列(ByVal 行 As Long) As Long
    スコア列 = m_スコア列.DataBodyRange(行).value
End Property

Private Property Let スコア列(ByVal 行 As Long, ByVal 値 As Long)
    If スコア列(行) <> 値 Then
        m_スコア列.DataBodyRange(行).value = 値
    End If
End Property

Private Property Get 達成率列(ByVal 行 As Long) As Double
    達成率列 = m_達成率列.DataBodyRange(行).value
End Property

Private Property Let 達成率列(ByVal 行 As Long, ByVal 値 As Double)
    If 達成率列(行) <> 値 Then
        m_達成率列.DataBodyRange(行).value = 値
    End If
End Property

Private Property Get 最低ライフ列(ByVal 行 As Long) As Long
    最低ライフ列 = m_最低ライフ列.DataBodyRange(行).value
End Property

Private Property Let 最低ライフ列(ByVal 行 As Long, ByVal 値 As Long)
    If 最低ライフ列(行) <> 値 Then
        m_最低ライフ列.DataBodyRange(行).value = 値
    End If
End Property

Private Property Get ルート文字列(ByVal 行 As Long) As String
    ルート文字列 = m_ルート文字列.DataBodyRange(行).value
End Property

Private Property Let ルート文字列(ByVal 行 As Long, ByVal 値 As String)
    If ルート文字列(行) <> 値 Then
        m_ルート文字列.DataBodyRange(行).value = 値
    End If
End Property

Private Property Get 早遅切替シート列(ByVal 行 As Long) As String
    早遅切替シート列 = m_早遅切替シート列.DataBodyRange(行).value
End Property

Private Property Let 早遅切替シート列(ByVal 行 As Long, ByVal 値 As String)
    If 早遅切替シート列(行) <> 値 Then
        m_早遅切替シート列.DataBodyRange(行).value = 値
    End If
End Property

Private Property Get 影響開始列(ByVal 行 As Long) As Long
    影響開始列 = m_影響開始列.DataBodyRange(行).value
End Property

Private Property Let 影響開始列(ByVal 行 As Long, ByVal 値 As Long)
    If 影響開始列(行) <> 値 Then
        m_影響開始列.DataBodyRange(行).value = 値
    End If
End Property

Private Property Get 影響終了列(ByVal 行 As Long) As Long
    影響終了列 = m_影響終了列.DataBodyRange(行).value
End Property

Private Property Let 影響終了列(ByVal 行 As Long, ByVal 値 As Long)
    If 影響終了列(行) <> 値 Then
        m_影響終了列.DataBodyRange(行).value = 値
    End If
End Property

Public Property Get Is簡易版() As Boolean
    Is簡易版 = m_is簡易版
End Property

Public Property Let Is簡易版(ByVal 値 As Boolean)
    m_is簡易版 = 値
End Property


