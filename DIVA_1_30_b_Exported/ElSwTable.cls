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

Private m_番号列 As ListColumn
Private m_行番号列 As ListColumn
Private m_評価列 As ListColumn
Private m_早遅列 As ListColumn
Private m_早遅フレーム列 As ListColumn
Private m_切替列 As ListColumn

Private m_評価リスト() As String
Private m_早遅リスト() As String
Private m_早遅フレームリスト() As String
Private m_切替リスト() As String

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
    結果シート名 = 譜面テーブル.所属シート.name & "_早遅切替リスト"
    Do Until Me.所属シート.name = 結果シート名
        Me.所属シート.name = 結果シート名
    Loop
    
    ' テーブル作成 ---------------------------------------------------------------------------------
    
    On Error GoTo テーブルを作成できない場合
    
    Set m_OwnTable = Me.所属シート.ListObjects.Add(xlSrcRange, Me.所属シート.Range(Me.所属シート.Cells(1, 1), Me.所属シート.Cells(1, 6)), , xlYes)
    
    On Error GoTo テーブル名変更エラー
    
    Dim 結果テーブル名 As String
    結果テーブル名 = 譜面テーブル.OwnTable.name & "_早遅切替リスト"
    Do Until Me.OwnTable.name = 結果テーブル名
        Me.OwnTable.name = 結果テーブル名
    Loop
    
    ' テーブル整形 ---------------------------------------------------------------------------------
    
    On Error GoTo 0
    
    Me.OwnTable.ShowTableStyleFirstColumn = True
    Me.OwnTable.TableStyle = 譜面テーブル.OwnTable.TableStyle
    
    Me.OwnTable.ListColumns(1).name = "No"
    Me.OwnTable.ListColumns(2).name = "行"
    Me.OwnTable.ListColumns(3).name = "評価"
    Me.OwnTable.ListColumns(4).name = "早遅"
    Me.OwnTable.ListColumns(5).name = "早遅Fr"
    Me.OwnTable.ListColumns(6).name = "切替"
    
    Me.OwnTable.ListColumns(1).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(2).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(3).Range.ColumnWidth = 8
    Me.OwnTable.ListColumns(4).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(5).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(6).Range.ColumnWidth = 6
    
    Set m_番号列 = Me.OwnTable.ListColumns(1)
    Set m_行番号列 = Me.OwnTable.ListColumns(2)
    Set m_評価列 = Me.OwnTable.ListColumns(3)
    Set m_早遅列 = Me.OwnTable.ListColumns(4)
    Set m_早遅フレーム列 = Me.OwnTable.ListColumns(5)
    Set m_切替列 = Me.OwnTable.ListColumns(6)
    
    アクティブシート.Activate
    
    Application.ScreenUpdating = True
    
    オブジェクト新規作成 = 0
    Exit Function
    
シート名変更エラー:
    Dim 結果シート名連番 As Long
    結果シート名連番 = 結果シート名連番 + 1
    結果シート名 = 譜面テーブル.所属シート.name & "_早遅切替リスト_" & 結果シート名連番
    Resume Next
    
テーブル名変更エラー:
    Dim 結果テーブル名連番 As Long
    結果テーブル名連番 = 結果テーブル名連番 + 1
    結果テーブル名 = 譜面テーブル.OwnTable.name & "_早遅切替リスト_" & 結果テーブル名連番
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
    Set m_行番号列 = Me.OwnTable.ListColumns("行")
    Set m_評価列 = Me.OwnTable.ListColumns("評価")
    Set m_早遅列 = Me.OwnTable.ListColumns("早遅")
    Set m_早遅フレーム列 = Me.OwnTable.ListColumns("早遅Fr")
    Set m_切替列 = Me.OwnTable.ListColumns("切替")
    
    On Error GoTo 0
    
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
    Set m_行番号列 = Nothing
    Set m_評価列 = Nothing
    Set m_早遅列 = Nothing
    Set m_早遅フレーム列 = Nothing
    Set m_切替列 = Nothing
    
    Erase m_評価リスト
    Erase m_早遅リスト
    Erase m_早遅フレームリスト
    Erase m_切替リスト
    
End Sub

Public Function オブジェクト最終整形(ByVal 譜面テーブル As IChartTable)
    
    Me.OwnTable.HeaderRowRange.HorizontalAlignment = xlCenter
    
    m_評価列.DataBodyRange.HorizontalAlignment = xlCenter
    m_早遅列.DataBodyRange.HorizontalAlignment = xlCenter
    
    m_早遅列.DataBodyRange.Validation.Add 譜面テーブル.早遅手動指定テーブル列.DataBodyRange.Validation.Type, Formula1:=譜面テーブル.早遅手動指定テーブル列.DataBodyRange.Validation.Formula1
    m_切替列.DataBodyRange.Validation.Add 譜面テーブル.切替判定テーブル列.DataBodyRange.Validation.Type, Formula1:=譜面テーブル.切替判定テーブル列.DataBodyRange.Validation.Formula1
    
    Me.OwnTable.ShowAutoFilterDropDown = False
    
End Function

Public Function 早遅切替情報読み込み(ByVal 譜面テーブル As IChartTable, _
    Optional ByVal 開始行 As Long = 1, _
    Optional ByVal 終了行 As Long = -1, _
    Optional ByVal ステータスバー進捗出力 As Boolean = True)
    
    If ステータスバー進捗出力 Then
        Application.StatusBar = "早遅と切替の一覧を譜面からシートに出力しています...... (準備中)"
    End If
    
    If 終了行 = -1 Then
        終了行 = 譜面テーブル.データ行数
    End If
    
    Dim 行 As Long
    Dim is早遅切替指定 As Boolean
    
    For 行 = 1 To 譜面テーブル.データ行数
        
        If ステータスバー進捗出力 Then
            Application.StatusBar = "早遅と切替の一覧を譜面からシートに出力しています...... ( " & 行 & " / " & 譜面テーブル.データ行数 & " 行目 )"
        End If
        
        If 譜面テーブル.評価列(行) <> 譜面テーブル.フレームずれ別評価(0) Then
            is早遅切替指定 = True
        ElseIf 譜面テーブル.早遅手動指定列(行) <> "" Then
            is早遅切替指定 = True
        ElseIf 譜面テーブル.早遅フレーム手動指定列(行) <> "" Then
            is早遅切替指定 = True
        ElseIf 譜面テーブル.切替判定文字列(行) <> "" Then
            is早遅切替指定 = True
        Else
            is早遅切替指定 = False
        End If
        
        If is早遅切替指定 Then
            Me.出力行追加
            行番号列(データ行数) = 行
            If 譜面テーブル.評価列(行) <> 譜面テーブル.フレームずれ別評価(0) Then
                評価列(データ行数) = 譜面テーブル.評価列(行)
            End If
            早遅列(データ行数) = 譜面テーブル.早遅手動指定列(行)
            早遅フレーム列(データ行数) = 譜面テーブル.早遅フレーム手動指定列(行)
            切替列(データ行数) = 譜面テーブル.切替判定文字列(行)
        End If
        
    Next 行
    
    DoEvents
    
    If ステータスバー進捗出力 Then
        Application.StatusBar = "早遅と切替の一覧を譜面からシートに出力しています...... (終了中)"
    End If
    
    Me.オブジェクト最終整形 譜面テーブル
    Me.早遅切替リスト設定 譜面テーブル
    
End Function

Public Function 早遅切替情報書き出し(ByVal 譜面テーブル As IChartTable, _
    Optional ByVal 開始行 As Long = 1, _
    Optional ByVal 終了行 As Long = -1, _
    Optional ByVal ステータスバー進捗出力 As Boolean = True)
    
    If ステータスバー進捗出力 Then
        Application.StatusBar = "早遅と切替の一覧をシートから譜面に反映しています...... (準備中)"
    End If
    
    Me.早遅切替リスト設定 譜面テーブル
    
    If ステータスバー進捗出力 Then
        DoEvents
    End If
    
    If 終了行 = -1 Then
        終了行 = 譜面テーブル.データ行数
    End If
    
    Dim 行 As Long
    
    For 行 = 開始行 To 終了行
        
        If ステータスバー進捗出力 Then
            Application.StatusBar = "早遅と切替の一覧をシートから譜面に反映しています...... ( " & 行 & " / " & 譜面テーブル.データ行数 & " 行目 )"
        End If
        
        If 譜面テーブル.評価列(行) <> m_評価リスト(行) Then
            譜面テーブル.評価列(行) = m_評価リスト(行)
        End If
        If 譜面テーブル.早遅手動指定列(行) <> m_早遅リスト(行) Then
            譜面テーブル.早遅手動指定列(行) = m_早遅リスト(行)
        End If
        If 譜面テーブル.早遅フレーム手動指定列(行) <> m_早遅フレームリスト(行) Then
            譜面テーブル.早遅フレーム手動指定列(行) = m_早遅フレームリスト(行)
        End If
        If 譜面テーブル.切替判定文字列(行) <> m_切替リスト(行) Then
            譜面テーブル.切替判定文字列(行) = m_切替リスト(行)
        End If
        
    Next 行
    
    If ステータスバー進捗出力 Then
        DoEvents
        Application.StatusBar = "早遅と切替の一覧をシートから譜面に反映しています...... (終了)"
    End If
    
End Function

Public Function 出力行追加()
    Me.OwnTable.ListRows.Add
    番号列(Me.データ行数) = Me.データ行数
End Function

Public Function 早遅切替リスト設定(ByVal 譜面テーブル As IChartTable, Optional ByVal 強制再設定 As Boolean = False)
    
    '既にリストが設定されている場合はしない
    If Not 強制再設定 Then
        If Not Not m_評価リスト Then
            If Not Not m_早遅リスト Then
                If Not Not m_早遅フレームリスト Then
                    If Not Not m_切替リスト Then
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    ReDim m_評価リスト(譜面テーブル.データ行数)
    ReDim m_早遅リスト(譜面テーブル.データ行数)
    ReDim m_早遅フレームリスト(譜面テーブル.データ行数)
    ReDim m_切替リスト(譜面テーブル.データ行数)
    
    Dim 対象譜面行 As Long
    Dim 行 As Long
    
    For 行 = 1 To Me.データ行数
        
        対象譜面行 = 行番号列(行)
        
        If m_評価リスト(対象譜面行) <> 評価列(行) Then
            m_評価リスト(対象譜面行) = 評価列(行)
        End If
        If m_早遅リスト(対象譜面行) <> 早遅列(行) Then
            m_早遅リスト(対象譜面行) = 早遅列(行)
        End If
        If m_早遅フレームリスト(対象譜面行) <> 早遅フレーム列(行) Then
            m_早遅フレームリスト(対象譜面行) = 早遅フレーム列(行)
        End If
        If m_切替リスト(対象譜面行) <> 切替列(行) Then
            m_切替リスト(対象譜面行) = 切替列(行)
        End If
        
    Next 行
    
    For 行 = 1 To 譜面テーブル.データ行数
        If m_評価リスト(行) = "" Then
            m_評価リスト(行) = 譜面テーブル.フレームずれ別評価(0)
        End If
    Next
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

Private Property Get 行番号列(ByVal 行 As Long) As Long
    行番号列 = m_行番号列.DataBodyRange(行).value
End Property

Private Property Let 行番号列(ByVal 行 As Long, ByVal 値 As Long)
    If 行番号列(行) <> 値 Then
        m_行番号列.DataBodyRange(行).value = 値
    End If
End Property

Private Property Get 評価列(ByVal 行 As Long) As String
    評価列 = m_評価列.DataBodyRange(行).value
End Property

Private Property Let 評価列(ByVal 行 As Long, ByVal 値 As String)
    If 評価列(行) <> 値 Then
        m_評価列.DataBodyRange(行).value = 値
    End If
End Property

Private Property Get 早遅列(ByVal 行 As Long) As String
    早遅列 = m_早遅列.DataBodyRange(行).value
End Property

Private Property Let 早遅列(ByVal 行 As Long, ByVal 値 As String)
    If 早遅列(行) <> 値 Then
        m_早遅列.DataBodyRange(行).value = 値
    End If
End Property

Private Property Get 早遅フレーム列(ByVal 行 As Long) As String
    早遅フレーム列 = m_早遅フレーム列.DataBodyRange(行).value
End Property

Private Property Let 早遅フレーム列(ByVal 行 As Long, ByVal 値 As String)
    If 早遅フレーム列(行) <> 値 Then
        m_早遅フレーム列.DataBodyRange(行).value = 値
    End If
End Property

Private Property Get 切替列(ByVal 行 As Long) As String
    切替列 = m_切替列.DataBodyRange(行).value
End Property

Private Property Let 切替列(ByVal 行 As Long, ByVal 値 As String)
    If 切替列(行) <> 値 Then
        m_切替列.DataBodyRange(行).value = 値
    End If
End Property

