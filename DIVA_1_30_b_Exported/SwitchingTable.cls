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

Private m_番号列 As ListColumn
Private m_解析ブロック列 As ListColumn
Private m_ブロック開始列 As ListColumn
Private m_ブロック終了列 As ListColumn
Private m_最大スコア差列 As ListColumn
Private m_ブロックスコア列 As ListColumn
Private m_切替スコア差列 As ListColumn
Private m_切替文字列 As ListColumn
Private m_MAX可能性文字列 As ListColumn
'Private m_切替列() As listColumn

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
    結果シート名 = 譜面テーブル.所属シート.name & "_切替"
    Do Until Me.所属シート.name = 結果シート名
        Me.所属シート.name = 結果シート名
    Loop
    
    ' テーブル作成 ---------------------------------------------------------------------------------
    
    On Error GoTo テーブルを作成できない場合
    
    Set m_OwnTable = Me.所属シート.ListObjects.Add(xlSrcRange, Me.所属シート.Range(Me.所属シート.Cells(1, 1), Me.所属シート.Cells(1, 9)), , xlYes)
    
    On Error GoTo テーブル名変更エラー
    
    Dim 結果テーブル名 As String
    結果テーブル名 = 譜面テーブル.OwnTable.name & "_切替"
    Do Until Me.OwnTable.name = 結果テーブル名
        Me.OwnTable.name = 結果テーブル名
    Loop
    
    On Error GoTo 0
    
    ' テーブル整形 ---------------------------------------------------------------------------------
    
    Me.OwnTable.ShowTableStyleFirstColumn = True
    Me.OwnTable.TableStyle = 譜面テーブル.OwnTable.TableStyle
    
    Me.OwnTable.ListColumns(1).name = "No"
    Me.OwnTable.ListColumns(2).name = "HBlock"
    Me.OwnTable.ListColumns(3).name = "開始行"
    Me.OwnTable.ListColumns(4).name = "終了行"
    Me.OwnTable.ListColumns(5).name = "最大差"
    Me.OwnTable.ListColumns(6).name = "スコア"
    Me.OwnTable.ListColumns(7).name = "切替差"
    Me.OwnTable.ListColumns(8).name = "切替"
    Me.OwnTable.ListColumns(9).name = "MAX可能性"
    
    Me.OwnTable.ListColumns(1).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(2).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(3).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(4).Range.ColumnWidth = 6
    Me.OwnTable.ListColumns(5).Range.ColumnWidth = 8
    Me.OwnTable.ListColumns(6).Range.ColumnWidth = 8
    Me.OwnTable.ListColumns(7).Range.ColumnWidth = 8
    Me.OwnTable.ListColumns(8).Range.ColumnWidth = 24
    Me.OwnTable.ListColumns(9).Range.ColumnWidth = 48
    
    Set m_番号列 = Me.OwnTable.ListColumns(1)
    Set m_解析ブロック列 = Me.OwnTable.ListColumns(2)
    Set m_ブロック開始列 = Me.OwnTable.ListColumns(3)
    Set m_ブロック終了列 = Me.OwnTable.ListColumns(4)
    Set m_最大スコア差列 = Me.OwnTable.ListColumns(5)
    Set m_ブロックスコア列 = Me.OwnTable.ListColumns(6)
    Set m_切替スコア差列 = Me.OwnTable.ListColumns(7)
    Set m_切替文字列 = Me.OwnTable.ListColumns(8)
    Set m_MAX可能性文字列 = Me.OwnTable.ListColumns(9)
    
    アクティブシート.Activate
    
    Application.ScreenUpdating = True
    
    オブジェクト新規作成 = 0
    Exit Function
    
シート名変更エラー:
    Dim 結果シート名連番 As Long
    結果シート名連番 = 結果シート名連番 + 1
    結果シート名 = 譜面テーブル.所属シート.name & "_切替_" & 結果シート名連番
    Resume Next
    
テーブル名変更エラー:
    Dim 結果テーブル名連番 As Long
    結果テーブル名連番 = 結果テーブル名連番 + 1
    結果テーブル名 = 譜面テーブル.OwnTable.name & "_切替_" & 結果テーブル名連番
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
    Set m_解析ブロック列 = Me.OwnTable.ListColumns("HBlock")
    Set m_ブロック開始列 = Me.OwnTable.ListColumns("開始行")
    Set m_ブロック終了列 = Me.OwnTable.ListColumns("終了行")
    Set m_最大スコア差列 = Me.OwnTable.ListColumns("最大差")
    Set m_ブロックスコア列 = Me.OwnTable.ListColumns("スコア")
    Set m_切替スコア差列 = Me.OwnTable.ListColumns("切替差")
    Set m_切替文字列 = Me.OwnTable.ListColumns("切替")
    Set m_MAX可能性文字列 = Me.OwnTable.ListColumns("MAX可能性")
    
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

Public Function オブジェクト最終整形() As Long
    
    If Me.データ行数 > 0 Then
        m_最大スコア差列.DataBodyRange(データ行数).FormulaArray = "=[@スコア]-MAX(([HBlock]=[@HBlock])*[スコア])"
        m_切替スコア差列.DataBodyRange(データ行数).FormulaArray = "=[@スコア]-MAX(([HBlock]=[@HBlock])*([切替]="""")*[スコア])"
    End If
    
    Dim 列 As Long
    For 列 = 9 + 1 To Me.OwnTable.ListColumns.count
        Me.OwnTable.ListColumns(列).name = "行" & (列 - 9)
        Me.OwnTable.ListColumns(列).Range.ColumnWidth = 5
    Next 列
    
    Me.OwnTable.HeaderRowRange.HorizontalAlignment = xlCenter
    
End Function

Public Function 最善切替情報取得(ByVal 譜面テーブル As IChartTable, Optional ByVal 自動再計算 As Boolean = False) As Switching
    
    Set 最善切替情報取得 = New Switching
    最善切替情報取得.オブジェクト設定 譜面テーブル
    Dim 行 As Long
    Dim 列 As Long
    Dim 切替対象行 As Long
    
    If Not 自動再計算 Then
        Me.OwnTable.Range.Calculate
    End If
    
    For 行 = 1 To Me.データ行数
        If m_最大スコア差列.DataBodyRange(行).value = 0 Then
            最善切替情報取得.切替文字列 = Def.文字列連結(最善切替情報取得.切替文字列, Me.切替文字列(行), vbCrLf)
            For 列 = 9 + 1 To Me.OwnTable.ListColumns.count
            切替対象行 = Me.OwnTable.ListColumns(列).DataBodyRange(行).value
                If 切替対象行 > 0 Then
                    最善切替情報取得.切替フラグ(切替対象行) = True
                End If
            Next
        End If
    Next 行
    
    For 行 = 1 To 譜面テーブル.データ行数
        If 譜面テーブル.切替判定列(行) Then
            最善切替情報取得.切替フラグ(行) = True
        End If
    Next
    
End Function

Public Sub オブジェクト解放()
    
    Set m_OwnTable = Nothing
    Set m_Sheet = Nothing
    
    Set m_番号列 = Nothing
    Set m_解析ブロック列 = Nothing
    Set m_ブロック開始列 = Nothing
    Set m_ブロック終了列 = Nothing
    Set m_最大スコア差列 = Nothing
    Set m_ブロックスコア列 = Nothing
    Set m_切替スコア差列 = Nothing
    Set m_切替文字列 = Nothing
    Set m_MAX可能性文字列 = Nothing
    
End Sub

Public Function 切替行入力(ByVal 入力行数 As Long, ByVal 切替数 As Long)
    Me.所属シート.Cells(Me.データ行数 + 1, 切替数 + 9).value = 入力行数
End Function

Public Function 出力行追加()
    Me.OwnTable.ListRows.Add
    番号列(Me.データ行数) = Me.データ行数
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

Public Property Get 最大ブロック開始行() As Long
    最大ブロック開始行 = Application.WorksheetFunction.Max(m_ブロック開始列.DataBodyRange)
End Property

Public Property Get 最大ブロック終了行() As Long
    最大ブロック終了行 = Application.WorksheetFunction.Max(m_ブロック終了列.DataBodyRange)
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

Public Property Get 解析ブロック列(ByVal 行 As Long) As Long
    解析ブロック列 = m_解析ブロック列.DataBodyRange(行).value
End Property

Public Property Let 解析ブロック列(ByVal 行 As Long, ByVal 値 As Long)
    If 解析ブロック列(行) <> 値 Then
        m_解析ブロック列.DataBodyRange(行).value = 値
    End If
End Property

Public Property Get ブロック開始列(ByVal 行 As Long) As Long
    ブロック開始列 = m_ブロック開始列.DataBodyRange(行).value
End Property

Public Property Let ブロック開始列(ByVal 行 As Long, ByVal 値 As Long)
    If ブロック開始列(行) <> 値 Then
        m_ブロック開始列.DataBodyRange(行).value = 値
    End If
End Property

Public Property Get ブロック終了列(ByVal 行 As Long) As Long
    ブロック終了列 = m_ブロック終了列.DataBodyRange(行).value
End Property

Public Property Let ブロック終了列(ByVal 行 As Long, ByVal 値 As Long)
    If ブロック終了列(行) <> 値 Then
        m_ブロック終了列.DataBodyRange(行).value = 値
    End If
End Property

Public Property Get ブロックスコア列(ByVal 行 As Long) As Long
    ブロックスコア列 = m_ブロックスコア列.DataBodyRange(行).value
End Property

Public Property Let ブロックスコア列(ByVal 行 As Long, ByVal 値 As Long)
    If ブロックスコア列(行) <> 値 Then
        m_ブロックスコア列.DataBodyRange(行).value = 値
    End If
End Property

Public Property Get 切替文字列(ByVal 行 As Long) As String
    切替文字列 = m_切替文字列.DataBodyRange(行).value
End Property

Public Property Let 切替文字列(ByVal 行 As Long, ByVal 値 As String)
    If 切替文字列(行) <> 値 Then
        m_切替文字列.DataBodyRange(行).value = 値
    End If
End Property

Public Property Get MAX可能性文字列(ByVal 行 As Long) As String
    MAX可能性文字列 = m_MAX可能性文字列.DataBodyRange(行).value
End Property

Public Property Let MAX可能性文字列(ByVal 行 As Long, ByVal 値 As String)
    If MAX可能性文字列(行) <> 値 Then
        m_MAX可能性文字列.DataBodyRange(行).value = 値
    End If
End Property

