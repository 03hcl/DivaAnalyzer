VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChartTable2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

#Const 再計算ログ = False

Implements IChartTable

Private m_自動再計算 As Boolean

Private m_OwnTable As ListObject
Private m_Sheet As Worksheet

Private m_最早COOLフレーム As Double
Private m_最遅COOLフレーム As Double
Private m_最早SADフレーム As Double
Private m_最遅SADフレーム As Double

Private m_データ行数 As Long

Private m_テンポチェンジ列() As Boolean

Private m_フレーム列() As Long

Private m_ノーツ列() As String
Private m_スライドノーツ列() As String
Private m_ノーツ個数列() As Long
Private m_チェイン列() As Long

Private m_ノーツ番号列() As Long
Private m_評価列() As String
Private m_コンボ列() As Long
Private m_評価位置列() As Long

Private m_早遅手動指定列() As String
Private m_早遅フレーム手動指定列() As String
Private m_切替判定文字列() As String
'Private m_切替判定列() As Boolean

Private m_早遅文字フレーム列() As Long
Private m_早遅フレーム列() As Long
Private m_ライフ列() As Long

Private m_ボタン押下判定列() As Boolean
Private m_ホールド終了押し直し判定列() As Boolean

Private m_基本スコア列() As Long
Private m_コンボボーナス列() As Long
Private m_ライフボーナス列() As Long
Private m_チェーンスライド基本スコア列() As Long
Private m_チェーンスライドボーナススコア列() As Long
Private m_チェーンスライドボーナス列() As Long

Private m_ホールドフレーム列() As Long
Private m_ホールド個別ボーナス列() As Long
Private m_ホールドボーナス列() As Long

Private m_ホールド個別開始フレーム列() As Long
Private m_ホールド開始フレーム列() As Long

Private m_補正点列() As Long

Private m_スコア列() As Long
Private m_ホールドスコア列() As Long
'Private m_チェーンスライドスコア列() As Long
'Private m_特定ボーナス除外スコア列() As Long

'Private m_達成率列() As Double

Private m_個別ホールド可能性フレーム列() As Long
Private m_ホールド可能性フレーム列() As Long
Private m_ホールドブロック切り分け列() As Long
Private m_ホールドブロック列() As Long


Private m_切替判定テーブル列 As ListColumn
Private m_早遅手動指定テーブル列 As ListColumn

Private m_フレームずれオフセット As Long
Private m_フレームずれ別評価リスト() As String
Private m_最大ホールド接続フレームずれ As Long

Private m_難易度テーブル As ListObject
Private m_評価テーブル As ListObject
Private m_COOL点 As Long

Private m_開始ライフ量 As Long
Private m_最大ライフ量 As Long
Private m_SAFETYライフ量 As Long

Private m_難易度文字列 As String

Public Function IChartTable_オブジェクト設定(Optional ByVal テーブル名 As String = "") As Long
    
    IChartTable_オブジェクト解放
    
    ' テーブル設定 ---------------------------------------------------------------------------------
    
    On Error GoTo テーブルを発見できない場合
    
    If Def.テーブルとシートの検索設定(m_OwnTable, m_Sheet, テーブル名) < 0 Then
        GoTo テーブルを発見できない場合
    End If
    
    ' メンバ設定 -----------------------------------------------------------------------------------
    
    Dim 定数シート As Worksheet
    Dim 定数シート名 As String
    
    Set 定数シート = IChartTable_所属シート
    
    On Error Resume Next
    定数シート名 = 定数シート.Names("DataSheetName").RefersToRange.value
    
    On Error GoTo テーブル定数を発見できない場合
    
    m_最早COOLフレーム = 定数シート.Names("EarliestCoolFrame").RefersToRange.value
    m_最遅COOLフレーム = 定数シート.Names("LatestCoolFrame").RefersToRange.value
    m_最早SADフレーム = 定数シート.Names("EarliestSadFrame").RefersToRange.value
    m_最遅SADフレーム = 定数シート.Names("LatestSadFrame").RefersToRange.value
    
    On Error GoTo テーブルのフォーマット形式が異なる場合
    
    m_データ行数 = IChartTable_OwnTable.ListRows.count
    
    ReDim m_テンポチェンジ列(IChartTable_データ行数)
    
    ReDim m_フレーム列(IChartTable_データ行数)

    ReDim m_ノーツ列(Def.マーク数, IChartTable_データ行数)
    ReDim m_スライドノーツ列(スライドマーク数, IChartTable_データ行数)
    ReDim m_ノーツ個数列(IChartTable_データ行数)
    ReDim m_チェイン列(IChartTable_データ行数)
    
    ReDim m_ノーツ番号列(IChartTable_データ行数)
    ReDim m_評価列(IChartTable_データ行数)
    ReDim m_コンボ列(IChartTable_データ行数)
    ReDim m_評価位置列(IChartTable_データ行数)
    
    ReDim m_早遅手動指定列(IChartTable_データ行数)
    ReDim m_早遅フレーム手動指定列(IChartTable_データ行数)
    ReDim m_切替判定文字列(IChartTable_データ行数)
'    ReDim m_切替判定列(IChartTable_データ行数)
    
    ReDim m_早遅文字フレーム列(IChartTable_データ行数)
    ReDim m_早遅フレーム列(IChartTable_データ行数)
    ReDim m_ライフ列(IChartTable_データ行数)
    
    ReDim m_ボタン押下判定列(IChartTable_データ行数)
    ReDim m_ホールド終了押し直し判定列(IChartTable_データ行数)
    
    ReDim m_基本スコア列(IChartTable_データ行数)
    ReDim m_コンボボーナス列(IChartTable_データ行数)
    ReDim m_ライフボーナス列(IChartTable_データ行数)
    ReDim m_チェーンスライド基本スコア列(IChartTable_データ行数)
    ReDim m_チェーンスライドボーナススコア列(IChartTable_データ行数)
    ReDim m_チェーンスライドボーナス列(IChartTable_データ行数)
    
    ReDim m_ホールドフレーム列(IChartTable_データ行数)
    ReDim m_ホールド個別ボーナス列(Def.マーク数, IChartTable_データ行数)
    ReDim m_ホールドボーナス列(IChartTable_データ行数)
    
    ReDim m_ホールド個別開始フレーム列(Def.マーク数, IChartTable_データ行数)
    ReDim m_ホールド開始フレーム列(IChartTable_データ行数)
    
    ReDim m_補正点列(IChartTable_データ行数)
    
    ReDim m_スコア列(IChartTable_データ行数)
    ReDim m_ホールドスコア列(IChartTable_データ行数)
'    ReDim m_チェーンスライドスコア列(IChartTable_データ行数)
'    ReDim m_特定ボーナス除外スコア列(IChartTable_データ行数)
    
'    ReDim m_達成率列(IChartTable_データ行数)
    
    ReDim m_個別ホールド可能性フレーム列(Def.マーク数, IChartTable_データ行数)
    ReDim m_ホールド可能性フレーム列(IChartTable_データ行数)
    ReDim m_ホールドブロック切り分け列(IChartTable_データ行数)
    ReDim m_ホールドブロック列(IChartTable_データ行数)
    
    Dim 行 As Long
    Dim マーク As Long
    
    For 行 = 1 To IChartTable_データ行数
        
        On Error Resume Next
        m_テンポチェンジ列(行) = IChartTable_OwnTable.ListColumns("change").DataBodyRange(行).value
        On Error GoTo テーブルのフォーマット形式が異なる場合
        
        m_フレーム列(行) = IChartTable_OwnTable.ListColumns("フレーム").DataBodyRange(行).value
        For マーク = 1 To Def.マーク数
            m_ノーツ列(マーク, 行) = IChartTable_OwnTable.ListColumns(Def.マーク文字(マーク)).DataBodyRange(行).value
        Next
        For マーク = 1 To スライドマーク数
            m_スライドノーツ列(マーク, 行) = IChartTable_OwnTable.ListColumns(Def.スライドマーク文字(マーク)).DataBodyRange(行).value
        Next
        
        m_評価列(行) = IChartTable_OwnTable.ListColumns("評価").DataBodyRange(行).value
        m_早遅手動指定列(行) = IChartTable_OwnTable.ListColumns("el").DataBodyRange(行).value
        m_早遅フレーム手動指定列(行) = IChartTable_OwnTable.ListColumns("補el").DataBodyRange(行).value
        m_切替判定文字列(行) = IChartTable_OwnTable.ListColumns("切替").DataBodyRange(行).value
        
        On Error Resume Next
        m_補正点列(行) = IChartTable_OwnTable.ListColumns("補正点").DataBodyRange(行).value
        On Error GoTo テーブルのフォーマット形式が異なる場合
        
    Next
    
    Set m_切替判定テーブル列 = IChartTable_OwnTable.ListColumns("切替")
    Set m_早遅手動指定テーブル列 = IChartTable_OwnTable.ListColumns("el")
    
    On Error GoTo テーブル定数を発見できない場合
    
    m_フレームずれオフセット = 定数シート.Names("FrameGapList").RefersToRange.Cells(1).value
    ReDim m_フレームずれ別評価リスト(定数シート.Names("FrameGapRatingList").RefersToRange.Rows.count)
    
    Dim 評価リスト行 As Long
    For 評価リスト行 = 1 To UBound(m_フレームずれ別評価リスト)
        m_フレームずれ別評価リスト(評価リスト行) = 定数シート.Names("FrameGapRatingList").RefersToRange.Cells(1).Offset(評価リスト行 - 1, 0).value
    Next 評価リスト行
    
    m_最大ホールド接続フレームずれ = IChartTable_最遅SADフレーム - IChartTable_最早SADフレーム + 1
    
    m_COOL点 = ThisWorkbook.Names("CoolScore").RefersToRange.value
    Set m_難易度テーブル = ThisWorkbook.Worksheets("Difficulty").ListObjects("DifficultyTable")
    Set m_評価テーブル = ThisWorkbook.Worksheets("Rating").ListObjects("RatingTable")
    
    m_開始ライフ量 = ThisWorkbook.Names("DefaultLife").RefersToRange.value
    m_最大ライフ量 = ThisWorkbook.Names("MaximumLife").RefersToRange.value
    m_SAFETYライフ量 = ThisWorkbook.Names("SafetyLife").RefersToRange.value
    
    m_難易度文字列 = IChartTable_所属シート.Names("Difficulty").RefersToRange.value
    
    IChartTable_再計算 1, IChartTable_データ行数
    IChartTable_リザルト再計算
        
    IChartTable_オブジェクト設定 = 0
    Exit Function
    
    ' エラーラベル ---------------------------------------------------------------------------------
    
テーブルを発見できない場合:
    IChartTable_オブジェクト設定 = -1
    GoTo エラー処理
    
テーブルのフォーマット形式が異なる場合:
    IChartTable_オブジェクト設定 = -2
    GoTo エラー処理
    
テーブル定数を発見できない場合:
    Dim ws As Worksheet
    If 定数シート.name = 定数シート名 Then
        GoTo テーブル定数を設定できない場合
    End If
    For Each ws In Worksheets
        If ws.name = 定数シート名 Then
            Set 定数シート = ws
            Resume
        End If
    Next ws
    GoTo テーブル定数を設定できない場合
    
テーブル定数を設定できない場合:
    IChartTable_オブジェクト設定 = -3
    GoTo エラー処理
    
エラー処理:
    IChartTable_オブジェクト解放
    Exit Function
    
End Function

Public Sub IChartTable_オブジェクト解放()
    
    Set m_OwnTable = Nothing
    Set m_Sheet = Nothing
    
    m_最早COOLフレーム = 0
    m_最遅COOLフレーム = 0
    m_最早SADフレーム = 0
    m_最遅SADフレーム = 0
    
    m_データ行数 = 0
    
    Erase m_テンポチェンジ列
    
    Erase m_フレーム列

    Erase m_ノーツ列
    Erase m_スライドノーツ列
    Erase m_ノーツ個数列
    Erase m_チェイン列
    
    Erase m_ノーツ番号列
    Erase m_評価列
    Erase m_コンボ列
    Erase m_評価位置列
    
    Erase m_早遅手動指定列
    Erase m_早遅フレーム手動指定列
    Erase m_切替判定文字列
'    Erase m_切替判定列
    
    Erase m_早遅文字フレーム列
    Erase m_早遅フレーム列
    Erase m_ライフ列
    
    Erase m_ボタン押下判定列
    Erase m_ホールド終了押し直し判定列
    
    Erase m_基本スコア列
    Erase m_コンボボーナス列
    Erase m_ライフボーナス列
    Erase m_チェーンスライド基本スコア列
    Erase m_チェーンスライドボーナススコア列
    Erase m_チェーンスライドボーナス列
    
    Erase m_ホールドフレーム列
    Erase m_ホールド個別ボーナス列
    Erase m_ホールドボーナス列
    
    Erase m_ホールド個別開始フレーム列
    Erase m_ホールド開始フレーム列
    
    Erase m_補正点列
    
    Erase m_スコア列
    Erase m_ホールドスコア列
'    Erase m_チェーンスライドスコア列
'    Erase m_特定ボーナス除外スコア列
    
'    Erase m_達成率列
    
    Erase m_個別ホールド可能性フレーム列
    Erase m_ホールド可能性フレーム列
    Erase m_ホールドブロック切り分け列
    Erase m_ホールドブロック列
    
    Set m_切替判定テーブル列 = Nothing
    Set m_早遅手動指定テーブル列 = Nothing
    
    m_フレームずれオフセット = 0
    Erase m_フレームずれ別評価リスト
    m_最大ホールド接続フレームずれ = 0
    
    Set m_難易度テーブル = Nothing
    Set m_評価テーブル = Nothing
    m_COOL点 = 0
    
    m_開始ライフ量 = 0
    m_最大ライフ量 = 0
    m_SAFETYライフ量 = 0
    
    m_難易度文字列 = ""
    
End Sub

Public Function IChartTable_再計算(ByVal 開始行 As Long, Optional ByVal 終了行 As Long = -1, Optional ByVal 自動再計算 As Boolean = False)
    
    If 終了行 = -1 Then
        終了行 = 開始行
    End If
    
    If (Not 自動再計算) And 開始行 <= 終了行 Then
        
        #If 再計算ログ Then
            Dim 開始時間 As Double
            Dim 終了時間 As Double
            開始時間 = Def.GetTickCount
        #End If
        
        Dim 行 As Long
        Dim マーク As Long
        Dim マーク2 As Long
        
        For 行 = 開始行 To 終了行
            
            ' {=SUM((テーブル[@[△]:[→]]<>"")*1)}
            m_ノーツ個数列(行) = 0
            For マーク = 1 To Def.マーク数
                If m_ノーツ列(マーク, 行) <> "" Then
                    m_ノーツ個数列(行) = m_ノーツ個数列(行) + 1
                End If
            Next マーク
            For マーク = 1 To スライドマーク数
                If m_スライドノーツ列(マーク, 行) <> "" Then
                    m_ノーツ個数列(行) = m_ノーツ個数列(行) + 1
                End If
            Next マーク
            
            ' =SUM(テーブル[@[←]:[→]])
            m_チェイン列(行) = 0
            For マーク = 1 To スライドマーク数
                If IsNumeric(m_スライドノーツ列(マーク, 行)) Then
                    m_チェイン列(行) = m_チェイン列(行) + CLng(m_スライドノーツ列(マーク, 行))
                End If
            Next マーク
            
            ' =([@Notes]>0)+IF(ISNUMBER($AA2),$AA2,0)
            If 行 = 1 Then
                m_ノーツ番号列(行) = 0
            Else
                m_ノーツ番号列(行) = m_ノーツ番号列(行 - 1)
            End If
            If m_ノーツ個数列(行) > 0 Then
                m_ノーツ番号列(行) = m_ノーツ番号列(行) + 1
            End If
            
            ' =MATCH([@評価],RatingTable[Sign],0)
            m_評価位置列(行) = Application.WorksheetFunction.Match(m_評価列(行), m_評価テーブル.ListColumns("Sign").DataBodyRange, 0)
            
            ' =IF([@Notes]>0,IF(INDEX(RatingTable[Combo],[@評価pos]),$AC1+1,0),IF(ISNUMBER($AC1),$AC1,0))
            If 行 = 1 Then
                m_コンボ列(行) = 0
            ElseIf m_ノーツ個数列(行) > 0 Then
                If Application.WorksheetFunction.index(m_評価テーブル.ListColumns("Combo").DataBodyRange, m_評価位置列(行)) Then
                    m_コンボ列(行) = m_コンボ列(行 - 1) + 1
                Else
                    m_コンボ列(行) = 0
                End If
            Else
                m_コンボ列(行) = m_コンボ列(行 - 1)
            End If
            
            ' =IF([@el]=EarlyCoolString,EarliestCoolFrame,IF([@el]=LateCoolString,LatestCoolFrame,IF([@el]=JustCoolString,0,"")))
            If m_早遅手動指定列(行) = Def.早COOL文字 Then
                m_早遅文字フレーム列(行) = IChartTable_最早COOLフレーム
            ElseIf m_早遅手動指定列(行) = Def.遅COOL文字 Then
                m_早遅文字フレーム列(行) = IChartTable_最遅COOLフレーム
            Else
                m_早遅文字フレーム列(行) = 0
            End If
            
            ' =IF([@change],IF(ISNUMBER($AH763),$AH763,0),IF([@補el]<>"",[@補el],IF([@elFr]<>"",[@elFr],0)))
            If m_テンポチェンジ列(行) Then
                If 行 = 1 Then
                    m_早遅フレーム列(行) = 0
                Else
                    m_早遅フレーム列(行) = m_早遅フレーム列(行 - 1)
                End If
            Else
                If m_早遅フレーム手動指定列(行) <> "" Then
                    m_早遅フレーム列(行) = CLng(m_早遅フレーム手動指定列(行))
                Else
                    m_早遅フレーム列(行) = m_早遅文字フレーム列(行)
                End If
            End If
            
            ' =MAX(
            '      MIN(
            '          IF(ISNUMBER($AD763),$AD763,DefaultLife)+
            '          IF(
            '              AND([@Notes]>0,$AD763>0),
            '              INDEX(RatingTable[#データ],[@評価pos],MATCH(Difficulty,RatingTable[#見出し],0))
            '              ,0),
            '          MaximumLife
            '      ),
            '      IF(
            '         [@フレーム]+IF([@評価]=Rating!$D$12,LatestSadFrame,[@補Fr])>INDEX(DifficultyTable[SAFETY2],MATCH(Difficulty,DifficultyTable[Difficulty],0)),
            '         0,
            '         SafetyLife
            '      )
            '  )
            If 行 = 1 Then
                m_ライフ列(行) = m_開始ライフ量
            Else
                m_ライフ列(行) = m_ライフ列(行 - 1)
            End If
            If m_ノーツ個数列(行) > 0 Then
                Dim isライフ増加 As Boolean
                If 行 = 1 Then
                    isライフ増加 = True
                ElseIf m_ライフ列(行 - 1) > 0 Then
                    isライフ増加 = True
                Else
                    isライフ増加 = False
                End If
                
                If isライフ増加 Then
                    m_ライフ列(行) = m_ライフ列(行) + Application.WorksheetFunction.index( _
                        m_評価テーブル.DataBodyRange, _
                        m_評価位置列(行), _
                        Application.WorksheetFunction.Match(m_難易度文字列, m_評価テーブル.HeaderRowRange, 0))
                End If
            End If
            m_ライフ列(行) = Application.WorksheetFunction.Min(m_ライフ列(行), m_最大ライフ量)
            
            Dim 判別フレーム As Long
            If m_評価列(行) = Def.WORST文字 Then
                判別フレーム = m_フレーム列(行) + IChartTable_最遅SADフレーム
            Else
                判別フレーム = m_フレーム列(行) + m_早遅フレーム列(行)
            End If
            If 判別フレーム > Application.WorksheetFunction.index( _
                m_難易度テーブル.ListColumns("SAFETY2").DataBodyRange, _
                Application.WorksheetFunction.Match(m_難易度文字列, m_難易度テーブル.ListColumns("Difficulty").DataBodyRange, 0)) Then
                m_ライフ列(行) = Application.WorksheetFunction.Max(m_ライフ列(行), 0)
            Else
                m_ライフ列(行) = Application.WorksheetFunction.Max(m_ライフ列(行), m_SAFETYライフ量)
            End If
            
            ' =INDEX(RatingTable[Pull],[@評価pos])
            m_ボタン押下判定列(行) = Application.WorksheetFunction.index(m_評価テーブル.ListColumns("Pull").DataBodyRange, m_評価位置列(行), 0)
            
            ' =OR(($AX757:$BA757>0)*(ナイトメア_パーティーナイト_EXEX[@[△]:[○]]<>""))
            m_ホールド終了押し直し判定列(行) = False
            For マーク = 1 To Def.マーク数
                If m_ノーツ列(マーク, 行) <> "" Then
                    If 行 = 1 Then
                        m_ホールド終了押し直し判定列(行) = True
                        Exit For
                    ElseIf m_ホールド個別開始フレーム列(マーク, 行 - 1) > 0 Then
                        m_ホールド終了押し直し判定列(行) = True
                        Exit For
                    End If
                End If
            Next マーク
            
            Dim isコンボ継続 As Boolean
            isコンボ継続 = (Application.WorksheetFunction.index(m_評価テーブル.ListColumns("Combo").DataBodyRange, m_評価位置列(行)) = True)
            
            ' =IFERROR(INDEX(RatingTable[Score],[@評価pos])*IF(INDEX(RatingTable[Combo],[@評価pos]),[@Notes],[@Notes]>0),0)
            m_基本スコア列(行) = Application.WorksheetFunction.index(m_評価テーブル.ListColumns("Score").DataBodyRange, m_評価位置列(行))
            If isコンボ継続 Then
                m_基本スコア列(行) = m_基本スコア列(行) * m_ノーツ個数列(行)
            Else
                If m_ノーツ個数列(行) = 0 Then
                    m_基本スコア列(行) = 0
                End If
            End If
            
            ' =IF([@Notes]>0,MIN(INT([@Combo]/10),5)*50,0)
            If m_ノーツ個数列(行) > 0 Then
                m_コンボボーナス列(行) = Application.WorksheetFunction.Min(m_コンボ列(行) \ 10, 5) * 50
            Else
                m_コンボボーナス列(行) = 0
            End If
            
            ' =IF([@Notes]>0,IF(INDEX(RatingTable[Combo],[@評価pos]),(IF([@Life]=MaximumLife,MAX([@chains],0),0)+IF(AND([@評価]=Rating!$D$4,$AJ769=MaximumLife),1,0))*10,0),0)
            m_ライフボーナス列(行) = 0
            If m_ノーツ個数列(行) > 0 Then
                If isコンボ継続 Then
                    If m_ライフ列(行) = m_最大ライフ量 Then
                        m_ライフボーナス列(行) = m_チェイン列(行)
                    End If
                    If m_評価列(行) = IChartTable_フレームずれ別評価(0) And m_ライフ列(行 - 1) = m_最大ライフ量 Then
                        m_ライフボーナス列(行) = m_ライフボーナス列(行) + 1
                    End If
                    m_ライフボーナス列(行) = m_ライフボーナス列(行) * 10
                End If
            End If
            
            ' =[@chains]*500
            m_チェーンスライド基本スコア列(行) = m_チェイン列(行) * 500
            
            ' =IF([@chains]>0,[@chains]*([@chains]+1)*5+[@pull]*1000,0)
            If m_チェイン列(行) > 0 Then
                m_チェーンスライドボーナススコア列(行) = m_チェイン列(行) * (m_チェイン列(行) + 1) * 5
                If m_ボタン押下判定列(行) Then
                    m_チェーンスライドボーナススコア列(行) = m_チェーンスライドボーナススコア列(行) + 1000
                End If
            Else
                m_チェーンスライドボーナススコア列(行) = 0
            End If
            
            ' =[@csScore]+[@csb]
            m_チェーンスライドボーナス列(行) = m_チェーンスライド基本スコア列(行) + m_チェーンスライドボーナススコア列(行)
            
            ' =([@フレーム]+[@補Fr])-IFERROR($BC763+AND([@pull],[@repull])+1,0)
            m_ホールドフレーム列(行) = m_フレーム列(行) + m_早遅フレーム列(行)
            If 行 > 1 Then
                m_ホールドフレーム列(行) = m_ホールドフレーム列(行) - m_ホールド開始フレーム列(行 - 1) - 1
                If m_ボタン押下判定列(行) And m_ホールド終了押し直し判定列(行) Then
                    m_ホールドフレーム列(行) = m_ホールドフレーム列(行) - 1
                End If
            End If
            
            ' {=IF(IFERROR($AY763>0,FALSE)
            '      ,IF([@holdFr]>300
            '          ,INT($BC763+300-$AY763)*10+1500
            '          ,IF(OR(AND([@pull],[@repull]),[@切替],AND(OR(ナイトメア_パーティーナイト_EXEX[@[△]:[○]]=HoldMarker),[@holdFr]=300))
            '              ,($BC763+[@holdFr]-$AY763)*10
            '              ,0
            '           )
            '       )
            '      ,0
            '   )
            ' }
            For マーク = 1 To Def.マーク数
                m_ホールド個別ボーナス列(マーク, 行) = 0
                If 行 > 1 Then
                    If m_ホールド個別開始フレーム列(マーク, 行 - 1) > 0 Then
                        Dim isホールド確定 As Boolean
                        isホールド確定 = False
                        
                        If m_ホールドフレーム列(行) > 300 Then
                            m_ホールド個別ボーナス列(マーク, 行) = 1500
                            isホールド確定 = True
                        ElseIf m_ボタン押下判定列(行) And m_ホールド終了押し直し判定列(行) Then
                            isホールド確定 = True
                        ElseIf IChartTable_切替判定列(行) Then
                            isホールド確定 = True
                        Else
                            If m_ホールドフレーム列(行) = 300 Then
                                For マーク2 = 1 To Def.マーク数
                                    If m_ノーツ列(マーク2, 行) = Def.HOLD文字 Then
                                        isホールド確定 = True
                                        Exit For
                                    End If
                                Next マーク2
                            End If
                        End If
                        
                        If isホールド確定 Then
                            m_ホールド個別ボーナス列(マーク, 行) = m_ホールド個別ボーナス列(マーク, 行) + _
                                (m_ホールド開始フレーム列(行 - 1) + Application.WorksheetFunction.Min(m_ホールドフレーム列(行), 300) - m_ホールド個別開始フレーム列(マーク, 行 - 1)) * 10
                        End If
                    End If
                End If
            Next マーク
            
            ' =SUM(ナイトメア_パーティーナイト_EXEX[@[△HBonus]:[○HBonus]])
            m_ホールドボーナス列(行) = 0
            For マーク = 1 To Def.マーク数
                m_ホールドボーナス列(行) = m_ホールドボーナス列(行) + m_ホールド個別ボーナス列(マーク, 行)
            Next マーク
            
            ' =IF(AND([@△]=HoldMarker,[@pull]),[@フレーム]+[@補Fr],IF([@△HBonus]>0,0,IF(ISNUMBER($AY757),$AY757,0)))
            For マーク = 1 To Def.マーク数
                If m_ノーツ列(マーク, 行) = Def.HOLD文字 And m_ボタン押下判定列(行) Then
                    m_ホールド個別開始フレーム列(マーク, 行) = m_フレーム列(行) + m_早遅フレーム列(行)
                Else
                    If m_ホールド個別ボーナス列(マーク, 行) > 0 Then
                        m_ホールド個別開始フレーム列(マーク, 行) = 0
                    ElseIf 行 = 1 Then
                        m_ホールド個別開始フレーム列(マーク, 行) = 0
                    Else
                        m_ホールド個別開始フレーム列(マーク, 行) = m_ホールド個別開始フレーム列(マーク, 行 - 1)
                    End If
                End If
            Next マーク
            
            ' =MAX(ナイトメア_パーティーナイト_EXEX[@[△H_fr]:[○H_fr]])
            m_ホールド開始フレーム列(行) = 0
            For マーク = 1 To Def.マーク数
                If m_ホールド個別開始フレーム列(マーク, 行) > m_ホールド開始フレーム列(行) Then
                    m_ホールド開始フレーム列(行) = m_ホールド個別開始フレーム列(マーク, 行)
                End If
            Next マーク
            
            ' =SUM(ナイトメア_パーティーナイト_EXEX[@[Score]:[LBonus]],[@CSBonus],[@HBonus],[@補正点])+IF(ISNUMBER($BE748),$BE748,0)
            If 行 = 1 Then
                m_スコア列(行) = 0
            Else
                m_スコア列(行) = m_スコア列(行 - 1)
            End If
            m_スコア列(行) = m_スコア列(行) + m_基本スコア列(行) + m_コンボボーナス列(行) + m_ライフボーナス列(行) + _
                m_チェーンスライドボーナス列(行) + m_ホールドボーナス列(行) + m_補正点列(行)
            
            ' =[@HBonus]+IF(ISNUMBER($BF5),$BF5,0)
            If 行 = 1 Then
                m_ホールドスコア列(行) = 0
            Else
                m_ホールドスコア列(行) = m_ホールドスコア列(行 - 1)
            End If
            m_ホールドスコア列(行) = m_ホールドスコア列(行) + m_ホールドボーナス列(行)
            
            ' {=IF(AND([@△]=HoldMarker,[@pull])
            '      ,[@フレーム]+[@補Fr]
            '      ,IF(ISNUMBER($BJ752)
            '         ,IF(OR(AND(MAX($BJ752:$BM752*(ナイトメア_パーティーナイト_EXEX[@[△]:[○]]<>""))>=$BJ752,[@pull])
            '                ,[@フレーム]+[@補Fr]>$BN752+301
            '                ,[@切替])
            '            ,0,
            '            IF(ISNUMBER($BJ752),$BJ752,0))
            '         ,0
            '       )
            '   )
            ' }
            For マーク = 1 To Def.マーク数
                If m_ノーツ列(マーク, 行) = Def.HOLD文字 And m_ボタン押下判定列(行) Then
                    m_個別ホールド可能性フレーム列(マーク, 行) = m_フレーム列(行) + m_早遅フレーム列(行)
                Else
                    If 行 = 1 Then
                        m_個別ホールド可能性フレーム列(マーク, 行) = 0
                    ElseIf m_フレーム列(行) + m_早遅フレーム列(行) > m_ホールド可能性フレーム列(行 - 1) + Def.未MAXフレーム最大値 + 1 Then
                        m_個別ホールド可能性フレーム列(マーク, 行) = 0
                    ElseIf IChartTable_切替判定列(行) Then
                        m_個別ホールド可能性フレーム列(マーク, 行) = 0
                    Else
                        Dim ホールド続行フレーム As Long
                        ホールド続行フレーム = 0
                        For マーク2 = 1 To Def.マーク数
                            If m_ノーツ列(マーク2, 行) <> "" And m_個別ホールド可能性フレーム列(マーク2, 行 - 1) > ホールド続行フレーム Then
                                ホールド続行フレーム = m_個別ホールド可能性フレーム列(マーク2, 行 - 1)
                            End If
                        Next マーク2
                        
                        If m_ボタン押下判定列(行) And ホールド続行フレーム >= m_個別ホールド可能性フレーム列(マーク, 行 - 1) Then
                            m_個別ホールド可能性フレーム列(マーク, 行) = 0
                        Else
                            m_個別ホールド可能性フレーム列(マーク, 行) = m_個別ホールド可能性フレーム列(マーク, 行 - 1)
                        End If
                    End If
                End If
            Next マーク
            
            ' =MAX(ナイトメア_パーティーナイト_EXEX[@[△Holdable]:[○Holdable]])
            m_ホールド可能性フレーム列(行) = 0
            For マーク = 1 To Def.マーク数
                If m_個別ホールド可能性フレーム列(マーク, 行) > m_ホールド可能性フレーム列(行) Then
                    m_ホールド可能性フレーム列(行) = m_個別ホールド可能性フレーム列(マーク, 行)
                End If
            Next マーク
            
            ' {=IF(ISNUMBER($BO2),$BO2,0)+
            '   IF([@holdable]=0
            '      ,0
            '      ,IF(OR((ナイトメア_パーティーナイト_EXEX[@[△Holdable]:[○Holdable]]>0)*(ナイトメア_パーティーナイト_EXEX[@[△Holdable]:[○Holdable]]<[@フレーム]+[@補Fr]),[@change])
            '        ,0
            '        ,1
            '       )
            '   )
            ' }
            If 行 = 1 Then
                m_ホールドブロック切り分け列(行) = 0
            Else
                m_ホールドブロック切り分け列(行) = m_ホールドブロック切り分け列(行 - 1)
            End If
            
            Dim ブロックインクリメント As Boolean
            ブロックインクリメント = True
            If m_ホールド可能性フレーム列(行) = 0 Then
                ブロックインクリメント = False
            ElseIf m_テンポチェンジ列(行) Then
                ブロックインクリメント = False
            Else
                For マーク = 1 To Def.マーク数
                    If m_個別ホールド可能性フレーム列(マーク, 行) > 0 And m_個別ホールド可能性フレーム列(マーク, 行) < m_フレーム列(行) + m_早遅フレーム列(行) Then
                        ブロックインクリメント = False
                        Exit For
                    End If
                Next マーク
            End If
            
            If ブロックインクリメント Then
                m_ホールドブロック切り分け列(行) = m_ホールドブロック切り分け列(行) + 1
            End If
            
            ' =IF([@holdable]>0,[@block],0)
            If m_ホールド可能性フレーム列(行) > 0 Then
                m_ホールドブロック列(行) = m_ホールドブロック切り分け列(行)
            Else
                m_ホールドブロック列(行) = 0
            End If
            
        Next 行
        
        #If 再計算ログ Then
            終了時間 = Def.GetTickCount
            Debug.Print IChartTable_OwnTable.name & vbTab & 開始行 & vbTab & 終了行 & vbTab & _
                "再計算実行" & vbTab & 終了時間 - 開始時間 & "ミリ秒"
        #End If
        
    End If
    
End Function

Public Function IChartTable_リザルト再計算() As Def.結果セット
    
    Dim ノーツ数 As Long
    
    Dim 計算ノーツ基本点 As Long
    Dim 計算コンボボーナス点 As Long
    Dim 計算チェーンスライド基本点 As Long
    
    Dim 実測ノーツ基本点 As Long
    Dim 実測コンボボーナス点 As Long
    Dim 実測チェーンスライド基本点 As Long
    
    Dim 実測ライフボーナス点 As Long
    Dim 実測ホールド関連ボーナス点 As Long
    
    Dim 実測チェーンスライドボーナス点 As Long
    
    Dim 達成率基準点 As Long
    Dim ホールド達成率係数 As Double
    
    Dim 最大コンボ数 As Long
    
    Dim リザルト As Def.結果セット
    
    Dim 行 As Long
    
    ' {=SUM((テーブル[Notes]>0)*1)}
    ノーツ数 = 0
    For 行 = 1 To IChartTable_データ行数
        If m_ノーツ個数列(行) > 0 Then
            ノーツ数 = ノーツ数 + 1
        End If
    Next
    
    ' =SUM(テーブル[Notes])*Rating!$E$4
    計算ノーツ基本点 = Application.WorksheetFunction.Sum(m_ノーツ個数列) * m_COOL点
    ' {=SUM((ComboBonusTable[Combo]<=$AJ$814)*ComboBonusTable[Bonus])}
    計算コンボボーナス点 = _
        Application.WorksheetFunction.Max(Application.WorksheetFunction.Min(ノーツ数 - 9, 10), 0) * 50 + _
        Application.WorksheetFunction.Max(Application.WorksheetFunction.Min(ノーツ数 - 19, 10), 0) * 100 + _
        Application.WorksheetFunction.Max(Application.WorksheetFunction.Min(ノーツ数 - 29, 10), 0) * 150 + _
        Application.WorksheetFunction.Max(Application.WorksheetFunction.Min(ノーツ数 - 39, 10), 0) * 200 + _
        Application.WorksheetFunction.Max(ノーツ数 - 49, 0) * 250
    ' {=SUM((テーブル[chains]>0)*テーブル[chains]*500)}
    計算チェーンスライド基本点 = Application.WorksheetFunction.Sum(m_チェイン列) * 500
    
    ' =SUM(テーブル[Score])
    ' {=SUM((テーブル[Notes]>0)*テーブル[Score])}
    実測ノーツ基本点 = Application.WorksheetFunction.Sum(m_基本スコア列)
    ' =SUM(テーブル[CBonus])
    実測コンボボーナス点 = Application.WorksheetFunction.Sum(m_コンボボーナス列)
    ' =SUM(テーブル[csScore])
    実測チェーンスライド基本点 = Application.WorksheetFunction.Sum(m_チェーンスライド基本スコア列)
    
    ' =SUM(テーブル[LBonus])
    実測ライフボーナス点 = Application.WorksheetFunction.Sum(m_ライフボーナス列)
    ' =SUM(テーブル[HBonus])
    実測ホールド関連ボーナス点 = Application.WorksheetFunction.Sum(m_ホールドボーナス列)
    
    ' =SUM(テーブル[csb])
    実測チェーンスライドボーナス点 = Application.WorksheetFunction.Sum(m_チェーンスライドボーナススコア列)
    
    達成率基準点 = 計算ノーツ基本点 + 計算コンボボーナス点 + 計算チェーンスライド基本点
    ' =INDEX(DifficultyTable[ホールド達成率係数],MATCH(Difficulty,DifficultyTable[Difficulty],0))
    ホールド達成率係数 = Application.WorksheetFunction.index(m_難易度テーブル.ListColumns("ホールド達成率係数").DataBodyRange, _
        Application.WorksheetFunction.Match(m_難易度文字列, m_難易度テーブル.ListColumns("Difficulty").DataBodyRange, 0))
    
    リザルト.達成率 = Application.WorksheetFunction.RoundDown( _
        CDbl(実測ノーツ基本点 + 実測コンボボーナス点 + 実測チェーンスライド基本点 + _
        実測ライフボーナス点 + _
        Application.WorksheetFunction.Min(達成率基準点 * 0.05, 実測ホールド関連ボーナス点 * ホールド達成率係数)) _
        / 達成率基準点, 4)
    リザルト.スコア = 実測ノーツ基本点 + 実測コンボボーナス点 + 実測チェーンスライド基本点 + _
        実測ライフボーナス点 + 実測ホールド関連ボーナス点 + 実測チェーンスライドボーナス点
    
    最大コンボ数 = Application.WorksheetFunction.Max(m_コンボ列)
    
    If Application.WorksheetFunction.Min(m_ライフ列) = 0 Then
        リザルト.クリアランク = Def.MISSTAKE文字
    Else
        If ノーツ数 = 最大コンボ数 And リザルト.達成率 >= Application.WorksheetFunction.index(m_難易度テーブル.ListColumns("PERFECT").DataBodyRange, _
            Application.WorksheetFunction.Match(m_難易度文字列, m_難易度テーブル.ListColumns("Difficulty").DataBodyRange, 0)) Then
            リザルト.クリアランク = m_難易度テーブル.ListColumns("PERFECT").name
        Else
            ' =INDEX(DifficultyTable[[#見出し],[MISS×TAKE]:[EXCELLENT]],MATCH(MeasuredAchievementRate,INDEX(DifficultyTable[[MISS×TAKE]:[EXCELLENT]],MATCH(Difficulty,DifficultyTable[Difficulty],0),0),1))
            リザルト.クリアランク = m_難易度テーブル.ListColumns(m_難易度テーブル.ListColumns(Def.MISSTAKE文字).index - 1 + _
                Application.WorksheetFunction.Match( _
                    リザルト.達成率, _
                    Application.WorksheetFunction.index( _
                        Range(m_難易度テーブル.ListColumns(Def.MISSTAKE文字).DataBodyRange, m_難易度テーブル.ListColumns("EXCELLENT").DataBodyRange), _
                        Application.WorksheetFunction.Match( _
                            m_難易度文字列, m_難易度テーブル.ListColumns("Difficulty").DataBodyRange, _
                            0), _
                        0), _
                    1) _
                ).name
        End If
    End If
    
    IChartTable_リザルト再計算 = リザルト
    
End Function

Public Function IChartTable_評価リスト取得( _
    ByVal 余裕フレーム As Long, _
    Optional ByVal 開始早ずれ許容フレーム As Long = -Def.未MAXフレーム最大値, _
    Optional ByVal 開始遅ずれ許容フレーム As Long = Def.未MAXフレーム最大値, _
    Optional ByVal 終了早ずれ許容フレーム As Long = -Def.未MAXフレーム最大値, _
    Optional ByVal 終了遅ずれ許容フレーム As Long = Def.未MAXフレーム最大値) _
    As Def.評価セット()
    
    If 開始早ずれ許容フレーム < IChartTable_デフォルト早ずれ許容フレーム Then
        開始早ずれ許容フレーム = IChartTable_デフォルト早ずれ許容フレーム
    End If
    If 開始遅ずれ許容フレーム > IChartTable_デフォルト遅ずれ許容フレーム Then
        開始遅ずれ許容フレーム = IChartTable_デフォルト遅ずれ許容フレーム
    End If
    If 終了早ずれ許容フレーム < IChartTable_デフォルト早ずれ許容フレーム Then
        終了早ずれ許容フレーム = IChartTable_デフォルト早ずれ許容フレーム
    End If
    If 終了遅ずれ許容フレーム > IChartTable_デフォルト遅ずれ許容フレーム Then
        終了遅ずれ許容フレーム = IChartTable_デフォルト遅ずれ許容フレーム
    End If
    
    If 開始早ずれ許容フレーム > 開始遅ずれ許容フレーム Or 終了早ずれ許容フレーム > 終了遅ずれ許容フレーム Then
        Exit Function
    End If
    If 終了遅ずれ許容フレーム - 開始早ずれ許容フレーム < 余裕フレーム Or 終了早ずれ許容フレーム - 開始遅ずれ許容フレーム > 余裕フレーム Then
        Exit Function
    End If
    
    Dim 開始フレームずれ As Long
    Dim 終了フレームずれ As Long
    
    開始フレームずれ = Application.WorksheetFunction.Min(開始遅ずれ許容フレーム, _
        Application.WorksheetFunction.Max(開始早ずれ許容フレーム, _
        (IChartTable_最遅COOLフレーム - 1) - Application.WorksheetFunction.Max(余裕フレーム, 0)))
    終了フレームずれ = Application.WorksheetFunction.Min(終了遅ずれ許容フレーム, _
        Application.WorksheetFunction.Max(終了早ずれ許容フレーム, _
        開始フレームずれ + 余裕フレーム))
    開始フレームずれ = 終了フレームずれ - 余裕フレーム
    
    Dim 評価数 As Long
    評価数 = 0
    
    Dim 評価リスト() As Def.評価セット
    Dim 評価index As Long
    Dim isリスト追加 As Boolean
    
    Dim 評価可能性リスト() As Def.評価セット
    Dim 評価可能性index As Long
    
    Do
        評価可能性リスト = IChartTable_評価可能性リスト取得(開始フレームずれ, 終了フレームずれ)
        If Not Not 評価可能性リスト Then
            For 評価可能性index = 1 To UBound(評価可能性リスト)
                isリスト追加 = True
                For 評価index = 1 To 評価数
                    If 評価リスト(評価index).開始評価 = 評価可能性リスト(評価可能性index).開始評価 And _
                        評価リスト(評価index).終了評価 = 評価可能性リスト(評価可能性index).終了評価 Then
                        isリスト追加 = False
                        If (Not 評価リスト(評価index).正常判定枠) And 評価可能性リスト(評価可能性index).正常判定枠 Then
                            評価リスト(評価index).開始フレームずれ = 評価可能性リスト(評価可能性index).開始フレームずれ
                            評価リスト(評価index).終了フレームずれ = 評価可能性リスト(評価可能性index).終了フレームずれ
                            評価リスト(評価index).正常判定枠 = 評価可能性リスト(評価可能性index).正常判定枠
                        End If
                    End If
                Next 評価index
                If isリスト追加 Then
                    評価数 = 評価数 + 1
                    ReDim Preserve 評価リスト(評価数)
                    評価リスト(評価数) = 評価可能性リスト(評価可能性index)
                End If
            Next 評価可能性index
        End If
        開始フレームずれ = 開始フレームずれ + 1
        終了フレームずれ = 終了フレームずれ + 1
    Loop Until 開始フレームずれ > 開始遅ずれ許容フレーム Or _
        終了フレームずれ > 終了遅ずれ許容フレーム Or _
        (余裕フレーム < 0 And 終了フレームずれ > IChartTable_最早COOLフレーム + 1) Or _
        (余裕フレーム >= 0 And 開始フレームずれ > IChartTable_最早COOLフレーム + 1)

    Dim 入れ替え後評価リスト() As Def.評価セット
    Dim 入れ替え後評価index As Long
    If Not Not 評価リスト Then
        ReDim 入れ替え後評価リスト(UBound(評価リスト))
        入れ替え後評価index = 0
        For 評価index = 1 To UBound(評価リスト)
            If 評価リスト(評価index).正常判定枠 Then
                入れ替え後評価index = 入れ替え後評価index + 1
                入れ替え後評価リスト(入れ替え後評価index) = 評価リスト(評価index)
            End If
        Next 評価index
        For 評価index = 1 To UBound(評価リスト)
            If Not 評価リスト(評価index).正常判定枠 Then
                入れ替え後評価index = 入れ替え後評価index + 1
                入れ替え後評価リスト(入れ替え後評価index) = 評価リスト(評価index)
            End If
        Next 評価index
    End If
    IChartTable_評価リスト取得 = 入れ替え後評価リスト
    
End Function

Private Function IChartTable_評価可能性リスト取得(ByVal 開始フレームずれ As Long, ByVal 終了フレームずれ As Long) As Def.評価セット()
    
    IChartTable_評価リスト追加 IChartTable_評価可能性リスト取得, 開始フレームずれ, 終了フレームずれ, IChartTable_フレームずれ別評価(開始フレームずれ), IChartTable_フレームずれ別評価(終了フレームずれ)
    
    If 開始フレームずれ < 0 Then
        IChartTable_評価リスト追加 IChartTable_評価可能性リスト取得, 開始フレームずれ, 終了フレームずれ, IChartTable_フレームずれ別評価(開始フレームずれ + 1), IChartTable_フレームずれ別評価(終了フレームずれ)
    ElseIf 開始フレームずれ > 0 Then
        IChartTable_評価リスト追加 IChartTable_評価可能性リスト取得, 開始フレームずれ, 終了フレームずれ, IChartTable_フレームずれ別評価(開始フレームずれ - 1), IChartTable_フレームずれ別評価(終了フレームずれ)
    End If
    
    If 終了フレームずれ < 0 Then
        IChartTable_評価リスト追加 IChartTable_評価可能性リスト取得, 開始フレームずれ, 終了フレームずれ, IChartTable_フレームずれ別評価(開始フレームずれ), IChartTable_フレームずれ別評価(終了フレームずれ + 1)
    ElseIf 終了フレームずれ > 0 Then
        IChartTable_評価リスト追加 IChartTable_評価可能性リスト取得, 開始フレームずれ, 終了フレームずれ, IChartTable_フレームずれ別評価(開始フレームずれ), IChartTable_フレームずれ別評価(終了フレームずれ - 1)
    End If
    
End Function

Private Function IChartTable_評価リスト追加(ByRef 評価リスト() As Def.評価セット, 開始フレームずれ As Long, ByVal 終了フレームずれ As Long, ByVal 開始評価 As String, 終了評価 As String)
    
    If 開始評価 = "" Or 終了評価 = "" Then
        Exit Function
    End If
    
    Dim index As Long
    If Not Not 評価リスト Then
        index = UBound(評価リスト) + 1
    Else
        index = 1
    End If
    ReDim Preserve 評価リスト(index)
    
    評価リスト(index).開始フレームずれ = 開始フレームずれ
    評価リスト(index).終了フレームずれ = 終了フレームずれ
    評価リスト(index).開始評価 = 開始評価
    評価リスト(index).終了評価 = 終了評価
    評価リスト(index).正常判定枠 = (IChartTable_フレームずれ別評価(開始フレームずれ) = 開始評価 And IChartTable_フレームずれ別評価(終了フレームずれ) = 終了評価)
    
End Function
' **************************************************************************************************
' テーブル自身
' **************************************************************************************************

Public Property Get IChartTable_OwnTable() As ListObject
    Set IChartTable_OwnTable = m_OwnTable
End Property

Public Property Get IChartTable_所属シート() As Worksheet
    Set IChartTable_所属シート = m_Sheet
End Property

' **************************************************************************************************
' テーブル定数
' **************************************************************************************************

Public Property Get IChartTable_データ行数() As Long
    IChartTable_データ行数 = m_データ行数
End Property

Public Property Get IChartTable_最早COOLフレーム() As Double
    IChartTable_最早COOLフレーム = m_最早COOLフレーム
End Property

Public Property Get IChartTable_最遅COOLフレーム() As Double
    IChartTable_最遅COOLフレーム = m_最遅COOLフレーム
End Property

Public Property Get IChartTable_最早SADフレーム() As Double
    IChartTable_最早SADフレーム = m_最早SADフレーム
End Property

Public Property Get IChartTable_最遅SADフレーム() As Double
    IChartTable_最遅SADフレーム = m_最遅SADフレーム
End Property

Public Property Get IChartTable_デフォルト早ずれ許容フレーム() As Long
    IChartTable_デフォルト早ずれ許容フレーム = IChartTable_最早SADフレーム - 1
End Property

Public Property Get IChartTable_デフォルト遅ずれ許容フレーム() As Long
    IChartTable_デフォルト遅ずれ許容フレーム = IChartTable_最遅SADフレーム + 1
End Property

Public Property Get IChartTable_最低ライフ() As Long
    IChartTable_最低ライフ = Application.WorksheetFunction.Min(m_ライフ列)
End Property

' **************************************************************************************************
' 切替判定用プロパティ
' **************************************************************************************************

Public Property Get IChartTable_ホールドスコア列(ByVal 行 As Long) As Long
    IChartTable_ホールドスコア列 = m_ホールドスコア列(行)
End Property

Public Property Get IChartTable_切替判定列(ByVal 行 As Long) As Boolean
    If IChartTable_切替判定文字列(行) = "" Then
        IChartTable_切替判定列 = False
    Else
        IChartTable_切替判定列 = IChartTable_切替判定文字列(行)
    End If
End Property

Public Property Let IChartTable_切替判定列(ByVal 行 As Long, ByVal 値 As Boolean)
    Dim 文字列 As String
    If 値 Then
        文字列 = "TRUE"
    Else
        文字列 = ""
    End If
    If IChartTable_切替判定文字列(行) <> 文字列 Then
        IChartTable_切替判定文字列(行) = 文字列
        If IChartTable_Is自動再計算 Then
            IChartTable_再計算 1, IChartTable_データ行数
        End If
    End If
End Property

Public Property Get IChartTable_切替判定テーブル列() As ListColumn
    Set IChartTable_切替判定テーブル列 = m_切替判定テーブル列
End Property

Public Property Get IChartTable_切替判定文字列(ByVal 行 As Long) As String
    IChartTable_切替判定文字列 = m_切替判定文字列(行)
End Property

Public Property Let IChartTable_切替判定文字列(ByVal 行 As Long, ByVal 値 As String)
    If IChartTable_切替判定文字列(行) <> 値 Then
        m_切替判定文字列(行) = 値
        If IChartTable_Is自動再計算 Then
            IChartTable_再計算 1, IChartTable_データ行数
        End If
    End If
End Property

Public Property Get IChartTable_ホールド可能性判定列(ByVal マーク As Long, ByVal 行 As Long) As Long
    IChartTable_ホールド可能性判定列 = m_個別ホールド可能性フレーム列(マーク, 行)
End Property

Public Property Get IChartTable_ホールドブロック列(ByVal 行 As Long) As Long
    IChartTable_ホールドブロック列 = m_ホールドブロック列(行)
End Property

' **************************************************************************************************
' 切替判定・早遅判定共通プロパティ
' **************************************************************************************************

Public Property Get IChartTable_ノーツ番号列(ByVal 行 As Long) As Long
    IChartTable_ノーツ番号列 = m_ノーツ番号列(行)
End Property

Public Property Get IChartTable_コンボ列(ByVal 行 As Long) As Long
    IChartTable_コンボ列 = m_コンボ列(行)
End Property

Public Property Get IChartTable_ホールド個別開始フレーム列(ByVal マーク As Long, ByVal 行 As Long) As Long
    IChartTable_ホールド個別開始フレーム列 = m_ホールド個別開始フレーム列(マーク, 行)
End Property

' **************************************************************************************************
' 早遅判定用プロパティ
' **************************************************************************************************

Public Property Get IChartTable_早遅手動指定列(ByVal 行 As Long) As String
    IChartTable_早遅手動指定列 = m_早遅手動指定列(行)
End Property

Public Property Let IChartTable_早遅手動指定列(ByVal 行 As Long, ByVal 値 As String)
    If IChartTable_早遅手動指定列(行) <> 値 Then
        m_早遅手動指定列(行) = 値
        If IChartTable_Is自動再計算 Then
            IChartTable_再計算 1, IChartTable_データ行数
        End If
    End If
End Property

Public Property Get IChartTable_早遅手動指定テーブル列() As ListColumn
    Set IChartTable_早遅手動指定テーブル列 = m_早遅手動指定テーブル列
End Property

Public Property Get IChartTable_早遅フレーム手動指定列(ByVal 行 As Long) As String
    IChartTable_早遅フレーム手動指定列 = m_早遅フレーム手動指定列(行)
End Property

Public Property Let IChartTable_早遅フレーム手動指定列(ByVal 行 As Long, ByVal 値 As String)
    If IChartTable_早遅フレーム手動指定列(行) <> 値 Then
        m_早遅フレーム手動指定列(行) = 値
        If IChartTable_Is自動再計算 Then
            IChartTable_再計算 1, IChartTable_データ行数
        End If
    End If
End Property

Public Property Get IChartTable_早遅フレーム列(ByVal 行 As Long) As Long
    IChartTable_早遅フレーム列 = m_早遅フレーム列(行)
End Property

Public Property Get IChartTable_ホールド終了押し直し判定列(ByVal 行 As Long) As Boolean
    IChartTable_ホールド終了押し直し判定列 = m_ホールド終了押し直し判定列(行)
End Property

Public Property Get IChartTable_ホールドフレーム列(ByVal 行 As Long) As Long
    IChartTable_ホールドフレーム列 = m_ホールドフレーム列(行)
End Property

Public Property Get IChartTable_ホールドボーナス列(ByVal 行 As Long) As Long
    IChartTable_ホールドボーナス列 = m_ホールドボーナス列(行)
End Property

Public Property Get IChartTable_ホールド開始フレーム列(ByVal 行 As Long) As Long
    IChartTable_ホールド開始フレーム列 = m_ホールド開始フレーム列(行)
End Property

' **************************************************************************************************
' 結果からの早遅判定用プロパティ
' **************************************************************************************************

Public Property Get IChartTable_スコア列(ByVal 行 As Long) As Long
    IChartTable_スコア列 = m_スコア列(行)
End Property

' **************************************************************************************************
' スコアタ判定用プロパティ
' **************************************************************************************************

Public Property Get IChartTable_フレーム列(ByVal 行 As Long) As Long
    IChartTable_フレーム列 = m_フレーム列(行)
End Property

Public Property Get IChartTable_ノーツ列(ByVal マーク As Long, ByVal 行 As Long) As String
    IChartTable_ノーツ列 = m_ノーツ列(マーク, 行)
End Property

Public Property Get IChartTable_スライドノーツ列(ByVal マーク As Long, ByVal 行 As Long) As String
    IChartTable_スライドノーツ列 = m_スライドノーツ列(マーク, 行)
End Property

Public Property Get IChartTable_評価列(ByVal 行 As Long) As String
    IChartTable_評価列 = m_評価列(行)
End Property

Public Property Let IChartTable_評価列(ByVal 行 As Long, ByVal 値 As String)
    If IChartTable_評価列(行) <> 値 Then
        m_評価列(行) = 値
        If IChartTable_Is自動再計算 Then
            IChartTable_再計算 1, IChartTable_データ行数
        End If
    End If
End Property

Public Property Get IChartTable_ライフ列(ByVal 行 As Long) As Long
    IChartTable_ライフ列 = m_ライフ列(行)
End Property

Public Property Get IChartTable_フレームずれ別評価(ByVal フレームずれ As Long) As String
    On Error GoTo 範囲外
    IChartTable_フレームずれ別評価 = m_フレームずれ別評価リスト(フレームずれ - m_フレームずれオフセット + 1)
    On Error GoTo 0
    Exit Property
範囲外:
    IChartTable_フレームずれ別評価 = ""
    Exit Property
End Property

Public Property Get IChartTable_最大ホールド接続フレームずれ() As Long
    IChartTable_最大ホールド接続フレームずれ = m_最大ホールド接続フレームずれ
End Property

' **************************************************************************************************

Public Property Get IChartTable_Is自動再計算() As Boolean
    IChartTable_Is自動再計算 = m_自動再計算
End Property

Public Property Let IChartTable_Is自動再計算(ByVal 値 As Boolean)
    m_自動再計算 = 値
    If IChartTable_Is自動再計算 Then
        IChartTable_再計算 1, IChartTable_データ行数
    End If
End Property
