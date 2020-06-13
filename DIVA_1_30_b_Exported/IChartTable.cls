VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IChartTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Public Function オブジェクト設定(Optional ByVal テーブル名 As String = "") As Long
    
End Function

Public Sub オブジェクト解放()
    
End Sub

Public Function 再計算(ByVal 開始行 As Long, Optional ByVal 終了行 As Long = -1, Optional ByVal 自動再計算 As Boolean = False)
    
End Function

Public Function リザルト再計算() As Def.結果セット
    
End Function

Public Function 評価リスト取得( _
    ByVal 余裕フレーム As Long, _
    Optional ByVal 開始早ずれ許容フレーム As Long = -Def.未MAXフレーム最大値, _
    Optional ByVal 開始遅ずれ許容フレーム As Long = Def.未MAXフレーム最大値, _
    Optional ByVal 終了早ずれ許容フレーム As Long = -Def.未MAXフレーム最大値, _
    Optional ByVal 終了遅ずれ許容フレーム As Long = Def.未MAXフレーム最大値) _
    As Def.評価セット()
    
End Function

Private Function 評価可能性リスト取得(ByVal 開始フレームずれ As Long, ByVal 終了フレームずれ As Long) As Def.評価セット()
    
End Function

Private Function 評価リスト追加(ByRef 評価リスト() As Def.評価セット, 開始フレームずれ As Long, ByVal 終了フレームずれ As Long, ByVal 開始評価 As String, 終了評価 As String)
    
End Function
' **************************************************************************************************
' テーブル自身
' **************************************************************************************************

Public Property Get OwnTable() As ListObject
End Property

Public Property Get 所属シート() As Worksheet
End Property

' **************************************************************************************************
' テーブル定数
' **************************************************************************************************

Public Property Get データ行数() As Long
End Property

Public Property Get 最早COOLフレーム() As Double
End Property

Public Property Get 最遅COOLフレーム() As Double
End Property

Public Property Get 最早SADフレーム() As Double
End Property

Public Property Get 最遅SADフレーム() As Double
End Property

Public Property Get デフォルト早ずれ許容フレーム() As Long
End Property

Public Property Get デフォルト遅ずれ許容フレーム() As Long
End Property

Public Property Get 最低ライフ() As Long
End Property

' **************************************************************************************************
' 切替判定用プロパティ
' **************************************************************************************************

Public Property Get ホールドスコア列(ByVal 行 As Long) As Long
End Property

Public Property Get 切替判定列(ByVal 行 As Long) As Boolean
End Property

Public Property Let 切替判定列(ByVal 行 As Long, ByVal 値 As Boolean)
End Property

Public Property Get 切替判定テーブル列() As ListColumn
End Property

Public Property Get 切替判定文字列(ByVal 行 As Long) As String
End Property

Public Property Let 切替判定文字列(ByVal 行 As Long, ByVal 値 As String)
End Property

Public Property Get ホールド可能性判定列(ByVal マーク As Long, ByVal 行 As Long) As Long
End Property

Public Property Get ホールドブロック列(ByVal 行 As Long) As Long
End Property

' **************************************************************************************************
' 切替判定・早遅判定共通プロパティ
' **************************************************************************************************

Public Property Get ノーツ番号列(ByVal 行 As Long) As Long
End Property

Public Property Get コンボ列(ByVal 行 As Long) As Long
End Property

Public Property Get ホールド個別開始フレーム列(ByVal マーク As Long, ByVal 行 As Long) As Long
End Property

' **************************************************************************************************
' 早遅判定用プロパティ
' **************************************************************************************************

Public Property Get 早遅手動指定列(ByVal 行 As Long) As String
End Property

Public Property Let 早遅手動指定列(ByVal 行 As Long, ByVal 値 As String)
End Property

Public Property Get 早遅手動指定テーブル列() As ListColumn
End Property

Public Property Get 早遅フレーム手動指定列(ByVal 行 As Long) As String
End Property

Public Property Let 早遅フレーム手動指定列(ByVal 行 As Long, ByVal 値 As String)
End Property

Public Property Get 早遅フレーム列(ByVal 行 As Long) As Long
End Property

Public Property Get ホールド終了押し直し判定列(ByVal 行 As Long) As Boolean
End Property

Public Property Get ホールドフレーム列(ByVal 行 As Long) As Long
End Property

Public Property Get ホールドボーナス列(ByVal 行 As Long) As Long
End Property

Public Property Get ホールド開始フレーム列(ByVal 行 As Long) As Long
End Property

' **************************************************************************************************
' 結果からの早遅判定用プロパティ
' **************************************************************************************************

Public Property Get スコア列(ByVal 行 As Long) As Long
End Property

' **************************************************************************************************
' スコアタ判定用プロパティ
' **************************************************************************************************

Public Property Get フレーム列(ByVal 行 As Long) As Long
End Property

Public Property Get ノーツ列(ByVal マーク As Long, ByVal 行 As Long) As String
End Property

Public Property Get スライドノーツ列(ByVal マーク As Long, ByVal 行 As Long) As String
End Property

Public Property Get 評価列(ByVal 行 As Long) As String
End Property

Public Property Let 評価列(ByVal 行 As Long, ByVal 値 As String)
End Property

Public Property Get ライフ列(ByVal 行 As Long) As Long
End Property

Public Property Get フレームずれ別評価(ByVal フレームずれ As Long) As String
End Property

Public Property Get 最大ホールド接続フレームずれ() As Long
End Property

Public Property Get Is自動再計算() As Boolean
End Property

Public Property Let Is自動再計算(ByVal 値 As Boolean)
End Property
