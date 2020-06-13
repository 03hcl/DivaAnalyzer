VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Switching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private m_譜面テーブル As IChartTable

Private m_切替文字列 As String
Private m_切替フラグ() As Boolean

Public Function オブジェクト設定(ByVal 譜面テーブル As IChartTable)
    
    Set m_譜面テーブル = 譜面テーブル
    
    m_切替文字列 = ""
    ReDim m_切替フラグ(m_譜面テーブル.データ行数)
    
    Dim 行 As Long
    For 行 = 1 To 譜面テーブル.データ行数
        Me.切替フラグ(行) = False
    Next 行
    
End Function

Public Function オブジェクト解放()
    
    Set m_譜面テーブル = Nothing
    
    m_切替文字列 = ""
    ReDim m_切替フラグ(LBound(m_切替フラグ))
    m_切替フラグ(LBound(m_切替フラグ)) = False
    
End Function

Public Function 切替早遅情報反映(Optional ByVal 自動再計算 As Boolean = False) As OutputString
    
    If 自動再計算 Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    
    Application.StatusBar = "切替と早遅の情報を譜面テーブルに反映しています......"
    
    Dim 行 As Long
    For 行 = 1 To 譜面テーブル.データ行数
        譜面テーブル.切替判定列(行) = Me.切替フラグ(行)
    Next
    
    譜面テーブル.再計算 1, 譜面テーブル.データ行数, 自動再計算
    
    Dim 開始行 As Long
    開始行 = 1
    Do While 譜面テーブル.スコア列(開始行) = 譜面テーブル.スコア列(1)
        開始行 = 開始行 + 1
    Loop
    
    Dim 終了行 As Long
    終了行 = 譜面テーブル.データ行数
    Do While 譜面テーブル.スコア列(終了行 - 1) = 譜面テーブル.スコア列(譜面テーブル.データ行数)
        終了行 = 終了行 - 1
    Loop
    
    Set 切替早遅情報反映 = Analyzing.指定範囲の早遅自動判定(開始行, 終了行, 自動再計算)
    
    Application.StatusBar = "切替と早遅の情報を譜面テーブルに反映しました。"
    
End Function

' **************************************************************************************************
' プロパティ
' **************************************************************************************************

Public Property Get 切替文字列() As String
    切替文字列 = m_切替文字列
End Property

Public Property Let 切替文字列(ByVal value As String)
    m_切替文字列 = value
End Property

Public Property Get 切替フラグ(ByVal index As Long) As Boolean
    切替フラグ = m_切替フラグ(index)
End Property

Public Property Let 切替フラグ(ByVal index As Long, ByVal value As Boolean)
    m_切替フラグ(index) = value
End Property


