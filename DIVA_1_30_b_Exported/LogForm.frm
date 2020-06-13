VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LogForm 
   Caption         =   "ログ情報 (×ボタンで一時停止)"
   ClientHeight    =   7236
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   12828
   OleObjectBlob   =   "LogForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "LogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'Public 親出力ログ As OutputString
'Private ログラベル() As Control
Private ログラベル数 As Long

Public Sub 出力(ByVal 出力文字列 As String, Optional ByVal 行番号 As Long = -1)
    ログラベル作成 行番号
    If 行番号 > 0 Then
        Me.Controls(ログラベル名取得(行番号)).Caption = 出力文字列
        DoEvents
    End If
End Sub

Public Sub フォーム出力終了()
    
End Sub

Private Function ログラベル作成(ByVal 行番号 As Long)
    Dim 行 As Long
    For 行 = ログラベル数 + 1 To 行番号
'    For 行 = UBound(ログラベル) + 1 To 行番号
'        ReDim ログラベル(行)
        ログラベル初期化 行
    Next 行
End Function

Private Function ログラベル初期化(ByVal 行番号 As Long)
    
    Dim 現在コントロール As Control
'    If ログラベル(行番号) Is Nothing Then
'        Set ログラベル(行番号) = Me.Controls.Add("Forms.Label.1", ログラベル名取得(行番号), False)
'    End If
    If ログラベル数 < 行番号 Then
        ログラベル数 = 行番号
        Set 現在コントロール = Me.Controls.Add("Forms.Label.1", ログラベル名取得(行番号), False)
    End If
    
'    ログラベル(行番号).Width = 600
'    ログラベル(行番号).Height = 12
'    ログラベル(行番号).Left = 0
'    ログラベル(行番号).Top = ログラベル(行番号).Height * (行番号 - 1)
'
''    ログラベル(行番号).Font.Name = "Migu 1M"
'    ログラベル(行番号).Font.Name = ThisWorkbook.Theme.ThemeFontScheme.MajorFont(msoThemeEastAsian).Name
'    ログラベル(行番号).Font.Size = ログラベル(行番号).Height
'
'    ログラベル(行番号).BackColor = RGB(255, 255, 221)
'
'    ログラベル(行番号).Visible = True
    
    現在コントロール.Width = 600
    現在コントロール.Height = 12
    現在コントロール.Left = 0
    現在コントロール.Top = 現在コントロール.Height * (行番号 - 1)
    
    現在コントロール.Font.name = ThisWorkbook.Theme.ThemeFontScheme.MajorFont(msoThemeEastAsian).name
    現在コントロール.Font.Size = 現在コントロール.Height
    現在コントロール.Caption = ""
    
    現在コントロール.BackColor = RGB(255, 255, 221)
    
    現在コントロール.Visible = True
    
End Function

Private Function ログラベル名取得(ByVal 行番号 As Long) As String
    ログラベル名取得 = "LogLabel" & 行番号
End Function

Private Sub UserForm_Initialize()
    ログラベル数 = 0
'    ReDim ログラベル(1)
'    ログラベル初期化 1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        MsgBox "一時停止しています。OKを押すと再開します。"
        Cancel = 1
    End If
End Sub

Private Sub UserForm_Terminate()
'    Set 親出力ログ = Nothing
'    Erase ログラベル
End Sub
