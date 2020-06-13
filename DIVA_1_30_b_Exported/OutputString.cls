VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutputString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private m_表示文字列 As String
Private m_ログ出力用文字列 As String

Public Property Get 表示文字列() As String
    表示文字列 = m_表示文字列
End Property

Public Property Let 表示文字列(ByVal value As String)
    m_表示文字列 = value
End Property

Public Property Get ログ出力用文字列() As String
    ログ出力用文字列 = m_ログ出力用文字列
End Property

Public Property Let ログ出力用文字列(ByVal value As String)
    m_ログ出力用文字列 = value
End Property

