VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private m_出力ファイル As TextStream
Private m_出力フォーム As LogForm

Public Function ファイル出力開始(Optional ByVal ファイルパス As String = "") As Long
    
    Dim myFso As New FileSystemObject
    
    If ファイルパス = "" Then
        ファイルパス = "log_" & Format(Now, "yyyy-mm-dd-hhnnss") & ".txt"
    End If
        
    On Error GoTo ログファイルの作成に失敗した場合
    Set m_出力ファイル = myFso.CreateTextFile(ファイルパス, True)
'    Set m_出力ファイル = myFso.OpenTextFile(ファイルパス, ForWriting, True)
    On Error GoTo 0
    
    If Isファイル出力 Then
        ファイルパス = myFso.GetFile(ファイルパス).Path
        Debug.Print "ログは " & ファイルパス & " に出力されます｡"
    Else
        ファイルパス = ""
        Debug.Print "ログはファイル出力されません。"
    End If
    
    ThisWorkbook.Names("logFilePath").RefersToRange.value = ファイルパス
    
    Set myFso = Nothing
    
    ファイル出力開始 = 0
    Exit Function
    
ログファイルの作成に失敗した場合:
    
    Dim answer As Long
    answer = MsgBox("ログファイルを作成できません。" & vbCrLf & "ログのファイル出力を行わずに処理を続行しますか？", vbOKCancel + vbExclamation)
    
    If answer = vbOK Then
        Set m_出力ファイル = Nothing
        Resume Next
    End If
    
    ファイル出力開始 = -1
    Exit Function
    
End Function

Public Sub ファイル出力終了()
    If Me.Isファイル出力 Then
        m_出力ファイル.Close
        Set m_出力ファイル = Nothing
    End If
End Sub

Public Function フォーム出力開始()
    Set m_出力フォーム = New LogForm
    m_出力フォーム.Show
End Function

Public Sub フォーム出力終了()
    If Me.Isフォーム出力 Then
        Unload m_出力フォーム
        Set m_出力フォーム = Nothing
    End If
End Sub

Public Sub 出力(ByVal 出力文字列 As String, Optional ByVal isステータスバー出力 As Boolean = True, Optional ByVal 出力行番号 As Long = -1)
    Debug.Print 出力文字列
    If isステータスバー出力 Then
        Application.StatusBar = Replace(Replace(出力文字列, vbCrLf, " "), vbTab, " ")
    End If
    If Me.Isファイル出力 Then
        m_出力ファイル.WriteLine 出力文字列
    End If
    If Isフォーム出力 Then
        m_出力フォーム.出力 出力文字列, 出力行番号
    End If
End Sub

Public Sub フォーム文字列削除(ByVal 行番号 As Long)
    If Isフォーム出力 Then
        m_出力フォーム.出力 "", 行番号
    End If
End Sub

Public Property Get Isファイル出力() As Boolean
    Isファイル出力 = (Not m_出力ファイル Is Nothing)
End Property

Public Property Get Isフォーム出力() As Boolean
    Isフォーム出力 = (Not m_出力フォーム Is Nothing)
End Property

Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()
    Me.ファイル出力終了
    Me.フォーム出力終了
End Sub
