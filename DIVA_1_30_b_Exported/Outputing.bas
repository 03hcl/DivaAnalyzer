Attribute VB_Name = "Outputing"
Option Explicit
Option Base 1

Public Sub 譜面データをテキスト形式で外部出力()
    
    Def.マーク文字設定
    Def.文字定数設定
    
    If Def.譜面テーブル設定() < 0 Then
        Exit Sub
    End If
    
    If MsgBox("現在のテーブルで処理を開始します。よろしいですか？" & vbCrLf & _
        "テーブル名: " & Def.譜面テーブル.OwnTable.name, vbOKCancel + vbInformation) <> vbOK Then
        MsgBox "処理を中止しました。", vbCritical
        Exit Sub
    End If
    
    Dim 出力用 As ProcessLog
    
    Set 出力用 = New ProcessLog
    If 出力用.ファイル出力開始("Chart_" & Def.譜面テーブル.OwnTable.name & "_" & Format(Now, "yyyy-mm-dd-hhnnss") & ".txt") < 0 Then
        MsgBox "ファイルに出力できません。"
        Exit Sub
    End If
    
    出力用.出力 "Difficulty=" & Def.譜面テーブル.所属シート.Names("Difficulty").RefersToRange.value
    出力用.出力 "Notes=" & Def.譜面テーブル.ノーツ番号列(Def.譜面テーブル.データ行数)
    出力用.出力 "Duration=" & Def.譜面テーブル.所属シート.Names("Duration").RefersToRange.value
    
    Dim フレームずれ As Long
    
    For フレームずれ = Def.譜面テーブル.最早SADフレーム To Def.譜面テーブル.最遅SADフレーム
        If Def.譜面テーブル.フレームずれ別評価(フレームずれ) <> "" Then
            出力用.出力 "FrameGapRating=" & フレームずれ & "," & Def.譜面テーブル.フレームずれ別評価(フレームずれ)
        End If
    Next
    
    Dim 現在ノーツ番号 As Long
    現在ノーツ番号 = 0
    
    Dim 行 As Long
    Dim マーク As Long
    Dim 出力譜面文字列 As String
    
    For 行 = 1 To Def.譜面テーブル.データ行数
        
        If Def.譜面テーブル.ノーツ番号列(行) > 現在ノーツ番号 Then
            現在ノーツ番号 = Def.譜面テーブル.ノーツ番号列(行)
            出力譜面文字列 = Def.譜面テーブル.フレーム列(行) & ","
            For マーク = 1 To Def.マーク数
                出力譜面文字列 = 出力譜面文字列 & Def.譜面テーブル.ノーツ列(マーク, 行) & ","
            Next マーク
            For マーク = 1 To Def.スライドマーク数
                出力譜面文字列 = 出力譜面文字列 & Def.譜面テーブル.スライドノーツ列(マーク, 行) & ","
            Next マーク
'            出力用.出力 "Note" & 現在ノーツ番号 & "=" & 出力譜面文字列
            出力用.出力 "Note=" & 出力譜面文字列
        End If
        
    Next 行
    
    出力用.ファイル出力終了
    
    MsgBox "外部出力が完了しました。"
    
    Rescue
    
End Sub

