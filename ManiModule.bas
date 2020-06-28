Attribute VB_Name = "MainModule"
Option Explicit

Public gstrPrinterName As String 'プリンタ名
Public gstrCmdDefault As String 'HPGL変換プログラム名のデフォルト

Sub Main()

    '2重起動をチェック
    If App.PrevInstance Then
        MsgBox "すでに起動されています！"
        Exit Sub
    End If

    'HPGL変換プログラムのデフォルト
    gstrCmdDefault = Environ("COMSPEC") & " /C " & "gawk -f nc2hplib.awk -f convert.awk"

    If Command <> "" Then
        Load frmMain
    Else
        frmSetting.Show
    End If

End Sub

Public Function fMyPath() As String

    'プログラム終了まで　MyPath　の内容を保持
    Static MyPath As String
    '途中でディレクトリ-が変更されても起動ディレクトリ-を確保
    If Len(MyPath) = 0& Then
        MyPath = App.Path   'ディレクトリ-を取得
        'ルートディレクトリーかの判断
        If Right$(MyPath, 1&) <> "\" Then
            MyPath = MyPath & "\"
        End If
    End If
    fMyPath = MyPath

End Function

Public Function fTempPath() As String

    'プログラム終了まで　TempPath　の内容を保持
    Static TempPath As String
    '途中でディレクトリ-が変更されてもTempディレクトリ-を確保
    If Len(TempPath) = 0& Then
        TempPath = Environ("TEMP")         'ディレクトリ-を取得
        'ルートディレクトリーかの判断
        If Right$(TempPath, 1&) <> "\" Then
            TempPath = TempPath & "\"
        End If
    End If
    fTempPath = TempPath

End Function

