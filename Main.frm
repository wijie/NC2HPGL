VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000009&
   Caption         =   "NC2HPGL"
   ClientHeight    =   435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1665
   ControlBox      =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   435
   ScaleWidth      =   1665
   StartUpPosition =   2  '画面の中央
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "変換中．．．"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim intFNo1 As Integer
    Dim bytBuf() As Byte
    Dim strHPGL As String
    Dim strFileName() As String
    Dim strTmpFName As String
    Dim strDir As String
    Dim strCmd As String 'HPGL変換プログラム名
    Dim strDocName As String
    Dim strTmpDocName() As String 'テンポラリ用配列
    Dim i As Integer

    Me.Show
    Me.Enabled = False
    DoEvents

    strDir = fMyPath()
    strFileName = Split(Command, " ", -1)

    '使用するプリンタの設定
    gstrPrinterName = GetSetting("NC2HPGL", _
                                 "Settings", _
                                 "Printer")
    If gstrPrinterName = "" Then
        Unload Me
        MsgBox "NC2HPGL.EXEを起動してプリンタを選択して下さい。", _
               vbCritical
        End
    End If

    'ファイルが存在するか調べる
    For i = 0 To UBound(strFileName)
        If Dir(strFileName(i), vbNormal) = "" Then
            Unload Me
            MsgBox strFileName(i) & "が見つかりません。", _
                   vbCritical
            End
        End If
    Next

    'NCを.\ncにコピー
    For i = 0 To UBound(strFileName)
        FileCopy strFileName(i), _
                 strDir & "\nc\" & Dir(strFileName(i))
    Next

    'スプーラに表示するドキュメント名
    strTmpDocName = Split(Dir(strFileName(0)), ".", -1)
    strDocName = strTmpDocName(0)

    'テンポラリファイル名
    strTmpFName = fTempPath() & "TEMP.HP"

    'HPGL変換プログラムの設定
    strCmd = GetSetting("NC2HPGL", _
                        "Settings", _
                        "Command", _
                        gstrCmdDefault)

    'カレントディレクトリをEXEファイルのあるディレクトリにセット
    ChDir strDir

    'HPGL変換プログラムの実行
    Call ShellWait(strCmd, vbMinimizedFocus)

    'TEMP.HPを読み込む
    intFNo1 = FreeFile
    Open strTmpFName For Binary As #intFNo1
    ReDim bytBuf(LOF(intFNo1))
    Get #intFNo1, , bytBuf
    Close #intFNo1
    strHPGL = StrConv(bytBuf, vbUnicode)
    Call PutHPGL(strHPGL, strDocName) 'スプーラに送り込む

    If Dir(strTmpFName) <> "" Then
        Kill strTmpFName ' テンポラリファイルを削除
    End If

    End

End Sub
