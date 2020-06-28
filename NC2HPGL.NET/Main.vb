Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMain
	Inherits System.Windows.Forms.Form
#Region "Windows フォーム デザイナによって生成されたコード"
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'スタートアップ フォームについては、最初に作成されたインスタンスが既定インスタンスになります。
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'この呼び出しは、Windows フォーム デザイナで必要です。
		InitializeComponent()
	End Sub
	'Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows フォーム デザイナで必要です。
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Label1 As System.Windows.Forms.Label
	'メモ : 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使って修正しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.Label1 = New System.Windows.Forms.Label
		Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
		Me.Text = "NC2HPGL"
		Me.ClientSize = New System.Drawing.Size(111, 29)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.ControlBox = False
		Me.Icon = CType(resources.GetObject("frmMain.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmMain"
		Me.Label1.BackColor = System.Drawing.SystemColors.ActiveCaptionText
		Me.Label1.Text = "変換中．．．"
		Me.Label1.Size = New System.Drawing.Size(65, 17)
		Me.Label1.Location = New System.Drawing.Point(24, 8)
		Me.Label1.TabIndex = 0
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(Label1)
	End Sub
#End Region 
#Region "アップグレード ウィザードのサポート コード"
	Private Shared m_vb6FormDefInstance As frmMain
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmMain
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmMain()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	
	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim intFNo1 As Short
		Dim bytBuf() As Byte
		Dim strHPGL As String
		Dim strFileName() As String
		Dim strTmpFName As String
		Dim strDir As String
		Dim strCmd As String 'HPGL変換プログラム名
		Dim strDocName As String
		Dim strTmpDocName() As String 'テンポラリ用配列
		Dim i As Short
		
		Me.Show()
		Me.Enabled = False
		System.Windows.Forms.Application.DoEvents()
		
		strDir = fMyPath()
		strFileName = Split(VB.Command(), " ", -1)
		
		'使用するプリンタの設定
		gstrPrinterName = GetSetting("NC2HPGL", "Settings", "Printer")
		If gstrPrinterName = "" Then
			Me.Close()
			MsgBox("NC2HPGL.EXEを起動してプリンタを選択して下さい。", MsgBoxStyle.Critical)
			End
		End If
		
		'ファイルが存在するか調べる
		For i = 0 To UBound(strFileName)
			'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
			If Dir(strFileName(i), FileAttribute.Normal) = "" Then
				Me.Close()
				MsgBox(strFileName(i) & "が見つかりません。", MsgBoxStyle.Critical)
				End
			End If
		Next 
		
		'NCを.\ncにコピー
		For i = 0 To UBound(strFileName)
			'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
			FileCopy(strFileName(i), strDir & "\nc\" & Dir(strFileName(i)))
		Next 
		
		'スプーラに表示するドキュメント名
		'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		strTmpDocName = Split(Dir(strFileName(0)), ".", -1)
		strDocName = strTmpDocName(0)
		
		'テンポラリファイル名
		strTmpFName = fTempPath() & "TEMP.HP"
		
		'HPGL変換プログラムの設定
		strCmd = GetSetting("NC2HPGL", "Settings", "Command", gstrCmdDefault)
		
		'カレントディレクトリをEXEファイルのあるディレクトリにセット
		ChDir(strDir)
		
		'HPGL変換プログラムの実行
		Call ShellWait(strCmd, AppWinStyle.MinimizedFocus)
		
		'TEMP.HPを読み込む
		intFNo1 = FreeFile
		FileOpen(intFNo1, strTmpFName, OpenMode.Binary)
		ReDim bytBuf(LOF(intFNo1))
		'UPGRADE_WARNING: Get は、FileGet にアップグレードされ、新しい動作が指定されました。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(intFNo1, bytBuf)
		FileClose(intFNo1)
		'UPGRADE_ISSUE: 定数 vbUnicode はアップグレードされませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2070"'
		strHPGL = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(bytBuf), vbUnicode)
		Call PutHPGL(strHPGL, strDocName) 'スプーラに送り込む
		
		'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Dir(strTmpFName) <> "" Then
			Kill(strTmpFName) ' テンポラリファイルを削除
		End If
		
		End
		
	End Sub
End Class