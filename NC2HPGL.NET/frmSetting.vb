Option Strict Off
Option Explicit On
Friend Class frmSetting
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
	Public WithEvents txtCommand As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmbPrinter As System.Windows.Forms.ComboBox
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'メモ : 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使って修正しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSetting))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.txtCommand = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmbPrinter = New System.Windows.Forms.ComboBox
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Form1"
		Me.ClientSize = New System.Drawing.Size(312, 192)
		Me.Location = New System.Drawing.Point(3, 22)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmSetting"
		Me.txtCommand.AutoSize = False
		Me.txtCommand.Size = New System.Drawing.Size(281, 18)
		Me.txtCommand.Location = New System.Drawing.Point(16, 104)
		Me.txtCommand.TabIndex = 4
		Me.txtCommand.AcceptsReturn = True
		Me.txtCommand.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCommand.BackColor = System.Drawing.SystemColors.Window
		Me.txtCommand.CausesValidation = True
		Me.txtCommand.Enabled = True
		Me.txtCommand.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCommand.HideSelection = True
		Me.txtCommand.ReadOnly = False
		Me.txtCommand.Maxlength = 0
		Me.txtCommand.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCommand.MultiLine = False
		Me.txtCommand.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCommand.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCommand.TabStop = True
		Me.txtCommand.Visible = True
		Me.txtCommand.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCommand.Name = "txtCommand"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "ｷｬﾝｾﾙ"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 33)
		Me.cmdCancel.Location = New System.Drawing.Point(216, 144)
		Me.cmdCancel.TabIndex = 3
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.cmdOK.Size = New System.Drawing.Size(81, 33)
		Me.cmdOK.Location = New System.Drawing.Point(112, 144)
		Me.cmdOK.TabIndex = 2
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.cmbPrinter.Size = New System.Drawing.Size(281, 20)
		Me.cmbPrinter.Location = New System.Drawing.Point(16, 40)
		Me.cmbPrinter.TabIndex = 0
		Me.cmbPrinter.Text = "cmbPrinter"
		Me.cmbPrinter.BackColor = System.Drawing.SystemColors.Window
		Me.cmbPrinter.CausesValidation = True
		Me.cmbPrinter.Enabled = True
		Me.cmbPrinter.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbPrinter.IntegralHeight = True
		Me.cmbPrinter.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbPrinter.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbPrinter.Sorted = False
		Me.cmbPrinter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbPrinter.TabStop = True
		Me.cmbPrinter.Visible = True
		Me.cmbPrinter.Name = "cmbPrinter"
		Me.Label2.Text = "HPGL変換コマンド"
		Me.Label2.Size = New System.Drawing.Size(97, 17)
		Me.Label2.Location = New System.Drawing.Point(16, 80)
		Me.Label2.TabIndex = 5
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.Text = "利用可能なプリンタ"
		Me.Label1.Size = New System.Drawing.Size(105, 17)
		Me.Label1.Location = New System.Drawing.Point(16, 16)
		Me.Label1.TabIndex = 1
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(txtCommand)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmbPrinter)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
	End Sub
#End Region 
#Region "アップグレード ウィザードのサポート コード"
	Private Shared m_vb6FormDefInstance As frmSetting
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmSetting
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmSetting()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		
		Me.Close()
		
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		
		SaveSetting("NC2HPGL", "Settings", "Printer", cmbPrinter.Text)
		
		SaveSetting("NC2HPGL", "Settings", "Command", txtCommand.Text)
		
		Me.Close()
		
	End Sub
	
	Private Sub frmSetting_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		'UPGRADE_ISSUE: Printer オブジェクト はアップグレードされませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2068"'
		Dim objPrn As Printer
		
		Text = "環境設定"
		
		With cmbPrinter
			.Text = GetSetting("NC2HPGL", "Settings", "Printer")
			'        .Locked = True
		End With
		
		'UPGRADE_ISSUE: Printers コレクション はアップグレードされませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2068"'
		For	Each objPrn In Printers
			'UPGRADE_ISSUE: Printer プロパティ objPrn.DeviceName はアップグレードされませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2069"'
			cmbPrinter.Items.Add(objPrn.DeviceName)
		Next objPrn
		
		txtCommand.Text = GetSetting("NC2HPGL", "Settings", "Command", gstrCmdDefault)
		
	End Sub
End Class