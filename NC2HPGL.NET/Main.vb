Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMain
	Inherits System.Windows.Forms.Form
#Region "Windows �t�H�[�� �f�U�C�i�ɂ���Đ������ꂽ�R�[�h"
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'�X�^�[�g�A�b�v �t�H�[���ɂ��ẮA�ŏ��ɍ쐬���ꂽ�C���X�^���X������C���X�^���X�ɂȂ�܂��B
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'���̌Ăяo���́AWindows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
		InitializeComponent()
	End Sub
	'Form �́A�R���|�[�l���g�ꗗ�Ɍ㏈�������s���邽�߂� dispose ���I�[�o�[���C�h���܂��B
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Label1 As System.Windows.Forms.Label
	'���� : �ȉ��̃v���V�[�W���� Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
	'Windows �t�H�[�� �f�U�C�i���g���ĕύX�ł��܂��B
	'�R�[�h �G�f�B�^���g���ďC�����Ȃ��ł��������B
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
		Me.Label1.Text = "�ϊ����D�D�D"
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
#Region "�A�b�v�O���[�h �E�B�U�[�h�̃T�|�[�g �R�[�h"
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
		Dim strCmd As String 'HPGL�ϊ��v���O������
		Dim strDocName As String
		Dim strTmpDocName() As String '�e���|�����p�z��
		Dim i As Short
		
		Me.Show()
		Me.Enabled = False
		System.Windows.Forms.Application.DoEvents()
		
		strDir = fMyPath()
		strFileName = Split(VB.Command(), " ", -1)
		
		'�g�p����v�����^�̐ݒ�
		gstrPrinterName = GetSetting("NC2HPGL", "Settings", "Printer")
		If gstrPrinterName = "" Then
			Me.Close()
			MsgBox("NC2HPGL.EXE���N�����ăv�����^��I�����ĉ������B", MsgBoxStyle.Critical)
			End
		End If
		
		'�t�@�C�������݂��邩���ׂ�
		For i = 0 To UBound(strFileName)
			'UPGRADE_WARNING: Dir �ɐV�������삪�w�肳��Ă��܂��B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
			If Dir(strFileName(i), FileAttribute.Normal) = "" Then
				Me.Close()
				MsgBox(strFileName(i) & "��������܂���B", MsgBoxStyle.Critical)
				End
			End If
		Next 
		
		'NC��.\nc�ɃR�s�[
		For i = 0 To UBound(strFileName)
			'UPGRADE_WARNING: Dir �ɐV�������삪�w�肳��Ă��܂��B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
			FileCopy(strFileName(i), strDir & "\nc\" & Dir(strFileName(i)))
		Next 
		
		'�X�v�[���ɕ\������h�L�������g��
		'UPGRADE_WARNING: Dir �ɐV�������삪�w�肳��Ă��܂��B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		strTmpDocName = Split(Dir(strFileName(0)), ".", -1)
		strDocName = strTmpDocName(0)
		
		'�e���|�����t�@�C����
		strTmpFName = fTempPath() & "TEMP.HP"
		
		'HPGL�ϊ��v���O�����̐ݒ�
		strCmd = GetSetting("NC2HPGL", "Settings", "Command", gstrCmdDefault)
		
		'�J�����g�f�B���N�g����EXE�t�@�C���̂���f�B���N�g���ɃZ�b�g
		ChDir(strDir)
		
		'HPGL�ϊ��v���O�����̎��s
		Call ShellWait(strCmd, AppWinStyle.MinimizedFocus)
		
		'TEMP.HP��ǂݍ���
		intFNo1 = FreeFile
		FileOpen(intFNo1, strTmpFName, OpenMode.Binary)
		ReDim bytBuf(LOF(intFNo1))
		'UPGRADE_WARNING: Get �́AFileGet �ɃA�b�v�O���[�h����A�V�������삪�w�肳��܂����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(intFNo1, bytBuf)
		FileClose(intFNo1)
		'UPGRADE_ISSUE: �萔 vbUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2070"'
		strHPGL = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(bytBuf), vbUnicode)
		Call PutHPGL(strHPGL, strDocName) '�X�v�[���ɑ��荞��
		
		'UPGRADE_WARNING: Dir �ɐV�������삪�w�肳��Ă��܂��B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Dir(strTmpFName) <> "" Then
			Kill(strTmpFName) ' �e���|�����t�@�C�����폜
		End If
		
		End
		
	End Sub
End Class