Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module MainModule
	
	Public gstrPrinterName As String '�v�����^��
	Public gstrCmdDefault As String 'HPGL�ϊ��v���O�������̃f�t�H���g
	
	'UPGRADE_WARNING: Sub Main() �����������Ƃ��ɃA�v���P�[�V�����͏I�����܂��B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1047"'
	Public Sub Main()
		
		'2�d�N�����`�F�b�N
		If (UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0) Then
			MsgBox("���łɋN������Ă��܂��I")
			Exit Sub
		End If
		
		'HPGL�ϊ��v���O�����̃f�t�H���g
		gstrCmdDefault = Environ("COMSPEC") & " /C " & "gawk -f nc2hplib.awk -f convert.awk"
		
		If VB.Command() <> "" Then
			'UPGRADE_ISSUE: Load �X�e�[�g�����g �̓T�|�[�g����Ă��܂���B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1039"'
			Load(frmMain)
		Else
			frmSetting.DefInstance.Show()
		End If
		
	End Sub
	
	Public Function fMyPath() As String
		
		'�v���O�����I���܂Ł@MyPath�@�̓��e��ێ�
		Static MyPath As String
		'�r���Ńf�B���N�g��-���ύX����Ă��N���f�B���N�g��-���m��
		If Len(MyPath) = 0 Then
			MyPath = VB6.GetPath '�f�B���N�g��-���擾
			'���[�g�f�B���N�g���[���̔��f
			If Right(MyPath, 1) <> "\" Then
				MyPath = MyPath & "\"
			End If
		End If
		fMyPath = MyPath
		
	End Function
	
	Public Function fTempPath() As String
		
		'�v���O�����I���܂Ł@TempPath�@�̓��e��ێ�
		Static TempPath As String
		'�r���Ńf�B���N�g��-���ύX����Ă�Temp�f�B���N�g��-���m��
		If Len(TempPath) = 0 Then
			TempPath = Environ("TEMP") '�f�B���N�g��-���擾
			'���[�g�f�B���N�g���[���̔��f
			If Right(TempPath, 1) <> "\" Then
				TempPath = TempPath & "\"
			End If
		End If
		fTempPath = TempPath
		
	End Function
End Module