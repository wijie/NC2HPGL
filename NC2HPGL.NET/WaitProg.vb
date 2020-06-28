Option Strict Off
Option Explicit On
Module WaitProg
	
	'ハンドルを取得
	Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
	
	'終了ステータスを取得
	Public Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Integer, ByRef lpExitCode As Integer) As Integer
	
	'ハンドルを閉じる
	Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Integer) As Integer
	
	Const PROCESS_QUERY_INFORMATION As Integer = &H400
	Const STILL_ACTIVE As Integer = &H103
	
	Public Sub ShellWait(ByVal strProg As String, ByVal bytStyle As Byte)
		
		Dim lngProcHandle As Integer
		Dim lngExitCode As Integer
		Dim lngReturnCode As Integer
		Dim lngTaskID As Integer
		
		lngTaskID = Shell(strProg, bytStyle)
		System.Windows.Forms.Application.DoEvents()
		
		'ハンドルを取得する
		lngProcHandle = OpenProcess(PROCESS_QUERY_INFORMATION, 1, lngTaskID)
		'終わるまで待つ
		Do 
			lngReturnCode = GetExitCodeProcess(lngProcHandle, lngExitCode)
			System.Windows.Forms.Application.DoEvents()
		Loop While lngExitCode = STILL_ACTIVE
		'ハンドルを閉じる
		CloseHandle(lngProcHandle)
	End Sub
End Module