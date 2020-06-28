Attribute VB_Name = "WaitProg"
Option Explicit

'ハンドルを取得
Declare Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

'終了ステータスを取得
Public Declare Function GetExitCodeProcess Lib "kernel32.dll" ( _
    ByVal hProcess As Long, _
    lpExitCode As Long) As Long

'ハンドルを閉じる
Declare Function CloseHandle Lib "kernel32.dll" ( _
    ByVal hObject As Long) As Long

Const PROCESS_QUERY_INFORMATION = &H400&
Const STILL_ACTIVE = &H103&

Public Sub ShellWait(ByVal strProg As String, _
                     ByVal bytStyle As Byte)

    Dim lngProcHandle As Long
    Dim lngExitCode As Long
    Dim lngReturnCode As Long
    Dim lngTaskID As Long

    lngTaskID = Shell(strProg, bytStyle)
    DoEvents

    'ハンドルを取得する
    lngProcHandle = _
        OpenProcess(PROCESS_QUERY_INFORMATION, 1, lngTaskID)
    '終わるまで待つ
    Do
        lngReturnCode = _
            GetExitCodeProcess(lngProcHandle, lngExitCode)
        DoEvents
    Loop While lngExitCode = STILL_ACTIVE
    'ハンドルを閉じる
    CloseHandle lngProcHandle
End Sub
