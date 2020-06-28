Attribute VB_Name = "MainModule"
Option Explicit

Public gstrPrinterName As String '�v�����^��
Public gstrCmdDefault As String 'HPGL�ϊ��v���O�������̃f�t�H���g

Sub Main()

    '2�d�N�����`�F�b�N
    If App.PrevInstance Then
        MsgBox "���łɋN������Ă��܂��I"
        Exit Sub
    End If

    'HPGL�ϊ��v���O�����̃f�t�H���g
    gstrCmdDefault = Environ("COMSPEC") & " /C " & "gawk -f nc2hplib.awk -f convert.awk"

    If Command <> "" Then
        Load frmMain
    Else
        frmSetting.Show
    End If

End Sub

Public Function fMyPath() As String

    '�v���O�����I���܂Ł@MyPath�@�̓��e��ێ�
    Static MyPath As String
    '�r���Ńf�B���N�g��-���ύX����Ă��N���f�B���N�g��-���m��
    If Len(MyPath) = 0& Then
        MyPath = App.Path   '�f�B���N�g��-���擾
        '���[�g�f�B���N�g���[���̔��f
        If Right$(MyPath, 1&) <> "\" Then
            MyPath = MyPath & "\"
        End If
    End If
    fMyPath = MyPath

End Function

Public Function fTempPath() As String

    '�v���O�����I���܂Ł@TempPath�@�̓��e��ێ�
    Static TempPath As String
    '�r���Ńf�B���N�g��-���ύX����Ă�Temp�f�B���N�g��-���m��
    If Len(TempPath) = 0& Then
        TempPath = Environ("TEMP")         '�f�B���N�g��-���擾
        '���[�g�f�B���N�g���[���̔��f
        If Right$(TempPath, 1&) <> "\" Then
            TempPath = TempPath & "\"
        End If
    End If
    fTempPath = TempPath

End Function

