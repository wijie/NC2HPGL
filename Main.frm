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
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "�ϊ����D�D�D"
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
    Dim strCmd As String 'HPGL�ϊ��v���O������
    Dim strDocName As String
    Dim strTmpDocName() As String '�e���|�����p�z��
    Dim i As Integer

    Me.Show
    Me.Enabled = False
    DoEvents

    strDir = fMyPath()
    strFileName = Split(Command, " ", -1)

    '�g�p����v�����^�̐ݒ�
    gstrPrinterName = GetSetting("NC2HPGL", _
                                 "Settings", _
                                 "Printer")
    If gstrPrinterName = "" Then
        Unload Me
        MsgBox "NC2HPGL.EXE���N�����ăv�����^��I�����ĉ������B", _
               vbCritical
        End
    End If

    '�t�@�C�������݂��邩���ׂ�
    For i = 0 To UBound(strFileName)
        If Dir(strFileName(i), vbNormal) = "" Then
            Unload Me
            MsgBox strFileName(i) & "��������܂���B", _
                   vbCritical
            End
        End If
    Next

    'NC��.\nc�ɃR�s�[
    For i = 0 To UBound(strFileName)
        FileCopy strFileName(i), _
                 strDir & "\nc\" & Dir(strFileName(i))
    Next

    '�X�v�[���ɕ\������h�L�������g��
    strTmpDocName = Split(Dir(strFileName(0)), ".", -1)
    strDocName = strTmpDocName(0)

    '�e���|�����t�@�C����
    strTmpFName = fTempPath() & "TEMP.HP"

    'HPGL�ϊ��v���O�����̐ݒ�
    strCmd = GetSetting("NC2HPGL", _
                        "Settings", _
                        "Command", _
                        gstrCmdDefault)

    '�J�����g�f�B���N�g����EXE�t�@�C���̂���f�B���N�g���ɃZ�b�g
    ChDir strDir

    'HPGL�ϊ��v���O�����̎��s
    Call ShellWait(strCmd, vbMinimizedFocus)

    'TEMP.HP��ǂݍ���
    intFNo1 = FreeFile
    Open strTmpFName For Binary As #intFNo1
    ReDim bytBuf(LOF(intFNo1))
    Get #intFNo1, , bytBuf
    Close #intFNo1
    strHPGL = StrConv(bytBuf, vbUnicode)
    Call PutHPGL(strHPGL, strDocName) '�X�v�[���ɑ��荞��

    If Dir(strTmpFName) <> "" Then
        Kill strTmpFName ' �e���|�����t�@�C�����폜
    End If

    End

End Sub
