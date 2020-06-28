VERSION 5.00
Begin VB.Form frmSetting 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "Form1"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox txtCommand 
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   4215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Text            =   "cmbPrinter"
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "HPGL変換コマンド"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "利用可能なプリンタ"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    SaveSetting "NC2HPGL", _
                "Settings", _
                "Printer", _
                cmbPrinter.Text

    SaveSetting "NC2HPGL", _
                "Settings", _
                "Command", _
                txtCommand.Text

    Unload Me

End Sub

Private Sub Form_Load()

    Dim objPrn As Printer

    Caption = "環境設定"

    With cmbPrinter
        .Text = GetSetting("NC2HPGL", _
                           "Settings", _
                           "Printer")
'        .Locked = True
    End With

    For Each objPrn In Printers
        cmbPrinter.AddItem objPrn.DeviceName
    Next

    txtCommand.Text = GetSetting("NC2HPGL", _
                                 "Settings", _
                                 "Command", _
                                 gstrCmdDefault)

End Sub
