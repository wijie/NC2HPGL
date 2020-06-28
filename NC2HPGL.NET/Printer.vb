Option Strict Off
Option Explicit On
Module Ploter
	
	Private Structure DOCINFO
		Dim pDocName As String
		Dim pOutputFile As String
		Dim pDatatype As String
	End Structure
	
	Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Integer) As Integer
	Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Integer) As Integer
	Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Integer) As Integer
	Private Declare Function OpenPrinter Lib "winspool.drv"  Alias "OpenPrinterA"(ByVal pPrinterName As String, ByRef phPrinter As Integer, ByVal pDefault As Integer) As Integer
	'UPGRADE_WARNING: 構造体 DOCINFO に、この Declare ステートメントの引数としてマーシャリング属性を渡す必要があります。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	Private Declare Function StartDocPrinter Lib "winspool.drv"  Alias "StartDocPrinterA"(ByVal hPrinter As Integer, ByVal Level As Integer, ByRef pDocInfo As DOCINFO) As Integer
	Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Integer) As Integer
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1016"'
	Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Integer, ByRef pBuf As Any, ByVal cdBuf As Integer, ByRef pcWritten As Integer) As Integer
	
	Public Sub PutHPGL(ByVal strHPGL As String, ByVal strDocName As String)
		
		Dim lhPrinter As Integer
		Dim lReturn As Integer
		Dim lpcWritten As Integer
		Dim lDoc As Integer
		Dim sWrittenData As String
		Dim MyDocInfo As DOCINFO
		lReturn = OpenPrinter(gstrPrinterName, lhPrinter, 0)
		If lReturn = 0 Then
			MsgBox("The Printer Name you typed wasn't recognized.")
			Exit Sub
		End If
		MyDocInfo.pDocName = strDocName
		MyDocInfo.pOutputFile = vbNullString
		MyDocInfo.pDatatype = vbNullString
		lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
		Call StartPagePrinter(lhPrinter)
		'    sWrittenData = "How's that for Magic !!!!" & vbFormFeed
		sWrittenData = strHPGL
		lReturn = WritePrinter(lhPrinter, sWrittenData, Len(sWrittenData), lpcWritten)
		lReturn = EndPagePrinter(lhPrinter)
		lReturn = EndDocPrinter(lhPrinter)
		lReturn = ClosePrinter(lhPrinter)
		
	End Sub
End Module