Attribute VB_Name = "Ploter"
Option Explicit

Private Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal _
    hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal _
    hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal _
    hPrinter As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
    "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
    ByVal pDefault As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias _
    "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
    pDocInfo As DOCINFO) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal _
    hPrinter As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal _
    hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
    pcWritten As Long) As Long

Public Sub PutHPGL(ByVal strHPGL As String, _
                    ByVal strDocName As String)

    Dim lhPrinter As Long
    Dim lReturn As Long
    Dim lpcWritten As Long
    Dim lDoc As Long
    Dim sWrittenData As String
    Dim MyDocInfo As DOCINFO
    lReturn = OpenPrinter(gstrPrinterName, lhPrinter, 0)
    If lReturn = 0 Then
        MsgBox "The Printer Name you typed wasn't recognized."
        Exit Sub
    End If
    MyDocInfo.pDocName = strDocName
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)
'    sWrittenData = "How's that for Magic !!!!" & vbFormFeed
    sWrittenData = strHPGL
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)

End Sub
