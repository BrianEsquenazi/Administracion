Attribute VB_Name = "Module3"
Option Explicit
Declare Function EnumPorts Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal Level As Long, lpbPorts As Byte, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Type PORT_INFO_2
        pPortName As Long
        pMonitorName As Long
        pDescription As Long
        fPortType As Long
        Reserved As Long
End Type
Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
   hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Declare Function CopyPointer2String Lib "KERNEL32" _
   Alias "lstrcpyA" ( _
   ByVal NewString As String, ByVal OldString As Long) As Long
Public Const PORT_TYPE_NET_ATTACHED = &H8
Public Const PORT_TYPE_READ = &H2
Public Const PORT_TYPE_REDIRECTED = &H4
Public Const PORT_TYPE_WRITE = &H1
Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, pJob As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Type JOB_INFO_1
        JobId As Long
        pPrinterName As Long
        pMachineName As Long
        pUserName As Long
        pDocument As Long
        pDatatype As Long
        pStatus As Long
        Status As Long
        Priority As Long
        Position As Long
        TotalPages As Long
        PagesPrinted As Long
        Submitted As SYSTEMTIME
End Type
Type PRINTER_DEFAULTS
        pDatatype As Long
        pDevMode As Long
        DesiredAccess As Long
End Type
Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, PDefault As PRINTER_DEFAULTS) As Long
Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Const JOB_STATUS_DELETING = &H4
Public Const JOB_STATUS_ERROR = &H2
Public Const JOB_STATUS_OFFLINE = &H20
Public Const JOB_STATUS_PAPEROUT = &H40
Public Const JOB_STATUS_PAUSED = &H1
Public Const JOB_STATUS_PRINTED = &H80
Public Const JOB_STATUS_PRINTING = &H10
Public Const JOB_STATUS_SPOOLING = &H8
Public Const JOB_STATUS_USER_INTERVENTION = 1024
Public Const MIN_PRIORITY = 1
Public Const MAX_PRIORITY = 99
Public Const DEF_PRIORITY = 1


Function PointerToString(p As Long) As String
   Dim s As String
   s = String(255, Chr$(0))
   CopyPointer2String s, p
   PointerToString = Left(s, InStr(s, Chr$(0)) - 1)
End Function


