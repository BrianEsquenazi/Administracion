VERSION 5.00
Begin VB.Form PrgRecuperoIva 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recupero de Iva"
   ClientHeight    =   6825
   ClientLeft      =   3105
   ClientTop       =   990
   ClientWidth     =   8055
   FillColor       =   &H00800000&
   Icon            =   "RecuperoIva.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4032.436
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   7563.21
   ShowInTaskbar   =   0   'False
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   3975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   1920
      TabIndex        =   0
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Image Cancela 
      Height          =   480
      Left            =   4200
      MouseIcon       =   "RecuperoIva.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "RecuperoIva.frx":074C
      ToolTipText     =   "Menu Principal"
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image Acepta 
      Height          =   480
      Left            =   2880
      MouseIcon       =   "RecuperoIva.frx":0F8E
      MousePointer    =   99  'Custom
      Picture         =   "RecuperoIva.frx":1298
      ToolTipText     =   "Confirma el Proceso"
      Top             =   5160
      Width           =   480
   End
End
Attribute VB_Name = "PrgRecuperoIva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WPasoUnifica As String

Dim ZZCarpeta As String
Dim ZZFecha As String
Dim ZZFechaI As String
Dim ZZDespacho As String
Dim ZZPosicion As String
Dim ZZItem As String
Dim ZZCuit As String
Dim ZZRazon As String
Dim ZZIvaDespacho As String
Dim ZZIvacomp As String
Dim ZZIvaTotal As String
Dim ZZPeriodo As String
Dim ZZIva As String
Dim ZZKg As String
Dim ZZTotalKg As String


Private Sub Acepta_Click()

    WPasoUnifica = Trim(WPasoUnifica)
    ZZLargo = Len(WPasoUnifica)
    
    If Right$(WPasoUnifica, 3) = "xls" Then
    
        WpasoII = Mid$(WPasoUnifica, 1, ZZLargo - 3) + "txt"
    
        Open WpasoII For Output As #1
        
        Set appExcel = CreateObject("Excel.application")
        
        ruta = WPasoUnifica
        LugarPlanilla = 5
    
        If Len(Dir(ruta)) > 0 Then
        
        
            Set objLibro = appExcel.workbooks.Open(ruta)
            
            Do
            
                LugarPlanilla = LugarPlanilla + 1
                
                ZZCarpeta = appExcel.cells(LugarPlanilla, 1).Value
                ZZFecha = appExcel.cells(LugarPlanilla, 2).Value
                ZZFechaI = appExcel.cells(LugarPlanilla, 3).Value
                ZZDespacho = appExcel.cells(LugarPlanilla, 4).Value
                ZZPosicion = appExcel.cells(LugarPlanilla, 5).Value
                ZZItem = appExcel.cells(LugarPlanilla, 6).Value
                ZZCuit = appExcel.cells(LugarPlanilla, 7).Value
                ZZRazon = appExcel.cells(LugarPlanilla, 8).Value
                ZZIvaDespacho = appExcel.cells(LugarPlanilla, 9).Value
                ZZIvacomp = appExcel.cells(LugarPlanilla, 10).Value
                ZZIvaTotal = appExcel.cells(LugarPlanilla, 11).Value
                ZZPeriodo = appExcel.cells(LugarPlanilla, 12).Value
                ZZIva = appExcel.cells(LugarPlanilla, 13).Value
                ZZKg = appExcel.cells(LugarPlanilla, 14).Value
                ZZTotalKg = appExcel.cells(LugarPlanilla, 15).Value
                
                If Val(ZZCarpeta) = 0 Then Exit Do
                
                If Val(ZZIvaTotal) <> 0 Then
                    
                    ZZDespacho = Trim(ZZDespacho) + Space$(100)
                    ZZDespacho = Left$(ZZDespacho, 16)
                    
                    ZZPosicion = Trim(ZZPosicion) + Space$(100)
                    ZZPosicion = Left$(ZZPosicion, 12)
                    
                    Call Ceros(ZZItem, 4)
                    
                    Call Ceros(ZZCuit, 11)
                    
                    ZZRazon = Trim(ZZRazon) + Space$(100)
                    ZZRazon = Left$(ZZRazon, 50)
                    
                    ZZIvaTotal = Str$(Int(Val(ZZIvaTotal) * 100))
                    Call Ceros(ZZIvaTotal, 15)
                    
                    ZZIvaDespacho = Str$(Int(Val(ZZIvaDespacho) * 100))
                    Call Ceros(ZZIvaDespacho, 15)
                    
                    ZZIvacomp = Str$(Int(Val(ZZIvacomp) * 100))
                    Call Ceros(ZZIvacomp, 15)
                    
                    ZZPeriodo = Right$(ZZPeriodo, 4) + Mid$(ZZPeriodo, 4, 2)
                    
                    
                    ZZCampo1 = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left(ZZFecha, 2)
                    ZZCampo2 = ZZDespacho
                    ZZCampo3 = ZZPosicion
                    ZZCampo4 = ZZItem
                    ZZCampo5 = ZZCuit
                    ZZCampo6 = ZZRazon
                    ZZCampo7 = ZZIvaTotal
                    ZZCampo8 = ZZIvaDespacho
                    ZZCampo9 = ZZIvacomp
                    ZZCampo10 = ZZPeriodo
                    ZZCampo11 = "D"
                    
                    WImpre = ZZCampo1 + ZZCampo2 + ZZCampo3 + ZZCampo4 + ZZCampo5 + ZZCampo6 + ZZCampo7 + ZZCampo8 + ZZCampo9 + ZZCampo10 + ZZCampo11
                    Print #1, WImpre
                    
                End If
                
            Loop
                
            appExcel.Quit
            Set appExcel = Nothing
            
        End If
        
        Close #1
        
        m$ = "Proceso Finalizado"
        A% = MsgBox(m$, 0, "Recupero de Iva")
        
        
    End If

End Sub

Private Sub Cancela_Click()
    PrgRecuperoIva.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive    ' Establece la ruta del directorio.
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path  ' Establece la ruta del archivo.
End Sub

Private Sub File1_dblClick()
    On Error GoTo WError
    
    WDrive = Drive1.Drive
    WDir = Dir1.Path
    WLon = Len(WDir)
    If Right$(WDir, 1) = "\" Then
        WDir = Mid(WDir, 1, WLon - 1)
    End If
    XNombre = WDir + "\"
    
    WPasoUnifica = XNombre + File1.filename
    
    Call Acepta_Click
    
    Exit Sub
    
WError:
    Rem MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    Resume Next
End Sub

Private Sub File1_Click()

    On Error GoTo WError
    
    WDrive = Drive1.Drive
    WDir = Dir1.Path
    WLon = Len(WDir)
    If Right$(WDir, 1) = "\" Then
        WDir = Mid(WDir, 1, WLon - 1)
    End If
    XNombre = WDir + "\"
    
    WPasoUnifica = XNombre + File1.filename
    
    Exit Sub
    
WError:
    Rem MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    Resume Next
    
End Sub

