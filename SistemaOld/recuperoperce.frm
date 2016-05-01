VERSION 5.00
Begin VB.Form PrgRecuperoPerce 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recupero de Pecepciones de IB Aduana"
   ClientHeight    =   6825
   ClientLeft      =   3105
   ClientTop       =   990
   ClientWidth     =   8055
   FillColor       =   &H00800000&
   Icon            =   "recuperoperce.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4032.436
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   7563.209
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
      MouseIcon       =   "recuperoperce.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "recuperoperce.frx":074C
      ToolTipText     =   "Menu Principal"
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image Acepta 
      Height          =   480
      Left            =   2880
      MouseIcon       =   "recuperoperce.frx":0F8E
      MousePointer    =   99  'Custom
      Picture         =   "recuperoperce.frx":1298
      ToolTipText     =   "Confirma el Proceso"
      Top             =   5160
      Width           =   480
   End
End
Attribute VB_Name = "PrgRecuperoPerce"
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

Dim WClave As String
Dim WFecha As String
Dim WTipo As String
Dim WNumero As String

Dim WNeto As Double
Dim WIva1 As Double
Dim WIva2 As Double
Dim WImpoIb As Double
Dim WImpoIbTucu As Double
Dim WImpoIbCiudad As Double
Dim WTotal As Double

Dim WIbTucu As Integer

Dim WPorceIb As Double
Dim WPorceIbCiudad As Double

Dim XNeto As String
Dim XImpoIb As String
Dim XPorceIb As String
Dim XImpoIbCiudad As String
Dim XPorceIbCiudad As String

Dim WCuit As String
Dim ZEntra(1000) As String
Dim WNroIbTucu As String
Dim WNombre As String
Dim WDomicilio As String
Dim WPuerta As String
Dim WLocalidad As String
Dim WProvincia As String
Dim WPostal As String
Dim Provincia(100) As String
Dim XTotal As String
Dim WOtros As Double
Dim XOtros As String
Dim WIva As Double
Dim XIva As String
Dim ZZAlicuota As Double

Dim WNroIbCiudad As String
Dim WNroIbCiudadII As String

Dim WRegimen As String
Dim WImporte As String

Dim XNOmbre As String


Private Sub Acepta_Click()

    WPasoUnifica = Trim(WPasoUnifica)
    ZZLargo = Len(WPasoUnifica)
    
    If Right$(WPasoUnifica, 3) = "xls" Then
    
        Rem WpasoII = Mid$(WPasoUnifica, 1, ZZLargo - 3) + "txt"
        Rem Open WpasoII For Output As #1
        
        ZNombre = XNOmbre + "aduana.Txt"
        Open ZNombre For Output As #1
    
        
        Set appExcel = CreateObject("Excel.application")
        
        ruta = WPasoUnifica
        LugarPlanilla = 1
    
        If Len(Dir(ruta)) > 0 Then
        
        
            Set objLibro = appExcel.workbooks.Open(ruta)
            
            Do
            
                LugarPlanilla = LugarPlanilla + 1
                
                WCuit = appExcel.cells(LugarPlanilla, 1).Value
                WRazon = appExcel.cells(LugarPlanilla, 2).Value
                WImpuesto = appExcel.cells(LugarPlanilla, 3).Value
                WDesImpuesto = appExcel.cells(LugarPlanilla, 4).Value
                WRegimen = appExcel.cells(LugarPlanilla, 5).Value
                WDesRegimen = appExcel.cells(LugarPlanilla, 6).Value
                WFecha = appExcel.cells(LugarPlanilla, 7).Value
                WNroCertificado = appExcel.cells(LugarPlanilla, 8).Value
                WDesOperacion = appExcel.cells(LugarPlanilla, 9).Value
                WImporte = appExcel.cells(LugarPlanilla, 10).Value
                WNroComprobante = appExcel.cells(LugarPlanilla, 11).Value
                WFechaComprob = appExcel.cells(LugarPlanilla, 12).Value
                WDesComprobante = appExcel.cells(LugarPlanilla, 13).Value
                
                If Trim(WCuit) = "" Then Exit Do
                
                
                
            
                Rem WCuit = !Cuit
                Rem WFecha = !Fecha
                Rem WRegimen = !regimen
                Call Ceros(WRegimen, 3)
                Rem WImporte = !Importe
                Call Ceros(WImporte, 15)
                WNroComprobante = Left$(WNroComprobante + Space$(30), 30)
                
            
        
                Rem Tipo de Operacion
                Campo1 = "2"
        
                Rem Cuit
                Campo2 = WCuit
        
                Rem fecha
                Campo3 = WFecha
        
                Rem Codigo de Regimen
                Campo4 = WRegimen
                
            
                Rem Importe
                Campo5 = WImporte
        
                Rem Numero del Comprobante
                Campo6 = WNroComprobante
        
                WImpre = Campo1 + Campo2 + Campo3 + Campo4 + Campo5 + Campo6 + Campo7 + Campo8 + Campo9 + Campo10 + Campo11 + Campo12 + Campo13 + Campo14 + Campo15 + Campo16 + Campo17 + Campo18 + Campo19 + Campo20 + Campo21
                
                
            
                Print #1, WImpre
                
                
                
            Loop
                
            appExcel.Quit
            Set appExcel = Nothing
            
        End If
        
        Close #1
        
        m$ = "Proceso Finalizado"
        A% = MsgBox(m$, 0, "Recupero de Percep. de IB Aduana")
        
        
    End If

End Sub

Private Sub Cancela_Click()
    PrgRecuperoPerce.Hide
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
    XNOmbre = WDir + "\"
    
    WPasoUnifica = XNOmbre + File1.filename
    
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
    XNOmbre = WDir + "\"
    
    WPasoUnifica = XNOmbre + File1.filename
    
    Exit Sub
    
WError:
    Rem MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    Resume Next
    
End Sub

Private Sub Form_Load()
    File1.Pattern = "*.xls"
End Sub
