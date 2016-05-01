VERSION 5.00
Begin VB.Form Prglee7 
   Caption         =   "Generacion de traspaso de datos"
   ClientHeight    =   4620
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4620
   ScaleWidth      =   6390
   Begin VB.Frame Frame2 
      Caption         =   "Control de Grabacion"
      Height          =   1815
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   4215
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cerrar"
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "Prglee7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WLinea As String

Private Sub Acepta_Click()
 
    Call Proceso
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Prglee5.Hide
    Unload Me
    Menu.SetFocus
End Sub


Private Sub Proceso()

    Rem On Error GoTo Error

Rem movimientos de envases dada


    Open "c:\prueba\ventas\" + WEmpresa + "men.txt" For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, Linea
        
        WTipo = Mid$(Linea, 25, 1)
        WCodigo = Mid$(Linea, 1, 6)
        WRenglon = Mid$(Linea, 7, 2)
        WCliente = Mid$(Linea, 27, 6)
        WFecha = Mid$(Linea, 13, 2) + "/" + Mid$(Linea, 11, 2) + "/19" + Mid$(Linea, 9, 2)
        WEnvase = Val(Mid$(Linea, 15, 3))
        WCantidad = Val(Mid$(Linea, 19, 6))
        WMovimiento = Mid$(Linea, 26, 1)
        If WMovimiento = "E" Then
            WMovimiento = "S"
                Else
            WMovimiento = "E"
        End If
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        WClave = WTipo + WCodigo + WRenglon
        
        With rstMovEnv
        
            .Index = "clave"
            .Seek "=", WClave
            If .NoMatch Then
                .AddNew
                
                !Tipo = WTipo
                !Codigo = WCodigo
                !Renglon = WRenglon
                !Cliente = WCliente
                !Fecha = WFecha
                !Envase = WEnvase
                !Cantidad = WCantidad
                !Movimiento = WMovimiento
                !fechaord = WFechaord
                !Clave = WClave
                
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.



    Call Cancela_click


    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Exit Sub
    
    
Error:
Stop
     coderr = Err
     
     Resume Next
     
End Sub




