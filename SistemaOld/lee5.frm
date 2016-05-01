VERSION 5.00
Begin VB.Form Prglee5 
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
Attribute VB_Name = "Prglee5"
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
Stop

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
        Wmovimiento = Mid$(Linea, 26, 1)
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
                !Movimiento = Wmovimiento
                !fechaord = WFechaord
                !Clave = WClave
                
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.


Rem terminado DADA


    Open "c:\prueba\ventas\" + WEmpresa + "ter.txt" For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, Linea
        
        WCodigo = Mid$(Linea, 1, 12)
        WLinea = Val(Mid$(Linea, 13, 4))
        WUnidad = Mid$(Linea, 18, 5)
        WInicial = Val(Mid$(Linea, 29, 10))
        WEntradas = Val(Mid$(Linea, 39, 11))
        WSalidas = Val(Mid$(Linea, 50, 11))
        WMinimo = Val(Mid$(Linea, 61, 11))
        WDeposito = ""
        WProceso = Val(Mid$(Linea, 72, 11))
        WPedido = Val(Mid$(Linea, 83, 11))
        WEnvase1 = Val(Mid$(Linea, 95, 3))
        WEnvase2 = Val(Mid$(Linea, 99, 3))
        WEnvase3 = Val(Mid$(Linea, 103, 3))
        WEnvase4 = Val(Mid$(Linea, 107, 3))
        WEnvase5 = Val(Mid$(Linea, 111, 3))
        WEnvase6 = Val(Mid$(Linea, 115, 3))
        WEnvase = Val(Mid$(Linea, 119, 3))
        WDescripcion = Mid$(Linea, 135, 30)
        
        With rstTerminado
        
            .Index = "Codigo"
            .Seek "=", WCodigo
            If .NoMatch Then
                .AddNew
                !Codigo = WCodigo
                !Descripcion = WDescripcion
                !Linea = WLinea
                !Unidad = WUnidad
                !Inicial = Val(WInicial)
                !Entradas = Val(WEntradas)
                !Salidas = Val(WSalidas)
                !MInimo = Val(WMinimo)
                !Deposito = ""
                !Pedido = WPedido
                Rem !Envase = WEnvase
                !Envase1 = WEnvase1
                !Envase2 = WEnvase2
                !Envase3 = WEnvase3
                !Envase4 = WEnvase4
                !Envase5 = WEnvase5
                !Envase6 = WEnvase6
                !Proceso = WProceso
                .Update
                    Else
                .Edit
                !Codigo = WCodigo
                !Descripcion = WDescripcion
                !Linea = WLinea
                !Unidad = WUnidad
                !Inicial = Val(WInicial)
                !Entradas = Val(WEntradas)
                !Salidas = Val(WSalidas)
                !MInimo = Val(WMinimo)
                !Deposito = ""
                !Pedido = WPedido
                Rem !Envase = WEnvase
                !Envase1 = WEnvase1
                !Envase2 = WEnvase2
                !Envase3 = WEnvase3
                !Envase4 = WEnvase4
                !Envase5 = WEnvase5
                !Envase6 = WEnvase6
                !Proceso = WProceso
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.


    'recibos

    
    Open "c:\prueba\adminis\" + WEmpresa + "CYA.txt" For Input As #1
    
    coderr = 0
    
    Do
                Line Input #1, WLinea
                
                If Mid$(WLinea, 1, 1) = "1" Then
                
                    WRecibo = Mid$(WLinea, 2, 5)
                    Call Ceros(WRecibo, 6)
                    WRenglon = Mid$(WLinea, 7, 2)
                    WFecha = Mid$(WLinea, 9, 2) + "/" + Mid$(WLinea, 11, 2) + "/19" + Mid$(WLinea, 13, 2)
                    WTiporec = "1"
                    WCliente = Mid$(WLinea, 16, 6)
                    WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    WRetganancias = 0
                    WRetIva = 0
                    WRetOtra = 0
                    WRetencion = 0
                    If Mid$(WLinea, 15, 1) = "2" Then
                        WTiporeg = "1"
                        WTipo1 = Mid$(WLinea, 27, 2)
                        WLetra1 = "A"
                        WPUnto1 = "0000"
                        WNumero1 = Mid$(WLinea, 29, 6)
                        WImporte1 = Val(Mid$(WLinea, 35, 10))
                        WTipo2 = ""
                        WNUmero2 = ""
                        WFecha2 = ""
                        WFechaord2 = ""
                        WBanco2 = ""
                        WImporte2 = 0
                        WEstado2 = ""
                        WObservaciones = ""
                        XEmpresa = 1
                        WClave = WRecibo + WRenglon
                        WImporte = 0
                            Else
                        WTiporeg = "2"
                        WTipo1 = ""
                        WLetra1 = ""
                        WPUnto1 = ""
                        WNumero1 = ""
                        WImporte1 = 0
                        WTipo2 = Mid$(WLinea, 27, 2)
                        WNUmero2 = Mid$(WLinea, 29, 6)
                        If Mid$(WLinea, 49, 2) = "00" Then
                            WFecha2 = Mid$(WLinea, 45, 2) + "/" + Mid$(WLinea, 47, 2) + "/20" + Mid$(WLinea, 49, 2)
                                Else
                            WFecha2 = Mid$(WLinea, 45, 2) + "/" + Mid$(WLinea, 47, 2) + "/19" + Mid$(WLinea, 49, 2)
                        End If
                        WFechaord2 = Right$(WFecha2, 4) + Mid$(WFecha2, 4, 2) + Left$(WFecha2, 2)
                        WBanco2 = Mid$(WLinea, 52, 100)
                        WImporte2 = Val(Mid$(WLinea, 35, 10))
                        WEstado2 = Mid$(WLinea, 51, 1)
                        WObservaciones = ""
                        XEmpresa = 1
                        WClave = WRecibo + WRenglon
                        WImporte = 0
                    End If
                
                    With rstRecibos
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Recibo = WRecibo
                            !Renglon = WRenglon
                            !Cliente = Left$(WCliente, 6)
                            !Fecha = WFecha
                            !fechaord = WFechaor
                            !Tiporec = WTiporec
                            !Retganancias = WRetganancias
                            !RetIva = WRetIva
                            !RetOtra = WRetOtra
                            !Retencion = WRetencion
                            !Tiporeg = WTiporeg
                            !Tipo1 = WTipo1
                            !Letra1 = WLetra1
                            !Punto1 = WPUnto1
                            !Numero1 = WNumero1
                            !Importe1 = WImporte1
                            !Tipo2 = WTipo2
                            !Numero2 = WNUmero2
                            !Fecha2 = WFecha2
                            !FechaOrd2 = WFechaord2
                            !banco2 = WBanco2
                            !Importe2 = WImporte2
                            !Estado2 = WEstado2
                            !Observaciones = WObservaciones
                            !Empresa = 1
                            !Clave = WClave
                            !Importe = WImporte
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Recibo = WRecibo
                            !Renglon = WRenglon
                            !Cliente = Left$(WCliente, 6)
                            !Fecha = WFecha
                            !fechaord = WFechaor
                            !Tiporec = WTiporec
                            !Retganancias = WRetganancias
                            !RetIva = WRetIva
                            !RetOtra = WRetOtra
                            !Retencion = WRetencion
                            !Tiporeg = WTiporeg
                            !Tipo1 = WTipo1
                            !Letra1 = WLetra1
                            !Punto1 = WPUnto1
                            !Numero1 = WNumero1
                            !Importe1 = WImporte1
                            !Tipo2 = WTipo2
                            !Numero2 = WNUmero2
                            !Fecha2 = WFecha2
                            !FechaOrd2 = WFechaord2
                            !banco2 = WBanco2
                            !Importe2 = WImporte2
                            !Estado2 = WEstado2
                            !Observaciones = WObservaciones
                            !Empresa = 1
                            !Clave = WClave
                            !Importe = WImporte
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
            End If
            If EOF(1) Then Exit Do
    Loop
    
    Close #1





    Call Cancela_click


    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Exit Sub
    
    
Error:
Stop
     coderr = Err
     
     Resume Next
     
End Sub




