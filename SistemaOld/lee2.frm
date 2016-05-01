VERSION 5.00
Begin VB.Form Prglee2 
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
Attribute VB_Name = "Prglee2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WLinea As String
Private WImporte As Double
Private WTeorico As Double
Private WReal As Double
Private WCantidad As Double

Private Sub Acepta_Click()
 
    Call Proceso
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Prglee2.Hide
    Unload Me
    Menu.SetFocus
End Sub


Private Sub Proceso()

    On Error GoTo Error
    
Stop

   Rem HOJA DE PRODUCCION
        
    'ensayos
        
    coderr = 0
    
    Open "c:\prueba\cotiza\" + WEmpresa + "hoja.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
        
                WHoja = Mid$(WLinea, 1, 6)
                WRenglon = Mid$(WLinea, 7, 2)
                WFecha = Mid$(WLinea, 9, 2) + "/" + Mid$(WLinea, 11, 2) + "/19" + Mid$(WLinea, 13, 2)
                WProducto = Mid$(WLinea, 40, 2) + "-" + Mid$(WLinea, 42, 5) + "-" + Mid$(WLinea, 47, 3)
                WCantidad = Val(Mid$(WLinea, 25, 8))
                WTipo = Mid$(WLinea, 33, 1)
                WLote = Val(Mid$(WLinea, 34, 6))
                If WTipo = "M" Then
                    WArticulo = Mid$(WLinea, 15, 2) + "-" + Mid$(WLinea, 17, 3) + "-" + Mid$(WLinea, 22, 3)
                    WTerminado = "  -     -   "
                        Else
                    WArticulo = "  -   -   "
                    WTerminado = Mid$(WLinea, 15, 2) + "-" + Mid$(WLinea, 17, 5) + "-" + Mid$(WLinea, 22, 3)
                End If
                WTeorico = Val(Mid$(WLinea, 50, 8))
                WReal = Val(Mid$(WLinea, 58, 8))
                If Val(Mid$(WLinea, 66, 6)) = 0 Then
                    WfechaIngord = Space$(8)
                    WfechaIng = "  /  /    "
                        Else
                    WfechaIng = Mid$(WLinea, 66, 2) + "/" + Mid$(WLinea, 68, 2) + "/19" + Mid$(WLinea, 70, 2)
                    WfechaIngord = Right$(WfechaIng, 4) + Mid$(WfechaIng, 4, 2) + Left$(WfechaIng, 2)
                End If
                WClave = WHoja + WRenglon
                Call Redondeo(WTeorico)
                Call Redondeo(WReal)
                Call Redondeo(WReal)
                
                With rstHoja
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Hoja = WHoja
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Producto = WProducto
                            !Teorico = WTeorico
                            !Real = WReal
                            !fechaIng = WfechaIng
                            !fechaIngord = WfechaIngord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Lote = WLote
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Hoja = WHoja
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Producto = WProducto
                            !Teorico = WTeorico
                            !Real = WReal
                            !fechaIng = WfechaIng
                            !fechaIngord = WfechaIngord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Lote = WLote
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1
                
   Rem cotizaciones
   
    coderr = 0
    
    Open "c:\prueba\cotiza\" + WEmpresa + "cot.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
   
                WCotiza = Mid$(WLinea, 1, 6)
                WRenglon = Mid$(WLinea, 7, 2)
                WFecha = Mid$(WLinea, 13, 2) + "/" + Mid$(WLinea, 11, 2) + "/19" + Mid$(WLinea, 9, 2)
                WProveedor = Mid$(WLinea, 15, 11)
                WArticulo = Mid$(WLinea, 26, 2) + "-" + Mid$(WLinea, 28, 3) + "-" + Mid$(WLinea, 31, 3)
                WPrecio = Val(Mid$(WLinea, 34, 8))
                WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                WCondicion = Mid$(WLinea, 42, 15)
                WObservaciones = Mid$(WLinea, 57, 20)
                WClave = WCotiza + WRenglon
                    
                With rstCotiza
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Cotiza = WCotiza
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Proveedor = WProveedor
                            !Articulo = WArticulo
                            !Precio = WPrecio
                            !FechaOrd = WFechaord
                            !Condicion = WCondicion
                            !Observaciones = WObservaciones
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Cotiza = WCotiza
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Proveedor = WProveedor
                            !Articulo = WArticulo
                            !Precio = WPrecio
                            !FechaOrd = WFechaord
                            !Condicion = WCondicion
                            !Observaciones = WObservaciones
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1
                
    
   Rem INFORME DE RECEPCION
        
   
    coderr = 0
    
    Open "c:\prueba\cotiza\" + WEmpresa + "inf.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
   
                WInforme = Mid$(WLinea, 1, 6)
                WRenglon = Mid$(WLinea, 7, 2)
                WFecha = Mid$(WLinea, 9, 2) + "/" + Mid$(WLinea, 11, 2) + "/19" + Mid$(WLinea, 13, 2)
                WRemito = Mid$(WLinea, 15, 6)
                WProveedor = Mid$(WLinea, 21, 11)
                WOrden = Mid$(WLinea, 32, 6)
                WArticulo = Mid$(WLinea, 38, 2) + "-" + Mid$(WLinea, 40, 3) + "-" + Mid$(WLinea, 43, 3)
                WCantidad = Val(Mid$(WLinea, 46, 8))
                WResta = Val(Mid$(WLinea, 54, 8))
                WClave = WInforme + WRenglon
                WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
            
                With rstInforme
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Informe = WInforme
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Proveedor = WProveedor
                            !Remito = WRemito
                            !Orden = WOrden
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Resta = WResta
                            !Clave = WClave
                            !FechaOrd = WFechaord
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Informe = WInforme
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Proveedor = WProveedor
                            !Remito = WRemito
                            !Orden = WOrden
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Resta = WResta
                            !Clave = WClave
                            !FechaOrd = WFechaord
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1
    
   Rem ORDENES DE COMPRA
   
    coderr = 0
    
    Open "c:\prueba\cotiza\" + WEmpresa + "ord.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
                
                WOrden = Mid$(WLinea, 1, 6)
                WRenglon = Mid$(WLinea, 7, 2)
                WFecha = Mid$(WLinea, 9, 2) + "/" + Mid$(WLinea, 11, 2) + "/19" + Mid$(WLinea, 13, 2)
                WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                WProveedor = Mid$(WLinea, 15, 11)
                WArticulo = Mid$(WLinea, 26, 2) + "-" + Mid$(WLinea, 28, 3) + "-" + Mid$(WLinea, 31, 3)
                WCantidad = Val(Mid$(WLinea, 34, 8))
                WPrecio = Val(Mid$(WLinea, 42, 8))
                WFecha1 = Mid$(WLinea, 50, 2) + "/" + Mid$(WLinea, 52, 2) + "/19" + Mid$(WLinea, 54, 2)
                WFecha2 = Mid$(WLinea, 56, 2) + "/" + Mid$(WLinea, 58, 2) + "/19" + Mid$(WLinea, 60, 2)
                WCondicion = Mid$(WLinea, 62, 15)
                WRecibida = Val(Mid$(WLinea, 77, 8))
                WClave = WOrden + WRenglon
                
                With rstOrden
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Orden = WOrden
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Proveedor = WProveedor
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Precio = WPrecio
                            !Fecha1 = WFecha1
                            !Fecha2 = WFecha2
                            !Condicion = WCondicion
                            !Recibida = WRecibida
                            !Clave = WClave
                            !Fechaentrega = "  /  /    "
                            !Saldo = 0
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Orden = WOrden
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Proveedor = WProveedor
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Precio = WPrecio
                            !Fecha1 = WFecha1
                            !Fecha2 = WFecha2
                            !Condicion = WCondicion
                            !Recibida = WRecibida
                            !Fechaentrega = "  /  /    "
                            !Saldo = 0
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1


    
   Rem LAUDO DE LIBERACION
   
    coderr = 0
    
    Open "c:\prueba\cotiza\" + WEmpresa + "lau.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
                
                WLaudo = Mid$(WLinea, 1, 6)
                WRenglon = Mid$(WLinea, 7, 2)
                WFecha = Mid$(WLinea, 9, 2) + "/" + Mid$(WLinea, 11, 2) + "/19" + Mid$(WLinea, 13, 2)
                WArticulo = Mid$(WLinea, 15, 2) + "-" + Mid$(WLinea, 17, 3) + "-" + Mid$(WLinea, 20, 3)
                WLiberada = Val(Mid$(WLinea, 23, 8))
                WDevuelta = Val(Mid$(WLinea, 31, 8))
                WOrden = Mid$(WLinea, 39, 6)
                WMarca = Mid$(WLinea, 45, 1)
                WLote = Mid$(WLinea, 46, 6)
                WRechazo = Mid$(WLinea, 52, 6)
                WInforme = Mid$(WLinea, 58, 6)
                WActualiza = Mid$(WLinea, 64, 1)
                WClave = WLaudo + WRenglon
                
                With rstLaudo
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Laudo = WLaudo
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Orden = Val(WOrden)
                            !Articulo = WArticulo
                            !Liberada = WLiberada
                            !Devuelta = WDevuelta
                            !Liberada = WLiberada
                            !Lote = Val(WLote)
                            !Rechazo = Val(WRechazo)
                            !Actualiza = WNuevo
                            !Marca = WMarca
                            !Informe = Val(WInforme)
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Laudo = WLaudo
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Orden = Val(WOrden)
                            !Articulo = WArticulo
                            !Liberada = WLiberada
                            !Devuelta = WDevuelta
                            !Liberada = WLiberada
                            !Lote = Val(WLote)
                            !Rechazo = Val(WRechazo)
                            !Actualiza = WNuevo
                            !Marca = WMarca
                            !Informe = Val(WInforme)
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1

    
   Rem moviimentos varios
   
    coderr = 0
    
    Open "c:\prueba\cotiza\" + WEmpresa + "var.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
                
                WCodigo = Mid$(WLinea, 1, 1) + Mid$(WLinea, 3, 5)
                WRenglon = Mid$(WLinea, 8, 2)
                WFecha = Mid$(WLinea, 10, 2) + "/" + Mid$(WLinea, 12, 2) + "/19" + Mid$(WLinea, 14, 2)
                WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                WTipo = Mid$(WLinea, 16, 1)
                If WTipo = "M" Then
                    WArticulo = Mid$(WLinea, 17, 2) + "-" + Mid$(WLinea, 19, 3) + "-" + Mid$(WLinea, 24, 3)
                    WTerminado = "  -     -   "
                        Else
                    WArticulo = "  -   -   "
                    WTerminado = Mid$(WLinea, 17, 2) + "-" + Mid$(WLinea, 19, 5) + "-" + Mid$(WLinea, 24, 3)
                End If
                WCantidad = Val(Mid$(WLinea, 27, 8))
                Select Case Val(Mid$(WLinea, 1, 1))
                    Case 1
                        WTipomov = 3
                    Case 2
                        WTipomov = 4
                    Case 3
                        WTipomov = 1
                    Case 4
                        WTipomov = 2
                    Case Else
                        Stop
                End Select
                Select Case Val(WTipomov)
                    Case 1, 3
                        WMovi = "E"
                    Case Else
                        WMovi = "S"
                End Select
                WObservaciones = Mid$(WLinea, 35, 50)
                WClave = WCodigo + WRenglon
                
                With rstMovvar
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Codigo = WCodigo
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Movi = WMovi
                            !Tipomov = WTipomov
                            !Observaciones = WObservaciones
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = WCodigo
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Movi = WMovi
                            !Tipomov = WTipomov
                            !Observaciones = WObservaciones
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
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




