VERSION 5.00
Begin VB.Form Prglee3 
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
Attribute VB_Name = "Prglee3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WLinea As String
Private WDeposito As String
Private WRecibo As String

Private Sub Acepta_Click()
 
    Call Proceso
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Prglee3.Hide
    Unload Me
    Menu.SetFocus
End Sub


Private Sub Proceso()

    On Error GoTo Error
    
    'bancos
        Stop
    coderr = 0
    
    Open "c:\prueba\adminis\" + WEmpresa + "ban.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
            
                WBanco = Mid$(WLinea, 4, 2)
                WDescripcion = Mid$(WLinea, 6, 30)
                WCuenta = Mid$(WLinea, 36, 13)
                    
                With rstBanco
                        .Index = "Banco"
                        .Seek "=", Val(WBanco)
                        If .NoMatch Then
                            .AddNew
                            !Banco = WBanco
                            !Nombre = WDescripcion
                            !Cuenta = WCuenta
                            !Empresa = 1
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Banco = WBanco
                            !Nombre = WDescripcion
                            !Cuenta = WCuenta
                            !Empresa = 1
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1

    'cuentas

    
    Open "c:\prueba\adminis\" + WEmpresa + "cuen.txt" For Input As #1
    
    coderr = 0
    
    Do
                Line Input #1, WLinea
            
                WCodigo = Mid$(WLinea, 1, 10)
                WDescripcion = Mid$(WLinea, 11, 40)
                    
                With rstCuenta
                        .Index = "Cuenta"
                        .Seek "=", WCodigo
                        If .NoMatch Then
                            .AddNew
                            !Cuenta = WCodigo
                            !Descripcion = WDescripcion
                            !nivel = 1
                            !Empresa = 1
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Cuenta = WCodigo
                            !Descripcion = WDescripcion
                            !nivel = 1
                            !Empresa = 1
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1


    'proveedor

    
    Open "c:\prueba\adminis\" + WEmpresa + "prv.txt" For Input As #1
    
    coderr = 0
    
    Do
                Line Input #1, WLinea
            
                WProveedor = Mid$(WLinea, 5, 11)
                WNombre = Mid$(WLinea, 16, 30)
                WDireccion = Mid$(WLinea, 46, 25)
                WLocalidad = Mid$(WLinea, 71, 20)
                WPostal = Mid$(WLinea, 91, 4)
                WProvincia = Mid$(WLinea, 95, 1)
                WTelefono = Mid$(WLinea, 96, 7)
                WIva = Mid$(WLinea, 103, 1)
                WCuit = Mid$(WLinea, 119, 15)
                WObservaciones = Mid$(WLinea, 149, 30)
                WCuenta = Mid$(WLinea, 179, 10)
                WEMail = ""
                WDias = ""
                WTipo = "1"
                WNombreCheque = ""
                    
                With rstProveedor
                        .Index = "proveedor"
                        .Seek "=", WProveedor
                        If .NoMatch Then
                            .AddNew
                            !Proveedor = WProveedor
                            !Nombre = WNombre
                            !Direccion = WDireccion
                            !Localidad = WLocalidad
                            !Postal = WPostal
                            !Cuit = WCuit
                            !Telefono = WTelefono
                            !email = WEMail
                            !Observaciones = WObservaciones
                            !Dias = WDias
                            !Tipo = WTipo
                            !Iva = WIva
                            !Provincia = WProvincia
                            !Cuenta = Left$(WCuenta, 10)
                            !NombreCheque = WNombreCheque
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Proveedor = WProveedor
                            !Nombre = WNombre
                            !Direccion = WDireccion
                            !Localidad = WLocalidad
                            !Postal = WPostal
                            !Cuit = WCuit
                            !Telefono = WTelefono
                            !email = WEMail
                            !Observaciones = WObservaciones
                            !Dias = WDias
                            !Tipo = WTipo
                            !Iva = WIva
                            !Provincia = WProvincia
                            !Cuenta = Left$(WCuenta, 10)
                            !NombreCheque = WNombreCheque
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1



    'ctacxte

    
    Open "c:\prueba\adminis\" + WEmpresa + "c.txt" For Input As #1
    
    coderr = 0
    
    Do
                Line Input #1, WLinea
                
                WProveedor = Mid$(WLinea, 1, 11)
                
                
                WTipo = Mid$(WLinea, 12, 2)
                WLetra = "A"
                WPunto = "0000"
                WNumero = Mid$(WLinea, 14, 6)
                WFecha = Mid$(WLinea, 24, 2) + "/" + Mid$(WLinea, 22, 2) + "/19" + Mid$(WLinea, 20, 2)
                WEstado = Mid$(WLinea, 26, 1)
                WTotal = Mid$(WLinea, 27, 10)
                WSaldo = Mid$(WLinea, 37, 10)
                If Mid$(WLinea, 41, 2) <> "00" Then
                    WVencimiento = Mid$(WLinea, 51, 2) + "/" + Mid$(WLinea, 49, 2) + "/19" + Mid$(WLinea, 47, 2)
                        Else
                    WVencimiento = "00/00/0000"
                End If

                WNroInterno = Mid$(WLinea, 53, 6)
                
                If Val(Mid$(WLinea, 63, 2)) <> 0 Then
                    WVencimiento1 = Mid$(WLinea, 63, 2) + "/" + Mid$(WLinea, 61, 2) + "/19" + Mid$(WLinea, 59, 2)
                        Else
                    WVencimiento1 = "00/00/0000"
                End If
                        
                WClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                WOrdfecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                WOrdvencimiento = Right$(WVencimiento, 4) + Mid$(WVencimiento, 4, 2) + Left$(WVencimiento, 2)
                    
                With rstCtaCtePrv
                        .Index = "ctacte"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Proveedor = WProveedor
                            !Letra = WLetra
                            !Tipo = WTipo
                            !Punto = WPunto
                            !Numero = WNumero
                            !Fecha = WFecha
                            !Estado = WEstado
                            !Vencimiento = WVencimiento
                            !Vencimiento1 = WVencimiento1
                            !NroInterno = Val(WNroInterno)
                            !Total = Val(WTotal)
                            !Saldo = Val(WSaldo)
                            !Clave = WProveedor + WLetra + WTipo + WPunto + WNumero
                            !OrdFecha = WOrdfecha
                            !OrdVencimiento = WOrdvencimiento
                            Select Case Val(WTipo)
                                Case 1
                                    !Impre = "FC"
                                Case 2
                                    !Impre = "ND"
                                Case 3
                                    !Impre = "NC"
                                    !Total = Abs(!Total) * -1
                                    !Saldo = Abs(!Saldo) * -1
                                Case 4
                                    !Impre = "AN"
                                    !Total = Abs(!Total) * -1
                                    !Saldo = Abs(!Saldo) * -1
                                Case 5
                                    !Impre = "OP"
                                Case 10
                                    !Impre = "FR"
                                Case Else
                                    !Impre = ""
                            End Select
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Proveedor = WProveedor
                            !Letra = WLetra
                            !Tipo = WTipo
                            !Punto = WPunto
                            !Numero = WNumero
                            !Fecha = WFecha
                            !Estado = WEstado
                            !Vencimiento = WVencimiento
                            !Vencimiento1 = WVencimiento1
                            !NroInterno = Val(WNroInterno)
                            !Total = Val(WTotal)
                            !Saldo = Val(WSaldo)
                            !Clave = WProveedor + WLetra + WTipo + WPunto + WNumero
                            !OrdFecha = WOrdfecha
                            !OrdVencimiento = WOrdvencimiento
                            Select Case Val(WTipo)
                                Case 1
                                    !Impre = "FC"
                                Case 2
                                    !Impre = "ND"
                                Case 3
                                    !Impre = "NC"
                                    !Total = Abs(!Total) * -1
                                    !Saldo = Abs(!Saldo) * -1
                                Case 4
                                    !Impre = "AN"
                                    !Total = Abs(!Total) * -1
                                    !Saldo = Abs(!Saldo) * -1
                                Case 5
                                    !Impre = "OP"
                                Case 10
                                    !Impre = "FR"
                                Case Else
                                    !Impre = ""
                            End Select
                           .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1


    'iva compras

    
    Open "c:\prueba\adminis\" + WEmpresa + "iva.txt" For Input As #1
    
    coderr = 0
    
    Do
                Line Input #1, WLinea
                
                WProveedor = Mid$(WLinea, 1, 11)
                Call Ceros(WProveedor, 11)
                WTipo = Mid$(WLinea, 12, 2)
                WNroInterno = Mid$(WLinea, 14, 6)
                WLetra = "A"
                WPunto = "0000"
                WFecha = Mid$(WLinea, 24, 2) + "/" + Mid$(WLinea, 22, 2) + "/19" + Mid$(WLinea, 20, 2)
                WNumero = Mid$(WLinea, 26, 6)
                WNeto = Val(Mid$(WLinea, 42, 10))
                WIva21 = Val(Mid$(WLinea, 52, 10))
                WIVa5 = Val(Mid$(WLinea, 62, 10))
                WIVa27 = 0
                WIB = 0
                WExento = Val(Mid$(WLinea, 72, 10))
                If Mid$(WLinea, 98, 2) = "00" Then
                    WVencimiento = Mid$(WLinea, 106, 2) + "/" + Mid$(WLinea, 100, 2) + "/20" + Mid$(WLinea, 98, 2)
                        Else
                    WVencimiento = Mid$(WLinea, 106, 2) + "/" + Mid$(WLinea, 100, 2) + "/19" + Mid$(WLinea, 98, 2)
                End If
                If Mid$(WLinea, 98, 2) = "00" Then
                    WVencimiento1 = Mid$(WLinea, 108, 2) + "/" + Mid$(WLinea, 106, 2) + "/20" + Mid$(WLinea, 104, 2)
                        Else
                    WVencimiento1 = Mid$(WLinea, 108, 2) + "/" + Mid$(WLinea, 106, 2) + "/19" + Mid$(WLinea, 104, 2)
                End If
                WPeriodo = WFecha
                
                Select Case Val(WTipo)
                    Case 1
                        WImpre = "FC"
                    Case 2
                        WImpre = "ND"
                    Case 3
                        WImpre = "NC"
                        WNeto = WNeto * -1
                        WIva21 = WIva21 * -1
                        WIVa5 = WIVa5 * -1
                        WIVa27 = WIVa27 * -1
                        WIB = WIB * -1
                        WExento = WExento * -1
                    Case Else
                        WImpre = "  "
                End Select
                WOrdfecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                WContado = "2"
                XEmpresa = 1
                
                With rstIvaComp
                        .Index = "NroInterno"
                        .Seek "=", Val(WNroInterno)
                        If .NoMatch Then
                            .AddNew
                            !NroInterno = WNroInterno
                            !Proveedor = WProveedor
                            !Tipo = WTipo
                            !Letra = WLetra
                            !Punto = WPunto
                            !Numero = WNumero
                            !Fecha = WFecha
                            !Vencimiento = WVencimiento
                            !Vencimiento1 = WVencimiento1
                            !Periodo = WPeriodo
                            !Neto = WNeto
                            !Iva21 = WIva21
                            !Iva5 = WIVa5
                            !Iva27 = WIVa27
                            !Ib = WIB
                            !Exento = WExento
                            !Impre = WImpre
                            !OrdFecha = WOrdfecha
                            !Contado = WContado
                            !Empresa = XEmpresa
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !NroInterno = WNroInterno
                            !Proveedor = WProveedor
                            !Tipo = WTipo
                            !Letra = WLetra
                            !Punto = WPunto
                            !Numero = WNumero
                            !Fecha = WFecha
                            !Vencimiento = WVencimiento
                            !Vencimiento1 = WVencimiento1
                            !Periodo = WPeriodo
                            !Neto = WNeto
                            !Iva21 = WIva21
                            !Iva5 = WIVa5
                            !Iva27 = WIVa27
                            !Ib = WIB
                            !Exento = WExento
                            !Impre = WImpre
                            !OrdFecha = WOrdfecha
                            !Contado = WContado
                            !Empresa = XEmpresa
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1


    'imputaciones DADA

    
    Open "c:\prueba\adminis\" + WEmpresa + "imp.txt" For Input As #1
    
    coderr = 0
    
    Do
                Line Input #1, WLinea
                
                WTipomovi = "2"
                WProveedor = Mid$(WLinea, 1, 11)
                Call Ceros(WProveedor, 11)
                WCuenta = Mid$(WLinea, 12, 13)
                WTipocomp = Mid$(WLinea, 25, 2)
                WNroInterno = Mid$(WLinea, 27, 6)
                WRenglon = Mid$(WLinea, 33, 2)
                WLetracomp = "A"
                WPuntocomp = "0000"
                WNrocomp = Mid$(WLinea, 41, 6)
                WFecha = Mid$(WLinea, 39, 2) + "/" + Mid$(WLinea, 37, 2) + "/19" + Mid$(WLinea, 35, 2)
                WObservaciones = ""
                Importe = Val(Mid$(WLinea, 47, 10))
                If Importe > 0 Then
                    WDebito = Importe
                    WCredito = 0
                        Else
                    WDebito = 0
                    WCredito = Abs(Importe)
                End If
                WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                WTitulo = "Compras"
                XEmpresa = 1
                WClave = WTipomovi + WNroInterno + WRenglon
                
                With rstImputac
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Tipomovi = WTipomovi
                            !NroInterno = WNroInterno
                            !Proveedor = WProveedor
                            !TipoComp = WTipocomp
                            !LetraComp = WLetracomp
                            !PuntoComp = WPuntocomp
                            !NroComp = WNrocomp
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Observaciones = WObservaciones
                            !Cuenta = Left$(WCuenta, 10)
                            !Debito = WDebito
                            !Credito = WCredito
                            !FechaOrd = WFechaord
                            !Titulo = WTitulos
                            !Empresa = XEmpresa
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Tipomovi = WTipomovi
                            !NroInterno = WNroInterno
                            !Proveedor = WProveedor
                            !TipoComp = WTipocomp
                            !LetraComp = WLetracomp
                            !PuntoComp = WPuntocomp
                            !NroComp = WNrocomp
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Observaciones = WObservaciones
                            !Cuenta = Left$(WCuenta, 10)
                            !Debito = WDebito
                            !Credito = WCredito
                            !FechaOrd = WFechaord
                            !Titulo = WTitulos
                            !Empresa = XEmpresa
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1

 Stop
    'deposito

    
    Open "c:\prueba\adminis\" + WEmpresa + "depo.txt" For Input As #1
    
    coderr = 0
    
    Do
                Line Input #1, WLinea
                
                If Mid$(WLinea, 1, 1) = "2" And Mid$(WLinea, 15, 1) = "2" Then
                
                    WDeposito = Mid$(WLinea, 2, 5)
                    Call Ceros(WDeposito, 6)
                    WRenglon = Mid$(WLinea, 7, 2)
                    WBanco = 1
                    WImporte = 0
                    WFecha = Mid$(WLinea, 9, 2) + "/" + Mid$(WLinea, 11, 2) + "/19" + Mid$(WLinea, 13, 2)
                    WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    If Mid$(WLinea, 36, 2) = "00" Then
                        WAcredita = Mid$(WLinea, 32, 2) + "/" + Mid$(WLinea, 34, 2) + "/20" + Mid$(WLinea, 36, 2)
                            Else
                        WAcredita = Mid$(WLinea, 32, 2) + "/" + Mid$(WLinea, 34, 2) + "/19" + Mid$(WLinea, 36, 2)
                    End If
                    WAcreditaord = Right$(WAcredita, 4) + Mid$(WAcredita, 4, 2) + Left$(WAcredita, 2)
                    WNUmero2 = Mid$(WLinea, 16, 6)
                    If Val(WNUmero2) = 0 Then
                        WTipo2 = 1
                            Else
                        WTipo2 = 2
                    End If
                    If Mid$(WLinea, 36, 2) = "00" Then
                        WFecha2 = Mid$(WLinea, 32, 2) + "/" + Mid$(WLinea, 34, 2) + "/20" + Mid$(WLinea, 36, 2)
                            Else
                        WFecha2 = Mid$(WLinea, 32, 2) + "/" + Mid$(WLinea, 34, 2) + "/19" + Mid$(WLinea, 36, 2)
                    End If
                    WObservaciones2 = ""
                    WImporte2 = Mid$(WLinea, 22, 10)
                    XEmpresa = 1
                    WClave = WDeposito + WRenglon
                
                    With rstDepositos
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Deposito = WDeposito
                            !Renglon = WRenglon
                            !Banco = WBanco
                            !Importe = WImporte
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Acredita = WAcredita
                            !AcreditaOrd = WAcreditaord
                            !Tipo2 = WTipo2
                            !Numero2 = WNUmero2
                            !Fecha2 = WFecha2
                            !Observaciones2 = WObservaciones
                            !Importe2 = Val(WImporte2)
                            !Empresa = 1
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Deposito = WDeposito
                            !Renglon = WRenglon
                            !Banco = WBanco
                            !Importe = WImporte
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Acredita = WAcredita
                            !AcreditaOrd = WAcreditaord
                            !Tipo2 = WTipo2
                            !Numero2 = WNUmero2
                            !Fecha2 = WFecha2
                            !Observaciones2 = WObservaciones
                            !Importe2 = Val(WImporte2)
                            !Empresa = 1
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
            End If
            If EOF(1) Then Exit Do
    Loop
    
    Close #1


    'recibos

    
    Open "c:\prueba\adminis\" + WEmpresa + "rec.txt" For Input As #1
    
    coderr = 0
    
    Do
                Line Input #1, WLinea
                
                If Mid$(WLinea, 1, 1) = "0" Or Mid$(WLinea, 1, 1) = "1" Then
                
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
                            !FechaOrd = WFechaor
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
                            !FechaOrd = WFechaor
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


    'orden de pago

    
    Open "c:\prueba\adminis\" + WEmpresa + "ord.txt" For Input As #1
    
    coderr = 0
    
    Do
                Line Input #1, WLinea
                
                WOrden = Mid$(WLinea, 1, 6)
                WRenglon = Mid$(WLinea, 7, 2)
                WProveedor = Mid$(WLinea, 9, 11)
                WFecha = Mid$(WLinea, 24, 2) + "/" + Mid$(WLinea, 22, 2) + "/19" + Mid$(WLinea, 20, 2)
                WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                WImporte = Val(Mid$(WLinea, 26, 9))
                WObservaciones = Mid$(WLinea, 35, 50)
                WTipoord = Mid$(WLinea, 85, 1)
                WTiporeg = Mid$(WLinea, 86, 1)
                WTipo1 = Mid$(WLinea, 87, 2)
                WLetra1 = "A"
                WPUnto1 = "0000"
                WNumero1 = Mid$(WLinea, 89, 6)
                WImporte1 = Val(Mid$(WLinea, 95, 9))
                WObservaciones2 = Mid$(WLinea, 104, 20)
                If WObservaciones2 = Null Then
                    WObservaciones2 = ""
                End If
                WTipo2 = Mid$(WLinea, 124, 2)
                WNUmero2 = Mid$(WLinea, 126, 6)
                WFecha2 = Mid$(WLinea, 132, 2) + "/" + Mid$(WLinea, 134, 2) + "/19" + Mid$(WLinea, 136, 2)
                WFechaord2 = Right$(WFecha2, 4) + Mid$(WFecha2, 4, 2) + Left$(WFecha2, 2)
                WBanco2 = Mid$(WLinea, 138, 2)
                WImporte2 = Val(Mid$(WLinea, 140, 9))
                XEmpresa = 1
                WClave = WOrden + WRenglon
                
                With rstPagos
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Orden = WOrden
                            !Renglon = WRenglon
                            !Proveedor = WProveedor
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Importe = WImporte
                            !Observaciones = WObservaciones
                            !TipoOrd = WTipoord
                            !Tiporeg = WTiporeg
                            !Tipo1 = WTipo1
                            !Letra1 = WLetra1
                            !Punto1 = WPUnto1
                            !Numero1 = WNumero1
                            !Importe1 = WImporte1
                            !Observaciones2 = WObservaciones2
                            !Tipo2 = WTipo2
                            !Numero2 = WNUmero2
                            !Fecha2 = WFecha2
                            !FechaOrd2 = WFechaord2
                            !banco2 = WBanco2
                            !Importe2 = WImporte2
                            !Empresa = 1
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Orden = WOrden
                            !Renglon = WRenglon
                            !Proveedor = WProveedor
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Importe = WImporte
                            !Observaciones = WObservaciones
                            !TipoOrd = WTipoord
                            !Tiporeg = WTiporeg
                            !Tipo1 = WTipo1
                            !Letra1 = WLetra1
                            !Punto1 = WPUnto1
                            !Numero1 = WNumero1
                            !Importe1 = WImpore1
                            !Observaciones2 = WObserbvaciones2
                            !Tipo2 = WTipo2
                            !Numero2 = WNUmero2
                            !Fecha2 = WFecha2
                            !FechaOrd2 = WFechaord2
                            !banco2 = WBanco2
                            !Importe2 = WImporte2
                            !Empresa = 1
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




