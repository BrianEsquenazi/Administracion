VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProcesoCiudad 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Traspaso de Percepcion y Retencion de la Ciudad de Bs As."
   ClientHeight    =   5775
   ClientLeft      =   3060
   ClientTop       =   1425
   ClientWidth     =   7290
   LinkTopic       =   "Form2"
   ScaleHeight     =   5775
   ScaleWidth      =   7290
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   4815
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   6
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   5
         Top             =   4200
         Width           =   1215
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   1920
         TabIndex        =   3
         Top             =   1920
         Width           =   2055
      End
      Begin VB.DriveListBox Drive 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   1440
         Width           =   2055
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1920
         TabIndex        =   4
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Desde fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
   End
End
Attribute VB_Name = "PrgProcesoCiudad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstClientes As Recordset
Dim spClientes As String
Dim rstCtaCte As Recordset
Dim spCtaCte As String
Dim rstRecibo As Recordset
Dim spRecibo As String
Dim XParam As String
Dim Vector(10000, 15) As String
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
Dim WIbCiudad As String

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


Private Sub Drive_Change()
    Dir1.Path = Drive.Drive
End Sub

Private Sub Acepta_Click()

    WDrive = Drive.Drive
    WDir = Dir1.Path
    
    XNombre = WDir + "\" + "ciudadII.Txt"
    Open XNombre For Output As #1
    
    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    
    GoTo dada
    
    
    
    
    spCtaCte = "ModificaCtacteImporteIva0"
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem Procesa las cobranzas
    
    Renglon = 0
    Erase Vector
    
    XParam = "'" + WDesde + "','" _
                 + WHasta + "'"
    spRecibo = "ListaRecibosDifeI" + XParam
    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibo.RecordCount > 0 Then
        With rstRecibo
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstRecibo!Clave
                    Vector(Renglon, 2) = rstRecibo!Fecha
                    Vector(Renglon, 3) = rstRecibo!Tipo1
                    Vector(Renglon, 4) = rstRecibo!Numero1
                    Vector(Renglon, 5) = rstRecibo!Cliente
                    Vector(Renglon, 6) = ""
                    Vector(Renglon, 7) = ""
                    Vector(Renglon, 8) = ""
                    Vector(Renglon, 9) = ""
                    Vector(Renglon, 10) = ""
                    Vector(Renglon, 11) = ""
                    Vector(Renglon, 12) = ""
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstRecibo.Close
    End If
     
    For Cicla = 1 To Renglon
    
        WClave = Vector(Cicla, 1)
        WFecha = Vector(Cicla, 2)
        WTipo = Vector(Cicla, 3)
        WNumero = Vector(Cicla, 4)
        
        ClaveCtacte = WTipo + WNumero + "01"
        spCtaCte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCte.RecordCount > 0 Then
            WNeto = IIf(IsNull(rstCtaCte!Neto), "0", rstCtaCte!Neto)
            WIva1 = IIf(IsNull(rstCtaCte!Iva1), "0", rstCtaCte!Iva1)
            WIva2 = IIf(IsNull(rstCtaCte!Iva2), "0", rstCtaCte!Iva2)
            WImpoIb = IIf(IsNull(rstCtaCte!impoib), "0", rstCtaCte!impoib)
            WImpoIbTucu = IIf(IsNull(rstCtaCte!ImpoIbTucu), "0", rstCtaCte!ImpoIbTucu)
            WImpoIbCiudad = IIf(IsNull(rstCtaCte!ImpoIbCiudad), "0", rstCtaCte!ImpoIbCiudad)
            WTotal = IIf(IsNull(rstCtaCte!Total), "0", rstCtaCte!Total)
            If WImpoIbCiudad = 0 Then
                Vector(Cicla, 1) = ""
                Vector(Cicla, 2) = ""
                Vector(Cicla, 3) = ""
                Vector(Cicla, 4) = ""
                Vector(Cicla, 5) = ""
                Vector(Cicla, 6) = ""
                Vector(Cicla, 7) = ""
                Vector(Cicla, 8) = ""
                Vector(Cicla, 9) = ""
                Vector(Cicla, 10) = ""
                Vector(Cicla, 11) = ""
                Vector(Cicla, 12) = ""
                    Else
                Vector(Cicla, 6) = Str$(WNeto)
                Vector(Cicla, 7) = Str$(WIva1)
                Vector(Cicla, 8) = Str$(WIva2)
                Vector(Cicla, 9) = Str$(WImpoIb)
                Vector(Cicla, 10) = Str$(WImpoIbTucu)
                Vector(Cicla, 11) = Str$(WImpoIbCiudad)
                Vector(Cicla, 12) = Str$(WTotal)
            End If
            rstCtaCte.Close
                Else
            Vector(Cicla, 1) = ""
            Vector(Cicla, 2) = ""
            Vector(Cicla, 3) = ""
            Vector(Cicla, 4) = ""
            Vector(Cicla, 5) = ""
            Vector(Cicla, 6) = ""
            Vector(Cicla, 7) = ""
            Vector(Cicla, 8) = ""
            Vector(Cicla, 9) = ""
            Vector(Cicla, 10) = ""
            Vector(Cicla, 11) = ""
            Vector(Cicla, 12) = ""
        End If
        
    Next Cicla
    
    For Cicla = 1 To Renglon
    
        WClave = Vector(Cicla, 1)
        If WClave <> "" Then
        
            WTipo = Vector(Cicla, 3)
            WNumero = Vector(Cicla, 4)
            WRecibo = Val(Left$(WClave, 6))
            WSale = "N"
        
            XParam = "'" + WTipo + "','" _
                         + WNumero + "'"
            spRecibo = "ListaRecibosFactura " + XParam
            Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibo.RecordCount > 0 Then
                With rstRecibo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Val(rstRecibo!Recibo) < Val(WRecibo) Then
                                WSale = "S"
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstRecibo.Close
            End If
            
            If WSale = "S" Then
                Vector(Cicla, 1) = ""
                Vector(Cicla, 2) = ""
                Vector(Cicla, 3) = ""
                Vector(Cicla, 4) = ""
                Vector(Cicla, 5) = ""
                Vector(Cicla, 6) = ""
                Vector(Cicla, 7) = ""
                Vector(Cicla, 8) = ""
                Vector(Cicla, 9) = ""
                Vector(Cicla, 10) = ""
                Vector(Cicla, 11) = ""
                Vector(Cicla, 12) = ""
            End If
            
        End If
        
    Next Cicla
    
    
    For Cicla = 1 To Renglon
    
        WClave = Vector(Cicla, 1)
        
        If WClave <> "" Then
        
            WRecibo = "00" + Left$(Vector(Cicla, 1), 6)
        
            WClave = Vector(Cicla, 1)
            WFecha = Vector(Cicla, 2)
            WTipo = Vector(Cicla, 3)
            WNumero = Vector(Cicla, 4)
            WCliente = Vector(Cicla, 5)
            WNeto = Val(Vector(Cicla, 6))
            WIva1 = Val(Vector(Cicla, 7))
            WIva2 = Val(Vector(Cicla, 8))
            WImpoIb = Val(Vector(Cicla, 9))
            WImpoIbTucu = Val(Vector(Cicla, 10))
            WImpoIbCiudad = Val(Vector(Cicla, 11))
            WTotal = Val(Vector(Cicla, 12))
            Call Redondeo(WImpoIbCiudad)
            
            If WImpoIbCiudad > 0 Then
        
                WCuit = ""
                WNroIbTucu = ""
                WNombre = ""
                WDomicilio = ""
                WPuerta = "00000"
                WLocalidad = ""
                WProvincia = "0"
                WPostal = ""
            
                spClientes = "ConsultaClientes " + "'" + WCliente + "'"
                Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
                If rstClientes.RecordCount > 0 Then
                    WCuit = Left$(rstClientes!Cuit, 13)
                    WNroIbCiudad = IIf(IsNull(rstClientes!NroIbCiudad), "", rstClientes!NroIbCiudad)
                    WNombre = rstClientes!Razon
                    WDomicilio = rstClientes!Direccion
                    WLocalidad = rstClientes!Localidad
                    WPuerta = "00000"
                    WProvincia = rstClientes!Provincia
                    WPostal = rstClientes!Postal
                    WPorceIb = 3
                    WIbCiudad = IIf(IsNull(rstClientes!IbCiudadII), "0", rstClientes!IbCiudadII)
                    rstClientes.Close
                End If
            
                Call Ceros(WNumero, 8)
                WNumero = Left$("0001" + WNumero + Space$(16), 16)
        
                Call Redondeo(WImpoIbCiudad)
                XImpoIbCiudad = Str$(WImpoIbCiudad)
                Call Ceros(XImpoIbCiudad, 16)
                Auxi = XImpoIbCiudad
                Auxi = MascaraII("#############.##", Auxi)
                XImpoIbCiudad = Auxi
        
                Call Redondeo(WTotal)
                XTotal = Str$(WTotal)
                Call Ceros(XTotal, 16)
                Auxi = XTotal
                Auxi = MascaraII("#############.##", Auxi)
                XTotal = Auxi
        
                Call Redondeo(WNeto)
                XNeto = Str$(WNeto)
                Call Ceros(XNeto, 16)
                Auxi = XNeto
                Auxi = MascaraII("#############.##", Auxi)
                XNeto = Auxi
        
                Call Redondeo(WPorceIb)
                XPorceIb = Str$(WPorceIb)
                Call Ceros(XPorceIb, 5)
                Auxi = XPorceIb
                Auxi = MascaraII("##.##", Auxi)
                XPorceIb = Auxi
        
                WOtros = WImpoIb + WImpoTucu
                Call Redondeo(WOtros)
                XOtros = Str$(WOtros)
                Call Ceros(XOtros, 16)
                Auxi = XOtros
                Auxi = MascaraII("#############.##", Auxi)
                XOtros = Auxi
        
                WIva = WIva1 + WIva2
                Call Redondeo(WIva)
                XIva = Str$(WIva)
                Call Ceros(XIva, 16)
                Auxi = XIva
                Auxi = MascaraII("#############.##", Auxi)
                XIva = Auxi
        
                Call Eval
                Call Ceros(WCuit, 11)
                Call Ceros(WPostal, 4)
                Call Ceros(WNroIbCiudad, 10)
        
                WNombre = Left$(WNombre + Space$(30), 30)
        
        
        
                Rem Tipo de Operacion
                Campo1 = "2"
        
                Rem Codigo de Norma
                Campo2 = "014"
        
                Rem fecha
                Campo3 = WFecha
        
                Rem tipo de comprobante
                Select Case Val(WTipo)
                    Case 1, 77
                        Campo4 = "01"
                    Case Else
                        Campo4 = "02"
                End Select
            
                Rem Letra de comprobante
                Campo5 = "A"
        
                Rem Numero del Comprobante
                Campo6 = WNumero
        
                Rem fecha
                Campo7 = WFecha
        
                Rem monto de la retencion
                Campo8 = XTotal
        
                Rem numero de la retencion
                Campo9 = Space(16)
        
                Rem tipo de documento
                Campo10 = "3"
        
                Rem documento
                Campo11 = WCuit
        
                Rem Codigo de Situacion de I.B.
                Select Case Val(WIbCiudad)
                    Case 1
                        Campo12 = "1"
                    Case 2
                        Campo12 = "2"
                    Case 3
                        Campo12 = "4"
                    Case 4
                        Campo12 = "5"
                    Case Else
                        Campo12 = "0"
                End Select
                
                Rem Numero  de I.B.
                If Val(WIbCiudad) <> 4 Then
                    Campo13 = WNroIbCiudad + " "
                        Else
                    Campo13 = WCuit
                End If
        
                Rem Conducion de Iva
                Campo14 = "1"
        
                Rem Razon
                Campo15 = WNombre
        
                Rem otros conceptos
                Campo16 = XOtros
        
                Rem iva
                Campo17 = XIva
        
                Rem neto
                Campo18 = XNeto
        
                Rem alicuota
                
                ZZAlicuota = Val(XImpoIbCiudad) / (Val(XNeto) / 100)
                Call Redondeo(ZZAlicuota)
                If ZZAlicuota = 4.5 Then
                    ZZAlicuota = 6
                End If
                If ZZAlicuota = 2 Then
                    ZZAlicuota = 3
                End If
                If ZZAlicuota = 2.5 Then
                    ZZAlicuota = 3
                End If
                
                WImpoIbCiudad = Val(XNeto) * (ZZAlicuota / 100)
                Call Redondeo(WImpoIbCiudad)
                XImpoIbCiudad = Str$(WImpoIbCiudad)
                Call Ceros(XImpoIbCiudad, 12)
                Auxi = XImpoIbCiudad
                Auxi = MascaraII("#############.##", Auxi)
                XImpoIbCiudad = Auxi
                
                XPorceIb = Str$(ZZAlicuota)
                Call Ceros(XPorceIb, 5)
                Auxi = XPorceIb
                Auxi = MascaraII("##.##", Auxi)
                XPorceIb = Auxi
                
                
                
                If ZZAlicuota = 6 Then
                    Campo2 = "016"
                End If
                
                Campo19 = XPorceIb
        
                Rem retencion
                Campo20 = XImpoIbCiudad
                Campo21 = XImpoIbCiudad
        
                WImpre = Campo1 + Campo2 + Campo3 + Campo4 + Campo5 + Campo6 + Campo7 + Campo8 + Campo9 + Campo10 + Campo11 + Campo12 + Campo13 + Campo14 + Campo15 + Campo16 + Campo17 + Campo18 + Campo19 + Campo20 + Campo21
                
                
            
                Print #1, WImpre
                
            End If
        End If
        
    Next Cicla
    
    
dada:
    Rem GoTo dada
    Rem dada
    
    Renglon = 0
    Erase Vector
    
    spPagos = "ListaPagos"
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
            
        With rstPagos
            .MoveFirst
            Do
            
                If WDesde <= !FechaOrd And !FechaOrd <= WHasta Then
                
                    ZZRetIbCiudad = IIf(IsNull(rstPagos!RetIbCiudad), "0", rstPagos!RetIbCiudad)
                    If ZZRetIbCiudad <> 0 And !Renglon = 1 Then
                
                        WOrden = !Orden
                        WRenglon = !Renglon
                        WProveedor = !Proveedor
                        WFecha = !Fecha
                        WFechaord = !FechaOrd
                        WImporte = !Importe
                        WRetencion = !Retencion
                        WObservaciones = !Observaciones
                        WCuenta = !Cuenta
                        WTipoord = !TipoOrd
                        WTiporeg = !Tiporeg
                        WTipo1 = !Tipo1
                        WLetra1 = !Letra1
                        WPunto1 = !Punto1
                        WNumero1 = !Numero1
                        WImporte1 = !Importe
                        WRetotra = ZZRetIbCiudad
                        WObservaciones2 = !Observaciones2
                        WTipo2 = !Tipo2
                        WNumero2 = !Numero2
                        WFecha2 = !Fecha2
                        WFechaord2 = !FechaOrd2
                        WBanco2 = !Banco2
                        WImporte2 = !Importe2
                        WClave = !Clave
                        WNroIbCiudadII = !CertificadoIbCiudad
                        
                        Renglon = Renglon + 1
                        Vector(Renglon, 1) = WOrden
                        Vector(Renglon, 2) = WProveedor
                        Vector(Renglon, 3) = WFecha
                        Vector(Renglon, 4) = Str$(WImporte)
                        Vector(Renglon, 5) = Str$(WRetencion)
                        Vector(Renglon, 6) = Str$(WRetotra)
                        Vector(Renglon, 7) = WNroIbCiudadII
                            
                    End If
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
    End If
    
    
    
    For Cicla = 1 To Renglon
    
        WOrden = Vector(Cicla, 1)
        WProveedor = Vector(Cicla, 2)
        WFecha = Vector(Cicla, 3)
        WTotal = Val(Vector(Cicla, 4))
        WImpoIb = Val(Vector(Cicla, 5))
        WImpoIbCiudad = Val(Vector(Cicla, 6))
        Call Redondeo(WImpoIbCiudad)
        WNroIbCiudadII = Vector(Cicla, 7)
        Call Ceros(WNroIbCiudadII, 16)
            
        If WImpoIbCiudad > 0 Then
        
                WCuit = ""
                WNroIbTucu = ""
                WNombre = ""
                WDomicilio = ""
                WPuerta = "00000"
                WLocalidad = ""
                WProvincia = "0"
                WPostal = ""
            
                spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    WNombre = RstProveedor!Nombre
                    WPorceIb = 2.5
                    WTipoiva = Val(RstProveedor!Iva)
                    WCuit = IIf(IsNull(RstProveedor!NroIb), "", RstProveedor!NroIb)
                    Call Eval
                    WNroIbCiudad = WCuit
                    WCuit = RstProveedor!Cuit
                    RstProveedor.Close
                End If
            
                Call Ceros(WNumero, 8)
                WNumero = Left$("0001" + WNumero + Space$(16), 16)
        
                Call Redondeo(WImpoIbCiudad)
                XImpoIbCiudad = Str$(WImpoIbCiudad)
                Call Ceros(XImpoIbCiudad, 16)
                Auxi = XImpoIbCiudad
                Auxi = MascaraII("#############.##", Auxi)
                XImpoIbCiudad = Auxi
        
                Call Redondeo(WTotal)
                XTotal = Str$(WTotal)
                Call Ceros(XTotal, 16)
                Auxi = XTotal
                Auxi = MascaraII("#############.##", Auxi)
                XTotal = Auxi
                
                If WTipoiva = 2 Then
                    WNeto = WTotal / 1.21
                    Call Redondeo(WNeto)
                        Else
                    WNeto = WTotal
                    Call Redondeo(WNeto)
                End If
                
                WIva1 = WTotal - WNeto
                WIva2 = 0
                Call Redondeo(WIva1)
        
                Call Redondeo(WNeto)
                XNeto = Str$(WNeto)
                Call Ceros(XNeto, 16)
                Auxi = XNeto
                Auxi = MascaraII("#############.##", Auxi)
                XNeto = Auxi
        
                Call Redondeo(WPorceIb)
                XPorceIb = Str$(WPorceIb)
                Call Ceros(XPorceIb, 5)
                Auxi = XPorceIb
                Auxi = MascaraII("#############.##", Auxi)
                XPorceIb = Auxi
        
                WOtros = 0
                Call Redondeo(WOtros)
                XOtros = Str$(WOtros)
                Call Ceros(XOtros, 16)
                Auxi = XOtros
                Auxi = MascaraII("#############.##", Auxi)
                XOtros = Auxi
        
                WIva = WIva1 + WIva2
                Call Redondeo(WIva)
                XIva = Str$(WIva)
                Call Ceros(XIva, 16)
                Auxi = XIva
                Auxi = MascaraII("#############.##", Auxi)
                XIva = Auxi
        
                Call Eval
                Call Ceros(WCuit, 11)
                Call Ceros(WPostal, 4)
                Call Ceros(WNroIbCiudad, 10)
        
                WNombre = Left$(WNombre + Space$(30), 30)
        
        
        
                Rem Tipo de Operacion
                Campo1 = "1"
        
                Rem Codigo de Norma
                Campo2 = "008"
        
                Rem fecha
                Campo3 = WFecha
        
                Rem tipo de comprobante
                Select Case Val(WTipo)
                    Case 1, 77
                        Campo4 = "01"
                    Case Else
                        Campo4 = "02"
                End Select
            
                Rem Letra de comprobante
                Campo5 = "A"
        
                Rem Numero del Comprobante
                Campo6 = WNumero
        
                Rem fecha
                Campo7 = WFecha
        
                Rem monto de la retencion
                Campo8 = XTotal
        
                Rem numero de la retencion
                Campo9 = WNroIbCiudadII
        
                Rem tipo de documento
                Campo10 = "2"
        
                Rem documento
                Campo11 = WCuit
        
                Rem Codigo de Situacion de I.B.
                Campo12 = "2"
        
                Rem Numero  de I.B.
                Campo13 = WNroIbCiudad + " "
        
                Rem Conducion de Iva
                Campo14 = "1"
        
                Rem Razon
                Campo15 = WNombre
        
                Rem otros conceptos
                Campo16 = XOtros
        
                Rem iva
                Campo17 = XIva
        
                Rem neto
                Campo18 = XNeto
        
                Rem alicuota
                
                ZZAlicuota = Val(XImpoIbCiudad) / (Val(XNeto) / 100)
                Call Redondeo(ZZAlicuota)
                If ZZAlicuota = 4.5 Then
                    ZZAlicuota = 6
                End If
                If ZZAlicuota = 2 Then
                    ZZAlicuota = 2.5
                End If
                
                WImpoIbCiudad = Val(XNeto) * (ZZAlicuota / 100)
                Call Redondeo(WImpoIbCiudad)
                XImpoIbCiudad = Str$(WImpoIbCiudad)
                Call Ceros(XImpoIbCiudad, 12)
                Auxi = XImpoIbCiudad
                Auxi = MascaraII("#############.##", Auxi)
                XImpoIbCiudad = Auxi
                
                XPorceIb = Str$(ZZAlicuota)
                Call Ceros(XPorceIb, 5)
                Auxi = XPorceIb
                Auxi = MascaraII("##.##", Auxi)
                XPorceIb = Auxi
                
                Campo19 = XPorceIb
                If ZZAlicuota = 6 Then
                    Campo2 = "016"
                End If
        
                Rem retencion
                Campo20 = XImpoIbCiudad
                Campo21 = XImpoIbCiudad
        
                WImpre = Campo1 + Campo2 + Campo3 + Campo4 + Campo5 + Campo6 + Campo7 + Campo8 + Campo9 + Campo10 + Campo11 + Campo12 + Campo13 + Campo14 + Campo15 + Campo16 + Campo17 + Campo18 + Campo19 + Campo20 + Campo21
            
                Print #1, WImpre
                
        End If
        
    Next Cicla
    
    Close #1
    
    Call Cancela_Click
        
End Sub

Private Sub Cancela_Click()
    Desde.SetFocus
    PrgProcesoCiudad.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Private Sub Nombre_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()

    Provincia(0) = "Capital Federal"
    Provincia(1) = "Buenos Aires"
    Provincia(2) = "Catamarca"
    Provincia(3) = "Cordoba"
    Provincia(4) = "Corrientes"
    Provincia(5) = "Chaco"
    Provincia(6) = "Chubut"
    Provincia(7) = "Entre Rios"
    Provincia(8) = "Formosa"
    Provincia(9) = "Jujuy"
    Provincia(10) = "La Pampa"
    Provincia(11) = "La Rioja"
    Provincia(12) = "Mendoza"
    Provincia(13) = "Misiones"
    Provincia(14) = "Neuquen"
    Provincia(15) = "Rio Negro"
    Provincia(16) = "Salta"
    Provincia(17) = "San Juan"
    Provincia(18) = "San Luis"
    Provincia(19) = "Santa Cruz"
    Provincia(20) = "Santa Fe"
    Provincia(21) = "Santiago del Estero"
    Provincia(22) = "Tucuman"
    Provincia(23) = "Tierra del Fuego"
    Provincia(24) = "Exterior"
    Provincia(25) = ""


    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
End Sub

Private Sub Eval()

    Es = WCuit

    x = ""
    MinusOk = 1                'a minus sign is okay only once, and only
                                'if it preceeds the first numeric character
    DecOk = 1                  'only the first decimal point is okay

    For XX = 1 To Len(Es)

        Y = Mid$(Es, XX, 1)

        If Y = "-" And MinusOk = 1 Then
               x = x + Y: MinusOk = 0

        ElseIf Y = "." And DecOk = 1 Then
               x = x + Y: DecOk = 0

        ElseIf Y >= "0" And Y <= "9" Then
               x = x + Y: MinusOk = 0

        End If

    Next

    WCuit = x

End Sub


