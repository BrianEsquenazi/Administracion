VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProcesoCiudadNuevo 
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
Attribute VB_Name = "PrgProcesoCiudadNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstClientes As Recordset
Dim spClientes As String
Dim rstCtacte As Recordset
Dim spCtacte As String
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
Dim ZZZImporte As Double

Dim ZZZNeto As Double

Dim WIbTucu As Integer
Dim WIbCiudad As String

Dim WPorceIb As Double
Dim WPorceIbCiudad As Double

Dim XNeto As String
Dim XXNeto As String
Dim XImpoIb As String
Dim XPorceIb As String
Dim XImpoIbCiudad As String
Dim XXXImpoIbCiudad As String
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

Dim ZZEntra(10000, 2) As String
Dim ZZEntraLugar As Integer

Private Sub Drive_Change()
    Dir1.Path = Drive.Drive
End Sub

Private Sub Acepta_Click()

    WDrive = Drive.Drive
    WDir = Dir1.Path
    
    XNOmbre = WDir + "\" + "ciudad.Txt"
    Open XNOmbre For Output As #1
    
    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    Erase ZZEntra
    ZZEntraLugar = 0
    
    Rem Procesa las ventas
    
    Renglon = 0
    Erase Vector
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCte"
    ZSql = ZSql + " Where CtaCte.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and CtaCte.OrdFecha <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and CtaCte.ImpoIbCiudad <> 0"
    ZSql = ZSql + " Order by CtaCte.OrdFecha"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
        
        With rstCtacte
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    
                    WNeto = IIf(IsNull(rstCtacte!Neto), "0", rstCtacte!Neto)
                    WIva1 = IIf(IsNull(rstCtacte!Iva1), "0", rstCtacte!Iva1)
                    WIva2 = IIf(IsNull(rstCtacte!Iva2), "0", rstCtacte!Iva2)
                    WImpoIb = IIf(IsNull(rstCtacte!impoib), "0", rstCtacte!impoib)
                    WImpoIbTucu = IIf(IsNull(rstCtacte!ImpoIbTucu), "0", rstCtacte!ImpoIbTucu)
                    WImpoIbCiudad = IIf(IsNull(rstCtacte!ImpoIbCiudad), "0", rstCtacte!ImpoIbCiudad)
                    WTotal = IIf(IsNull(rstCtacte!Total), "0", rstCtacte!Total)
                    
                    Call Redondeo(WNeto)
                    Call Redondeo(WIva1)
                    Call Redondeo(WIva2)
                    Call Redondeo(WImpoIb)
                    Call Redondeo(WImpoIbTucu)
                    Call Redondeo(WImpoIbCiudad)
                    
                    
                    
                    aaaa = rstCtacte!Numero
                    
                    
                    Rem alicuota
                    ZZAlicuota = WImpoIbCiudad / (WNeto / 100)
                    Call Redondeo(ZZAlicuota)
            
                    ZZZNeto = WImpoIbCiudad / ZZAlicuota * 100
                    Call Redondeo(ZZZNeto)
                    
                    
                    WTotal = ZZZNeto + WIva1 + WIva2 + WImpoIb + WImpoIbTucu + WImpoIbCiudad
                    
                    
                    
                    Vector(Renglon, 1) = rstCtacte!Clave
                    Vector(Renglon, 2) = rstCtacte!Fecha
                    Vector(Renglon, 3) = rstCtacte!Tipo
                    Vector(Renglon, 4) = rstCtacte!Numero
                    Vector(Renglon, 5) = rstCtacte!Cliente
                    Vector(Renglon, 6) = Str$(ZZZNeto)
                    Vector(Renglon, 7) = Str$(WIva1)
                    Vector(Renglon, 8) = Str$(WIva2)
                    Vector(Renglon, 9) = Str$(WImpoIb)
                    Vector(Renglon, 10) = Str$(WImpoIbTucu)
                    Vector(Renglon, 11) = Str$(WImpoIbCiudad)
                    Vector(Renglon, 12) = Str$(WTotal)
                    Vector(Renglon, 13) = Str$(ZZAlicuota)
                    
                    
                    
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCtacte.Close
    End If
     
    For Cicla = 1 To Renglon
    
        WClave = Vector(Cicla, 1)
        
        If WClave <> "" Then
        
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
            WAlicuota = Val(Vector(Cicla, 13))
            
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
                    WPorceIb = rstClientes!PorceIbCaba
                    WIbCiudad = IIf(IsNull(rstClientes!IbCiudadII), "0", rstClientes!IbCiudadII)
                    rstClientes.Close
                End If
            
                Call Ceros(WNumero, 12)
                WNumero = Left$("0001" + WNumero + Space$(16), 16)
        
                Call Redondeo(WImpoIbCiudad)
                XImpoIbCiudad = Str$(WImpoIbCiudad)
                Call Ceros(XImpoIbCiudad, 16)
                Auxi = XImpoIbCiudad
                XXXImpoIbCiudad = XImpoIbCiudad
                Auxi = MascaraII("#############.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                XImpoIbCiudad = Auxi1
                
                
        
                Call Redondeo(WTotal)
                Rem XTotal = Str$(WTotal - WImpoIbCiudad)
                XTotal = Str$(WTotal)
                Call Ceros(XTotal, 16)
                Auxi = XTotal
                Auxi = MascaraII("#############.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                XTotal = Auxi1
                
                
                XXNeto = Str$(WNeto)
                
                
                WOtros = WImpoIb + WImpoIbTucu + WImpoIbCiudad
                Rem WNeto = WNeto - WOtros
        
                Call Redondeo(WNeto)
                XNeto = Str$(WNeto)
                Call Ceros(XNeto, 16)
                Auxi = XNeto
                Auxi = MascaraII("#############.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                XNeto = Auxi1
        
                Call Redondeo(WPorceIb)
                XPorceIb = Str$(WPorceIb)
                Call Ceros(XPorceIb, 5)
                Auxi = WPorceIb
                Auxi = MascaraII("##.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                WPorceIb = Auxi1
        
                WOtros = WImpoIb + WImpoIbTucu + WImpoIbCiudad
                Call Redondeo(WOtros)
                XOtros = Str$(WOtros)
                Call Ceros(XOtros, 16)
                Auxi = XOtros
                Auxi = MascaraII("#############.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                XOtros = Auxi1
        
                WIva = WIva1 + WIva2
                Call Redondeo(WIva)
                XIva = Str$(WIva)
                Call Ceros(XIva, 16)
                Auxi = XIva
                Auxi = MascaraII("#############.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                XIva = Auxi1
        
                Call Eval
                Call Ceros(WCuit, 11)
                Call Ceros(WPostal, 4)
                Call Ceros(WNroIbCiudad, 10)
                Call Ceros(WIbCiudad, 1)
        
                WNombre = Left$(WNombre + Space$(30), 30)
                Auxi = UCase(WNombre)
                For ZZCiclo = 1 To 30
                    If Mid$(Auxi, ZZCiclo, 1) = "Ñ" Then
                        Auxi = Left$(Auxi, ZZCiclo - 1) + "N" + Mid$(Auxi, ZZCiclo + 1, 11)
                    End If
                Next ZZCiclo
                WNombre = Auxi
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
                    Campo13 = "0" + WNroIbCiudad
                        Else
                    Campo13 = WCuit
                End If
        
                Rem Conducion de Iva
                Campo14 = "1"
        
                Rem Razon
                Campo15 = WNombre
                Auxi = UCase(WNombre)
                For ZZCiclo = 1 To 30
                    If Mid$(Auxi, ZZCiclo, 1) = "Ñ" Then
                        Auxi = Left$(Auxi, ZZCiclo - 1) + "N" + Mid$(Auxi, ZZCiclo + 1, 11)
                    End If
                Next ZZCiclo
                WNombre = Auxi
        
                Rem otros conceptos
                Campo16 = XOtros
        
                Rem iva
                Campo17 = XIva
        
        
                Rem Rem alicuota
                Rem ZZAlicuota = Val(XXXImpoIbCiudad) / (Val(XXNeto) / 100)
                Rem Call Redondeo(ZZAlicuota)
                Rem XPorceIb = Str$(ZZAlicuota)
                Rem Call Ceros(XPorceIb, 5)
                Rem Auxi = XPorceIb
                Rem Auxi = MascaraII("##.##", Auxi)
                Rem Call Convierte_datos(Auxi, Auxi1)
                Rem XPorceIb = Auxi1
        
                Rem ZZZNeto = Val(XXXImpoIbCiudad) / ZZAlicuota * 100
                Rem Call Redondeo(ZZZNeto)
                Rem XNeto = Str$(ZZZNeto)
                Rem Call Ceros(XNeto, 16)
                Rem Auxi = XNeto
                Rem Auxi = MascaraII("#############.##", Auxi)
                Rem Call Convierte_datos(Auxi, Auxi1)
                Rem XNeto = Auxi1
        
        
                Rem neto
                Campo18 = XNeto
                
                Rem alicuota
                ZZAlicuota = WAlicuota
                Call Redondeo(ZZAlicuota)
                XPorceIb = Str$(ZZAlicuota)
                Call Ceros(XPorceIb, 5)
                Auxi = XPorceIb
                Auxi = MascaraII("##.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                XPorceIb = Auxi1
                
                
                If ZZAlicuota = 6 Then
                    Campo2 = "016"
                End If
                
                Campo19 = XPorceIb
        
                Rem retencion
                Campo20 = XImpoIbCiudad
                Campo21 = XImpoIbCiudad
        
                WImpre = Campo1 + Campo2 + Campo3 + Campo4 + Campo5 + Campo6 + Campo7 + Campo8 + Campo9 + Campo10 + Campo11 + Campo12 + Campo13 + Campo14 + Campo15 + Campo16 + Campo17 + Campo18 + Campo19 + Campo20 + Campo21
                
                ZZEntraLugar = ZZEntraLugar + 1
                ZZEntra(ZZEntraLugar, 1) = WFecha
                ZZEntra(ZZEntraLugar, 2) = WImpre
            
                Rem Print #1, WImpre
                
            End If
        End If
        
    Next Cicla
    
    
    
    
    
    
    
    
    
    
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
        WNumero = Vector(Cicla, 1)
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
                    WIbCiudad = IIf(IsNull(RstProveedor!IbCiudadII), "0", RstProveedor!IbCiudadII)
                    RstProveedor.Close
                End If
            
                Call Ceros(WNumero, 12)
                WNumero = Left$("0001" + WNumero + Space$(16), 16)
        
                Call Redondeo(WImpoIbCiudad)
                XImpoIbCiudad = Str$(WImpoIbCiudad)
                Call Ceros(XImpoIbCiudad, 16)
                Auxi = XImpoIbCiudad
                XXXImpoIbCiudad = Str$(WImpoIbCiudad)
                ZZZImpoIbCiudad = WImpoIbCiudad
                Auxi = MascaraII("#############.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                XImpoIbCiudad = Auxi1
        
                Call Redondeo(WTotal)
                XTotal = Str$(WTotal)
                Call Ceros(XTotal, 16)
                Auxi = XTotal
                Auxi = MascaraII("#############.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                XTotal = Auxi1
                
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
                Call Convierte_datos(Auxi, Auxi1)
                XNeto = Auxi1
        
                Call Redondeo(WPorceIb)
                XPorceIb = Str$(WPorceIb)
                Call Ceros(XPorceIb, 5)
                Auxi = XPorceIb
                Auxi = MascaraII("#############.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                XPorceIb = Auxi1
        
                WOtros = 0
                Call Redondeo(WOtros)
                XOtros = Str$(WOtros)
                Call Ceros(XOtros, 16)
                Auxi = XOtros
                Auxi = MascaraII("#############.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                XOtros = Auxi1
        
                WIva = WIva1 + WIva2
                Call Redondeo(WIva)
                XIva = Str$(WIva)
                Call Ceros(XIva, 16)
                Auxi = XIva
                Auxi = MascaraII("#############.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                XIva = Auxi1
        
                Call Eval
                Call Ceros(WCuit, 11)
                Call Ceros(WPostal, 4)
                Call Ceros(WNroIbCiudad, 10)
                Call Ceros(WIbCiudad, 1)
        
                WNombre = Left$(WNombre + Space$(30), 30)
                Auxi = UCase(WNombre)
                For ZZCiclo = 1 To 30
                    If Mid$(Auxi, ZZCiclo, 1) = "Ñ" Then
                        Auxi = Left$(Auxi, ZZCiclo - 1) + "N" + Mid$(Auxi, ZZCiclo + 1, 11)
                    End If
                Next ZZCiclo
                WNombre = Auxi
                WNombre = Left$(WNombre + Space$(30), 30)
        
        
        
                Rem Tipo de Operacion
                Campo1 = "1"
        
                Rem Codigo de Norma
                Campo2 = "008"
        
                Rem fecha
                Campo3 = WFecha
        
                Rem tipo de comprobante
                Rem Select Case Val(WTipo)
                Rem     Case 1, 77
                Rem         Campo4 = "01"
                Rem     Case Else
                Rem         Campo4 = "02"
                Rem End Select
                Campo4 = "03"
            
                Rem Letra de comprobante
                Campo5 = " "
        
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
                Campo13 = "0" + WNroIbCiudad
        
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
                ZZZBase = 0
                Renglon = 0
                Do
                
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    ClavePagos = WOrden + Auxi1
                
                    spPagos = "ConsultaPagos " + "'" + ClavePagos + "'"
                    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPagos.RecordCount > 0 Then
                        If Val(rstPagos!Tiporeg) = 1 Then
                            ZZZTipo = rstPagos!Tipo1
                            ZZZLetra = rstPagos!Letra1
                            ZZZPunto = rstPagos!Punto1
                            ZZZNumero = rstPagos!Numero1
                            ZZZImporte = rstPagos!Importe1
                            If WTipoiva = 2 Then
                                ZZZImporte = ZZZImporte / 1.21
                            End If
                            Call Redondeo(ZZZImporte)
                            Rem If ZZZImporte < 300 Then
                            Rem     ZZZImporte = 0
                            Rem End If
                            ZZZBase = ZZZBase + ZZZImporte
                        End If
                        rstPagos.Close
                            Else
                        Exit Do
                    End If
                Loop
                ZZAlicuota = Val(XXXImpoIbCiudad) / (ZZZBase / 100)
                Call Redondeo(ZZAlicuota)
                Rem If ZZAlicuota = 4.5 Then
                Rem     ZZAlicuota = 6
                Rem End If
                Rem If ZZAlicuota = 2 Then
                Rem     ZZAlicuota = 2.5
                Rem End If
                
                WImpoIbCiudad = WNeto * (ZZAlicuota / 100)
                Call Redondeo(WImpoIbCiudad)
                XImpoIbCiudad = Str$(WImpoIbCiudad)
                XXXImpoIbCiudad = Str$(WImpoIbCiudad)
                Call Ceros(XImpoIbCiudad, 12)
                Auxi = XImpoIbCiudad
                Auxi = MascaraII("#############.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                XImpoIbCiudad = Auxi1
                
                
                
                
                WTotal = WTotal - ZZZImpoIbCiudad + WImpoIbCiudad
                
                Call Redondeo(WTotal)
                XTotal = Str$(WTotal)
                Call Ceros(XTotal, 16)
                Auxi = XTotal
                Auxi = MascaraII("#############.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                XTotal = Auxi1
        
                Rem monto de la retencion
                Campo8 = XTotal
                
                
                
                XPorceIb = Str$(ZZAlicuota)
                Call Ceros(XPorceIb, 5)
                Auxi = XPorceIb
                Auxi = MascaraII("##.##", Auxi)
                Call Convierte_datos(Auxi, Auxi1)
                XPorceIb = Auxi1
                
                Campo19 = XPorceIb
                If ZZAlicuota = 6 Then
                    Campo2 = "016"
                End If
        
                Rem retencion
                Campo20 = XImpoIbCiudad
                Campo21 = XImpoIbCiudad
        
                WImpre = Campo1 + Campo2 + Campo3 + Campo4 + Campo5 + Campo6 + Campo7 + Campo8 + Campo9 + Campo10 + Campo11 + Campo12 + Campo13 + Campo14 + Campo15 + Campo16 + Campo17 + Campo18 + Campo19 + Campo20 + Campo21
            
                ZZEntraLugar = ZZEntraLugar + 1
                ZZEntra(ZZEntraLugar, 1) = WFecha
                ZZEntra(ZZEntraLugar, 2) = WImpre
            
            
                Rem Print #1, WImpre
                
        End If
        
    Next Cicla
    
    
    For Ciclo = 1 To ZZEntraLugar

        For dada = Ciclo + 1 To ZZEntraLugar

            If ZZEntra(Ciclo, 1) > ZZEntra(dada, 1) Then

                Auxi1 = ZZEntra(Ciclo, 1)
                Auxi2 = ZZEntra(Ciclo, 2)
                
                ZZEntra(Ciclo, 1) = ZZEntra(dada, 1)
                ZZEntra(Ciclo, 2) = ZZEntra(dada, 2)
                
                ZZEntra(dada, 1) = Auxi1
                ZZEntra(dada, 2) = Auxi2

            End If

        Next dada

    Next Ciclo
    
    
    

    For Ciclo = 1 To ZZEntraLugar
        Print #1, ZZEntra(Ciclo, 2)
    Next Ciclo
    
    
    
    
    
    
da:
    
    
    
    
    Close #1
    
    Call Cancela_Click
        
End Sub

Private Sub Cancela_Click()
    Desde.SetFocus
    PrgProcesoCiudadNuevo.Hide
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


