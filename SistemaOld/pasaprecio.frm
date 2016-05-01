VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPasaPrecio 
   Caption         =   "Traspaso de Precios"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cierre 
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Ejecuta Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin MSMask.MaskEdBox Terminado 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "AA-#####-###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Articulo 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
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
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin VB.Label DesArticulo 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label DesTerminado 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "PrgPasaPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WPasaVector(10000, 50) As String
Dim ZLugar As Integer

Dim ZCliente As String
Dim ZTerminado As String
Dim ZArticulo As String
Dim ZClave As String
        
Dim ZPrecio As String
Dim ZDescripcion As String
        
Dim ZFecha1 As String
Dim ZFactura1 As String
Dim ZPrecio1 As String
Dim ZCantidad1 As String
        
Dim ZFecha2 As String
Dim ZFactura2 As String
Dim ZPrecio2 As String
Dim ZCantidad2 As String
        
Dim ZFecha3 As String
Dim ZFactura3 As String
Dim ZPrecio3 As String
Dim ZCantidad3 As String
        
Dim ZFecha4 As String
Dim ZFactura4 As String
Dim ZPrecio4 As String
Dim ZCantidad4 As String
        
Dim ZFecha5 As String
Dim ZFactura5 As String
Dim ZPrecio5 As String
Dim ZCantidad5 As String
        
Dim ZDate As String
Dim ZFecha As String
Dim ZPago As String
Dim ZEstado As String


Private Sub Cierre_Click()
    PrgPasaPrecio.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Proceso_Click()

    ZLugar = 0
    Erase WPasaVector
        
    ZSql = " "
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Precios"
    ZSql = ZSql + " Where Precios.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " Order by Precios.Clave"
    
    rsPrecios = ZSql
    Set rstPrecios = db.OpenRecordset(rsPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        With rstPrecios
            .MoveFirst
            Do
                If .EOF = False Then
                    
                    ZLugar = ZLugar + 1
                        
                    WPasaVector(ZLugar, 1) = rstPrecios!Clave
                    WPasaVector(ZLugar, 2) = rstPrecios!Cliente
                    WPasaVector(ZLugar, 3) = rstPrecios!Terminado
                    WPasaVector(ZLugar, 4) = Str$(rstPrecios!Precio)
                    WPasaVector(ZLugar, 5) = rstPrecios!Descripcion
                    WPasaVector(ZLugar, 6) = rstPrecios!Fecha1
                    WPasaVector(ZLugar, 7) = rstPrecios!Factura1
                    WPasaVector(ZLugar, 8) = Str$(rstPrecios!Precio1)
                    WPasaVector(ZLugar, 9) = Str$(rstPrecios!Cantidad1)
                    WPasaVector(ZLugar, 10) = rstPrecios!Fecha2
                    WPasaVector(ZLugar, 11) = rstPrecios!Factura2
                    WPasaVector(ZLugar, 12) = Str$(rstPrecios!Precio2)
                    WPasaVector(ZLugar, 13) = Str$(rstPrecios!Cantidad2)
                    WPasaVector(ZLugar, 14) = rstPrecios!Fecha3
                    WPasaVector(ZLugar, 15) = rstPrecios!Factura3
                    WPasaVector(ZLugar, 16) = Str$(rstPrecios!Precio3)
                    WPasaVector(ZLugar, 17) = Str$(rstPrecios!Cantidad3)
                    WPasaVector(ZLugar, 18) = rstPrecios!Fecha4
                    WPasaVector(ZLugar, 19) = rstPrecios!Factura4
                    WPasaVector(ZLugar, 20) = Str$(rstPrecios!Precio4)
                    WPasaVector(ZLugar, 21) = Str$(rstPrecios!Cantidad4)
                    WPasaVector(ZLugar, 22) = rstPrecios!Fecha5
                    WPasaVector(ZLugar, 23) = rstPrecios!Factura5
                    WPasaVector(ZLugar, 24) = Str$(rstPrecios!Precio5)
                    WPasaVector(ZLugar, 25) = Str$(rstPrecios!Cantidad5)
                    WPasaVector(ZLugar, 26) = rstPrecios!WDate
                    WPasaVector(ZLugar, 27) = IIf(IsNull(rstPrecios!Fecha), "", rstPrecios!Fecha)
                    WPasaVector(ZLugar, 28) = IIf(IsNull(rstPrecios!Pago), "0", rstPrecios!Pago)
                    WPasaVector(ZLugar, 29) = IIf(IsNull(rstPrecios!Estado), "0", rstPrecios!Estado)
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPrecios.Close
    End If
    
    
    For Ciclo = 1 To ZLugar
    
        ZCliente = WPasaVector(Ciclo, 2)
        ZArticulo = Articulo.Text
        ZClave = ZCliente + ZArticulo
        
        ZPrecio = WPasaVector(Ciclo, 4)
        ZDescripcion = WPasaVector(Ciclo, 5)
        
        ZFecha1 = WPasaVector(Ciclo, 6)
        ZFactura1 = WPasaVector(Ciclo, 7)
        ZPrecio1 = WPasaVector(Ciclo, 8)
        ZCantidad1 = WPasaVector(Ciclo, 9)
        
        ZFecha2 = WPasaVector(Ciclo, 10)
        ZFactura2 = WPasaVector(Ciclo, 11)
        ZPrecio2 = WPasaVector(Ciclo, 12)
        ZCantidad2 = WPasaVector(Ciclo, 13)
        
        ZFecha3 = WPasaVector(Ciclo, 14)
        ZFactura3 = WPasaVector(Ciclo, 15)
        ZPrecio3 = WPasaVector(Ciclo, 16)
        ZCantidad3 = WPasaVector(Ciclo, 17)
        
        ZFecha4 = WPasaVector(Ciclo, 18)
        ZFactura4 = WPasaVector(Ciclo, 19)
        ZPrecio4 = WPasaVector(Ciclo, 20)
        ZCantidad4 = WPasaVector(Ciclo, 21)
        
        ZFecha5 = WPasaVector(Ciclo, 22)
        ZFactura5 = WPasaVector(Ciclo, 23)
        ZPrecio5 = WPasaVector(Ciclo, 24)
        ZCantidad5 = WPasaVector(Ciclo, 25)
        
        ZDate = WPasaVector(Ciclo, 26)
        ZFecha = WPasaVector(Ciclo, 27)
        ZPago = WPasaVector(Ciclo, 28)
        ZEstado = WPasaVector(Ciclo, 29)
    
        XParam = "'" + ZClave + "','" + ZCliente + "','" + ZArticulo + "','" + ZPrecio + "','" _
                     + ZFecha1 + "','" + ZFactura1 + "','" + ZPrecio1 + "','" + ZCantidad1 + "','" _
                     + ZFecha2 + "','" + ZFactura2 + "','" + ZPrecio2 + "','" + ZCantidad2 + "','" _
                     + ZFecha3 + "','" + ZFactura3 + "','" + ZPrecio3 + "','" + ZCantidad3 + "','" _
                     + ZFecha4 + "','" + ZFactura4 + "','" + ZPrecio4 + "','" + ZCantidad4 + "','" _
                     + ZFecha5 + "','" + ZFactura5 + "','" + ZPrecio5 + "','" + ZCantidad5 + "','" _
                     + ZDate + "','" + ZFecha + "','" + ZPago + "'"
        Set rstPreciosMp = db.OpenRecordset("AltaPreciosMp " + XParam, dbOpenSnapshot, dbSQLPassThrough)
    
        ZSql = ""
        ZSql = ZSql & "UPDATE PreciosMp SET "
        ZSql = ZSql & "Estado = " + "'" + ZEstado + "'"
        ZSql = ZSql & " Where Clave = " + "'" + ZClave + "'"
        spPreciosMp = ZSql
        Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    Articulo.Text = "  -   -   "
    Terminado.Text = "  -     -   "
    
    Terminado.SetFocus

End Sub

Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Terminado.Text = UCase(Terminado.Text)
        WTerminado = Terminado.Text
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesTerminado.Caption = rstTerminado!Descripcion
            rstTerminado.Close
            Articulo.SetFocus
                Else
            Terminado.SetFocus
        End If
    End If
End Sub

Private Sub Articulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Articulo.Text = UCase(Articulo.Text)
        WArticulo = Articulo.Text
        spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DesArticulo.Caption = rstArticulo!Descripcion
            rstArticulo.Close
            Terminado.SetFocus
                Else
            Articulo.SetFocus
        End If
    End If
End Sub

