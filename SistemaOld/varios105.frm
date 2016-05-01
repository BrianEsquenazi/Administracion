VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgVarios105 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Comprobantes por Concepto Varios AL 10.5 % DE IVA"
   ClientHeight    =   8310
   ClientLeft      =   315
   ClientTop       =   405
   ClientWidth     =   11535
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8310
   ScaleWidth      =   11535
   Visible         =   0   'False
   Begin VB.ComboBox ReteCiudad 
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
      ItemData        =   "varios105.frx":0000
      Left            =   9360
      List            =   "varios105.frx":0002
      TabIndex        =   63
      Top             =   720
      Width           =   1935
   End
   Begin VB.ComboBox ReteIb 
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
      ItemData        =   "varios105.frx":0004
      Left            =   9360
      List            =   "varios105.frx":0006
      TabIndex        =   59
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton ReImpre 
      Caption         =   "Impresion"
      Height          =   495
      Left            =   9720
      TabIndex        =   55
      Top             =   7440
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Moneda"
      Height          =   615
      Left            =   6360
      TabIndex        =   52
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Pesos 
         Caption         =   "$"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Dolares 
         Caption         =   "U$S"
         Height          =   255
         Left            =   720
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox NroFactura 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6120
      MaxLength       =   6
      TabIndex        =   51
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   120
      TabIndex        =   46
      Top             =   5880
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Frame Frame4 
      Height          =   1815
      Left            =   9240
      TabIndex        =   39
      Top             =   2160
      Width           =   1935
      Begin VB.OptionButton Expo 
         Caption         =   "Exportacion"
         Height          =   375
         Left            =   240
         TabIndex        =   58
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton Ajuste 
         Caption         =   "Dif. Cambio ($ y U$S)"
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Exenta 
         Caption         =   "Exenta (Ch.Rec)"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Normal 
         Caption         =   "Normal"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   9240
      TabIndex        =   35
      Top             =   1200
      Width           =   2055
      Begin VB.OptionButton Credito 
         Caption         =   "Nota de Credito"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Debito 
         Caption         =   "Nota de Debito"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Factura 
         Caption         =   "Factura Varias"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta de Datos"
      Height          =   495
      Left            =   9720
      TabIndex        =   33
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglones"
      Height          =   495
      Left            =   9720
      TabIndex        =   32
      Top             =   6480
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   615
      Left            =   360
      TabIndex        =   29
      Top             =   5160
      Width           =   8655
      Begin VB.TextBox WDescripcion 
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   34
         Text            =   " "
         Top             =   240
         Width           =   6135
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Text            =   " "
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox WImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   30
         Text            =   " "
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox Paridad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   28
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Calcula 
      Caption         =   "Calcula Datos"
      Height          =   495
      Left            =   9720
      TabIndex        =   26
      Top             =   6000
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   5640
      TabIndex        =   17
      Top             =   5760
      Width           =   3255
      Begin VB.Label Label17 
         Caption         =   "IB Ciudad"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label ImpoIbCiudad 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1200
         TabIndex        =   61
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label ImpoIbTucu 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1200
         TabIndex        =   57
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "IB Tucu."
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "IB Bs.As."
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   840
         Width           =   975
      End
      Begin VB.Label ImpoIb 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1200
         TabIndex        =   48
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Interes"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Descuento"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Dto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1200
         TabIndex        =   43
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Interes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1200
         TabIndex        =   42
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Total 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1200
         TabIndex        =   25
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Iva2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Iva1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Neto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1200
         TabIndex        =   22
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Iva 10.5%"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Iva 21%"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Neto"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9120
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   9720
      TabIndex        =   16
      Top             =   5520
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      Height          =   1815
      Left            =   2280
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Cliente 
      Height          =   285
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   11
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Numero 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   450
      Left            =   9720
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   450
      Left            =   9720
      TabIndex        =   5
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   450
      Left            =   9720
      TabIndex        =   4
      Top             =   5040
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3855
      Left            =   360
      OleObjectBlob   =   "varios105.frx":0008
      TabIndex        =   3
      Top             =   1200
      Width           =   8655
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8760
      TabIndex        =   2
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1815
      ItemData        =   "varios105.frx":09F6
      Left            =   120
      List            =   "varios105.frx":09FD
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label Label16 
      Caption         =   "Ret Ciudad"
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
      Left            =   8280
      TabIndex        =   64
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Ret I.B."
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
      Left            =   8280
      TabIndex        =   60
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Factura Asociada"
      Height          =   255
      Left            =   4680
      TabIndex        =   50
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Paridad"
      Height          =   255
      Left            =   2640
      TabIndex        =   27
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Vencimiento"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Comprobante"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgVarios105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 2 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WPlazo1 As Integer
Private WPlazo2 As Integer
Private WDias1 As Integer
Private WDias2 As Integer
Private WFecha As String
Private Wvencimiento As String
Private WVencimiento1 As String
Private WPago1 As Integer
Private WPago2 As Integer
Private WNeto As Double
Private XNeto As Double
Private WIva1 As Double
Private WIva2 As Double
Private WTotal As Double
Private WImpoDto As Double
Private WImpoInteres As Double
Private WDescuento As Double
Private WTasa As Double
Private WCodIva As String
Private Precio As Double
Private Cantidad As Double
Private WAnterior As Integer
Private WDescri As String
Private WTipo As String
Private WProvincia As String
Private WRubro As Integer
Private WVendedor As Integer
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WProv As String
Private WPostal As String
Private WImpiva As String
Private WCuit As String
Private WPago As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private WDirentrega As String
Private Auxiliar(100, 2) As String
Private Articulo As String
Private Auxi As String
Private Auxi1 As String
Private Renglon As Integer
Dim rstNumero As Recordset
Dim spNumero As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstDesccomp As Recordset
Dim spDesccomp As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstPago As Recordset
Dim spPago As String
Dim XParam As String
Dim Compara As Double
Private WCodIb As Integer
Private WCodIbTucu As Integer
Private WCodIbCiudad As Integer
Private WImpoIb As Double
Private WImpoIbTucu As Double
Private WImpoIbCiudad As Double
Private WPorceCm05Tucu As Double

Dim WNro As String
Dim ZTipo As String
Private WTexto1 As String
Private WTexto2 As String

Private Sub Calcula_FechaVto()

    spPago = "ConsultaPago " + "'" + Str$(WPago1) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WDias1 = rstPago!Dias
        WPlazo1 = rstPago!Plazo
        WTasa = rstPago!Tasa
        WDescuento = rstPago!Descuento
        WPago = rstPago!Nombre
        rstPago.Close
    End If
    
    WFecha = Fecha.Text
    Call Calcula_vencimiento(WFecha, WDias1, Wvencimiento)
    
    spPago = "ConsultaPago " + "'" + Str$(WPago2) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WDias2 = rstPago!Dias
        WPlazo2 = rstPago!Plazo
        rstPago.Close
   End If
    
    Call Calcula_vencimiento(WFecha, WDias2, WVencimiento1)

End Sub

Private Sub Ajuste_Click()
    If Expo.Value = False Then
            
        spNumero = "ConsultaNumero " + "'" + ZTipo + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = "1"
        End If
        
            Else
            
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
        
    End If
End Sub

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    WDescripcion.Text = ""
    WImporte.Text = ""
    WLinea.Text = ""
    
    WDescripcion.SetFocus

End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Clientes"

     Opcion.Visible = True
     
 End Sub

Private Sub Credito_Click()

    If Factura.Value = True Then
        ZTipo = "01"
    End If
    If Debito.Value = True Then
        ZTipo = "03"
    End If
    If Credito.Value = True Then
        ZTipo = "04"
    End If
    
    If Expo.Value = False Then
            
        spNumero = "ConsultaNumero " + "'" + ZTipo + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = "1"
        End If
        
            Else
            
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
        
    End If

End Sub

Private Sub Debito_Click()

    If Factura.Value = True Then
        ZTipo = "01"
    End If
    If Debito.Value = True Then
        ZTipo = "03"
    End If
    If Credito.Value = True Then
        ZTipo = "04"
    End If
    
    If Expo.Value = False Then
            
        spNumero = "ConsultaNumero " + "'" + ZTipo + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = "1"
        End If
        
            Else
            
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
        
    End If

End Sub

Private Sub Exenta_Click()
    If Expo.Value = False Then
            
        spNumero = "ConsultaNumero " + "'" + ZTipo + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = "1"
        End If
        
            Else
            
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
        
    End If
End Sub

Private Sub Expo_Click()
    If Expo.Value = False Then
            
        spNumero = "ConsultaNumero " + "'" + ZTipo + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = "1"
        End If
        
            Else
            
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
        
        Pesos.Value = False
        Dolares.Value = True
        Call Calcula_Click
        
    End If
End Sub

Private Sub Factura_Click()

    If Factura.Value = True Then
        ZTipo = "01"
    End If
    If Debito.Value = True Then
        ZTipo = "03"
    End If
    If Credito.Value = True Then
        ZTipo = "04"
    End If
    
    If Expo.Value = False Then
            
        spNumero = "ConsultaNumero " + "'" + ZTipo + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = "1"
        End If
        
            Else
            
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
        
    End If
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub



Private Sub Normal_Click()
    If Expo.Value = False Then
            
        spNumero = "ConsultaNumero " + "'" + ZTipo + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = "1"
        End If
        
            Else
            
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
        
    End If
End Sub

Private Sub reImpre_Click()
    If Expo.Value = False Then
        Call Impresion
            Else
        Call Impresion_Expo
    End If
End Sub

 Private Sub Opcion_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Rem XIndice = 0
    
    Select Case XIndice
        Case 0
            Ayuda.Visible = True
            Ayuda.Text = ""
            spClientes = "ListaClienteConsulta"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                With rstClientes
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstClientes!Cliente + " " + rstClientes!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstClientes!Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstClientes.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub DBGrid1_GotFocus()
    
    WCol = DBGrid1.Col
    WRow = DBGrid1.Row
    
    DBGrid1.Col = WCol
    DBGrid1.Row = WRow
    
    DBGrid1.Col = 0
    WDescri = DBGrid1.Text
    
    DBGrid1.Col = 1
    WImporte = DBGrid1.Text
    
    If WDescri = "" And Val(WImporte) = 0 Then
        WDescripcion.Text = ""
        WLinea.Text = ""
            Else
        WLinea.Text = DBGrid1.Row + 1
        WDescripcion.Text = DBGrid1.Text
    End If
    
    DBGrid1.Col = 0
    WDescripcion.Text = DBGrid1.Text

    DBGrid1.Col = 1
    If Val(DBGrid1.Text) <> 0 Then
        WImporte.Text = DBGrid1.Text
            Else
        WImporte.Text = ""
    End If
    
    WDescripcion.SetFocus
    
    If Fecha.Text = "  /  /    " Or Cliente.Text = "" Then
         Numero.SetFocus
    End If

End Sub

Private Sub Calcula_Click()

    WNeto = 0

    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 1
            WImporte = Val(DBGrid1.Text)
                    
            WNeto = WNeto + WImporte
                    
        Next iRow
            
    Next a
    
    Call Calcula_Importe
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
End Sub

Private Sub Calcula_Importe()

    Rem If Val(Paridad.Text) <> 0 Then
    Rem     WNeto = WNeto * Val(Paridad.Text)
    Rem End If
    If Exenta.Value = True Then
        Paridad.Text = "1"
    End If
    
    XNeto = WNeto
    WImpoDto = 0
    WImpoInteres = 0
    
    If Normal.Value = True Or Ajuste.Value = True Then
    
    If WDescuento <> 0 Then
        WImpoDto = WNeto * WDescuento / 100
        Call Redondeo(WImpoDto)
        WNeto = WNeto - WImpoDto
    End If
    
    If WTasa <> 0 Then
        WImpoInteres = (WNeto * WPlazo1 * WTasa) / 36000
        Call Redondeo(WImpoInteres)
        WNeto = WNeto + WImpoInteres
    End If
    
    End If
    
    WIva1 = 0
    WIva2 = 0
    WImpoIb = 0
    WImpoIbTucu = 0
    WImpoIbCiudad = 0
    
    If Normal.Value = True Or Ajuste.Value = True Then
    
        ZFechaCompa = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        If ZFechaCompa >= "20071201" Then
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                ZZIb = IIf(IsNull(rstCliente!Ib), "0", rstCliente!Ib)
                WPorceIb = IIf(IsNull(rstCliente!PorceIb), "0", rstCliente!PorceIb)
                rstCliente.Close
            End If
        
            If ZZIb <> 2 Then
                WImpoIb = WNeto * (WPorceIb / 100)
                Call Redondeo(WImpoIb)
            End If
    
                Else
    
            Select Case WCodIb
                Case 0, 1
                    Select Case Val(WCodIva)
                        Case 1
                            WImpoIb = WNeto * 0.025
                        Case 2, 4, 5, 6
                            WImpoIb = WNeto * 0.03
                        Case Else
                            WImpoIb = 0
                    End Select
                    Call Redondeo(WImpoIb)
                Case Else
                    WImpoIb = 0
            End Select
            
        End If
    
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WPorceCm05Tucu = IIf(IsNull(rstCliente!PorceCm05Tucu), "0", rstCliente!PorceCm05Tucu)
            rstCliente.Close
        End If
        If WPorceCm05Tucu = 0 Then
            WPorceCm05Tucu = 1
        End If
        Select Case WCodIbTucu
            Case 1, 2, 3
                WImpoIbTucu = WNeto * 0.0175 * WPorceCm05Tucu
                Call Redondeo(WImpoIbTucu)
                WImpoPorceIbTucu = 1.75
            Case 4
                WImpoIbTucu = WNeto * 0.035
                Call Redondeo(WImpoIbTucu)
                WImpoPorceIbTucu = 3.5
            Case 5
                WImpoIbTucu = WNeto * 0.025
                Call Redondeo(WImpoIbTucu)
                WImpoPorceIbTucu = 2.5
            Case Else
                WImpoIbTucu = 0
        End Select
        
        Select Case WCodIbCiudad
            Case 1
                WImpoIbCiudad = WNeto * 0.035
                Call Redondeo(WImpoIbCiudad)
            Case 2
                WImpoIbCiudad = WNeto * 0.06
                Call Redondeo(WImpoIbCiudad)
            Case Else
                WImpoIbCiudad = 0
        End Select
    
        If Pesos.Value = True Then
            Compara = WNeto
                Else
            Compara = WNeto * Val(Paridad.Text)
        End If
        Call Redondeo(Compara)
        If Compara < 100 Then
            WImpoIb = 0
        End If
        If Compara < 500 Then
            WImpoIbCiudad = 0
        End If
        
        If ReteIb.ListIndex = 2 Then
            WImpoIb = 0
        End If
        
        If ReteCiudad.ListIndex = 2 Then
            WImpoIbCiudad = 0
        End If
    
        Select Case Val(WCodIva)
            Case 2
                WIva1 = WNeto * 0.21
                WIva2 = WNeto * 0.105
                Call Redondeo(WIva1)
                Call Redondeo(WIva2)
            Case 3, 4, 5
                WIva1 = 0
                WIva2 = 0
            Case Else
                WIva1 = WNeto * 0.105
                Call Redondeo(WIva1)
        End Select
            
    End If
    
    If WNeto <> 0 Then
        Call Convierte1_datos(Str$(WNeto), Auxi)
        Neto.Caption = Pusing("###,###.##", Auxi)
            Else
        Neto.Caption = "0.00"
    End If
    
    If WImpoIb <> 0 Then
        Call Convierte1_datos(Str$(WImpoIb), Auxi)
        ImpoIb.Caption = Pusing("###,###.##", Auxi)
            Else
        ImpoIb.Caption = "0.00"
    End If
    
    If WImpoIbTucu <> 0 Then
        Call Convierte1_datos(Str$(WImpoIbTucu), Auxi)
        ImpoIbTucu.Caption = Pusing("###,###.##", Auxi)
            Else
        ImpoIbTucu.Caption = "0.00"
    End If
    
    If WImpoIbCiudad <> 0 Then
        Call Convierte1_datos(Str$(WImpoIbCiudad), Auxi)
        ImpoIbCiudad.Caption = Pusing("###,###.##", Auxi)
            Else
        ImpoIbCiudad.Caption = "0.00"
    End If
    
    If WImpoDto <> 0 Then
        Call Convierte1_datos(Str$(WImpoDto), Auxi)
        Dto.Caption = Pusing("###,###.##", Auxi)
            Else
        Dto.Caption = "0.00"
    End If
    
    If WImpoInteres <> 0 Then
        Call Convierte1_datos(Str$(WImpoInteres), Auxi)
        Interes.Caption = Pusing("###,###.##", Auxi)
            Else
        Interes.Caption = "0.00"
    End If
    
    If WIva1 <> 0 Then
        Call Convierte1_datos(Str$(WIva1), Auxi)
        Iva1.Caption = Pusing("###,###.##", Auxi)
            Else
        Iva1.Caption = "0.00"
    End If
    
    If WIva2 <> 0 Then
        Call Convierte1_datos(Str$(WIva2), Auxi)
        Iva2.Caption = Pusing("###,###.##", Auxi)
            Else
        Iva2.Caption = "0.00"
    End If
    
    WTotal = WNeto + WIva1 + WIva2 + WImpoIb + WImpoIbTucu + WImpoIbCiudad
    Call Convierte1_datos(Str$(WTotal), Auxi)
    Total.Caption = Pusing("###,###.##", Auxi)

End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstAuxiliar
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()

    If ReteIb.ListIndex = 0 Then
        m$ = "Se debe informar si se debe retener o no Ingresos Brutos"
        aa% = MsgBox(m$, 0, "MODULO DE FACTURACION")
        Exit Sub
    End If
    
    If ReteCiudad.ListIndex = 0 Then
        m$ = "Se debe informar si se debe retener o no Ingresos Brutos de Ciudad de Bs As"
        aa% = MsgBox(m$, 0, "MODULO DE FACTURACION")
        Exit Sub
    End If
    
        Pasa = "S"

        If Val(NroFactura.Text) <> 0 Then

        If Normal.Value = True Or Ajuste.Value = True Then
            WNro = NroFactura.Text
            Call Ceros(WNro, 8)
            WClaveCtacte = "01" + WNro + "01"
            spCtacte = "ConsultaCtacte " + "'" + WClaveCtacte + "'"
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtacte.RecordCount > 0 Then
                If rstCtacte!Cliente <> Cliente.Text Then
                    m$ = "El cliente de la factura no se corresponde al cliente informado"
                    a% = MsgBox(m$, 0, "Emision de Comprobantes Varios")
                    Pasa = "N"
                End If
                rstCtacte.Close
                    Else
                m$ = "No existe el numero de comprobante informado"
                a% = MsgBox(m$, 0, "Emision de Comprobantes Varios")
                Pasa = "N"
            End If
                Else
            NroFactura.Text = ""
        End If
        
        End If
        
        If Pasa = "S" Then

        Rem If Exenta.Value = True Then
            WPago1 = 1
            WPago2 = 1
            Call Calcula_FechaVto
        Rem End If

        Cliente.Text = UCase(Cliente.Text)
        
        Renglon = Renglon + 1
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
    
        DBGrid1.Col = 0
        DBGrid1.Text = ""

        Call Calcula_Click
        
        If Normal.Value = True Or Ajuste.Value = True Then
            If Val(WCodIva) = 3 Or Val(WCodIva) = 5 Then
                WImporte = WNeto
                WNeto = WNeto / 1.105
                Call Redondeo(WNeto)
                WIva1 = WImporte - WNeto
                WIva2 = 0
            End If
        End If
        
        If Factura.Value = True Then
            WTipo = "03"
            WImpre = "FV"
        End If
        If Debito.Value = True Then
            WTipo = "04"
            WImpre = "ND"
        End If
        If Credito.Value = True Then
            WTipo = "05"
            WImpre = "NC"
        End If
        
        If Exenta.Value = True Then
            WImpre = "CH"
        End If
        
        WNumero = Numero.Text
        WRenglon = "01"
        WCliente = Cliente.Text
        WFecha = Fecha.Text
        WEstado = "0"
        Rem Wvencimiento = Wvencimiento
        Rem WVencimiento1 = WVencimiento1
        Call Convierte_datos(Str$(Total), Auxi)
        If Pesos.Value = True Then
            If Credito.Value = False Then
                XTotal = Str$(WTotal)
                XTotalUs = Str$(WTotal / Val(Paridad.Text))
                XSaldo = Str$(WTotal)
                XSaldoUs = Str$(WTotal / Val(Paridad.Text))
                XNet = Str$(WNeto)
                XIva1 = Str$(WIva1)
                XIva2 = Str$(WIva2)
                XImpoIb = Str$(WImpoIb)
                XImpoIbTucu = Str$(WImpoIbTucu)
                XImpoIbCiudad = Str$(WImpoIbCiudad)
                XSeguro = ""
                XFlete = ""
                    Else
                XTotal = Str$(WTotal * -1)
                XTotalUs = Str$(WTotal * -1 / Val(Paridad.Text))
                XSaldo = Str$(WTotal * -1)
                XSaldoUs = Str$(WTotal * -1 / Val(Paridad.Text))
                XNet = Str$(WNeto * -1)
                XIva1 = Str$(WIva1 * -1)
                XIva2 = Str$(WIva2 * -1)
                XImpoIb = Str$(WImpoIb * -1)
                XImpoIbTucu = Str$(WImpoIbTucu * -1)
                XImpoIbCiudad = Str$(WImpoIbCiudad * -1)
                XSeguro = ""
                XFlete = ""
            End If
                Else
            If Credito.Value = False Then
                XTotal = Str$(WTotal * Val(Paridad.Text))
                XTotalUs = Str$(WTotal)
                XSaldo = Str$(WTotal * Val(Paridad.Text))
                XSaldoUs = Str$(WTotal)
                XNet = Str$(WNeto * Val(Paridad.Text))
                XIva1 = Str$(WIva1 * Val(Paridad.Text))
                XIva2 = Str$(WIva2 * Val(Paridad.Text))
                XImpoIb = Str$(WImpoIb * Val(Paridad.Text))
                XImpoIbTucu = Str$(WImpoIbTucu * Val(Paridad.Text))
                XImpoIbCiudad = Str$(WImpoIbCiudad * Val(Paridad.Text))
                XSeguro = ""
                XFlete = ""
                    Else
                XTotal = Str$(WTotal * -1 * Val(Paridad.Text))
                XTotalUs = Str$(WTotal * -1)
                XSaldo = Str$(WTotal * -1 * Val(Paridad.Text))
                XSaldoUs = Str$(WTotal * -1)
                XNet = Str$(WNeto * -1 * Val(Paridad.Text))
                XIva1 = Str$(WIva1 * -1 * Val(Paridad.Text))
                XIva2 = Str$(WIva2 * -1 * Val(Paridad.Text))
                XImpoIb = Str$(WImpoIb * -1 * Val(Paridad.Text))
                XImpoIbTucu = Str$(WImpoIbTucu * -1 * Val(Paridad.Text))
                XImpoIbCiudad = Str$(WImpoIbCiudad * -1 * Val(Paridad.Text))
                XSeguro = ""
                XFlete = ""
            End If
        End If
        
        If Expo.Value = True Then
            If Credito.Value = False Then
                XTotal = Str$(WTotal)
                XTotalUs = Str$(WTotal)
                XSaldo = Str$(WTotal)
                XSaldoUs = Str$(WTotal)
                XNet = Str$(WNeto * Val(Paridad.Text))
                XIva1 = Str$(WIva1 * Val(Paridad.Text))
                XIva2 = Str$(WIva2 * Val(Paridad.Text))
                XImpoIb = Str$(WImpoIb * Val(Paridad.Text))
                XImpoIbTucu = Str$(WImpoIbTucu * Val(Paridad.Text))
                XImpoIbCiudad = Str$(WImpoIbCiudad * Val(Paridad.Text))
                XSeguro = ""
                XFlete = ""
                    Else
                XTotal = Str$(WTotal * -1)
                XTotalUs = Str$(WTotal * -1)
                XSaldo = Str$(WTotal * -1)
                XSaldoUs = Str$(WTotal * -1)
                XNet = Str$(WNeto * -1 * Val(Paridad.Text))
                XIva1 = Str$(WIva1 * -1 * Val(Paridad.Text))
                XIva2 = Str$(WIva2 * -1 * Val(Paridad.Text))
                XImpoIb = Str$(WImpoIb * -1 * Val(Paridad.Text))
                XImpoIbTucu = Str$(WImpoIbTucu * -1 * Val(Paridad.Text))
                XImpoIbCiudad = Str$(WImpoIbCiudad * -1 * Val(Paridad.Text))
                XSeguro = ""
                XFlete = ""
            End If
        End If
            
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WOrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
        WOrdVencimiento1 = Right$(WVencimiento1, 4) + Mid$(WVencimiento1, 4, 2) + Left$(WVencimiento1, 2)
        WPedido = ""
        WRemito = ""
        WOrden = ""
        WParidad = Paridad.Text
        WProvincia = WProvincia
        XVendedor = Str$(WVendedor)
        XRubro = Str$(WRubro)
        WComprobante = ""
        WAceptada = ""
        WCosto = ""
        WImporte1 = ""
        WImporte2 = ""
        WImporte3 = ""
        WImporte4 = ""
        WImporte5 = ""
        WImporte6 = ""
        WImporte7 = ""
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        WClave = WTipo + Auxi + "01"
        XEmpresa = "1"
        WDate = Date$
        WNroFactura = NroFactura.Text
        WNroRecibo = ""
        
        XParam = "'" + WClave + "','" _
                    + WTipo + "','" + WNumero + "','" _
                    + WRenglon + "','" + WCliente + "','" _
                    + WFecha + "','" + WEstado + "','" _
                    + Wvencimiento + "','" + WVencimiento1 + "','" _
                    + XTotal + "','" + XTotalUs + "','" _
                    + XSaldo + "','" + XSaldoUs + "','" _
                    + WOrdFecha + "','" + WOrdVencimiento + "','" _
                    + WOrdVencimiento1 + "','" + WImpre + "','" _
                    + XEmpresa + "','" _
                    + XNet + "','" + XIva1 + "','" _
                    + XIva2 + "','" + WPedido + "','" _
                    + WRemito + "','" + WOrden + "','" _
                    + WParidad + "','" + WProvincia + "','" _
                    + XVendedor + "','" + XRubro + "','" _
                    + WComprobante + "','" + WAceptada + "','" _
                    + WCosto + "','" _
                    + WImporte1 + "','" + WImporte2 + "','" _
                    + WImporte3 + "','" + WImporte4 + "','" _
                    + WImporte5 + "','" + WImporte6 + "','" _
                    + WImporte7 + "','" + WDate + "','" _
                    + XSeguro + "','" + XFlete + "','" _
                    + XImpoIb + "','" + WNroFactura + "','" _
                    + WNroRecibo + "'"
                        
        spCtacte = "AltaCtacteVarios " + XParam
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        

        ZSql = ""
        ZSql = ZSql + "UPDATE CtaCte SET "
        ZSql = ZSql + " ImpoIbTucu = " + "'" + XImpoIbTucu + "',"
        ZSql = ZSql + " ImpoIbCiudad = " + "'" + XImpoIbCiudad + "'"
        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                     
        spCtacte = ZSql
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        If Pesos.Value = True Then
            WMoneda = "0"
                Else
            WMoneda = "1"
        End If
        spCtacte = "UPDATE Ctacte SET" _
                    + " Moneda = " + "'" + WMoneda + "'" _
                    + " Where Clave = " + "'" + WClave + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                        
        Renglon = 0
        WRenglon = 0
        DBGrid1.Refresh
        
        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRenglon = WRenglon + 1
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                WDescripcion = DBGrid1.Text
                    
                DBGrid1.Col = 1
                WImporte = DBGrid1.Text
                    
                If WDescripcion <> "" Or Val(WImporte) <> 0 Then
                    
                    Renglon = Renglon + 1
                    Auxi = Str$(Renglon)
                    Call Ceros(Auxi, 2)
                        
                    Auxi1 = Str$(Numero.Text)
                    Call Ceros(Auxi1, 8)
                        
                    If Factura.Value = True Then
                        WTipo = "03"
                    End If
                    If Debito.Value = True Then
                        WTipo = "04"
                    End If
                    If Credito.Value = True Then
                        WTipo = "05"
                    End If
                        
                    WNumero = Numero.Text
                    WRenglon = Str$(Renglon)
                    WDescripcion = WDescripcion
                    WImporte = WImporte
                    XEmpresa = "1"
                    
                    WClave = WTipo + Auxi1 + Auxi
                    WDate = Date$
                    
                    XParam = "'" + WClave + "','" _
                        + WTipo + "','" _
                        + WNumero + "','" _
                        + WRenglon + "','" _
                        + WDescripcion + "','" _
                        + WImporte + "','" _
                        + XEmpresa + "','" _
                        + WDate + "'"
                        
                    spDesccomp = "AltaDesccomp " + XParam
                    Set rstDesccomp = db.OpenRecordset(spDesccomp, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
                                        
            Next iRow
            
        Next a
        
        If Credito.Value = True Then
        
            If WIva1 <> 0 Then
                Articulo = "PT-99999-999"
        
                spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WLinea = rstTerminado!Linea
                    rstTerminado.Close
                        Else
                    WLinea = 50
                End If
        
                Renglon = 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                            
                Auxi1 = Str$(Numero.Text)
                Call Ceros(Auxi1, 8)
        
                XTipo = "02"
        
                WNumero = Numero.Text
                XRenglon = Str$(Renglon)
                WArticulo = Articulo
                XCantidad = "1"
                If Pesos.Value = True Then
                    If Normal.Value = True Then
                        XPrecio = Str$(WNeto)
                        XPrecioUs = Str$(WNeto / Val(Paridad.Text))
                        XImporte = Str$(WNeto * -1)
                        XImporteUs = Str$((WNeto / Val(Paridad.Text)) * -1)
                            Else
                        XPrecio = Str$(WNeto)
                        XPrecioUs = Str$(0)
                        XImporte = Str$(WNeto * -1)
                        XImporteUs = Str$(0)
                    End If
                        Else
                    If Normal.Value = True Then
                        XPrecio = Str$(WNeto * Val(Paridad.Text))
                        XPrecioUs = Str$(WNeto)
                        XImporte = Str$(WNeto * -1 * Val(Paridad.Text))
                        XImporteUs = Str$(WNeto * -1)
                            Else
                        XPrecio = Str$(WNeto)
                        XPrecioUs = Str$(0)
                        XImporte = Str$(WNeto * -1)
                        XImporteUs = Str$(0)
                    End If
                End If
                WCliente = Cliente.Text
                WParidad = Paridad.Text
                XVendedor = Str$(WVendedor)
                XRubro = Str$(WRubro)
                XLinea = Str$(WLinea)
                XCosto2 = ""
                XCosto1 = ""
                WCoeficiente = ""
                WPedido = ""
                WFecha = Fecha.Text
                WImporte1 = ""
                WImporte2 = ""
                WImporte3 = ""
                WImporte4 = ""
                WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                XArticulo = Left$(Articulo, 8)
                WRemito = ""
                WClave = XTipo + Auxi1 + Auxi
                WDate = Date$
                XCanti = ""
                XImpo = ""
                XImpoUs = ""
                XMarca = ""
                WLote1 = "0"
                WCanti1 = "0"
                WLote2 = "0"
                WCanti2 = "0"
                Wlote3 = "0"
                WCanti3 = "0"
                WLote4 = "0"
                WCanti5 = "0"
                WLote4 = "0"
                WCanti5 = "0"
                XTipoproDy = "T"
                XArticuloDy = "  -   -   "
                            
                XParam = "'" + WClave + "','" _
                    + XTipo + "','" + WNumero + "','" _
                    + XRenglon + "','" + WArticulo + "','" _
                    + XCantidad + "','" + XPrecio + "','" _
                    + XPrecioUs + "','" + XImporte + "','" _
                    + XImporteUs + "','" + WCliente + "','" _
                    + WParidad + "','" + XVendedor + "','" _
                    + XRubro + "','" + XLinea + "','" _
                    + XCosto1 + "','" + XCosto2 + "','" _
                    + WCoeficiente + "','" + WPedido + "','" _
                    + WFecha + "','" + WImporte1 + "','" _
                    + WImporte2 + "','" + WImporte3 + "','" _
                    + WImporte4 + "','" + WOrdFecha + "','" _
                    + XArticulo + "','" + WRemito + "','" _
                    + WDate + "','" + XCanti + "','" _
                    + XImpo + "','" + XImpoUs + "','" _
                    + XMarca + "','" _
                    + WLote1 + "','" + WCanti1 + "','" _
                    + WLote2 + "','" + WCanti2 + "','" _
                    + Wlote3 + "','" + WCanti3 + "','" _
                    + WLote4 + "','" + WCanti4 + "','" _
                    + WLote5 + "','" + WCanti5 + "','" _
                    + XTipoproDy + "','" + XArticuloDy + "'"
                    
                spEstadistica = "AltaEstadistica " + XParam
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
                Else
            
            If WIva1 <> 0 Then
        
                Articulo = "PT-99999-999"
        
                spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WLinea = rstTerminado!Linea
                    rstTerminado.Close
                        Else
                    WLinea = 50
                End If
        
                Renglon = 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                            
                Auxi1 = Str$(Numero.Text)
                Call Ceros(Auxi1, 8)
        
                XTipo = "01"
        
                WNumero = Numero.Text
                XRenglon = Str$(Renglon)
                WArticulo = Articulo
                XCantidad = "1"
                If Pesos.Value = True Then
                    If Normal.Value = True Then
                        XPrecio = Str$(WNeto)
                        XPrecioUs = Str$(WNeto / Val(Paridad.Text))
                        XImporte = Str$(WNeto)
                        XImporteUs = Str$(WNeto / Val(Paridad.Text))
                            Else
                        XPrecio = Str$(WNeto)
                        XPrecioUs = Str$(0)
                        XImporte = Str$(WNeto)
                        XImporteUs = Str$(0)
                    End If
                        Else
                    If Normal.Value = True Then
                        XPrecio = Str$(WNeto * Val(Paridad.Text))
                        XPrecioUs = Str$(WNeto)
                        XImporte = Str$(WNeto * Val(Paridad.Text))
                        XImporteUs = Str$(WNeto)
                            Else
                        XPrecio = Str$(WNeto)
                        XPrecioUs = Str$(0)
                        XImporte = Str$(WNeto)
                        XImporteUs = Str$(0)
                    End If
                End If
                WCliente = Cliente.Text
                WParidad = Paridad.Text
                XVendedor = Str$(WVendedor)
                XRubro = Str$(WRubro)
                XLinea = Str$(WLinea)
                XCosto2 = ""
                XCosto1 = ""
                WCoeficiente = ""
                WPedido = ""
                WFecha = Fecha.Text
                WImporte1 = ""
                WImporte2 = ""
                WImporte3 = ""
                WImporte4 = ""
                WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                XArticulo = Left$(Articulo, 8)
                WRemito = ""
                WClave = XTipo + Auxi1 + Auxi
                WDate = Date$
                XCanti = ""
                XImpo = ""
                XImpoUs = ""
                XMarca = ""
                WLote1 = "0"
                WCanti1 = "0"
                WLote2 = "0"
                WCanti2 = "0"
                Wlote3 = "0"
                WCanti3 = "0"
                WLote4 = "0"
                WCanti5 = "0"
                WLote4 = "0"
                WCanti5 = "0"
                XTipoproDy = "T"
                XArticuloDy = "  -   -   "
                            
                XParam = "'" + WClave + "','" _
                    + XTipo + "','" + WNumero + "','" _
                    + XRenglon + "','" + WArticulo + "','" _
                    + XCantidad + "','" + XPrecio + "','" _
                    + XPrecioUs + "','" + XImporte + "','" _
                    + XImporteUs + "','" + WCliente + "','" _
                    + WParidad + "','" + XVendedor + "','" _
                    + XRubro + "','" + XLinea + "','" _
                    + XCosto1 + "','" + XCosto2 + "','" _
                    + WCoeficiente + "','" + WPedido + "','" _
                    + WFecha + "','" + WImporte1 + "','" _
                    + WImporte2 + "','" + WImporte3 + "','" _
                    + WImporte4 + "','" + WOrdFecha + "','" _
                    + XArticulo + "','" + WRemito + "','" _
                    + WDate + "','" + XCanti + "','" _
                    + XImpo + "','" + XImpoUs + "','" _
                    + XMarca + "','" _
                    + WLote1 + "','" + WCanti1 + "','" _
                    + WLote2 + "','" + WCanti2 + "','" _
                    + Wlote3 + "','" + WCanti3 + "','" _
                    + WLote4 + "','" + WCanti4 + "','" _
                    + WLote5 + "','" + WCanti5 + "','" _
                    + XTipoproDy + "','" + XArticuloDy + "'"
                    
                spEstadistica = "AltaEstadistica " + XParam
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
        End If
        
        If Factura.Value = True Then
            ZTipo = "01"
        End If
        If Debito.Value = True Then
            ZTipo = "03"
        End If
        If Credito.Value = True Then
            ZTipo = "04"
        End If
        If Expo.Value = True Then
            ZTipo = "02"
        End If
                
        If Val(WCodIva) <> 3 And Val(WCodIva) <> 5 Then
            spNumero = "ConsultaNumero " + "'" + ZTipo + "'"
            Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
            If rstNumero.RecordCount > 0 Then
                WCodigo = ZTipo
                WNumero = Numero.Text
                XParam = "'" + WCodigo + "','" _
                            + WNumero + "'"
                spNumero = "ModificaNumero " + XParam
                Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
        
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                WAuxiliar = !Nombre
            End If
        End With
    
        Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
        Rem Listado.GroupSelectionFormula = "{Pedido.Pedido} in " + Pedido.Text + " to " + Pedido.Text
        Rem Listado.Destination = 1
        Rem Listado.Action = 1
        
        If Expo.Value = False Then
            Call Impresion
                Else
            Call Impresion_Expo
        End If
        
        Call Limpia_Click

        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
        
        Numero.SetFocus
        
        End If
        
End Sub


Private Sub Ingresa_Click()

    WLinea.Text = ""
    WDescripcion.Text = ""
    WImporte.Text = ""
    
    WDescripcion.SetFocus
    
End Sub


Private Sub Limpia_Click()

    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    ReteIb.ListIndex = 0
    ReteCiudad.ListIndex = 0
    
    WLinea.Text = ""
    WDescripcion.Text = ""
    WImporte.Text = ""
  
    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 1
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Total.Caption = ""
    ImpoIb.Caption = ""
    ImpoIbTucu.Caption = ""
    ImpoIbCiudad.Caption = ""
    Paridad.Text = ""
    Dto.Caption = ""
    Interes.Caption = ""
    
    Factura.Value = True
    Debito.Value = False
    Credito.Value = False
    Normal.Value = True
    Exenta.Value = False
    Expo.Value = False
    Ajuste.Value = False
    Pesos.Value = True
    Dolares.Value = False
    
    Graba.Enabled = True
    Borra.Enabled = True
    Ingresa.Enabled = True
    
    If Factura.Value = True Then
        ZTipo = "01"
    End If
    If Debito.Value = True Then
        ZTipo = "03"
    End If
    If Credito.Value = True Then
        ZTipo = "04"
    End If
    
    If Expo.Value = False Then
            
        spNumero = "ConsultaNumero " + "'" + ZTipo + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = "1"
        End If
        
            Else
            
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
        
    End If
    
    Numero.SetFocus

End Sub

Private Sub WDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WImporte.SetFocus
    End If
End Sub

Private Sub WImporte_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WImporte.Text = Pusing("###,###.##", WImporte.Text)
        Call Alta_Vector
        Call Ingresa_Click
        Call Calcula_Click
        WImporte.Text = ""
        WDescripcion.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spClientes = "ConsultaCliente " + "'" + Claveven$ + "'"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                Cliente.Text = rstClientes!Cliente
                DesCliente.Caption = rstClientes!Razon
                WPago1 = 1
                WPago2 = 1
                WVendedor = rstClientes!vendedor
                WProvincia = rstClientes!Provincia
                WRubro = rstClientes!Rubro
                WCodIva = rstClientes!Iva
                WCodIb = rstCliente!Ib
                WCodIbTucu = IIf(IsNull(rstCliente!IbTucu), "0", rstCliente!IbTucu)
                WCodIbCiudad = IIf(IsNull(rstCliente!IbCiudad), "0", rstCliente!IbCiudad)
                WRazon = rstClientes!Razon
                WDireccion = rstClientes!Direccion
                WLocalidad = rstClientes!Localidad
                WPostal = rstClientes!Postal
                WCuit = rstClientes!Cuit
                WDirentrega = rstClientes!DirEntrega
                rstClientes.Close
                Call Calcula_FechaVto
                Vencimiento.Text = Wvencimiento
            End If
            Ayuda.Visible = False
            
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 40 Then
                            DBGrid1.Row = DBGrid1.Row + 1
                            DBGrid1.Col = 0
                            KeyCode = 0
                        End If
                        Call Calcula_Click
                        DBGrid1.Row = WRow

                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
    End Select

    
End Sub


' Cuando el usuario hace clic en el icono Agregar, esta subrutina agrega una
' nueva fila a la variable RowBuf y un marcador a la variable NewRowBookmark
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows = mTotalRows + 1
ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
NewRowBookmark = mTotalRows - 1 'Establece el marcador a la última fila.

' El bucle siguiente agrega un nuevo registro a la base de datos.
For iCol = 0 To UBound(UserData, 1)
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
    Else
        ' Si no se establece ningún valor para la columna, usa DefaultValue
        UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
    End If
Next iCol

End Sub

' Esta subrutina elimina una fila basándose en su marcador.
Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

' Mueve todas las filas encima de la fila eliminada de
' la matriz.

For iRow = Bookmark + 1 To mTotalRows - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows = mTotalRows - 1

End Sub

' Se llama a esta subrutina cada vez que DBGrid quiere mostrar
' datos nuevos.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid está solicitando filas, así que se las damos

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' Si StartLocation es Null, empieza a leer por el final
' o por el principio del conjunto de datos.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Busca la posición para empezar a leer, basándose en el marcador
    ' StartLocation y en la variable iIncr
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfiere datos de nuestra matriz de conjunto de datos al objeto RowBuf
' que DBGrid utiliza para presentar los datos
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
    For iCol = 0 To UBound(UserData, 1)
        RowBuf.Value(iRow, iCol) = UserData(iCol, CurRow&)
    Next iCol
    ' Establece el marcador mediante CurRow&, que es también
    ' nuestro índice de matriz
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

' Esta subrutina actualiza los datos de la matriz después de
' haberse modificado.

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Se están actualizando los datos

' Actualiza cada columna de la matriz de conjuntos de datos
For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub


Private Sub Form_Load()

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
    
    Iva(1) = "Inscripto"
    Iva(2) = "No Inscripto"
    Iva(3) = "Consumidor Final"
    Iva(4) = "Exento"
    Iva(5) = "Monotributo"
    Iva(6) = "No Catalogado"
    
    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"
    
    ReteIb.Clear
    
    ReteIb.AddItem ""
    ReteIb.AddItem "Calcula"
    ReteIb.AddItem "No Calcula"
    
    ReteIb.ListIndex = 0
    
    ReteCiudad.Clear
    
    ReteCiudad.AddItem ""
    ReteCiudad.AddItem "Calcula"
    ReteCiudad.AddItem "No Calcula"
    
    ReteCiudad.ListIndex = 0
    

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 1, 0 To 40)

mTotalRows& = 40

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DBGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DBGrid1.Columns.Count - 1 To 0 Step -1
      DBGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 1
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 6000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Importe"
             DBGrid1.Columns(newcnt).Width = 2000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
         
Next i
 
    Rem DBGrid1.FirstRow = 0
    Rem DBGrid1.Col = 0
    Rem DBGrid1.Row = 0
    
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    WLinea.Text = ""
    WDescripcion.Text = ""
    WImporte.Text = ""
    Renglon = 0
    
    Factura.Value = True
    Debito.Value = False
    Credito.Value = False
    Normal.Value = True
    Exenta.Value = False
    Expo.Value = False
    Ajuste.Value = False
    Pesos.Value = True
    Dolares.Value = False

    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
            spCambios = "ConsultaCambio " + "'" + Fecha.Text + "'"
            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambios.RecordCount > 0 Then
                Paridad.Text = Pusing("###,###.##", Str$(rstCambios!Cambio))
                        Else
                Paridad.Text = ""
            End If
     
    If Factura.Value = True Then
        ZTipo = "01"
    End If
    If Debito.Value = True Then
        ZTipo = "03"
    End If
    If Credito.Value = True Then
        ZTipo = "04"
    End If
    
    If Expo.Value = False Then
            
        spNumero = "ConsultaNumero " + "'" + ZTipo + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = "1"
        End If
        
            Else
            
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    
    End If
     
    Numero.SetFocus
    
End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) = 0 Then

            Renglon = Renglon + 1
            
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
            
            WAnterior = DBGrid1.Row
                
            DBGrid1.Col = 0
            DBGrid1.Text = WDescripcion.Text
            
            If Val(WImporte.Text) <> 0 Then
                DBGrid1.Col = 1
                DBGrid1.Text = Pusing("###,###.##", WImporte.Text)
                    Else
                DBGrid1.Col = 1
                DBGrid1.Text = ""
            End If
            
            DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
            
            WAnterior = DBGrid1.Row
                
            DBGrid1.Col = 0
            DBGrid1.Text = WDescripcion.Text
            
            If Val(WImporte.Text) <> 0 Then
                DBGrid1.Col = 1
                DBGrid1.Text = Pusing("###,###.##", WImporte.Text)
                    Else
                DBGrid1.Col = 1
                DBGrid1.Text = ""
            End If
            
            DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Proceso_Click()

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 1
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    
    If Factura.Value = True Then
        WTipo = "03"
    End If
    If Debito.Value = True Then
        WTipo = "04"
    End If
    If Credito.Value = True Then
        WTipo = "05"
    End If
    
    XParam = "'" + WTipo + "','" _
                + Numero.Text + "'"
    
    spDesccomp = "ConsultaDesccomp1 " + XParam
    Set rstDesccomp = db.OpenRecordset(spDesccomp, dbOpenSnapshot, dbSQLPassThrough)
    If rstDesccomp.RecordCount > 0 Then
    
        With rstDesccomp
            .MoveFirst
            Do
                If .EOF = False Then
                
                Renglon = Renglon + 1
            
                Lugar1 = Int((Renglon - 1) / 10) * 10
                Lugar2 = Renglon - Lugar1
                
                DBGrid1.FirstRow = Lugar1
                DBGrid1.Row = Lugar2 - 1
                
                DBGrid1.Col = 0
                DBGrid1.Text = !Descripcion
                
                If !Importe <> 0 Then
                    DBGrid1.Col = 1
                    DBGrid1.Text = Pusing("###,###.##", Str$(!Importe))
                        Else
                    DBGrid1.Col = 1
                    DBGrid1.Text = ""
                End If
    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstDesccomp.Close
    End If
    
    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    Call Calcula_Click
    
    Graba.Enabled = False
    Borra.Enabled = False
    Ingresa.Enabled = False

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Factura.Value = True Then
            WTipo = "03"
        End If
        If Debito.Value = True Then
            WTipo = "04"
        End If
        If Credito.Value = True Then
            WTipo = "05"
        End If
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        ClaveCtacte = WTipo + Auxi + "01"
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
                Fecha.Text = rstCtacte!Fecha
                Cliente.Text = rstCtacte!Cliente
                Vencimiento.Text = rstCtacte!Vencimiento
                Paridad.Text = rstCtacte!Paridad
                WMoneda = IIf(IsNull(rstCtacte!Moneda), "0", rstCtacte!Moneda)
                If Val(WMoneda) = 0 Then
                    Pesos.Value = True
                    Dolares.Value = False
                        Else
                    Pesos.Value = False
                    Dolares.Value = True
                End If
                rstCtacte.Close
                
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    Cliente.Text = rstCliente!Cliente
                    DesCliente.Caption = rstCliente!Razon
                    WPago1 = 1
                    WPago2 = 1
                    WVendedor = rstCliente!vendedor
                    WProvincia = rstCliente!Provincia
                    WRubro = rstCliente!Rubro
                    WCodIva = rstCliente!Iva
                    WCodIb = rstCliente!Ib
                    WCodIbTucu = IIf(IsNull(rstCliente!IbTucu), "0", rstCliente!IbTucu)
                    WCodIbCiudad = IIf(IsNull(rstCliente!IbCiudad), "0", rstCliente!IbCiudad)
                    WRazon = rstCliente!Razon
                    WDireccion = rstCliente!Direccion
                    WLocalidad = rstCliente!Localidad
                    WPostal = rstCliente!Postal
                    WCuit = rstCliente!Cuit
                    WDirentrega = rstCliente!DirEntrega
                End If
                Call Proceso_Click
                    Else
                Rem .Index = "Numero"
                Rem .Seek "=", Val(Numero.Text)
                Rem If .NoMatch = False Then
                Rem     m$ = "Comprobante ya existente"
                Rem   A% = MsgBox(m$, 0, "Ingreso de comprobantes varias")
                Rem     Numero.SetFocus
                Rem        Else
                Rem    Graba.Enabled = True
                Rem    Borra.Enabled = True
                Rem    Ingresa.Enabled = True
                Rem    WNumero = Numero.Text
                Rem    Numero.Text = WNumero
                Rem    Fecha.SetFocus
                Rem End If
                Graba.Enabled = True
                Borra.Enabled = True
                Ingresa.Enabled = True
                WNumero = Numero.Text
                Numero.Text = WNumero
                Fecha.SetFocus
                
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Cliente.Text = rstCliente!Cliente
            DesCliente.Caption = rstCliente!Razon
            WPago1 = 1
            WPago2 = 1
            WVendedor = rstCliente!vendedor
            WProvincia = rstCliente!Provincia
            WRubro = rstCliente!Rubro
            WCodIva = rstCliente!Iva
            WCodIb = rstCliente!Ib
            WCodIbTucu = IIf(IsNull(rstCliente!IbTucu), "0", rstCliente!IbTucu)
            WCodIbCiudad = IIf(IsNull(rstCliente!IbCiudad), "0", rstCliente!IbCiudad)
            WRazon = rstCliente!Razon
            WDireccion = rstCliente!Direccion
            WLocalidad = rstCliente!Localidad
            WPostal = rstCliente!Postal
            WCuit = rstCliente!Cuit
            WDirentrega = rstCliente!DirEntrega
            rstCliente.Close
            Call Calcula_FechaVto
            Vencimiento.Text = Wvencimiento
            NroFactura.SetFocus
            If Exenta.Value = True Or Expo.Value = True Then
                NroFactura.Text = ""
                DBGrid1.FirstRow = 0
                DBGrid1.Col = 0
                DBGrid1.Row = 0
                DBGrid1.SetFocus
                    Else
                NroFactura.SetFocus
            End If
                Else
            Cliente.SetFocus
        End If
    End If
End Sub

Private Sub NroFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Pasa = "S"
        If Normal.Value = True Or Ajuste.Value = True Then
            WNro = NroFactura.Text
            Call Ceros(WNro, 8)
            WClaveCtacte = "01" + WNro + "01"
            spCtacte = "ConsultaCtacte " + "'" + WClaveCtacte + "'"
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtacte.RecordCount > 0 Then
                If rstCtacte!Cliente <> Cliente.Text Then
                    m$ = "El cliente de la factura no se corresponde al cliente informado"
                    a% = MsgBox(m$, 0, "Emision de Comprobantes Varios")
                    Pasa = "N"
                End If
                rstCtacte.Close
                    Else
                m$ = "No existe el numero de comprobante informado"
                a% = MsgBox(m$, 0, "Emision de Comprobantes Varios")
                Pasa = "N"
            End If
                Else
            NroFactura.Text = ""
        End If
        If Pasa = "S" Then
            DBGrid1.FirstRow = 0
            DBGrid1.Col = 0
            DBGrid1.Row = 0
            DBGrid1.SetFocus
                Else
            NroFactura.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            spCambios = "ConsultaCambio " + "'" + Fecha.Text + "'"
            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambios.RecordCount > 0 Then
                Paridad.Text = Pusing("###,###.##", Str$(rstCambios!Cambio))
                        Else
                Paridad.Text = ""
            End If
            If Val(Paridad.Text) <> 0 Then
                Call Calcula_FechaVto
                Vencimiento.Text = Wvencimiento
                Cliente.SetFocus
                    Else
                m$ = "No exsite paridad cargada para esta fecha"
                a% = MsgBox(m$, 0, "Emision de Comprobante varios")
                Fecha.SetFocus
            End If
                Else
            m$ = "Formato de fecha invalido"
            a% = MsgBox(m$, 0, "Emision de Comprobante varios")
            Fecha.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Vencimiento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi = "S" Then
            Remito.SetFocus
                Else
            Vencimiento.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Impresion()

    If Val(WEmpresa) = 1 Then
        Rem Open "dada.txt" For Output As #1
        Open "lpt1" For Output As #1
            Else
        Open "lpt1" For Output As #1
        Rem Open "DADA.TXT" For Output As #1
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
    End If

    Rem Width #1, 255

    Print #1, Chr$(27) + Chr$(40) + "19U";
    Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "1" + Chr$(72);
    Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)

    For XX% = 1 To 2
    
        If Val(WEmpresa) = 1 Then
            Rem Print #1, ""
        End If

        Select Case Val(WEmpresa)
            Case 1
                Print #1, ""
                Print #1, ""
                Print #1, ""
            Case Else
                If Debito.Value <> True Then
                    Print #1, ""
                    Print #1, ""
                    Print #1, ""
                        Else
                    Print #1, ""
                    Print #1, ""
                End If
        End Select
        If Factura.Value = True Then
            Print #1, Tab(55); "FACTURA"
        End If
        If Debito.Value = True Then
            If Val(WEmpresa) = 1 Then
                Print #1, Tab(55); "NOTA DE DEBITO"
                    Else
                Print #1, Tab(55); ""
            End If
        End If
        If Credito.Value = True Then
            Print #1, Tab(55); "NOTA DE CREDITO"
        End If
        Print #1, ""
        Print #1, Tab(59); Fecha.Text
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(8); WRazon
        Print #1, Tab(8); WDireccion
        Print #1, Tab(8); Left$(WLocalidad, 33);
        Print #1, Tab(55); Cliente.Text
        Print #1, Tab(8); Provincia(Val(WProv)); " ("; WPostal; ")"
        Print #1, ""
        Print #1, Tab(8); Iva(Val(WCodIva));
        Print #1, Tab(48); WCuit
        Print #1, ""
        Print #1, ""
        Print #1, Tab(5); WPago
        
        If Val(WEmpresa) <> 1 And Debito.Value = True Then
            Print #1, ""
        End If
        
        Print #1, ""
        Print #1, ""
        Print #1, ""

        Impre = 0

        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
            
                If Impre < 22 Then
                
                    WRow = iRow
                    DBGrid1.Row = WRow
                    
                    DBGrid1.Col = 0
                    Descri = DBGrid1.Text
                
                    DBGrid1.Col = 1
                    parcial = Val(DBGrid1.Text)
                    
                    If Descri <> "" Or parcial <> 0 Then
                        Print #1, Tab(15); Left$(Descri, 48);
                        If parcial <> 0 Then
                            Rem If Val(WCodIva) = 1 Or Val(WCodIva) = 2 Then
                            Rem     Print #1, Tab(68); Alinea("###,###.##", Str$(Parcial))
                            Rem        Else
                            Rem    Parcial = Str$(Val(Parcial) * 1.21)
                            Rem    Print #1, Tab(68); Alinea("###,###.##", Str$(Parcial))
                            Rem End If
                            If Pesos.Value = True Then
                                Print #1, Tab(65); "$"; Alinea("###,###.##", Str$(parcial))
                                    Else
                                Print #1, Tab(65); "U$S"; Alinea("###,###.##", Str$(parcial))
                            End If
                        End If
                        Print #1, ""
                        Impre = Impre + 1
                    End If
               End If
                    
            Next iRow
            
        Next a
        
        Rem OJO
        Rem For aa = Impre To 21
        For aa = Impre To 22
                Print #1, ""
        Next aa

        Rem M# = Total# / 100
        Rem GoSub 4630
        
        Rem dada
        Rem dada
        Rem dada
        Rem dada
        
        Select Case Val(WCodIva)
            Case 3, 5
                If Pesos.Value = True Then
        
                    Impotot = Val(Total.Caption) / Val(Paridad.Text)

                    Print #1, ""
                    Print #1, ""
        
                    Print #1, Tab(1); "EL PRESENTE COMPROBANTE EQUIVALE A U$S ";
                    Print #1, Alinea("###,###.##", Str$(Impotot))
                    Print #1, Tab(1); "CALCULADOS AL TIPO DE CAMBIO DE ";
                    Print #1, Alinea("##.##", Paridad.Text)
        
                    Print #1, ""
        
                        Else
        
                    ImporteIb = Val(ImpoIb.Caption) * Val(Paridad.Text)
                    ImporteIbTucu = Val(ImpoIbTucu.Caption) * Val(Paridad.Text)
                    ImporteIbCiudad = Val(ImpoIbCiudad.Caption) * Val(Paridad.Text)
                    ImpoNeto = Val(Neto.Caption) * Val(Paridad.Text)
                    ImpoIva = (Val(Iva1.Caption) + Val(Iva2.Caption)) * Val(Paridad.Text)
                    Impotot = Val(Total.Caption) * Val(Paridad.Text)
    
                    If Val(WEmpresa) = 1 Then
        
                        Print #1, Tab(1); "ESTE IMPORTE ESTA EXPRESADO EN DOLARES ESTADOUNIDENSES."
                        Print #1, Tab(1); "REEXPRESION EN PESOS AL SOLO EFECTO CONTABLE/IMPOSITIVO"
                        Print #1, Tab(1); "TIPO DE CAMBIO:";
                        Print #1, Alinea("##.##", Paridad.Text);
                        Rem Print #1, " I.V.A.:";
                        Rem Print #1, Alinea("#,###.##", Str$(ImpoIva));
                        Print #1, " TOTAL:";
                        Print #1, Alinea("###,###.##", Str$(Impotot))
                        Rem Print #1, Tab(1); "NETO:";
                        Rem Print #1, Alinea("###,###.##", Str$(ImpoNeto));
                        If ImporteIb <> 0 Then
                            Print #1, " I.BRUTOS BS. AS.:";
                            Print #1, Alinea("#,###.##", Str$(ImporteIb));
                        End If
                        If ImporteIbTucu <> 0 Then
                            Print #1, " I.BRUTOS TUCUMAN:";
                            Print #1, Alinea("#,###.##", Str$(ImporteIbTucu));
                        End If
                        If ImporteIbCiudad <> 0 Then
                            Print #1, " I.BRUTOS CIUDAD:";
                            Print #1, Alinea("#,###.##", Str$(ImporteIbCiudad))
                                Else
                            Print #1, ""
                        End If
                
                            Else
                    
                        Print #1, Tab(3); "ESTE IMPORTE ESTA EXPRESADO EN DOLARES ESTADOUNIDENSES."
                        Print #1, Tab(3); "REEXPRESION EN PESOS AL SOLO EFECTO CONTABLE/IMPOSITIVO"
                        Print #1, Tab(3); "TIPO DE CAMBIO:";
                        Print #1, Alinea("##.##", Paridad.Text);
                        Rem Print #1, " I.V.A.:";
                        Rem Print #1, Alinea("#,###.##", Str$(ImpoIva));
                        Print #1, " TOTAL:";
                        Print #1, Alinea("###,###.##", Str$(Impotot))
                        Rem Print #1, Tab(3); "IMPORTE NETO:";
                        Rem Print #1, Alinea("###,###.##", Str$(ImpoNeto));
                        If ImporteIb <> 0 Then
                            Print #1, " PERCEPCION DE INGRESOS BRUTOS:";
                            Print #1, Alinea("#,###.##", Str$(ImporteIb))
                                Else
                            Print #1, ""
                        End If
                
                    End If
        
                    Print #1, ""
                    Print #1, ""
        
                End If
        
                If Val(WEmpresa) <> 1 Then
                    Print #1, ""
                End If
        
                Rem If Pesos.Value = True Then
                Rem     Print #1, Tab(65); "$"; Alinea("###,###.##", Str$(XNeto))
                Rem         Else
                Rem     Print #1, Tab(65); "U$S"; Alinea("###,###.##", Str$(XNeto))
                Rem End If
                Print #1, ""

                If Val(Interes.Caption) <> 0 Then
                        Print #1, Tab(56); "Interes";
                        If Pesos.Value = True Then
                            Print #1, Tab(65); "$"; Alinea("###,###.##", Interes.Caption)
                                Else
                            Print #1, Tab(65); "U$S"; Alinea("###,###.##", Interes.Caption)
                        End If
                                                  Else
                        Print #1, ""
                End If

                If Val(Dto.Caption) <> 0 Then
                        Print #1, Tab(56); "Dto"; Alinea("##.##", Str$(WDescuento));
                        If Pesos.Value = True Then
                            Print #1, Tab(65); "$"; Alinea("###,###.##", Dto.Caption)
                                Else
                            Print #1, Tab(65); "U$S"; Alinea("###,###.##", Dto.Caption)
                        End If
                                Else
                        Print #1, ""
                End If

                Print #1, Tab(3); Left$(M1, 60)
                Print #1, Tab(3); Left$(M2, 55)
        
                If Pesos.Value = True Then
                    Print #1, Tab(65); "$"; Alinea("###,###.##", Total.Caption);
                        Else
                    Print #1, Tab(65); "U$S"; Alinea("###,###.##", Total.Caption);
                End If
                Print #1, Chr$(12)
            
            Case Else
                If Pesos.Value = True Then
        
                    Impotot = Val(Total.Caption) / Val(Paridad.Text)

                    Print #1, ""
                    Print #1, ""
        
                    Print #1, Tab(1); "EL PRESENTE COMPROBANTE EQUIVALE A U$S ";
                    Print #1, Alinea("###,###.##", Str$(Impotot))
                    Print #1, Tab(1); "CALCULADOS AL TIPO DE CAMBIO DE ";
                    Print #1, Alinea("##.##", Paridad.Text)
        
                    Print #1, ""
        
                        Else
        
                    ImporteIb = Val(ImpoIb.Caption) * Val(Paridad.Text)
                    ImporteIbTucu = Val(ImpoIbTucu.Caption) * Val(Paridad.Text)
                    ImporteIbCiudad = Val(ImpoIbCiudad.Caption) * Val(Paridad.Text)
                    ImpoNeto = Val(Neto.Caption) * Val(Paridad.Text)
                    ImpoIva = (Val(Iva1.Caption) + Val(Iva2.Caption)) * Val(Paridad.Text)
                    Impotot = Val(Total.Caption) * Val(Paridad.Text)
    
                    If Val(WEmpresa) = 1 Then
        
                        Print #1, Tab(1); "ESTE IMPORTE ESTA EXPRESADO EN DOLARES ESTADOUNIDENSES."
                        Print #1, Tab(1); "REEXPRESION EN PESOS AL SOLO EFECTO CONTABLE/IMPOSITIVO"
                        Print #1, Tab(1); "TIPO DE CAMBIO:";
                        Print #1, Alinea("##.##", Paridad.Text);
                        Print #1, " I.V.A.:";
                        Print #1, Alinea("#,###.##", Str$(ImpoIva));
                        Print #1, " TOTAL:";
                        Print #1, Alinea("###,###.##", Str$(Impotot))
                        Print #1, Tab(1); "NETO:";
                        Print #1, Alinea("###,###.##", Str$(ImpoNeto));
                        If ImporteIb <> 0 Then
                            Print #1, " I.BRUTOS BS. AS.:";
                            Print #1, Alinea("#,###.##", Str$(ImporteIb));
                        End If
                        If ImporteIbTucu <> 0 Then
                            Print #1, " I.BRUTOS TUCUMAN:";
                            Print #1, Alinea("#,###.##", Str$(ImporteIbTucu));
                        End If
                        If ImporteIbCiudad <> 0 Then
                            Print #1, " I.BRUTOS CIUDAD:";
                            Print #1, Alinea("#,###.##", Str$(ImporteIbCiudad))
                                Else
                            Print #1, ""
                        End If
                
                            Else
                    
                        Print #1, Tab(3); "ESTE IMPORTE ESTA EXPRESADO EN DOLARES ESTADOUNIDENSES."
                        Print #1, Tab(3); "REEXPRESION EN PESOS AL SOLO EFECTO CONTABLE/IMPOSITIVO"
                        Print #1, Tab(3); "TIPO DE CAMBIO:";
                        Print #1, Alinea("##.##", Paridad.Text);
                        Print #1, " I.V.A.:";
                        Print #1, Alinea("#,###.##", Str$(ImpoIva));
                        Print #1, " TOTAL:";
                        Print #1, Alinea("###,###.##", Str$(Impotot))
                        Print #1, Tab(3); "IMPORTE NETO:";
                        Print #1, Alinea("###,###.##", Str$(ImpoNeto));
                        If ImporteIb <> 0 Then
                            Print #1, " PERCEPCION DE INGRESOS BRUTOS:";
                            Print #1, Alinea("#,###.##", Str$(ImporteIb))
                                Else
                            Print #1, ""
                        End If
                
                    End If
        
                    Print #1, ""
                    Print #1, ""
        
                End If
        
                If Val(WEmpresa) <> 1 Then
                    Print #1, ""
                End If
        
                If Pesos.Value = True Then
                    Print #1, Tab(65); "$"; Alinea("###,###.##", Str$(XNeto))
                        Else
                    Print #1, Tab(65); "U$S"; Alinea("###,###.##", Str$(XNeto))
                End If

                If Val(Interes.Caption) <> 0 Then
                        Print #1, Tab(56); "Interes";
                        If Pesos.Value = True Then
                            Print #1, Tab(65); "$"; Alinea("###,###.##", Interes.Caption)
                                Else
                            Print #1, Tab(65); "U$S"; Alinea("###,###.##", Interes.Caption)
                        End If
                                                  Else
                        Print #1, ""
                End If

                If Val(Dto.Caption) <> 0 Then
                        Print #1, Tab(56); "Dto"; Alinea("##.##", Str$(WDescuento));
                        If Pesos.Value = True Then
                            Print #1, Tab(65); "$"; Alinea("###,###.##", Dto.Caption)
                                Else
                            Print #1, Tab(65); "U$S"; Alinea("###,###.##", Dto.Caption)
                        End If
                                Else
                        Print #1, ""
                End If

                Print #1, Tab(3); Left$(M1, 60);
                If Pesos.Value = True Then
                    Print #1, Tab(65); "$"; Alinea("###,###.##", Neto.Caption)
                        Else
                    Print #1, Tab(65); "U$S"; Alinea("###,###.##", Neto.Caption)
                End If
                Print #1, Tab(3); Left$(M2, 55);
                If Val(Iva1.Caption) <> 0 Then
                        Print #1, Tab(61); "10.5";
                        If Pesos.Value = True Then
                            Print #1, Tab(65); "$"; Alinea("###,###.##", Iva1.Caption)
                                Else
                            Print #1, Tab(65); "U$S"; Alinea("###,###.##", Iva1.Caption)
                        End If
                                Else
                        Print #1, ""
                End If

                If Val(WEmpresa) <> 1 And Debito.Value = True Then
                    Print #1, Tab(10); "";
                        Else
                    Select Case XX
                            Case 1
                                    Print #1, Tab(3); "ORIGINAL";
                            Case 2
                                    Print #1, Tab(3); "DUPLICADO";
                            Case 3
                                    Print #1, Tab(3); "TRIPLICADO";
                            Case Else
                    End Select
                End If
        
                If Val(ImpoIbCiudad.Caption) <> 0 Then
                    If Pesos.Value = True Then
                        Print #1, Tab(14); "P.Ciudad";
                        Print #1, Tab(23); " $ "; Alinea("##,###.##", ImpoIbCiudad.Caption);
                            Else
                        Print #1, Tab(14); "P.Ciudad";
                        Print #1, Tab(23); "U$S"; Alinea("##,###.##", ImpoIbCiudad.Caption);
                    End If
                End If
                If Val(ImpoIbTucu.Caption) <> 0 Then
                    If Pesos.Value = True Then
                        Print #1, Tab(36); "P.Tucuman";
                        Print #1, Tab(46); " $ "; Alinea("##,###.##", ImpoIbTucu.Caption);
                            Else
                        Print #1, Tab(36); "P.Tucuman";
                        Print #1, Tab(46); "U$S"; Alinea("##,###.##", ImpoIbTucu.Caption);
                    End If
                End If
                If Val(ImpoIb.Caption) <> 0 Then
                        Print #1, Tab(60); "I.B.";
                        If Pesos.Value = True Then
                            Print #1, Tab(65); " $ "; Alinea("##,###.##", ImpoIb.Caption)
                                Else
                            Print #1, Tab(65); "U$S"; Alinea("##,###.##", ImpoIb.Caption)
                        End If
                                Else
                        If Val(Iva2.Caption) <> 0 Then
                            Print #1, Tab(60); "10.5";
                            If Pesos.Value = True Then
                                Print #1, Tab(65); "$"; Alinea("##,###.##", Iva2.Caption)
                                    Else
                                Print #1, Tab(65); "U$S"; Alinea("##,###.##", Iva2.Caption)
                            End If
                                Else
                            Print #1, ""
                        End If
                End If

                If Pesos.Value = True Then
                    Print #1, Tab(65); "$"; Alinea("###,###.##", Total.Caption);
                        Else
                    Print #1, Tab(65); "U$S"; Alinea("###,###.##", Total.Caption);
                End If
                Print #1, Chr$(12)
        End Select

        Next XX%

        Close #1

End Sub



Sub Impresion_Expo()

    Open "LPT1" For Output As #99

    Print #99, Chr$(27) + Chr$(40) + "19U";
    Print #99, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
    Print #99, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
    Print #99, Chr$(27) + Chr$(38) + Chr$(108) + "1" + Chr$(72)
    Print #99, Chr$(27) + Chr$(40) + Chr$(115) + "12" + Chr$(72)

    For XX = 1 To 1

        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Rem Print #99, Tab(18); Left$(fecorden.Text, 2);
        Rem Print #99, Tab(21); Mid$(fecorden.Text, 4, 2);
        Rem Print #99, Tab(24); Right$(fecorden.Text, 2);
        Rem Print #99, Tab(27); Left$(NroOrden.Text, 6);
        Rem Print #99, Tab(37); Consignatario.Text;
        Print #99, Tab(68); Left$(Fecha.Text, 2);
        Print #99, Tab(71); Mid$(Fecha.Text, 4, 2);
        Print #99, Tab(74); Right$(Fecha.Text, 2)
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""

        Rem Print #99, Tab(45); Envio1.Text
        Print #99, ""
        Print #99, Tab(3); Left$(WRazon, 40)
        Rem Print #99, Tab(45); Envio2.Text

        Print #99, ""
        Print #99, Tab(3); Left$(WDireccion, 40)
        Rem Print #99, Tab(45); Pago1.Text
        Print #99, Tab(3); Left$(WLocalidad, 40)
        Rem Print #99, Tab(45); Pago2.Text
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, Tab(85); "USD"
        Print #99, ""
        Print #99, ""

        Suma1 = 0
        Suma2 = 0
        Suma3 = 0
        WRenglon = 0
        
        Impre = 0
        
        
        
        
        
        
        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
            
                If Impre < 22 Then
                
                    WRow = iRow
                    DBGrid1.Row = WRow
                    
                    DBGrid1.Col = 0
                    Descri = DBGrid1.Text
                
                    DBGrid1.Col = 1
                    parcial = Val(DBGrid1.Text)
                    
                    If Descri <> "" Or parcial <> 0 Then
                        Print #99, Tab(22); Left$(Descri, 48);
                        Print #99, Tab(79); "USD ";
                        Print #99, Tab(83); Alinea("###,###.##", Str$(parcial))
                        Impre = Impre + 1
                    End If
                    
               End If
                    
            Next iRow
            
        Next a
        
        For DA = Impre To 23
            Print #99, ""
        Next DA

        Print #99, Tab(5); "Todas las disputas que puedan surgir en el presente contrato seran finalmente arregladas"
        Print #99, Tab(5); "de acuerdo a las Reglas de Conciliacion y Arbitraje  de  la  Camara  Internacional   de"
        Print #99, Tab(5); "Comercio por uno o mas arbitros de acuerdo de dichas reglas"
        Print #99, Tab(5); "INCOTERMS 1990";
        
        Call Numtolet
        
        WTexto1 = UCase(WTexto1)
        WTexto2 = UCase(WTexto2)

        Print #99, Tab(22); "Son Dolares estadounidenses"
        Print #99, Tab(20); WTexto1

        Rem Print #99, Tab(2); Alinea("###", Str$(Suma1));
        Rem Print #99, Tab(20); WTexto2;
        Rem Print #99, Tab(60); Alinea("#####.#", Str$(Suma2));
        Rem Print #99, Tab(68); Alinea("#####", Str$(Suma3));
        Print #99, Tab(79); "USD ";
        Print #99, Tab(83); Alinea("###,###.##", Neto.Caption)
        Print #99, ""
        Print #99, Tab(79); "USD ";
        Print #99, Tab(83); Alinea("###,###.##", Neto.Caption)
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, Tab(79); "USD ";
        Print #99, Tab(83); Alinea("###,###.##", Total.Caption)

    Next XX
        
    Close #99
End Sub
        







Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
            
                    DA = Len(rstCliente!Razon) - WEspacios
                
                    For aa = 1 To DA
                        If Left$(Ayuda.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                            Auxi = rstCliente!Cliente
                            IngresaItem = Auxi + "    " + rstCliente!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
                            WIndice.AddItem IngresaItem
                            Exit For
                        End If
                    Next aa
                    .MoveNext
                    
                        Else
                        
                    Exit Do
                
                End If
            Loop
        End With
        rstCliente.Close
    End If
    End If

End Sub

Private Sub Numtolet()

    'Convertir en letras el número en Text1
    
    Dim Numero As String
    Dim Letras As String
    Dim sCentimos As String
    Dim sMoneda As String
            
    sMoneda = "dolares"
    sCentimos = "centavos"
    
    Numero = CStr(Val(Total.Caption))
    
    WTexto1 = Numero2Letra(Numero, , sMoneda & " ", sCentimos & " ")
    WTexto1 = WTexto1 + Space$(50)
    
    Pasa = 0
    
    For DA = 40 To 1 Step -1
        If Mid$(WTexto1, DA, 1) = Space$(1) Then
            Pasa = 1
        End If
        If Pasa = 1 Then
            If Mid$(WTexto1, DA, 1) <> Space$(1) Then
                Exit For
            End If
        End If
    Next DA
    
    WTexto2 = Mid$(WTexto1, DA + 2, 35)
    WTexto1 = Left$(WTexto1, DA)
    
End Sub


