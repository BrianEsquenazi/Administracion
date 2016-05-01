VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAltaMinuta 
   AutoRedraw      =   -1  'True
   Caption         =   "Generacion de Minuta de Cobranza"
   ClientHeight    =   8340
   ClientLeft      =   600
   ClientTop       =   435
   ClientWidth     =   10515
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   10515
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
      Height          =   615
      Left            =   5520
      TabIndex        =   26
      Top             =   7200
      Width           =   3615
      Begin VB.OptionButton Pesos 
         Caption         =   "Pesos"
         Height          =   255
         Left            =   720
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Dolares 
         Caption         =   "Dolares"
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Marca 
      Caption         =   "Marca"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   120
      TabIndex        =   23
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Emite 
      Caption         =   "Emite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1920
      TabIndex        =   21
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox ObservacionesII 
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   1
      Text            =   " "
      Top             =   960
      Width           =   5895
   End
   Begin VB.TextBox Observaciones 
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   19
      Text            =   " "
      Top             =   1320
      Width           =   5895
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   9
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   8
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   7
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   6
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   5
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   4
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   3
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3960
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   5295
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9340
      _Version        =   327680
      BackColor       =   16776960
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3720
      TabIndex        =   7
      Top             =   7200
      Width           =   1455
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4440
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin VB.TextBox Cliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   3
      Text            =   " "
      Top             =   120
      Width           =   1695
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
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
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label ACobrar 
      Alignment       =   1  'Right Justify
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
      Height          =   375
      Left            =   8400
      TabIndex        =   25
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "A Cobrar"
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
      Height          =   495
      Left            =   8400
      TabIndex        =   24
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Total"
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
      Height          =   375
      Left            =   8400
      TabIndex        =   22
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Observaciones"
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
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha"
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
      Left            =   240
      TabIndex        =   18
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Saldo 
      Alignment       =   1  'Right Justify
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
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label DesCliente 
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
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "PrgAltaMinuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private Importe1 As Double
Private Importe2 As Double
Private Importe3 As Double
Private WTipo As Integer
Private WSaldo As Double
Private Acumula As Double
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim XParam As String
Private WNume As String

Private Sub cmdClose_Click()
    PrgAltaMinuta.Hide
    Unload Me
    PrgMiraAgenda.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


Private Sub Dolares_Click()
    Call Proceso_Click
End Sub

Private Sub Emite_Click()

    ZCodigo = "1"
    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Minuta"
    spMinuta = ZSql
    Set rstminuta = db.OpenRecordset(spMinuta, dbOpenSnapshot, dbSQLPassThrough)
    If rstminuta.RecordCount > 0 Then
        rstminuta.MoveLast
        ZUltimo = IIf(IsNull(rstminuta!CodigoMayor), "0", rstminuta!CodigoMayor)
        ZCodigo = Str$(ZUltimo + 1)
        rstminuta.Close
    End If
    
    If Pesos.Value = True Then
        ZMoneda = "0"
            Else
        ZMoneda = "1"
    End If
    
    ZRenglon = "0"
    For Ciclo = 1 To 1000
        If WVector1.TextMatrix(Ciclo, 0) = "X" Then
        
            ZCliente = Cliente.Text
            ZRenglon = Str$(Val(ZRenglon) + 1)
            ZRazon = DesCliente.Caption
            ZFechaAlta = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZHorario = ""
            ZFecha = Fecha.Text
            ZFechaFactura = WVector1.TextMatrix(Ciclo, 3)
            ZTipo = WVector1.TextMatrix(Ciclo, 1)
            ZFactura = WVector1.TextMatrix(Ciclo, 2)
            ZImporte = WVector1.TextMatrix(Ciclo, 6)
            ZObservaciones = Observaciones.Text
            ZObservacionesII = ObservacionesII.Text
            ZDireccion = ""
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                ZDireccion = rstCliente!Direccion
                ZLocalidad = rstCliente!Localidad
                rstCliente.Close
            End If
            ZOrdFechaAlta = Right$(ZFechaAlta, 4) + Mid$(ZFechaAlta, 4, 2) + Left$(ZFechaAlta, 2)
            ZOrdFecha = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
            
            Auxi = ZCodigo
            Call Ceros(Auxi, 6)
            Auxi1 = ZRenglon
            Call Ceros(Auxi1, 2)
            ZClave = Auxi + Auxi1
            
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Minuta ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "FechaAlta ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Razon ,"
            ZSql = ZSql + "Direccion ,"
            ZSql = ZSql + "Localidad ,"
            ZSql = ZSql + "Horario ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "FechaFactura ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Factura ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "Moneda ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "OrdFechaAlta ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "ObservacionesII )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZClave + "',"
            ZSql = ZSql + "'" + ZCodigo + "',"
            ZSql = ZSql + "'" + ZRenglon + "',"
            ZSql = ZSql + "'" + ZFechaAlta + "',"
            ZSql = ZSql + "'" + ZCliente + "',"
            ZSql = ZSql + "'" + ZRazon + "',"
            ZSql = ZSql + "'" + ZDireccion + "',"
            ZSql = ZSql + "'" + ZLocalidad + "',"
            ZSql = ZSql + "'" + ZHorario + "',"
            ZSql = ZSql + "'" + ZFecha + "',"
            ZSql = ZSql + "'" + ZFechaFactura + "',"
            ZSql = ZSql + "'" + ZTipo + "',"
            ZSql = ZSql + "'" + ZFactura + "',"
            ZSql = ZSql + "'" + ZImporte + "',"
            ZSql = ZSql + "'" + ZMoneda + "',"
            ZSql = ZSql + "'" + ZOrdFecha + "',"
            ZSql = ZSql + "'" + ZOrdFechaAlta + "',"
            ZSql = ZSql + "'" + ZObservaciones + "',"
            ZSql = ZSql + "'" + ZObservacionesII + "')"

            spMinuta = ZSql
            Set rstminuta = db.OpenRecordset(spMinuta, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    Next Ciclo
    
    If Val(ZRenglon) = 0 Then
    
        ZCliente = Cliente.Text
        ZRenglon = Str$(Val(ZRenglon) + 1)
        ZRazon = DesCliente.Caption
        ZFechaAlta = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        ZHorario = ""
        ZFecha = Fecha.Text
        ZFechaFactura = ""
        ZTipo = ""
        ZFactura = ""
        ZImporte = ""
        ZObservaciones = Observaciones.Text
        ZObservacionesII = ObservacionesII.Text
        ZDireccion = ""
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZDireccion = rstCliente!Direccion
            ZLocalidad = rstCliente!Localidad
            rstCliente.Close
        End If
        ZOrdFechaAlta = Right$(ZFechaAlta, 4) + Mid$(ZFechaAlta, 4, 2) + Left$(ZFechaAlta, 2)
        ZOrdFecha = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
        
        Auxi = ZCodigo
        Call Ceros(Auxi, 6)
        Auxi1 = ZRenglon
        Call Ceros(Auxi1, 2)
        ZClave = Auxi + Auxi1
        
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Minuta ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "FechaAlta ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Razon ,"
        ZSql = ZSql + "Direccion ,"
        ZSql = ZSql + "Localidad ,"
        ZSql = ZSql + "Horario ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "FechaFactura ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Factura ,"
        ZSql = ZSql + "Importe ,"
        ZSql = ZSql + "Moneda ,"
        ZSql = ZSql + "OrdFecha ,"
        ZSql = ZSql + "OrdFechaAlta ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "ObservacionesII )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZClave + "',"
        ZSql = ZSql + "'" + ZCodigo + "',"
        ZSql = ZSql + "'" + ZRenglon + "',"
        ZSql = ZSql + "'" + ZFechaAlta + "',"
        ZSql = ZSql + "'" + ZCliente + "',"
        ZSql = ZSql + "'" + ZRazon + "',"
        ZSql = ZSql + "'" + ZDireccion + "',"
        ZSql = ZSql + "'" + ZLocalidad + "',"
        ZSql = ZSql + "'" + ZHorario + "',"
        ZSql = ZSql + "'" + ZFecha + "',"
        ZSql = ZSql + "'" + ZFechaFactura + "',"
        ZSql = ZSql + "'" + ZTipo + "',"
        ZSql = ZSql + "'" + ZFactura + "',"
        ZSql = ZSql + "'" + ZImporte + "',"
        ZSql = ZSql + "'" + ZMoneda + "',"
        ZSql = ZSql + "'" + ZOrdFecha + "',"
        ZSql = ZSql + "'" + ZOrdFechaAlta + "',"
        ZSql = ZSql + "'" + ZObservaciones + "',"
        ZSql = ZSql + "'" + ZObservacionesII + "')"

        spMinuta = ZSql
        Set rstminuta = db.OpenRecordset(spMinuta, dbOpenSnapshot, dbSQLPassThrough)
    
    End If
    
    DoEvents

   Rem by nan
   
    Listado.WindowTitle = "Minuta de Cobranza"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
            
    Listado.ReportFileName = "Minuta.rpt"
                
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.GroupSelectionFormula = "{Minuta.Codigo} in " + ZCodigo + " to " + ZCodigo
    Listado.SelectionFormula = "{Minuta.Codigo} in " + ZCodigo + " to " + ZCodigo

    Listado.SQLQuery = "SELECT Minuta.Codigo, Minuta.Renglon, Minuta.FechaAlta, Minuta.Cliente, Minuta.Razon, Minuta.Fecha, Minuta.FechaFactura, Minuta.Factura, Minuta.Importe, Minuta.Observaciones, Minuta.ObservacionesII, Minuta.Direccion, Minuta.Tipo, Minuta.Localidad, Minuta.Moneda  " _
                + "From " _
                + DSQ + ".dbo.Minuta Minuta " _
                + "Where " _
                + "Minuta.Codigo >= " + ZCodigo + " AND " _
                + "Minuta.Codigo <= " + ZCodigo

    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    Listado.Connect = Connect()
    Listado.Action = 1
  
    DoEvents
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Cliente SET "
    ZSql = ZSql + "FechaMinuta =  " + "'" + Fecha.Text + "',"
    ZSql = ZSql + "Fecha =  " + "'" + "  /  /    " + "',"
    ZSql = ZSql + "OrdFecha =  " + "'" + "" + "',"
    ZSql = ZSql + "Anotacion =  " + "'" + "" + "',"
    ZSql = ZSql + "Hora =  " + "'" + "" + "'"
    ZSql = ZSql + " Where Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    
    Call cmdClose_Click

End Sub

Private Sub Form_Load()

    Call Limpia_Vector
 
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = "  /  /    "
    Observaciones.Text = ""
    ObservacionesII.Text = ""
    
    Pesos.Value = True
    Dolares.Value = False
    
    Cliente.Text = WMuestra
    DoEvents
    Call Cliente_KeyPress(13)
    
End Sub

Private Sub Proceso_Click()

    Cliente.Text = UCase(Cliente.Text)
    
    WSalida = "N"
    
    Call Limpia_Vector

    Renglon = 0
    WSaldo = 0
    
    XParam = "'" + Cliente.Text + "'"
    spCtacte = "ListaCtacteCliente " + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
    
        With rstCtacte
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    WPasa = "N"
                    
                    If !Tipo < 50 Then
                        WPasa = "S"
                    End If
                        
                    If WPasa = "S" Then
                    
                        If Pesos.Value = True Then
                            If !Total > 0 Then
                                Importe1 = !Total
                                Importe2 = 0
                                    Else
                                Importe1 = 0
                                Importe2 = !Total
                            End If
                            Importe3 = !Saldo
                                Else
                            If !Totalus > 0 Then
                                Importe1 = !Totalus
                                Importe2 = 0
                                    Else
                                Importe1 = 0
                                Importe2 = !Totalus
                            End If
                            Importe3 = !Saldous
                        End If
                        Call Redondeo(Importe3)
                    
                        If Importe3 <> 0 Then
                    
                            Renglon = Renglon + 1
                
                            Select Case !Tipo
                                Case 1
                                    WVector1.TextMatrix(Renglon, 1) = "Fac"
                                Case 2
                                    WVector1.TextMatrix(Renglon, 1) = "Dev"
                                Case 3
                                    WVector1.TextMatrix(Renglon, 1) = "Fac"
                                Case 4
                                    Select Case Left$(!Impre, 2)
                                        Case "DC"
                                            WVector1.TextMatrix(Renglon, 1) = "D.C"
                                        Case "CH"
                                            WVector1.TextMatrix(Renglon, 1) = "CHR"
                                        Case Else
                                            WVector1.TextMatrix(Renglon, 1) = "N/D"
                                    End Select
                                Case 5
                                    Select Case Left$(!Impre, 2)
                                        Case "DC"
                                            WVector1.TextMatrix(Renglon, 1) = "D.C"
                                        Case "CH"
                                            WVector1.TextMatrix(Renglon, 1) = "CHR"
                                        Case Else
                                            WVector1.TextMatrix(Renglon, 1) = "N/C"
                                    End Select
                                Case 6
                                    WVector1.TextMatrix(Renglon, 1) = "Rec"
                                Case 7
                                    WVector1.TextMatrix(Renglon, 1) = "Ant"
                                Case 10
                                    WVector1.TextMatrix(Renglon, 1) = "FCR"
                                Case 50
                                    WVector1.TextMatrix(Renglon, 1) = "Doc"
                                Case Else
                            End Select
                            
                            WVector1.TextMatrix(Renglon, 2) = Pusing("######", Str$(!Numero))
                            WVector1.TextMatrix(Renglon, 3) = !Fecha
                    
                            If Importe1 <> 0 Then
                                WVector1.TextMatrix(Renglon, 4) = Pusing("###,###,###.##", Str$(Importe1))
                                    Else
                                WVector1.TextMatrix(Renglon, 4) = ""
                            End If
                    
                            If Importe2 <> 0 Then
                                WVector1.TextMatrix(Renglon, 5) = Pusing("###,###,###.##", Str$(Importe2))
                                    Else
                                WVector1.TextMatrix(Renglon, 5) = ""
                            End If
                    
                            If Importe3 <> 0 Then
                                WVector1.TextMatrix(Renglon, 6) = Pusing("###,###,###.##", Str$(Importe3))
                                    Else
                                WVector1.TextMatrix(Renglon, 6) = ""
                            End If
                            
                            WSaldo = WSaldo + Importe3
                    
                            WVector1.TextMatrix(Renglon, 7) = !Vencimiento
                            WVector1.TextMatrix(Renglon, 8) = !Vencimiento1
                        
                        End If
                        
                    End If
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            End If
            
        End With
    
    End If
    
    Saldo.Caption = Pusing("###,###,###.##", Str$(WSaldo))
    
    WVector1.Col = 1
    WVector1.Row = 1
    WVector1.TopRow = 1

End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        WCliente = Cliente.Text
        Cliente.Text = WCliente
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente.Caption = rstCliente!Razon
            Observaciones.Text = rstCliente!Horario
            rstCliente.Close
            Call Proceso_Click
        End If
    End If
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 10
    WVector1.FixedRows = 1
    WVector1.Rows = 1001
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                WVector1.Text = "Debito"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WVector1.Text = "Credito"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector1.Text = "Saldo"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector1.Text = "Vencimiento"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                WVector1.Text = "Vencimiento"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 9
                WVector1.Text = "Acumulado"
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            ObservacionesII.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub


Private Sub ObservacionesII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
End Sub

Private Sub Pesos_Click()
    Call Proceso_Click
End Sub

Private Sub WVector1_DblClick()
    If WVector1.TextMatrix(WVector1.Row, 0) = "X" Then
        WVector1.TextMatrix(WVector1.Row, 0) = ""
            Else
        WVector1.TextMatrix(WVector1.Row, 0) = "X"
    End If
End Sub

Private Sub Marca_Click()
    rowini = WVector1.Row
    RowFin = WVector1.RowSel
    For Ciclo = rowini To RowFin
        If WVector1.TextMatrix(Ciclo, 0) = "X" Then
            WVector1.TextMatrix(Ciclo, 0) = ""
                Else
            WVector1.TextMatrix(Ciclo, 0) = "X"
        End If
    Next Ciclo
    ZSuma = 0
    For Ciclo = 1 To 1000
        If WVector1.TextMatrix(Ciclo, 0) = "X" Then
            ZSuma = ZSuma + Val(WVector1.TextMatrix(Ciclo, 6))
        End If
    Next Ciclo
    ACobrar.Caption = Str$(ZSuma)
    ACobrar.Caption = Pusing("###,###.##", ACobrar.Caption)
End Sub


