VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsultaHojaRuta 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Hojas de ruta"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   495
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin VB.TextBox NroViaje 
      Alignment       =   1  'Right Justify
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
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   24
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox RetiraProv 
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
      Left            =   4320
      MaxLength       =   50
      TabIndex        =   23
      Top             =   840
      Width           =   7335
   End
   Begin VB.TextBox TotalPesos 
      Alignment       =   1  'Right Justify
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
      Left            =   7440
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TotalKilos 
      Alignment       =   1  'Right Justify
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
      Left            =   10440
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox WTexto1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
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
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   375
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   2280
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
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
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   4455
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7858
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.TextBox Camion 
      Alignment       =   1  'Right Justify
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
      Left            =   7440
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   18
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Chofer 
      Alignment       =   1  'Right Justify
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
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   15
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Hoja 
      Alignment       =   1  'Right Justify
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
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Ayuda 
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
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   6855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10560
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   2280
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "consultahojaruta.frx":0000
      Left            =   120
      List            =   "consultahojaruta.frx":0007
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   4200
      TabIndex        =   12
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   327680
      Enabled         =   0   'False
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
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro de Viaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Retira Prov."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   25
      Top             =   840
      Width           =   1455
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   8160
      MouseIcon       =   "consultahojaruta.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "consultahojaruta.frx":031F
      ToolTipText     =   "Impresion "
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Pesos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   22
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Kilos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   19
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label DesCamion 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   17
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Camion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chofer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label DesChofer 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   480
      Width           =   2895
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
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   9720
      MouseIcon       =   "consultahojaruta.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "consultahojaruta.frx":0E6B
      ToolTipText     =   "Salida"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   9000
      MouseIcon       =   "consultahojaruta.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "consultahojaruta.frx":19B7
      ToolTipText     =   "Limpia la pantalla"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Nro. Hoja"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgConsultaHojaRuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstChofer As Recordset
Dim spChofer As String
Dim rstHojaRuta As Recordset
Dim rsHojaRuta As String
Dim rstCamion As Recordset
Dim spCamion As String

Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Cantidad As Double
Dim Renglon As Integer
Dim ZCodigo As String
Dim ZOperador As String
Dim WVersion As String
Dim WRenglon As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
Dim ZArti(100) As String
Dim ZBajaHoja(100) As String
Dim ZProceso As String
Dim Provincia(100) As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Hoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        Call Limpia_Vector
        WRenglon = 0
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM HojaRuta"
        ZSql = ZSql + " Where HojaRuta.Hoja = " + "'" + Hoja.Text + "'"
        ZSql = ZSql + " Order by HojaRuta.Clave"
    
        rsHojaRuta = ZSql
        Set rstHojaRuta = db.OpenRecordset(rsHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
        If rstHojaRuta.RecordCount > 0 Then
        
            rstHojaRuta.Close
            
            Call Proceso_Click
            
            WVector1.TopRow = 1
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
            
                Else
                
            Hoja.SetFocus
            
        End If
        
    End If
    If KeyAscii = 27 Then
        Hoja.Text = ""
    End If
End Sub

Private Sub cmdClose1_Click()

    Call Limpia_Click
    PrgConsultaHojaRuta.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector

    Hoja.Text = ""
    Fecha.Text = "  /  /    "
    Camion.Text = ""
    Chofer.Text = ""
    DesCamion.Caption = ""
    DesChofer.Caption = ""
    TotalKilos.Text = ""
    NroViaje.Text = ""
    RetiraProv.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Hoja) as [HojaMayor]"
    ZSql = ZSql + " FROM HojaRuta"
    spHojaRuta = ZSql
    Set rstHojaRuta = db.OpenRecordset(spHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
    If rstHojaRuta.RecordCount > 0 Then
        rstHojaRuta.MoveLast
        WHojaMayor = IIf(IsNull(rstHojaRuta!HojaMayor), "0", rstHojaRuta!HojaMayor)
        Hoja.Text = Str$(WHojaMayor + 1)
        rstHojaRuta.Close
            Else
        Hoja.Text = "0"
    End If
    
    Renglon = 0
    
    WVector1.Col = 1
    WVector1.Row = 1
    
    Hoja.SetFocus
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    Provincia(0) = "Capital"
    Provincia(1) = "Bs.As"
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
    
    WVector1.Col = 1
    WVector1.Row = 1

    Hoja.Text = ""
    Fecha.Text = "  /  /    "
    Camion.Text = ""
    Chofer.Text = ""
    DesCamion.Caption = ""
    DesChofer.Caption = ""
    TotalKilos.Text = ""
    TotalPesos.Text = ""
    NroViaje.Text = ""
    RetiraProv.Text = ""
    
    Renglon = 0
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    WRenglon = 0
    
        Sql1 = "Select *"
        Sql2 = " FROM HojaRuta"
        Sql3 = " Where HojaRuta.Hoja = " + "'" + Hoja.Text + "'"
        Sql4 = " Order by HojaRuta.Clave"
        
        rsHojaRuta = Sql1 + Sql2 + Sql3 + Sql4
        Set rstHojaRuta = db.OpenRecordset(rsHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
        If rstHojaRuta.RecordCount > 0 Then
            With rstHojaRuta
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Fecha.Text = rstHojaRuta!Fecha
                        Camion.Text = rstHojaRuta!Camion
                        Chofer.Text = rstHojaRuta!Chofer
                        NroViaje.Text = IIf(IsNull(rstHojaRuta!NroViaje), "", rstHojaRuta!NroViaje)
                        RetiraProv.Text = IIf(IsNull(rstHojaRuta!RetiraProv), "", rstHojaRuta!RetiraProv)
                    
                        WRenglon = WRenglon + 1
                        WVector1.Row = WRenglon
                        Renglon = WRenglon
                    
                        WVector1.Col = 1
                        WVector1.Text = Str$(rstHojaRuta!Pedido)
                
                        WVector1.Col = 2
                        WVector1.Text = Trim(rstHojaRuta!Cliente)
                        
                        WVector1.Col = 3
                        WVector1.Text = Trim(rstHojaRuta!Razon)
                
                        WVector1.Col = 4
                        WVector1.Text = Str$(rstHojaRuta!Remito)
                
                        WVector1.Col = 5
                        WVector1.Text = Str$(rstHojaRuta!Kilos)
                        WVector1.Text = Pusing("###,###", WVector1.Text)
                        
                        WVector1.Col = 6
                        WVector1.Text = ""
                
                        WVector1.Col = 7
                        WVector1.Text = rstHojaRuta!seguridad
                        
                        WVector1.Col = 8
                        WVector1.Text = ""
                        
                        WVector1.Col = 9
                        WVector1.Text = ""
                
                        WVector1.Col = 10
                        WVector1.Text = rstHojaRuta!Clave
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstHojaRuta.Close
        End If
    
    For Ciclo = 1 To WRenglon
    
        ZZPedido = WVector1.TextMatrix(Ciclo, 1)
        ZZCliente = WVector1.TextMatrix(Ciclo, 2)
        ZZRemito = WVector1.TextMatrix(Ciclo, 4)
        
        spCliente = "ConsultaCliente " + "'" + ZZCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
        
            WVector1.TextMatrix(Ciclo, 8) = Provincia(rstCliente!Provincia)
            
            Erase ZDirEntrega
                
            WDirentrega = ""
            ZDirEntrega(1) = rstCliente!DirEntrega
            ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
            ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
            ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
            ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                
            rstCliente.Close
                
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Pedido = " + "'" + ZZPedido + "'"
        ZSql = ZSql + " Order by Pedido.Remito desc"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            ZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
            WDirentrega = ZDirEntrega(ZLugarDirEntrega)
            WVector1.TextMatrix(Ciclo, 9) = WDirentrega
            ZZRemito = IIf(IsNull(rstPedido!Remito), "", rstPedido!Remito)
            WVector1.TextMatrix(Ciclo, 4) = ZZRemito
            rstPedido.Close
        End If
        
        ZZNumero = 0
        ZZTipo = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CtaCte"
        ZSql = ZSql + " Where CtaCte.Pedido = " + "'" + Trim(ZZPedido) + "'"
        spCtacte = ZSql
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            With rstCtacte
                .MoveFirst
                Do
                    If .EOF = False Then
                                
                        If Val(ZZRemito) = Val(rstCtacte!Remito) Then
                            WVector1.TextMatrix(Ciclo, 6) = Str$(rstCtacte!Total)
                            WVector1.TextMatrix(Ciclo, 6) = Pusing("###,###.##", WVector1.TextMatrix(Ciclo, 6))
                            ZZNumero = rstCtacte!Numero
                            ZZTipo = rstCtacte!Tipo
                        End If
                                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCtacte.Close
        End If
        
        If Val(ZZNumero) <> 0 Then
            ZZSuma = 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Estadistica"
            ZSql = ZSql + " Where Estadistica.Tipo = " + "'" + ZZTipo + "'"
            ZSql = ZSql + " and Estadistica.Numero = " + "'" + Str$(ZZNumero) + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
                With rstEstadistica
                    .MoveFirst
                    Do
                        If .EOF = False Then
                                                
                            ZZSuma = ZZSuma + rstEstadistica!Cantidad
                                    
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEstadistica.Close
            End If
            If ZZSuma <> 0 Then
                WVector1.TextMatrix(Ciclo, 5) = Str$(ZZSuma)
                WVector1.TextMatrix(Ciclo, 5) = Pusing("###,###", WVector1.TextMatrix(Ciclo, 5))
            End If
        End If
    
    Next Ciclo
    
    Sql1 = "Select *"
    Sql2 = " FROM Camion"
    Sql3 = " Where Camion.Codigo = " + "'" + Camion.Text + "'"
    spCamion = Sql1 + Sql2 + Sql3
    Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
    If rstCamion.RecordCount > 0 Then
        DesCamion.Caption = Trim(rstCamion!Descripcion)
        rstCamion.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM Chofer"
    Sql3 = " Where Chofer.Codigo = " + "'" + Chofer.Text + "'"
    spChofer = Sql1 + Sql2 + Sql3
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    If rstChofer.RecordCount > 0 Then
        DesChofer.Caption = Trim(rstChofer!Descripcion)
        rstChofer.Close
    End If
    
    Call Calcula_Click

End Sub

Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub Lista_Click()

    For Ciclo = 1 To 100
    
        If Val(WVector1.TextMatrix(Ciclo, 1)) <> 0 Then
        
            ZZPedido = WVector1.TextMatrix(Ciclo, 1)
            ZZCliente = WVector1.TextMatrix(Ciclo, 2)
            ZZRazon = WVector1.TextMatrix(Ciclo, 3)
            ZZRemito = WVector1.TextMatrix(Ciclo, 4)
            ZZKilos = WVector1.TextMatrix(Ciclo, 5)
            ZZPesos = WVector1.TextMatrix(Ciclo, 6)
            ZZSeguridad = WVector1.TextMatrix(Ciclo, 7)
            ZZprovincia = WVector1.TextMatrix(Ciclo, 8)
            ZZDireccion = WVector1.TextMatrix(Ciclo, 9)
            ZZClave = WVector1.TextMatrix(Ciclo, 10)
            
            spCliente = "ConsultaCliente " + "'" + ZZCliente + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                ZZCuit = rstCliente!Cuit
                ZZPostal = rstCliente!Postal
                rstCliente.Close
            End If
            
            ZZFactura = ""
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCte"
            ZSql = ZSql + " Where CtaCte.Pedido = " + "'" + Trim(ZZPedido) + "'"
            ZSql = ZSql + " Order by CtaCte.ordfecha desc"
            spCtacte = ZSql
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtacte.RecordCount > 0 Then
                With rstCtacte
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Val(ZZRemito) = Val(rstCtacte!Remito) Then
                                ZZFactura = Str$(rstCtacte!Numero)
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCtacte.Close
            End If
            
            ZZCodigoArticulo = ""
            ZZPunto = ""
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pedido"
            ZSql = ZSql + " Where Pedido.Pedido = " + "'" + ZZPedido + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                Select Case rstPedido!TipoPedido
                    Case 1
                        WTipoPedido = "CO"
                        ZZCodigoArticulo = "320411"
                        ZZPunto = "1"
                    Case 3
                        WTipoPedido = "BI"
                        ZZCodigoArticulo = "340391"
                        ZZPunto = "4"
                    Case 4
                        WTipoPedido = "FA"
                        ZZCodigoArticulo = "291815"
                        ZZPunto = "1"
                    Case 5
                        WTipoPedido = "PG"
                        ZZCodigoArticulo = "340391"
                        ZZPunto = "1"
                    Case Else
                        WTipoPedido = "PT"
                        ZZCodigoArticulo = "340391"
                        ZZPunto = "4"
                End Select
                rstPedido.Close
            End If
        
            ZSql = ""
            ZSql = ZSql + "UPDATE HojaRuta SET "
            ZSql = ZSql + "Remito = " + "'" + ZZRemito + "',"
            ZSql = ZSql + "Kilos = " + "'" + ZZKilos + "',"
            ZSql = ZSql + "Pesos = " + "'" + ZZPesos + "',"
            ZSql = ZSql + "Razon = " + "'" + ZZRazon + "',"
            ZSql = ZSql + "Cuit = " + "'" + ZZCuit + "',"
            ZSql = ZSql + "Postal = " + "'" + ZZPostal + "',"
            ZSql = ZSql + "Direccion = " + "'" + ZZDireccion + "',"
            ZSql = ZSql + "Factura = " + "'" + ZZFactura + "',"
            ZSql = ZSql + "Punto = " + "'" + ZZPunto + "',"
            ZSql = ZSql + "Provincia = " + "'" + ZZprovincia + "',"
            ZSql = ZSql + "CodigoArticulo = " + "'" + ZZCodigoArticulo + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZZClave + "'"
                     
            spHojaRuta = ZSql
            Set rstHojaRuta = db.OpenRecordset(spHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    
    Next Ciclo



    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT HojaRuta.Hoja, HojaRuta.Renglon, HojaRuta.Fecha, HojaRuta.Chofer, HojaRuta.Camion, HojaRuta.Pedido, HojaRuta.Cliente, HojaRuta.Remito, HojaRuta.Seguridad, HojaRuta.Kilos, HojaRut.Pesos, HojaRuta.Razon, HojaRuta.Cuit, HojaRut.Postal, HojaRuta.Direccion, HojaRuta.Factura, HojaRuta.Punto, HojaRuta.CodigoArticulo, HojaRuta.Provincia, HojaRuta.NroViaje, HojaRuta.RetiraProv, " _
            + "Chofer.Descripcion, " _
            + "Camion.Descripcion, Camion.Patente " _
            + "From " _
            + DSQ + ".dbo.HojaRuta HojaRuta, " _
            + DSQ + ".dbo.Chofer Chofer, " _
            + DSQ + ".dbo.Camion Camion " _
            + "Where " _
            + "HojaRuta.Chofer = Chofer.Codigo AND " _
            + "HojaRuta.Camion = Camion.Codigo AND " _
            + "HojaRuta.Hoja >= " + Hoja.Text + " AND " _
            + "HojaRuta.Hoja <= " + Hoja.Text
            
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{HojaRuta.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.SelectionFormula = "{HojaRuta.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    Listado.ReportFileName = "HojaRutaCot.rpt"
    
    Listado.Action = 1

End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1
        Case 113
            WTexto1.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 123
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Col > 1 Then
                WVector1.Col = WVector1.Col - 1
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1
        Case 113
            WTexto2.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto3.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 10
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    Call Calcula_Click
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 11
    WVector1.FixedRows = 1
    WVector1.Rows = 101
    
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
    
    WVector1.ColWidth(0) = 100
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Pedido"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Cliente"
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Razon"
                WVector1.ColWidth(Ciclo) = 2500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Remito"
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Kilos"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "  $  "
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "Clase"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = "Provincia"
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = "Direccion"
                WVector1.ColWidth(Ciclo) = 7000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 10
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 20
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector1.Text
        Rem WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        Rem WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        Rem WTitulo(Ciclo).Width = WVector1.CellWidth
        Rem WTitulo(Ciclo).Height = WVector1.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub


Private Sub Calcula_Click()

    ZTotalKilos = 0
    ZTotalPesos = 0

    For a = 1 To 100
        ZTotalKilos = ZTotalKilos + Val(WVector1.TextMatrix(a, 5))
        ZTotalPesos = ZTotalPesos + Val(WVector1.TextMatrix(a, 6))
    Next a
    
    TotalKilos.Text = Str$(ZTotalKilos)
    TotalPesos.Text = Str$(ZTotalPesos)
    
    TotalKilos.Text = Pusing("###,###", TotalKilos.Text)
    TotalPesos.Text = Pusing("###,###.##", TotalPesos.Text)
    
End Sub

Private Sub Impresion()

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT HojaRuta.Hoja, HojaRuta.Renglon, HojaRuta.Fecha, HojaRuta.Chofer, HojaRuta.Camion, HojaRuta.Pedido, HojaRuta.Cliente, HojaRuta.Remito, HojaRuta.Seguridad, HojaRuta.Kilos, HojaRuta.ObservaI, HojaRuta.ObservaII, HojaRuta.Bultos, HojaRuta.Razon, " _
            + "Chofer.Descripcion, " _
            + "Camion.Descripcion, Camion.Patente " _
            + "From " _
            + DSQ + ".dbo.HojaRuta HojaRuta, " _
            + DSQ + ".dbo.Chofer Chofer, " _
            + DSQ + ".dbo.Camion Camion " _
            + "Where " _
            + "HojaRuta.Chofer = Chofer.Codigo AND " _
            + "HojaRuta.Camion = Camion.Codigo AND " _
            + "HojaRuta.Hoja >= " + Hoja.Text + " AND " _
            + "HojaRuta.Hoja <= " + Hoja.Text
            
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{HojaRuta.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.SelectionFormula = "{HojaRuta.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    If Val(WEmpresa) = 1 Then
        Listado.ReportFileName = "HojaRuta.rpt"
            Else
        Listado.ReportFileName = "HojaRutapelli.rpt"
    End If
    
    Listado.Action = 1
    
End Sub


