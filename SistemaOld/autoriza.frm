VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAutoriza 
   AutoRedraw      =   -1  'True
   Caption         =   "Autorizacion de Pedidos a Facturar"
   ClientHeight    =   7320
   ClientLeft      =   150
   ClientTop       =   690
   ClientWidth     =   11550
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11550
   Begin VB.Frame PantaCambiaFecha 
      Height          =   3615
      Left            =   1440
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   8655
      Begin VB.ComboBox Tipoped 
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
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox Problema 
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
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   28
         Text            =   " "
         Top             =   2040
         Width           =   5655
      End
      Begin VB.ComboBox Concepto 
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
         Left            =   2640
         TabIndex        =   27
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CommandButton ConfirmaVto 
         Caption         =   "CONFIRMA"
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
         Left            =   3360
         TabIndex        =   24
         Top             =   2760
         Width           =   1575
      End
      Begin MSMask.MaskEdBox VtoI 
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
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
      Begin MSMask.MaskEdBox VtoII 
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
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
      Begin VB.Label Label14 
         Caption         =   "Tipo Pedido"
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
         Left            =   4200
         TabIndex        =   30
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   480
         TabIndex        =   26
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Concepto"
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
         Left            =   480
         TabIndex        =   25
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nueva Fecha Vto."
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
         Left            =   480
         TabIndex        =   23
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Vto. Original"
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
         Left            =   480
         TabIndex        =   22
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame PantaBloqueo 
      Caption         =   "Bloqueo de Partidas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   2520
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   4695
      Begin VB.OptionButton Opcion3 
         Caption         =   "Bloquea Partida Entregada y Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   720
         TabIndex        =   18
         Top             =   1320
         Width           =   3735
      End
      Begin VB.OptionButton Opcion2 
         Caption         =   "Bloquea Solo Partida Entregada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   840
         Width           =   3255
      End
      Begin VB.OptionButton Opcion1 
         Caption         =   "No Bloquea Partida ni Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton ConfirmaBloqueo 
         Caption         =   "CONFIRMA"
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
         Left            =   1560
         TabIndex        =   15
         Top             =   2040
         Width           =   1575
      End
   End
   Begin VB.CommandButton Anula 
      Caption         =   "Anula Pedido"
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
      Left            =   8040
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Clave1 
      Caption         =   "  Ingreso de Clave de Seguridad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2640
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Cancelagraba 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Ingrese su Password"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton Graba 
      Caption         =   "&Graba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4560
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Autorizo 
      Caption         =   "Autorizo"
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
      Left            =   5640
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin MSMask.MaskEdBox HastaFecha 
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
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
   Begin MSMask.MaskEdBox DesdeFecha 
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
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
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   6375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11245
      _Version        =   327680
      Rows            =   4000
      Cols            =   12
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Lee datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta Fecha"
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
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde Fecha"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgAutoriza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstMuestra As Recordset
Dim spMuestra As String
Dim XParam As String
Dim TotalPedidos As Integer
Dim WGraba As String
Dim ZVector(100, 4) As String
Dim CargaPedido(100, 10) As String
Dim ZDirEntrega(10) As String
Private CargaEmpresa(10, 2) As String

Dim ZCodigo As String
Dim ZTerminado As String
Dim ZArticulo As String
Dim ZEnsayo As String
Dim ZNombre As String
Dim ZFecha As String
Dim ZFechaOrd As String
Dim ZCantidad As String
Dim ZCliente As String
Dim ZRazon As String
Dim ZDescriCliente As String
Dim ZVendedor As String
Dim ZDesVendedor As String
Dim ZObservaciones As String
Dim ZAutoriza As String
Dim ZImpresion As String
Dim ZPedido As String
Dim ZLugarDirEntrega As String
Dim ZDescriDirEntrega As String

Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer
Dim DiaFeriado(100) As String
Dim ZZFecEntrega As String

Private Sub Autorizo_Click()

    If Muestra.TextMatrix(Muestra.Row, 6) = "DEVOL" Then
        Opcion1.Value = True
        Opcion2.Value = False
        Opcion3.Value = False
        PantaBloqueo.Visible = True
        Exit Sub
    End If
    
    
    Muestra.Col = 1
    WPedido = Muestra.Text
    
    WTipoPedido = 0
    WTipoPed = 0
    spPedido = "ListaPedido " + "'" + WPedido + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        WTipoPedido = rstPedido!TipoPedido
        WTipoPed = rstPedido!Tipoped
        WFecEntrega = rstPedido!FecEntrega
        WFechaPedido = rstPedido!Fecha
        rstPedido.Close
    End If
    
    ZParcial = "N"
    spPedido = "ConsultaPedido1 " + "'" + WPedido + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                    If rstPedido!Facturado <> 0 Then
                        ZParcial = "S"
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    WFechaAutoriza = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    If WFechaPedido <> WFechaAutoriza And ZParcial = "N" Then
        VtoI.Text = WFecEntrega
        VtoII.Text = WFecEntrega
        Tipoped.ListIndex = WTipoPed
        If WTipoPed = 0 Then
            Call Calcula_FecEntrega
            Call Calcula_Feriado
            VtoII.Text = ZZFecEntrega
                Else
            ZZDia = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZZOrdDia = Right$(ZZDia, 4) + Mid$(ZZDia, 4, 2) + Left$(ZZDia, 2)
            ZZOrdEntrega = Right$(VtoII.Text, 4) + Mid$(VtoII.Text, 4, 2) + Left$(VtoII.Text, 2)
            If ZZOrdEntrega < ZZOrdDia Then
                VtoII.Text = ZZDia
            End If
        End If
        Concepto.ListIndex = 13
        Problema.Text = ""
        PantaCambiaFecha.Visible = True
        VtoII.SetFocus
    End If
    
    Muestra.Col = 8
    Muestra.Text = "Autorizado"
    Muestra.Col = 1
  
End Sub

Private Sub ConfirmaVto_Click()

    Call Valida_fecha(VtoII.Text, Auxi)
    If Auxi <> "S" Then
        Exit Sub
    End If


    If Concepto.ListIndex <> 0 And Concepto.ListIndex <> 14 Then
    
        WAtraso = "1"
        
        Sql1 = "Select Max(Numero) as [NumeroMayor]"
        Sql2 = " FROM Atraso"
        spAtraso = Sql1 + Sql2
        Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
        If rstAtraso.RecordCount > 0 Then
            WAtraso = Str$(rstAtraso!Numeromayor + 1)
            rstAtraso.Close
        End If
    
        WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        WFechaEntregaord = Right$(VtoII.Text, 4) + Mid$(VtoII.Text, 4, 2) + Left$(VtoII.Text, 2)
        
        ZZVersionPedido = ""
        spPedido = "ListaPedido " + "'" + Muestra.TextMatrix(Muestra.Row, 1) + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            ZZVersionPedido = Str$(rstPedido!Version)
            rstPedido.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Atraso ("
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "OrdFecha ,"
        ZSql = ZSql + "Pedido ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Problema ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "FechaEntrega ,"
        ZSql = ZSql + "OrdFechaEntrega ,"
        ZSql = ZSql + "DesCliente ,"
        ZSql = ZSql + "DesTerminado ,"
        ZSql = ZSql + "DesArticulo ,"
        ZSql = ZSql + "Concepto ,"
        ZSql = ZSql + "Solicitud ,"
        ZSql = ZSql + "Origen ,"
        ZSql = ZSql + "VersionPedido)"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WAtraso + "',"
        ZSql = ZSql + "'" + WFecha + "',"
        ZSql = ZSql + "'" + WOrdFecha + "',"
        ZSql = ZSql + "'" + Muestra.TextMatrix(Muestra.Row, 1) + "',"
        ZSql = ZSql + "'" + Muestra.TextMatrix(Muestra.Row, 3) + "',"
        ZSql = ZSql + "'" + "  -     -   " + "',"
        ZSql = ZSql + "'" + Problema.Text + "',"
        ZSql = ZSql + "'" + "  -   -   " + "',"
        ZSql = ZSql + "'" + VtoII.Text + "',"
        ZSql = ZSql + "'" + WOrdFechaEntrega + "',"
        ZSql = ZSql + "'" + Muestra.TextMatrix(Muestra.Row, 4) + "',"
        ZSql = ZSql + "'" + "" + " ',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + Str$(Concepto.ListIndex) + "',"
        ZSql = ZSql + "'" + "0" + "',"
        ZSql = ZSql + "'" + "1" + "',"
        ZSql = ZSql + "'" + ZZVersionPedido + "')"
        
        spAtraso = ZSql
        Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
        
    End If

    PantaCambiaFecha.Visible = False
    Muestra.Col = 11
    Muestra.Text = VtoII.Text
    Muestra.Col = 8
    Muestra.Text = "Autorizado"
    Muestra.Col = 1
    
End Sub

Private Sub ConfirmaBloqueo_Click()

    If Opcion1.Value = True Then
        Muestra.Col = 10
        Muestra.Text = "N"
    End If
    
    If Opcion2.Value = True Then
        Muestra.Col = 10
        Muestra.Text = "P"
    End If
    
    If Opcion3.Value = True Then
        Muestra.Col = 10
        Muestra.Text = "S"
    End If
    
    Muestra.Col = 8
    Muestra.Text = "Autorizado"
    Muestra.Col = 1
    
    PantaBloqueo.Visible = False

End Sub


Private Sub cmdClose_Click()
    With rstEmpresa
        .Close
    End With
    PrgAutoriza.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Anula_Click()

    Muestra.Col = 8
    Muestra.Text = "Anulado"
    Muestra.Col = 1

End Sub


Private Sub Form_Load()

    Concepto.Clear
    
    Concepto.AddItem ""
    Concepto.AddItem "Falta M.P.Local"
    Concepto.AddItem "Falta M.P. Importada"
    Concepto.AddItem "Cambio de Prioridades"
    Concepto.AddItem "Falta de Capacidad Disponible"
    Concepto.AddItem "Error del Sistema"
    Concepto.AddItem "Varios"
    Concepto.AddItem "Problemas Vehiculos"
    Concepto.AddItem "Problemas Logistica"
    Concepto.AddItem "Problemas Recepcion Cliente"
    Concepto.AddItem "Varios"
    Concepto.AddItem "Corte de Luz"
    Concepto.AddItem "Pedido por el Cliente"
    Concepto.AddItem "Falta de Pago"
    Concepto.AddItem "Confirmacion Pedido Parcial"
    Concepto.AddItem "Envase"
    
    Concepto.ListIndex = 0
    
    Tipoped.Clear
    
    Tipoped.AddItem "Normal"
    Tipoped.AddItem "a Fecha"
    Tipoped.AddItem "Fecha Limite"
    Tipoped.AddItem "Urgente"
    Tipoped.AddItem "Retira Cliente"
    Tipoped.AddItem "Muestra"
    Tipoped.AddItem "Muestra Retira"
    
    Tipoped.ListIndex = 0
    

    Call Limpia_Vector
    
    Muestra.Font.Bold = True
    
    Muestra.ColWidth(0) = 50
    Muestra.ColWidth(1) = 1000
    Muestra.ColWidth(2) = 1200
    Muestra.ColWidth(3) = 1000
    Muestra.ColWidth(4) = 2500
    Muestra.ColWidth(5) = 1200
    Muestra.ColWidth(6) = 1000
    Muestra.ColWidth(7) = 1000
    Muestra.ColWidth(8) = 1000
    Muestra.ColWidth(9) = 900
    Muestra.ColWidth(10) = 100
    Muestra.ColWidth(11) = 100
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
    Muestra.Col = 7
    Muestra.Text = "Importe"
    
    Muestra.Col = 8
    Muestra.Text = "Estado"
    
    Muestra.Col = 9
    Muestra.Text = "Impresa"
    
    Muestra.Col = 10
    Muestra.Text = ""
    
    Muestra.Col = 11
    Muestra.Text = ""
    
    Rem DesdeFecha.Text = "  /  /    "
    Rem HastaFecha.Text = "  /  /    "
    DesdeFecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    HastaFecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Rem DesdeFecha.SetFocus
    
End Sub

Private Sub Graba_Click()

    If WGraba <> "S" Then
        Call Ingresa_clave
            Else
            
        WGraba = ""
        WFechaAutoriza = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)

        For Ciclo = 1 To TotalPedidos
    
            Muestra.Row = Ciclo
            Muestra.Col = 8
            If Muestra.Text = "Autorizado" Then
        
                Muestra.Col = 6
                If Muestra.Text = "DEVOL" Then
        
                    Muestra.Col = 3
                    ZCliente = Muestra.Text
                    
                    Muestra.Col = 10
                    WBloqueo = Muestra.Text
                    
                    Muestra.Col = 1
                    WPedido = Muestra.Text
                    WMarca = "X"
                    XFecha = DesdeFecha.Text
                    XFechaOrd = Right$(DesdeFecha.Text, 4) + Mid$(DesdeFecha.Text, 4, 2) + Left$(DesdeFecha.Text, 2)
            
                    XParam = "'" + WPedido + "','" _
                             + WMarca + "','" _
                             + XFecha + "','" _
                             + XFechaOrd + "'"
                                           
                    spPedidoDevol = "ModificaPedidoDevolAutoriza " + XParam
                    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
                    
                    WImpreLaboI = ""
                    WImpreLaboII = ""
                    WImpreProdI = "N"
                    WImpreProdII = "N"
                    WImpreProdIII = "N"
                    WImpreProdIV = "N"
                    WImpresionII = "N"
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE PedidoDevol SET "
                    ZSql = ZSql + " ImpreLaboI =  " + "'" + WImpreLaboI + "',"
                    ZSql = ZSql + " ImpreLaboII =  " + "'" + WImpreLaboII + "',"
                    ZSql = ZSql + " ImpreProdI =  " + "'" + WImpreProdI + "',"
                    ZSql = ZSql + " ImpreProdII =  " + "'" + WImpreProdII + "',"
                    ZSql = ZSql + " ImpreProdIII =  " + "'" + WImpreProdIII + "',"
                    ZSql = ZSql + " ImpreProdIV =  " + "'" + WImpreProdIV + "',"
                    ZSql = ZSql + " ImpresionII =  " + "'" + WImpresionII + "',"
                    ZSql = ZSql + " Bloqueo =  " + "'" + WBloqueo + "'"
                    ZSql = ZSql + " Where Pedido = " + "'" + WPedido + "'"
                    spPedidoDevol = ZSql
                    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If WBloqueo = "S" Then
                    
                        Rem bloquea los articulos
                        
                        Erase ZVector
                        ZLugar = 0
                        
                        spPedidoDevol = "ListaPedidoDevol " + "'" + WPedido + "'"
                        Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
                        If rstPedidoDevol.RecordCount > 0 Then
                            With rstPedidoDevol
                                .MoveFirst
                                Do
                                    If .EOF = False Then
                
                                        ZLugar = ZLugar + 1
                
                                        ZVector(ZLugar, 1) = rstPedidoDevol!Terminado
                                        ZVector(ZLugar, 2) = Str$(rstPedidoDevol!Cantidad)
                                        ZVector(ZLugar, 3) = rstPedidoDevol!Partida
                        
                                        .MoveNext
                                            Else
                                        Exit Do
                                    End If
                                Loop
                            End With
                            rstPedidoDevol.Close
                        End If
    
                        For Cicla = 1 To ZLugar
                        
                            Terminado = ZVector(Cicla, 1)
                            Cantidad = ZVector(Cicla, 2)
                            Lote = ZVector(Cicla, 3)
                            
                            XEmpresa = Wempresa
                            
                            Select Case Val(XEmpresa)
                                Case 1, 3, 5, 6, 7, 10, 11
                                    CargaEmpresa(1, 1) = "0001"
                                    CargaEmpresa(1, 2) = "Empresa01"
                                    CargaEmpresa(2, 1) = "0003"
                                    CargaEmpresa(2, 2) = "Empresa03"
                                    CargaEmpresa(3, 1) = "0005"
                                    CargaEmpresa(3, 2) = "Empresa05"
                                    CargaEmpresa(4, 1) = "0006"
                                    CargaEmpresa(4, 2) = "Empresa06"
                                    CargaEmpresa(5, 1) = "0007"
                                    CargaEmpresa(5, 2) = "Empresa07"
                                    CargaEmpresa(6, 1) = "0010"
                                    CargaEmpresa(6, 2) = "Empresa10"
                                    CargaEmpresa(7, 1) = "0011"
                                    CargaEmpresa(7, 2) = "Empresa11"
                                    ZHasta = 7
                                Case Else
                                    CargaEmpresa(1, 1) = "0002"
                                    CargaEmpresa(1, 2) = "Empresa02"
                                    CargaEmpresa(2, 1) = "0004"
                                    CargaEmpresa(2, 2) = "Empresa04"
                                    CargaEmpresa(3, 1) = "0008"
                                    CargaEmpresa(3, 2) = "Empresa08"
                                    CargaEmpresa(4, 1) = "0009"
                                    CargaEmpresa(4, 2) = "Empresa09"
                                    ZHasta = 4
                            End Select
            
                            For ZCiclo = 1 To ZHasta
            
                                Wempresa = CargaEmpresa(ZCiclo, 1)
                                txtOdbc = CargaEmpresa(ZCiclo, 2)
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                                If Left$(Terminado, 2) <> "PT" And Left$(Terminado, 2) <> "PE" Then
        
                                    ZEntra = "N"
                                    
                                    If Left$(Terminado, 2) = "DY" Or Left$(Terminado, 2) = "DK" Then
                                        ZArti = "DY-" + Right$(Terminado, 7)
                                            Else
                                        If Left$(Terminado, 2) = "DS" Or Left$(Terminado, 2) = "NS" Then
                                            ZArti = "DS-" + Right$(Terminado, 7)
                                                Else
                                            If Left$(Terminado, 2) = "DQ" Or Left$(Terminado, 2) = "NQ" Then
                                                ZArti = "DQ-" + Right$(Terminado, 7)
                                                    Else
                                                ZArti = Left$(Terminado, 2) + "-" + Right$(Terminado, 7)
                                            End If
                                        End If
                                    End If
                                    
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM Laudo"
                                    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + ZArti + "'"
                                    ZSql = ZSql + " and Laudo.Lote = " + "'" + Lote + "'"
                                    ZSql = ZSql + " Order by Laudo.Laudo"
                                    spLaudo = ZSql
                                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstLaudo.RecordCount > 0 Then
                                
                                        rstLaudo.Close
                                    
                                        ZEntra = "S"
                                        ZMarcaEstado = "N"
                                    
                                        ZSql = ""
                                        ZSql = ZSql + "UPDATE Laudo SET "
                                        ZSql = ZSql + "Estado  = " + "'" + ZMarcaEstado + "'"
                                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + ZArti + "'"
                                        ZSql = ZSql + " and Laudo.Lote = " + "'" + Lote + "'"
                                        spLaudo = ZSql
                                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                    
                                    End If
                        
                                    If ZEntra = "N" Then
                                
                                        ZSql = ""
                                        ZSql = ZSql + "Select *"
                                        ZSql = ZSql + " FROM Guia"
                                        ZSql = ZSql + " Where Guia.Articulo = " + "'" + ZArti + "'"
                                        ZSql = ZSql + " and Guia.Lote = " + "'" + Lote + "'"
                                        spMovguia = ZSql
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstMovguia.RecordCount > 0 Then
                                            rstMovguia.Close
                                        
                                            ZMarcaEstado = "N"
                                        
                                            ZSql = ""
                                            ZSql = ZSql + "UPDATE Guia SET "
                                            ZSql = ZSql + "Estado  = " + "'" + ZMarcaEstado + "'"
                                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + ZArti + "'"
                                            ZSql = ZSql + " and Guia.Lote = " + "'" + Lote + "'"
                                            spMovguia = ZSql
                                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                        
                                        End If
                                    End If
            
                                        Else
                
                                    ZEntra = "N"
            
                                    XParam = "'" + Lote + "','" _
                                                 + Terminado + "'"
                                    spHoja = "ListaHojaProducto " + XParam
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstHoja.RecordCount > 0 Then
                                
                                        rstHoja.Close
                                    
                                        ZEntra = "S"
                                        ZMarcaEstado = "N"
                                    
                                        ZSql = ""
                                        ZSql = ZSql + "UPDATE Hoja SET "
                                        ZSql = ZSql + "Estado  = " + "'" + ZMarcaEstado + "'"
                                        ZSql = ZSql + " Where Hoja.Producto = " + "'" + Terminado + "'"
                                        ZSql = ZSql + " and Hoja.Hoja = " + "'" + Lote + "'"
                                        spHoja = ZSql
                                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                        
                                    End If
                
                                    If ZEntra = "N" Then
                                
                                        XParam = "'" + Terminado + "','" _
                                                     + Lote + "'"
                                                
                                        spMovguia = "ListaMovguiaLote1 " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstMovguia.RecordCount > 0 Then
                                    
                                            rstMovguia.Close
                                        
                                            ZMarcaEstado = "N"
                                        
                                            ZSql = ""
                                            ZSql = ZSql + "UPDATE Guia SET "
                                            ZSql = ZSql + "Estado  = " + "'" + ZMarcaEstado + "'"
                                            ZSql = ZSql + " Where Guia.Terminado = " + "'" + Terminado + "'"
                                            ZSql = ZSql + " and Guia.Lote = " + "'" + Lote + "'"
                                            spMovguia = ZSql
                                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                        
                                        End If
                                    End If
                
                                End If
                                
                            Next ZCiclo
                            
                            Call Conecta_Empresa
        
                        Next Cicla
                        
                            Else
                            
                        Rem graba la liberafcion automatica
                            
                        Rem Erase ZVector
                        Rem ZLugar = 0
                        
                        Rem spPedidoDevol = "ListaPedidoDevol " + "'" + WPedido + "'"
                        Rem Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
                        Rem If rstPedidoDevol.RecordCount > 0 Then
                        Rem     With rstPedidoDevol
                        Rem         .MoveFirst
                        Rem         Do
                        Rem             If .EOF = False Then
                        Rem
                        Rem                 ZLugar = ZLugar + 1
                        Rem
                        Rem                 ZVector(ZLugar, 1) = rstPedidoDevol!Terminado
                        Rem                 ZVector(ZLugar, 2) = Str$(rstPedidoDevol!Cantidad)
                        Rem                 ZVector(ZLugar, 3) = rstPedidoDevol!Partida
                        Rem
                        Rem                 .MoveNext
                        Rem                     Else
                        Rem                 Exit Do
                        Rem             End If
                        Rem         Loop
                        Rem     End With
                        Rem     rstPedidoDevol.Close
                        Rem End If
                        Rem
                        Rem For Cicla = 1 To ZLugar
                        Rem
                        Rem     ZTerminado = ZVector(Cicla, 1)
                        Rem     ZCantidad = ZVector(Cicla, 2)
                        Rem     ZLote = ZVector(Cicla, 3)
                        Rem     ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                        Rem
                        Rem     ZSql = ""
                        Rem     ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
                        Rem     ZSql = ZSql + " FROM LiberaTerminado"
                        Rem     spLiberaTerminado = ZSql
                        Rem     Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        Rem     If rstLiberaTerminado.RecordCount > 0 Then
                        Rem         rstLiberaTerminado.MoveLast
                        Rem         WCodigoMayor = IIf(IsNull(rstLiberaTerminado!CodigoMayor), "0", rstLiberaTerminado!CodigoMayor)
                        Rem         Lote = Str$(WCodigoMayor)
                        Rem         rstLiberaTerminado.Close
                        Rem             Else
                        Rem         Lote = "0"
                        Rem     End If
                        Rem
                        Rem     WCodigo = Str$(Val(Lote) + 1)
                        Rem     WProducto = ZTerminado
                        Rem     WFecha = ZFecha
                        Rem     WFechaOrd = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
                        Rem     WPartida = ZLote
                        Rem     WPartiOri = ""
                        Rem     WValor1 = ""
                        Rem     WValor2 = ""
                        Rem     WValor3 = ""
                        Rem     WValor4 = ""
                        Rem     WValor5 = ""
                        Rem     WValor6 = ""
                        Rem     WValor7 = ""
                        Rem     WValor8 = ""
                        Rem     WValor9 = ""
                        Rem     WValor10 = ""
                        Rem     WEnsayo = ""
                        Rem     WAspecto = ""
                        Rem     WObservaciones = "Liberacion Automatica"
                        Rem     WConfecciono = ""
                        Rem     WMarca = "N"
                        Rem     WCliente = ZCliente
                        Rem     WObserva = ""
                        Rem     WCantidad = ZCantidad
                        Rem     WOrigen = "L"
                        Rem     WTipo = "PT"
                        Rem     WImpreProdI = "N"
                        Rem     WImpreProdII = "N"
                        Rem     WImpreProdIII = "N"
                        Rem     WImpreVentas = "N"
                        Rem     WTipopro = ""
                        Rem
                        Rem     XTipoPro = ""
                        Rem     XCodigo = Val(Mid$(ZTerminado, 4, 5))
                        Rem     If Left$(ZTerminado, 2) = "DY" Or Left$(ZTerminado, 2) = "DW" Then
                        Rem         XTipoPro = "CO"
                        Rem             Else
                        Rem         If XCodigo >= 0 And XCodigo <= 999 Then
                        Rem             XTipoPro = "CO"
                        Rem                 Else
                        Rem             If XCodigo >= 11000 And XCodigo <= 12999 Then
                        Rem                 XTipoPro = "CO"
                        Rem                     Else
                        Rem                 If XCodigo >= 25000 And XCodigo <= 25999 Then
                        Rem                     XTipoPro = "FA"
                        Rem                         Else
                        Rem                     If XCodigo >= 2300 And XCodigo <= 2399 Then
                        Rem                         XTipoPro = "BI"
                        Rem                             Else
                        Rem                         XTipoPro = "PT"
                        Rem                     End If
                        Rem                 End If
                        Rem             End If
                        Rem         End If
                        Rem     End If
                        Rem
                        Rem     ZLinea = 0
                        Rem     spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
                        Rem     Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        Rem     If rstTerminado.RecordCount > 0 Then
                        Rem         ZLinea = rstTerminado!Linea
                        Rem         rstTerminado.Close
                        Rem     End If
                        Rem
                        Rem     Select Case ZLinea
                        Rem         Case 8
                        Rem             XTipoPro = "PG"
                        Rem         Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                        Rem             XTipoPro = "FA"
                        Rem         Case Else
                        Rem     End Select
                        Rem
                        Rem     WTipopro = XTipoPro
                        Rem
                        Rem     Rem Select Case WTipopro
                        Rem     Rem     Case "CO", "PG"
                        Rem     Rem         WImpreProdI = "S"
                        Rem     Rem     Case "BI", "PT"
                        Rem     Rem         WImpreProdII = "S"
                        Rem     Rem     Case "FA"
                        Rem     Rem         WImpreProdIII = "S"
                        Rem     Rem     Case Else
                        Rem     Rem End Select
                        Rem
                        Rem     ZSql = ""
                        Rem     ZSql = ZSql & "INSERT INTO LiberaTerminado ("
                        Rem     ZSql = ZSql & "Codigo, "
                        Rem     ZSql = ZSql & "Producto, "
                        Rem     ZSql = ZSql & "Fecha, "
                        Rem     ZSql = ZSql & "OrdFecha, "
                        Rem     ZSql = ZSql & "Partida, "
                        Rem     ZSql = ZSql & "PartiOri, "
                        Rem     ZSql = ZSql & "Valor1, "
                        Rem     ZSql = ZSql & "Valor2, "
                        Rem     ZSql = ZSql & "Valor3, "
                        Rem     ZSql = ZSql & "Valor4, "
                        Rem     ZSql = ZSql & "Valor5, "
                        Rem     ZSql = ZSql & "Valor6, "
                        Rem     ZSql = ZSql & "Valor7, "
                        Rem     ZSql = ZSql & "Valor8, "
                        Rem     ZSql = ZSql & "Valor9, "
                        Rem     ZSql = ZSql & "Valor10, "
                        Rem     ZSql = ZSql & "Ensayo, "
                        Rem     ZSql = ZSql & "Aspecto, "
                        Rem     ZSql = ZSql & "Observaciones, "
                        Rem     ZSql = ZSql & "Confecciono, "
                        Rem     ZSql = ZSql & "Marca, "
                        Rem     ZSql = ZSql & "Cliente, "
                        Rem     ZSql = ZSql & "Cantidad, "
                        Rem     ZSql = ZSql & "Observa, "
                        Rem     ZSql = ZSql & "Origen, "
                        Rem     ZSql = ZSql & "Tipo, "
                        Rem     ZSql = ZSql & "ImpreProdI, "
                        Rem     ZSql = ZSql & "ImpreProdII, "
                        Rem     ZSql = ZSql & "ImpreProdIII, "
                        Rem     ZSql = ZSql & "ImpreVentas, "
                        Rem     ZSql = ZSql & "TipoPro) "
                        Rem     ZSql = ZSql & "Values ("
                        Rem     ZSql = ZSql & "'" + WCodigo + "',"
                        Rem     ZSql = ZSql & "'" + WProducto + "',"
                        Rem     ZSql = ZSql & "'" + WFecha + "',"
                        Rem     ZSql = ZSql & "'" + WOrdFecha + "',"
                        Rem     ZSql = ZSql & "'" + WPartida + "',"
                        Rem     ZSql = ZSql & "'" + WPartiOri + "',"
                        Rem     ZSql = ZSql & "'" + WValor1 + "',"
                        Rem     ZSql = ZSql & "'" + WValor2 + "',"
                        Rem     ZSql = ZSql & "'" + WValor3 + "',"
                        Rem     ZSql = ZSql & "'" + WValor4 + "',"
                        Rem     ZSql = ZSql & "'" + WValor5 + "',"
                        Rem     ZSql = ZSql & "'" + WValor6 + "',"
                        Rem     ZSql = ZSql & "'" + WValor7 + "',"
                        Rem     ZSql = ZSql & "'" + WValor8 + "',"
                        Rem     ZSql = ZSql & "'" + WValor9 + "',"
                        Rem     ZSql = ZSql & "'" + WValor10 + "',"
                        Rem     ZSql = ZSql & "'" + WEnsayo + "',"
                        Rem     ZSql = ZSql & "'" + WAspecto + "',"
                        Rem     ZSql = ZSql & "'" + WObservaciones + "',"
                        Rem     ZSql = ZSql & "'" + WConfecciono + "',"
                        Rem     ZSql = ZSql & "'" + WMarca + "',"
                        Rem     ZSql = ZSql & "'" + WCliente + "',"
                        Rem     ZSql = ZSql & "'" + WCantidad + "',"
                        Rem     ZSql = ZSql & "'" + WObserva + "',"
                        Rem     ZSql = ZSql & "'" + WOrigen + "',"
                        Rem     ZSql = ZSql & "'" + WTipo + "',"
                        Rem     ZSql = ZSql & "'" + WImpreProdI + "',"
                        Rem     ZSql = ZSql & "'" + WImpreProdII + "',"
                        Rem     ZSql = ZSql & "'" + WImpreProdIII + "',"
                        Rem     ZSql = ZSql & "'" + WImpreVentas + "',"
                        Rem     ZSql = ZSql & "'" + WTipopro + "')"
                        Rem
                        Rem     spLiberaTerminado = ZSql
                        Rem     Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        Rem
                        Rem Next Cicla
                        
                    End If
                    
                        Else
                        
                    If Muestra.Text = "MUESTRA" Then
                    
                        Rem WMarca = "S"
                        Rem Muestra.Col = 1
                        Rem WMuestra = Muestra.Text
                        Rem WPedido = WMuestra
                        
                        Rem Sql1 = "UPDATE Muestra SET "
                        Rem Sql2 = " Autoriza =  " + "'" + WMarca + "'"
                        Rem Sql3 = " Where Pedido = " + "'" + WPedido + "'"
                        Rem spMuestra = Sql1 + Sql2 + Sql3
                        Rem Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                        
                        Muestra.Col = 1
                        WPedido = Muestra.Text
                        
                        LugarPedido = 0
                        Erase CargaPedido
                        spPedido = "ConsultaPedido1 " + "'" + WPedido + "'"
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                        If rstPedido.RecordCount > 0 Then
                            With rstPedido
                                .MoveFirst
                                Do
                                    If .EOF = False Then
                                        LugarPedido = LugarPedido + 1
                                        CargaPedido(LugarPedido, 1) = rstPedido!Terminado
                                        CargaPedido(LugarPedido, 2) = Str$(rstPedido!Cantidad)
                                        If Left$(rstPedido!Terminado, 2) = "ML" Then
                                            CargaPedido(LugarPedido, 3) = IIf(IsNull(rstPedido!NombreComercial), "", rstPedido!NombreComercial)
                                        End If
                                        CargaPedido(LugarPedido, 4) = IIf(IsNull(rstPedido!OrdenTrabajo), "", rstPedido!OrdenTrabajo)
                                        CargaPedido(LugarPedido, 5) = IIf(IsNull(rstPedido!Referencia), "", rstPedido!Referencia)
                                        ZFechaPedido = rstPedido!Fecha
                                        ZCliente = rstPedido!Cliente
                                        ZObservaciones = rstPedido!Observaciones
                                        ZLugarDirEntrega = rstPedido!DirEntrega
                                        .MoveNext
                                            Else
                                        Exit Do
                                    End If
                                Loop
                            End With
                            rstPedido.Close
                        End If
                        
                        ZRazon = ""
                        spCliente = "ConsultaCliente " + "'" + ZCliente + "'"
                        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                        If rstCliente.RecordCount > 0 Then
                            ZRazon = rstCliente!Razon
                            ZVendedor = rstCliente!vendedor
                            Erase ZDirEntrega
                            ZDirEntrega(1) = rstCliente!DirEntrega
                            ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                            ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                            ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                            ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                            rstCliente.Close
                        End If
                    
                        ZDescriDirEntrega = ZDirEntrega(Val(ZLugarDirEntrega))
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Vendedor"
                        ZSql = ZSql + " Where Vendedor.Vendedor = " + "'" + ZVendedor + "'"
                        spVendedor = ZSql
                        Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
                        If rstVendedor.RecordCount > 0 Then
                            ZDesVendedor = rstVendedor!Nombre
                            rstVendedor.Close
                        End If
                        
                        Rem Sql1 = "Select Max(Pedido) as [PedidoMayor]"
                        Rem Sql2 = " FROM Muestra"
                        Rem spMuestra = Sql1 + Sql2
                        Rem Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                        Rem If rstMuestra.RecordCount > 0 Then
                        Rem     rstMuestra.MoveLast
                        Rem     WPedidoMayor = IIf(IsNull(rstMuestra!PedidoMayor), "0", rstMuestra!PedidoMayor)
                        Rem     ZPedido = Mid$(Str$(WPedidoMayor + 1), 2, 8)
                        Rem     rstMuestra.Close
                        Rem         Else
                        Rem     ZPedido = "1"
                        Rem End If
                        ZPedido = WPedido
    
                        ZRenglon = 0
                        For CiclaPedido = 1 To LugarPedido
        
                            ZCodigoTerminado = CargaPedido(CiclaPedido, 1)
                            ZCantidad = CargaPedido(CiclaPedido, 2)
                            ZNombreComercial = Trim(CargaPedido(CiclaPedido, 3))
                            ZOrdenTrabajo = CargaPedido(CiclaPedido, 4)
                            ZReferencia = Trim(CargaPedido(CiclaPedido, 5))
                            
                            Select Case Left$(ZCodigoTerminado, 2)
                                Case "PT", "PE", "YQ", "YF", "YP", "YH"
                                    ZTerminado = ZCodigoTerminado
                                    ZArticulo = ""
                                Case Else
                                    ZTerminado = ""
                                    ZArticulo = Left$(ZCodigoTerminado, 3) + Right$(ZCodigoTerminado, 7)
                            End Select
        
                            ZEnsayo = ""
                            ZNombre = ""
            
                            spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                            If rstTerminado.RecordCount > 0 Then
                                ZNombre = rstTerminado!Descripcion
                                rstTerminado.Close
                            End If
    
                            spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
                            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstArticulo.RecordCount > 0 Then
                                ZNombre = rstArticulo!Descripcion
                                rstArticulo.Close
                            End If
    
                            ZAutoriza = "S"
                            ZImpresion = "S"
                            
                            If ZNombreComercial = "" Then
                                ZDescriCliente = ZNombre
                                    Else
                                ZDescriCliente = ZNombreComercial
                            End If
                            
                            If ZTerminado <> "" Then
                                ClavePrecios = ZCliente + ZTerminado
                                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                                If rstPrecios.RecordCount > 0 Then
                                    ZDescriCliente = rstPrecios!Descripcion
                                    rstPrecios.Close
                                End If
                            End If
            
                            Sql1 = "Select Max(Codigo) as [CodigoMayor]"
                            Sql2 = " FROM Muestra"
                            spMuestra = Sql1 + Sql2
                            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMuestra.RecordCount > 0 Then
                                rstMuestra.MoveLast
                                WCodigoMayor = IIf(IsNull(rstMuestra!CodigoMayor), "0", rstMuestra!CodigoMayor)
                                ZCodigo = Mid$(Str$(WCodigoMayor + 1), 2, 8)
                                rstMuestra.Close
                                    Else
                                ZCodigo = "1"
                            End If
                            
                            If ZArticulo = "ML-008-100" Then
                            
                                spArticulo = "ConsultaArticulo " + "'" + "ML-008-100" + "'"
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstArticulo.RecordCount > 0 Then
                                
                                    ZZDescripcion = Trim(rstArticulo!Descripcion)
                                    ZZUnidad = rstArticulo!Unidad
                                    ZZDeposito = rstArticulo!Deposito
                                    ZZInicial = Str$(rstArticulo!Inicial)
                                    ZZEntradas = Str$(rstArticulo!Entradas)
                                    ZZSalidas = Str$(rstArticulo!Salidas)
                                    ZZMInimo = Str$(rstArticulo!MInimo)
                                    ZZMinimo1 = IIf(IsNull(rstArticulo!Minimo1), "0", rstArticulo!Minimo1)
                                    ZZVenta = IIf(IsNull(rstArticulo!Venta), "0", rstArticulo!Venta)
                                    ZZEnvase = Str$(rstArticulo!Envase)
                                    
                                    ZZCosto1 = "0"
                                    ZZWCosto1 = "0"
                                    ZZZCosto1 = "0"
                                    ZZCosto6 = "0"
                                    ZZWCosto2 = "0"
                                    ZZWCosto3 = "0"
                                    ZZCosto4 = "0"
                                    
                                    ZZOrdenI = ""
                                    ZZOrdenII = ""
                                    ZZOrdenIII = ""
                                    ZZPtaOrdenI = ""
                                    ZZPtaOrdenII = ""
                                    ZZPtaOrdenIII = ""
                                    
                                    ZZRs = rstArticulo!Rs
                                    ZZFlete = Str$(rstArticulo!Flete)
                                    ZZMoneda = rstArticulo!Moneda
                                    ZZControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                                    ZZReventa = IIf(IsNull(rstArticulo!Reventa), "0", rstArticulo!Reventa)
                                    ZZSedronar = IIf(IsNull(rstArticulo!Sedronar), "0", rstArticulo!Sedronar)
                                    ZZTipoMp = IIf(IsNull(rstArticulo!TipoMp), "0", rstArticulo!TipoMp)
                                    ZZCodSedronar = IIf(IsNull(rstArticulo!CodSedronar), "", rstArticulo!CodSedronar)
                                    ZZDensidad = IIf(IsNull(rstArticulo!Densidad), "", rstArticulo!Densidad)
                                    ZZCodigoDy = IIf(IsNull(rstArticulo!CodigoDy), "", rstArticulo!CodigoDy)
                                    ZZLeyenda = IIf(IsNull(rstArticulo!Leyenda), "0", rstArticulo!Leyenda)
                                    ZZClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
                                    ZZIntervencion = IIf(IsNull(rstArticulo!Intervencion), "", rstArticulo!Intervencion)
                                    ZZNaciones = IIf(IsNull(rstArticulo!Naciones), "", rstArticulo!Naciones)
                                    ZZEmbalaje = IIf(IsNull(rstArticulo!Embalaje), "", rstArticulo!Embalaje)
                                    ZZMeses = IIf(IsNull(rstArticulo!Meses), "0", rstArticulo!Meses)
                                    ZZDerechos = IIf(IsNull(rstArticulo!Derechos), "0", rstArticulo!Derechos)
                                    ZZparance = IIf(IsNull(rstArticulo!Posarance), "0", rstArticulo!Posarance)
                                    
                                    ZZTipoCosto = IIf(IsNull(rstArticulo!TipoCosto), "0", rstArticulo!TipoCosto)
                                    ZZProveedor = rstArticulo!Proveedor
                                    
                                    rstArticulo.Close
            
                                            
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM Articulo"
                                    ZSql = ZSql + " Where Articulo.Codigo <= " + "'" + "ML-999-100" + "'"
                                    ZSql = ZSql + " Order by Articulo.Codigo"
                                    spArticulo = ZSql
                                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstArticulo.RecordCount > 0 Then
                                        With rstArticulo
                                            .MoveLast
                                            ZZCodigoNuevo = rstArticulo!Codigo
                                        End With
                                        rstArticulo.Close
                                    End If
                                                    
                                    ZZNroMuestra = Val(Mid$(ZZCodigoNuevo, 4, 3))
                                    If ZZNroMuestra < 100 Then
                                        ZZNroMuestra = 100
                                            Else
                                        ZZNroMuestra = ZZNroMuestra + 1
                                    End If
                                    Auxi = Str$(ZZNroMuestra)
                                    Call Ceros(Auxi, 3)
                                    
                                    ZArticulo = "ML-" + Auxi + "-100"
                                            
                                    XParam = "'" + ZArticulo + "','" _
                                                 + ZZDescripcion + "','" _
                                                 + "0" + "','" _
                                                 + "0" + "','" _
                                                 + "0" + "','" _
                                                 + "0" + "','" _
                                                 + "0" + "','" _
                                                 + "0" + "','" _
                                                 + "0" + "','" _
                                                 + ZZUnidad + "','" _
                                                 + "0" + "','" _
                                                 + "0" + "','" _
                                                 + ZZEnvase + "','" _
                                                 + ZZRs + "','" _
                                                 + "  /  /    " + "','" _
                                                 + "" + "','" _
                                                 + "" + "','" _
                                                 + ZZProveedor + "','" _
                                                 + "" + "','" + ZZFlete + "','" _
                                                 + ZZMoneda + "','" + Str$(ZZControla) + "','" _
                                                 + ZZDensidad + "','" + "" + "','" _
                                                 + "" + "','" + "" + "','" _
                                                 + "" + "','" _
                                                 + "" + "'"
                                                 
                                    Set rstArticulo = db.OpenRecordset("AltaArticuloII " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                                    
                                    
                                    ZSql = ""
                                    ZSql = ZSql & "UPDATE Articulo SET "
                                    ZSql = ZSql & "Descripcion = " + "'" + ZZDescripcion + "',"
                                    ZSql = ZSql & "Costo1 = " + "'" + "" + "',"
                                    ZSql = ZSql & "Costo2 = " + "'" + "" + "',"
                                    ZSql = ZSql & "Inicial = " + "'" + "0" + "',"
                                    ZSql = ZSql & "Entradas = " + "'" + "" + "',"
                                    ZSql = ZSql & "Salidas = " + "'" + "" + "',"
                                    ZSql = ZSql & "Minimo = " + "'" + "" + "',"
                                    ZSql = ZSql & "Laboratorio = " + "'" + "" + "',"
                                    ZSql = ZSql & "Unidad = " + "'" + ZZUnidad + "',"
                                    ZSql = ZSql & "Pedido = " + "'" + "" + "',"
                                    ZSql = ZSql & "Deposito = " + "'" + "" + "',"
                                    ZSql = ZSql & "Envase = " + "'" + ZZEnvase + "',"
                                    ZSql = ZSql & "Rs = " + "'" + ZZRs + "',"
                                    ZSql = ZSql & "Fecha = " + "'" + "" + "',"
                                    ZSql = ZSql & "Orden = " + "'" + "" + "',"
                                    ZSql = ZSql & "Dife = " + "'" + "" + "',"
                                    ZSql = ZSql & "Proveedor = " + "'" + ZZProveedor + "',"
                                    ZSql = ZSql & "WDate = " + "'" + "" + "',"
                                    ZSql = ZSql & "Flete = " + "'" + ZZFlete + "',"
                                    ZSql = ZSql & "Moneda = " + "'" + ZZMoneda + "',"
                                    ZSql = ZSql & "Controla = " + "'" + Str$(ZZControla) + "',"
                                    ZSql = ZSql & "Densidad = " + "'" + ZZDensidad + "',"
                                    ZSql = ZSql & "Costo3 = " + "'" + "" + "',"
                                    ZSql = ZSql & "WCosto1 = " + "'" + "" + "',"
                                    ZSql = ZSql & "WCosto2 = " + "'" + "" + "',"
                                    ZSql = ZSql & "WCosto3 = " + "'" + "" + "',"
                                    ZSql = ZSql & "Venta = " + "'" + "" + "'"
                                    ZSql = ZSql & " Where Codigo = " + "'" + ZArticulo + "'"
                                            
                                    spArticulo = ZSql
                                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                    
                    
                    
                                    XParam = "'" + ZArticulo + "','" _
                                                 + "" + "'"
                                                     
                                    spArticulo = "ModificaArticuloMinimo1 " + XParam
                                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                    
                                    WLeyenda = ""
                                    XParam = "'" + ZArticulo + "','" _
                                                 + WLeyenda + "'"
                                                     
                                    spArticulo = "ModificaArticuloLeyenda " + XParam
                                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                    
                                    ZSql = ""
                                    ZSql = ZSql & "UPDATE Articulo SET "
                                    ZSql = ZSql & "Responsable = " + "'" + "" + "',"
                                    ZSql = ZSql & "Reventa = " + "'" + Str$(ZZReventa) + "',"
                                    ZSql = ZSql & "CodSedronar = " + "'" + ZZCodSedronar + "',"
                                    ZSql = ZSql & "Sedronar = " + "'" + ZZSedronar + "',"
                                    ZSql = ZSql & "TipoMp = " + "'" + ZZTipoMp + "',"
                                    ZSql = ZSql & "Minimo1 = " + "'" + Str$(ZZMinimo1) + "',"
                                    ZSql = ZSql & "Leyenda = " + "'" + Str$(ZZLeyenda) + "',"
                                    ZSql = ZSql & "Clase = " + "'" + "" + "',"
                                    ZSql = ZSql & "Intervencion = " + "'" + "" + "',"
                                    ZSql = ZSql & "Naciones = " + "'" + "" + "',"
                                    ZSql = ZSql & "Embalaje = " + "'" + "" + "',"
                                    ZSql = ZSql & "Meses = " + "'" + "" + "',"
                                    ZSql = ZSql & "TipoCosto = " + "'" + ZZTipoCosto + "',"
                                    ZSql = ZSql & "CodigoDy = " + "'" + ZZCodigoDy + "'"
                                    ZSql = ZSql & " Where Codigo = " + "'" + ZArticulo + "'"
                                            
                                    spArticulo = ZSql
                                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                                End If
                                
                            End If
                            
                
                            ZFecha = ZFechaPedido
                            ZFechaOrd = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
                
                            ZNombre = Left$(ZNombre, 50)
                            ZRazon = Left$(ZRazon, 50)
                            ZDescriCliente = Left$(ZDescriCliente, 50)
                            ZObservaciones = Left$(ZObservaciones, 50)
                            If Trim(ZReferencia) <> "" Then
                                ZObservaciones = Left$(ZReferencia, 50)
                            End If
                            ZDescriDirEntrega = Left$(ZDescriDirEntrega, 50)
                            ZOrdenTrabajo = ZOrdenTrabajo
                            
                            ZSql = ""
                            ZSql = ZSql + "INSERT INTO Muestra ("
                            ZSql = ZSql + "Codigo ,"
                            ZSql = ZSql + "Producto ,"
                            ZSql = ZSql + "Articulo ,"
                            ZSql = ZSql + "Ensayo ,"
                            ZSql = ZSql + "Nombre ,"
                            ZSql = ZSql + "Fecha ,"
                            ZSql = ZSql + "OrdFecha ,"
                            ZSql = ZSql + "Cantidad ,"
                            ZSql = ZSql + "Cliente ,"
                            ZSql = ZSql + "Razon ,"
                            ZSql = ZSql + "DescriCliente ,"
                            ZSql = ZSql + "Vendedor ,"
                            ZSql = ZSql + "DesVendedor ,"
                            ZSql = ZSql + "Observaciones ,"
                            ZSql = ZSql + "OrdenTRabajo ,"
                            ZSql = ZSql + "Autoriza ,"
                            ZSql = ZSql + "Impresion ,"
                            ZSql = ZSql + "Pedido ,"
                            ZSql = ZSql + "DirEntrega ,"
                            ZSql = ZSql + "DescriDirEntrega) "
                            ZSql = ZSql + "Values ("
                            ZSql = ZSql + "'" + ZCodigo + "',"
                            ZSql = ZSql + "'" + ZTerminado + "',"
                            ZSql = ZSql + "'" + ZArticulo + "',"
                            ZSql = ZSql + "'" + ZEnsayo + "',"
                            ZSql = ZSql + "'" + ZNombre + "',"
                            ZSql = ZSql + "'" + ZFecha + "',"
                            ZSql = ZSql + "'" + ZFechaOrd + "',"
                            ZSql = ZSql + "'" + ZCantidad + "',"
                            ZSql = ZSql + "'" + ZCliente + "',"
                            ZSql = ZSql + "'" + ZRazon + "',"
                            ZSql = ZSql + "'" + ZDescriCliente + "',"
                            ZSql = ZSql + "'" + ZVendedor + "',"
                            ZSql = ZSql + "'" + ZDesVendedor + "',"
                            ZSql = ZSql + "'" + ZObservaciones + "',"
                            ZSql = ZSql + "'" + ZOrdenTrabajo + "',"
                            ZSql = ZSql + "'" + ZAutoriza + "',"
                            ZSql = ZSql + "'" + ZImpresion + "',"
                            ZSql = ZSql + "'" + ZPedido + "',"
                            ZSql = ZSql + "'" + ZLugarDirEntrega + "',"
                            ZSql = ZSql + "'" + ZDescriDirEntrega + "')"
            
                            spMuestra = ZSql
                            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
        
                        Next CiclaPedido
                        
                        
                        
                        Rem actualiza pedido
                
                        Muestra.Col = 1
                        WPedido = Muestra.Text
                        WMarca = "X"
                        XFecha = DesdeFecha.Text
                        XFechaOrd = Right$(DesdeFecha.Text, 4) + Mid$(DesdeFecha.Text, 4, 2) + Left$(DesdeFecha.Text, 2)
                
                        XParam = "'" + WPedido + "','" _
                                + WMarca + "','" _
                                + XFecha + "','" _
                                + XFechaOrd + "'"
                                           
                        spPedido = "ModificaPedidoAutoriza " + XParam
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                        
                        WTipoPedido = 0
                        WTipoPed = 0
                        spPedido = "ListaPedido " + "'" + WPedido + "'"
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                        If rstPedido.RecordCount > 0 Then
                            WTipoPedido = rstPedido!TipoPedido
                            WTipoPed = rstPedido!Tipoped
                            rstPedido.Close
                        End If
                
                        Select Case WTipoPedido
                            Case 1, 4
                                WMarca = "0"
                            Case Else
                                WMarca = "1"
                        End Select
                        If Val(Wempresa) = 8 Then
                            Rem WMarca = "2"
                            WMarca = "1"
                        End If
                
                        XParam = "'" + WPedido + "','" _
                                    + WMarca + "'"
                                 
                        spPedido = "ModificaPedidoProceso1 " + XParam
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                        
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Pedido SET "
                        ZSql = ZSql + " ImpreMuestra =  " + "'" + "N" + "'"
                        ZSql = ZSql + " Where Pedido = " + "'" + WPedido + "'"
                        spPedido = ZSql
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                        
                            Else
                            
                        Rem dada
                        Rem actualiza pedido
                        Rem dada
                        
                        Muestra.Col = 1
                        WPedido = Muestra.Text
                        
                        WTipoPedido = 0
                        WTipoPed = 0
                        WFechaPedido = "00/00/0000"
                        spPedido = "ListaPedido " + "'" + WPedido + "'"
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                        If rstPedido.RecordCount > 0 Then
                            WTipoPedido = rstPedido!TipoPedido
                            WTipoPed = rstPedido!Tipoped
                            WFechaPedido = rstPedido!Fecha
                            WCliente = rstPedido!Cliente
                            rstPedido.Close
                        End If
                
                        Muestra.Col = 1
                        WPedido = Muestra.Text
                        WMarca = "X"
                        XFecha = WFechaAutoriza
                        XFechaOrd = Right$(WFechaAutoriza, 4) + Mid$(WFechaAutoriza, 4, 2) + Left$(WFechaAutoriza, 2)
                
                        XParam = "'" + WPedido + "','" _
                                + WMarca + "','" _
                                + XFecha + "','" _
                                + XFechaOrd + "'"
                                           
                        spPedido = "ModificaPedidoAutoriza " + XParam
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                        
                        If Val(Wempresa) = 1 Then
                        
                            ZZCanti = 0
                                
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM MovEnv"
                            ZSql = ZSql + " Where MovEnv.Cliente = " + "'" + WCliente + "'"
                            ZSql = ZSql + " and MovEnv.FechaOrd >= " + "'" + "20120101" + "'"
                            ZSql = ZSql + " and MovEnv.Envase = " + "'" + "30" + "'"
                            spMovenv = ZSql
                            Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovenv.RecordCount > 0 Then
                           
                                With rstMovenv
                            
                                    .MoveFirst
                                    
                                    Do
                                    
                                        WCantidad = rstMovenv!Cantidad
                                        WMovi = rstMovenv!Movimiento
                    
                                        If WMovi = "E" Then
                                            ZZCanti = ZZCanti - WCantidad
                                                Else
                                            ZZCanti = ZZCanti + WCantidad
                                        End If
                                        
                                        .MoveNext
                                        
                                        If .EOF = True Then
                                            Exit Do
                                        End If
                                        
                                    Loop
                                End With
                                
                                rstMovenv.Close
                                
                            End If
                            
                            If ZZCanti > 0 Then
                                ZSql = ""
                                ZSql = ZSql + "UPDATE Pedido SET "
                                ZSql = ZSql + " MarcaEnvase =  " + "'" + "N" + "',"
                                ZSql = ZSql + " CantidadEnvase =  " + "'" + Str$(ZZCanti) + "'"
                                ZSql = ZSql + " Where Pedido = " + "'" + WPedido + "'"
                                ZSql = ZSql + " and Renglon = " + "'" + "1" + "'"
                                spPedido = ZSql
                                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                                
                        End If
                        
                        If WFechaPedido <> WFechaAutoriza Then
                        
                            Rem If WTipoped = 0 Then
                            Rem
                            Rem     XFec1 = WFechaAutoriza
                            Rem     strDia = Format$(XFec1, "dddd")
                            Rem     BDia = Format(XFec1, "w")
                            Rem     Select Case BDia
                            Rem         Case 2, 3, 4
                            Rem             SumaDia = 2
                            Rem         Case 5, 6
                            Rem             SumaDia = 4
                            Rem         Case 7
                            Rem             SumaDia = 3
                            Rem         Case 1
                            Rem             SumaDia = 2
                            Rem         Case Else
                            Rem     End Select
                            Rem     SumaDia = SumaDia + 1
                            Rem     Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                            Rem     XFecEntrega = XFec2
                            Rem     XOrdFecEntrega = Right$(XFecEntrega, 4) + Mid$(XFecEntrega, 4, 2) + Left$(XFecEntrega, 2)
                            Rem
                            Rem     ZSql = ""
                            Rem     ZSql = ZSql + "UPDATE Pedido SET "
                            Rem     ZSql = ZSql + " FecEntrega =  " + "'" + XFecEntrega + "',"
                            Rem     ZSql = ZSql + " OrdFecEntrega =  " + "'" + XOrdFecEntrega + "'"
                            Rem     ZSql = ZSql + " Where Pedido = " + "'" + WPedido + "'"
                            Rem     spPedido = ZSql
                            Rem     Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                            Rem
                            Rem         Else
                            Rem
                            Rem     Muestra.Col = 11
                            Rem     XFec2 = Trim(Muestra.Text)
                            Rem
                            Rem     If XFec2 <> "" Then
                            Rem
                            Rem         XFecEntrega = XFec2
                            Rem         XOrdFecEntrega = Right$(XFecEntrega, 4) + Mid$(XFecEntrega, 4, 2) + Left$(XFecEntrega, 2)
                            Rem
                            Rem         ZSql = ""
                            Rem         ZSql = ZSql + "UPDATE Pedido SET "
                            Rem         ZSql = ZSql + " FecEntrega =  " + "'" + XFecEntrega + "',"
                            Rem         ZSql = ZSql + " OrdFecEntrega =  " + "'" + XOrdFecEntrega + "'"
                            Rem         ZSql = ZSql + " Where Pedido = " + "'" + WPedido + "'"
                            Rem         spPedido = ZSql
                            Rem         Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                            Rem
                            Rem     End If
                            Rem
                            Rem End If
                            
                            Muestra.Col = 11
                            XFec2 = Trim(Muestra.Text)
                            
                            If XFec2 <> "" Then
                            
                                ZZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                                ZZFechaOrd = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
                                XFecEntrega = XFec2
                                XOrdFecEntrega = Right$(XFecEntrega, 4) + Mid$(XFecEntrega, 4, 2) + Left$(XFecEntrega, 2)
                                
                                spPedido = "ListaPedido " + "'" + WPedido + "'"
                                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                                If rstPedido.RecordCount > 0 Then
                                    ZZFechaOriginal = rstPedido!Fecha
                                    rstPedido.Close
                                End If
                                
                                ZSql = ""
                                ZSql = ZSql + "UPDATE Pedido SET "
                                ZSql = ZSql + " MarcaAutorizacion =  " + "'" + "S" + "',"
                                ZSql = ZSql + " Fecha =  " + "'" + ZZFecha + "',"
                                ZSql = ZSql + " FechaOrd =  " + "'" + ZZFechaOrd + "',"
                                ZSql = ZSql + " FecEntrega =  " + "'" + XFecEntrega + "',"
                                ZSql = ZSql + " OrdFecEntrega =  " + "'" + XOrdFecEntrega + "'"
                                ZSql = ZSql + " Where Pedido = " + "'" + WPedido + "'"
                                spPedido = ZSql
                                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                                
                            End If
                            
                        End If
                        
                        WTipoPedido = 0
                        spPedido = "ListaPedido " + "'" + WPedido + "'"
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                        If rstPedido.RecordCount > 0 Then
                            WTipoPedido = rstPedido!TipoPedido
                            rstPedido.Close
                        End If
                
                        Select Case WTipoPedido
                            Case 1, 4
                                WMarca = "0"
                            Case Else
                                WMarca = "1"
                        End Select
                        If Val(Wempresa) = 8 Then
                            Rem WMarca = "2"
                            WMarca = "1"
                        End If
                
                        XParam = "'" + WPedido + "','" _
                                    + WMarca + "'"
                                 
                        spPedido = "ModificaPedidoProceso1 " + XParam
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                        
                        
                    End If
                    
                End If
                
            
            End If
                
            If Muestra.Text = "Anulado" Then
            
                Muestra.Col = 6
                If Muestra.Text = "DEVOL" Then
            
                    Muestra.Col = 1
                    WPedido = Muestra.Text
                
                    T$ = "Anulacion de la Solicitud de Devolucion"
                    m$ = "Confirma la anulacion de la Solcitud Nro.:" + Muestra.Text
                    Respuesta% = MsgBox(m$, 32 + 4, T$)
                    If Respuesta% = 6 Then
        
                        Muestra.Col = 1
                        WPedido = Muestra.Text
                
                        WMarca = "X"
                        XParam = "'" + WPedido + "'"
                        spPedidoDevol = "ModificaPedidoDevolAnulacion " + XParam
                        Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
                    
                    End If
                    
                        Else
                        
                    If Muestra.Text = "MUESTRA" Then
                    
                        Muestra.Col = 1
                        WPedido = Muestra.Text
                
                        T$ = "Anulacion de Muestras"
                        m$ = "Confirma la anulacion de la Muestra Nro.:" + Muestra.Text
                        Respuesta% = MsgBox(m$, 32 + 4, T$)
                        If Respuesta% = 6 Then
            
                            Muestra.Col = 1
                            WPedido = Muestra.Text
                    
                            Sql1 = "DELETE Muestra"
                            Sql2 = " Where Codigo = " + "'" + WPedido + "'"
                            spMuestra = Sql1 + Sql2
                            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                        
                        End If
                        
                            Else
                        
                        Muestra.Col = 1
                        WPedido = Muestra.Text
                
                        T$ = "Anulacion de Pedido"
                        m$ = "Confirma la anulacion del pedido Nro.:" + Muestra.Text
                        Respuesta% = MsgBox(m$, 32 + 4, T$)
                        If Respuesta% = 6 Then
            
                            Muestra.Col = 1
                            WPedido = Muestra.Text
                    
                            WMarca = "X"
                            XParam = "'" + WPedido + "'"
                            spPedido = "ModificaPedidoAnulacion " + XParam
                            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                        
                            WMarca = "0"
                            XParam = "'" + WPedido + "','" _
                                    + WMarca + "'"
                            spPedido = "ModificaPedidoProceso1 " + XParam
                            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                    
                        End If
                    End If
                End If
                
            End If
    
        Next Ciclo
    
        Call cmdClose_Click
    End If

End Sub


Private Sub Proceso_Click()

    WSalida = "N"
        
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
    Muestra.Col = 7
    Muestra.Text = "Importe"
    
    Muestra.Col = 8
    Muestra.Text = "Estado"
    
    Muestra.Col = 9
    Muestra.Text = "Impresa"
    
    Muestra.Col = 10
    Muestra.Text = ""
    
    Muestra.Col = 11
    Muestra.Text = ""
    
    Renglon = 0
    WSaldo = 0
    
    WAno = Right$(DesdeFecha.Text, 4)
    WMes = Mid$(DesdeFecha.Text, 4, 2)
    WDia = Left$(DesdeFecha.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHasta = WAno + WMes + WDia
    
    Pasa = 0
    Pedido = ""
    Fecha = "  /  /    "
    Cliente = ""
    Razon = ""
    FEntrega = "  /  /    "
    Tipo = 0
    Importe = 0
    Estado = ""
    
    Rem
    Rem Pedidos de Venta
    Rem
    
    XParam = "'" + WDesde + "','" _
            + "N" + " '"
    spPedido = "ListaPedidoFechaMarca " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
    With rstPedido
    
        .MoveFirst
        If .NoMatch = False Then
            Do
                Rem If WDesde <= rstPedido!FechaOrd And WHasta >= rstPedido!FechaOrd Then
                
                aa = rstPedido!FechaOrd
                
                    If Pasa = 0 Then
                        corte = rstPedido!Pedido
                        Fecha = rstPedido!Fecha
                        Cliente = rstPedido!Cliente
                        FEntrega = rstPedido!FecEntrega
                        Tipo = rstPedido!Tipoped
                        Importe = 0
                        Estado = rstPedido!Autorizo
                        Impresa = rstPedido!Impresion
                        Pasa = 1
                    End If
                    
                    If corte <> rstPedido!Pedido Then
                    
                        Renglon = Renglon + 1
            
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = Pusing("######", Str$(corte))
                        
                        Muestra.Col = 2
                        Muestra.Text = Fecha
                
                        Muestra.Col = 3
                        Muestra.Text = Cliente
                        
                        Muestra.Col = 4
                        Muestra.Text = ""
                        
                        Muestra.Col = 5
                        Muestra.Text = FEntrega
                        
                        Select Case Tipo
                            Case 0
                                Muestra.Col = 6
                                Muestra.Text = "Normal"
                            Case 1
                                Muestra.Col = 6
                                Muestra.Text = "A Fecha"
                            Case 2
                                Muestra.Col = 6
                                Muestra.Text = "Fecha LImite"
                            Case 3
                                Muestra.Col = 6
                                Muestra.Text = "Urgente"
                            Case 4
                                Muestra.Col = 6
                                Muestra.Text = "Retira Cliente"
                            Case 5
                                Muestra.Col = 6
                                Muestra.Text = "MUESTRA"
                            Case Else
                                Muestra.Col = 6
                                Muestra.Text = ""
                        End Select
                        
                        Muestra.Col = 7
                        Muestra.Text = Pusing("###,###,###.##", Str$(Importe))
                        
                        If Estado = "X" Then
                            Muestra.Col = 8
                            Muestra.Text = "Autorizad"
                                Else
                            Muestra.Col = 8
                            Muestra.Text = ""
                        End If
                        
                        If Impresa = "X" Then
                            Muestra.Col = 9
                            Muestra.Text = "SI"
                                Else
                            Muestra.Col = 9
                            Muestra.Text = ""
                        End If
                        
                        Muestra.Col = 10
                        Muestra.Text = ""
                        
                        Muestra.Col = 11
                        Muestra.Text = ""
                            
                        corte = rstPedido!Pedido
                        Fecha = rstPedido!Fecha
                        Cliente = rstPedido!Cliente
                        FEntrega = rstPedido!FecEntrega
                        Tipo = rstPedido!Tipoped
                        Importe = 0
                        Estado = rstPedido!Autorizo
                        Impresa = rstPedido!Impresion
                        Pasa = 1
                        
                    
                    End If
                    
                    Importe = Importe + ((rstPedido!Cantidad - rstPedido!Facturado) * rstPedido!Precio)
                    
                Rem End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    
    If Pasa <> 0 Then
                    
        Renglon = Renglon + 1
            
        Muestra.Row = Renglon
                        
        Muestra.Col = 1
        Muestra.Text = Pusing("######", Str$(corte))
                        
        Muestra.Col = 2
        Muestra.Text = Fecha
                
        Muestra.Col = 3
        Muestra.Text = Cliente
                        
        Muestra.Col = 4
        Muestra.Text = ""
                        
        Muestra.Col = 5
        Muestra.Text = FEntrega
                        
        Muestra.Col = 6
        Muestra.Text = Str$(Tipo)
        
        Select Case Tipo
            Case 0
                Muestra.Col = 6
                Muestra.Text = "Normal"
            Case 1
                Muestra.Col = 6
                Muestra.Text = "A Fecha"
            Case 2
                Muestra.Col = 6
                Muestra.Text = "Fecha LImite"
            Case 3
                Muestra.Col = 6
                Muestra.Text = "Urgente"
            Case 4
                Muestra.Col = 6
                Muestra.Text = "Retira Cliente"
            Case 5
                Muestra.Col = 6
                Muestra.Text = "MUESTRA"
            Case Else
                Muestra.Col = 6
                Muestra.Text = ""
        End Select
                                
        Muestra.Col = 7
        Muestra.Text = Pusing("###,###,###.##", Str$(Importe))
                        
        If Estado = "X" Then
            Muestra.Col = 8
            Muestra.Text = "Autorizad"
                Else
            Muestra.Col = 8
            Muestra.Text = ""
        End If
        
        If Impresa = "X" Then
            Muestra.Col = 9
            Muestra.Text = "SI"
                Else
            Muestra.Col = 9
            Muestra.Text = ""
        End If
        
        Muestra.Col = 10
        Muestra.Text = ""
        
        Muestra.Col = 11
        Muestra.Text = ""
                        
    End If
    
    rstPedido.Close
    
    End If
    
    Pasa = 0
    
    XParam = "'" + WDesde + "','" _
            + WHasta + "'"
    spPedido = "ListaPedidoFecha " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
    With rstPedido
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                If WDesde <= rstPedido!FechaOrd And WHasta >= rstPedido!FechaOrd Then
                
                    If Pasa = 0 Then
                        corte = rstPedido!Pedido
                        Fecha = rstPedido!Fecha
                        Cliente = rstPedido!Cliente
                        FEntrega = rstPedido!FecEntrega
                        Tipo = rstPedido!Tipoped
                        Importe = 0
                        Estado = rstPedido!Autorizo
                        Impresa = rstPedido!Impresion
                        Pasa = 1
                    End If
                    
                    If corte <> rstPedido!Pedido Then
                    
                        Renglon = Renglon + 1
            
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = Pusing("######", Str$(corte))
                        
                        Muestra.Col = 2
                        Muestra.Text = Fecha
                
                        Muestra.Col = 3
                        Muestra.Text = Cliente
                        
                        Muestra.Col = 4
                        Muestra.Text = ""
                        
                        Muestra.Col = 5
                        Muestra.Text = FEntrega
                        
                        Select Case Tipo
                            Case 0
                                Muestra.Col = 6
                                Muestra.Text = "Normal"
                            Case 1
                                Muestra.Col = 6
                                Muestra.Text = "A Fecha"
                            Case 2
                                Muestra.Col = 6
                                Muestra.Text = "Fecha LImite"
                            Case 3
                                Muestra.Col = 6
                                Muestra.Text = "Urgente"
                            Case 4
                                Muestra.Col = 6
                                Muestra.Text = "Retira Cliente"
                            Case 5
                                Muestra.Col = 6
                                Muestra.Text = "MUESTRA"
                            Case Else
                                Muestra.Col = 6
                                Muestra.Text = ""
                        End Select
                        
                        Muestra.Col = 7
                        Muestra.Text = Pusing("###,###,###.##", Str$(Importe))
                        
                        If Estado = "X" Then
                            Muestra.Col = 8
                            Muestra.Text = "Autorizad"
                                Else
                            Muestra.Col = 8
                            Muestra.Text = ""
                        End If
                        
                        If Impresa = "X" Then
                            Muestra.Col = 9
                            Muestra.Text = "SI"
                                Else
                            Muestra.Col = 9
                            Muestra.Text = ""
                        End If
                        
                        Muestra.Col = 10
                        Muestra.Text = ""
                        
                        corte = rstPedido!Pedido
                        Fecha = rstPedido!Fecha
                        Cliente = rstPedido!Cliente
                        FEntrega = rstPedido!FecEntrega
                        Tipo = rstPedido!Tipoped
                        Importe = 0
                        Estado = rstPedido!Autorizo
                        Impresa = rstPedido!Impresion
                        Pasa = 1
                        
                    
                    End If
                    
                    Importe = Importe + ((rstPedido!Cantidad - rstPedido!Facturado) * rstPedido!Precio)
                    
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    
    If Pasa <> 0 Then
                    
        Renglon = Renglon + 1
            
        Muestra.Row = Renglon
                        
        Muestra.Col = 1
        Muestra.Text = Pusing("######", Str$(corte))
                        
        Muestra.Col = 2
        Muestra.Text = Fecha
                
        Muestra.Col = 3
        Muestra.Text = Cliente
                        
        Muestra.Col = 4
        Muestra.Text = ""
                        
        Muestra.Col = 5
        Muestra.Text = FEntrega
                        
        Muestra.Col = 6
        Muestra.Text = Str$(Tipo)
        
        Select Case Tipo
            Case 0
                Muestra.Col = 6
                Muestra.Text = "Normal"
            Case 1
                Muestra.Col = 6
                Muestra.Text = "A Fecha"
            Case 2
                Muestra.Col = 6
                Muestra.Text = "Fecha LImite"
            Case 3
                Muestra.Col = 6
                Muestra.Text = "Urgente"
            Case 4
                Muestra.Col = 6
                Muestra.Text = "Retira Cliente"
            Case 5
                Muestra.Col = 6
                Muestra.Text = "MUESTRA"
            Case Else
                Muestra.Col = 6
                Muestra.Text = ""
        End Select
                                
        Muestra.Col = 7
        Muestra.Text = Pusing("###,###,###.##", Str$(Importe))
                        
        If Estado = "X" Then
            Muestra.Col = 8
            Muestra.Text = "Autorizad"
                Else
            Muestra.Col = 8
            Muestra.Text = ""
        End If
        
        If Impresa = "X" Then
            Muestra.Col = 9
            Muestra.Text = "SI"
                Else
            Muestra.Col = 9
            Muestra.Text = ""
        End If
        
        Muestra.Col = 10
        Muestra.Text = ""
                        
    End If
    
    rstPedido.Close
    
    End If
    
    
    Rem
    Rem solicitud de devolucion
    Rem
    
    Pasa = 0
    
     XParam = "'" + WDesde + "','" _
            + "N" + " '"
    spPedidoDevol = "ListaPedidoDevolFechaMarca " + XParam
    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedidoDevol.RecordCount > 0 Then
    With rstPedidoDevol
    
        .MoveFirst
        If .NoMatch = False Then
            Do
                
                aa = rstPedidoDevol!FechaOrd
            
                If Pasa = 0 Then
                    corte = rstPedidoDevol!Pedido
                    Fecha = rstPedidoDevol!Fecha
                    Cliente = rstPedidoDevol!Cliente
                    Importe = 0
                    Estado = rstPedidoDevol!Autorizo
                    Impresa = rstPedidoDevol!Impresion
                    Pasa = 1
                End If
                    
                If corte <> rstPedidoDevol!Pedido Then
                    
                    Renglon = Renglon + 1
            
                    Muestra.Row = Renglon
                        
                    Muestra.Col = 1
                    Muestra.Text = Pusing("######", Str$(corte))
                        
                    Muestra.Col = 2
                    Muestra.Text = Fecha
                
                    Muestra.Col = 3
                    Muestra.Text = Cliente
                        
                    Muestra.Col = 4
                    Muestra.Text = ""
                        
                    Muestra.Col = 5
                    Muestra.Text = ""
                        
                    Muestra.Col = 6
                    Muestra.Text = "DEVOL"
                        
                    Muestra.Col = 7
                    Muestra.Text = Pusing("###,###,###.##", Str$(Importe))
                        
                    If Estado = "X" Then
                        Muestra.Col = 8
                        Muestra.Text = "Autorizad"
                            Else
                        Muestra.Col = 8
                        Muestra.Text = ""
                    End If
                        
                    If Impresa = "X" Then
                        Muestra.Col = 9
                        Muestra.Text = "SI"
                            Else
                        Muestra.Col = 9
                        Muestra.Text = ""
                    End If
                    
                    Muestra.Col = 10
                    Muestra.Text = ""
                        
                    corte = rstPedidoDevol!Pedido
                    Fecha = rstPedidoDevol!Fecha
                    Cliente = rstPedidoDevol!Cliente
                    Importe = 0
                    Estado = rstPedidoDevol!Autorizo
                    Impresa = rstPedidoDevol!Impresion
                    Pasa = 1
                        
                End If
                
                Importe = Importe + ((rstPedidoDevol!Cantidad - rstPedidoDevol!Facturado) * rstPedidoDevol!Precio)
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    
    If Pasa <> 0 Then
                    
        Renglon = Renglon + 1
            
        Muestra.Row = Renglon
                        
        Muestra.Col = 1
        Muestra.Text = Pusing("######", Str$(corte))
                        
        Muestra.Col = 2
        Muestra.Text = Fecha
                
        Muestra.Col = 3
        Muestra.Text = Cliente
                        
        Muestra.Col = 4
        Muestra.Text = ""
                        
        Muestra.Col = 5
        Muestra.Text = ""
                        
        Muestra.Col = 6
        Muestra.Text = "DEVOL"
        
        Muestra.Col = 7
        Muestra.Text = Pusing("###,###,###.##", Str$(Importe))
                        
        If Estado = "X" Then
            Muestra.Col = 8
            Muestra.Text = "Autorizad"
                Else
            Muestra.Col = 8
            Muestra.Text = ""
        End If
        
        If Impresa = "X" Then
            Muestra.Col = 9
            Muestra.Text = "SI"
                Else
            Muestra.Col = 9
            Muestra.Text = ""
        End If
        
        Muestra.Col = 10
        Muestra.Text = ""
                        
    End If
    
    rstPedidoDevol.Close
    
    End If
    
    Pasa = 0
    
    XParam = "'" + WDesde + "','" _
            + WHasta + "'"
    spPedidoDevol = "ListaPedidoDevolFecha " + XParam
    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedidoDevol.RecordCount > 0 Then
    
    With rstPedidoDevol
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                If WDesde <= rstPedidoDevol!FechaOrd And WHasta >= rstPedidoDevol!FechaOrd Then
                
                    If Pasa = 0 Then
                        corte = rstPedidoDevol!Pedido
                        Fecha = rstPedidoDevol!Fecha
                        Cliente = rstPedidoDevol!Cliente
                        Importe = 0
                        Estado = rstPedidoDevol!Autorizo
                        Impresa = rstPedidoDevol!Impresion
                        Pasa = 1
                    End If
                    
                    If corte <> rstPedidoDevol!Pedido Then
                    
                        Renglon = Renglon + 1
            
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = Pusing("######", Str$(corte))
                        
                        Muestra.Col = 2
                        Muestra.Text = Fecha
                
                        Muestra.Col = 3
                        Muestra.Text = Cliente
                        
                        Muestra.Col = 4
                        Muestra.Text = ""
                        
                        Muestra.Col = 5
                        Muestra.Text = ""
                        
                        Muestra.Col = 6
                        Muestra.Text = "DEVOL"
                        
                        Muestra.Col = 7
                        Muestra.Text = Pusing("###,###,###.##", Str$(Importe))
                        
                        If Estado = "X" Then
                            Muestra.Col = 8
                            Muestra.Text = "Autorizad"
                                Else
                            Muestra.Col = 8
                            Muestra.Text = ""
                        End If
                        
                        If Impresa = "X" Then
                            Muestra.Col = 9
                            Muestra.Text = "SI"
                                Else
                            Muestra.Col = 9
                            Muestra.Text = ""
                        End If
                        
                        Muestra.Col = 10
                        Muestra.Text = ""
                        
                        corte = rstPedidoDevol!Pedido
                        Fecha = rstPedidoDevol!Fecha
                        Cliente = rstPedidoDevol!Cliente
                        Importe = 0
                        Estado = rstPedidoDevol!Autorizo
                        Impresa = rstPedidoDevol!Impresion
                        Pasa = 1
                        
                    
                    End If
                    
                    Importe = Importe + ((rstPedidoDevol!Cantidad - rstPedidoDevol!Facturado) * rstPedidoDevol!Precio)
                    
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    
    If Pasa <> 0 Then
                    
        Renglon = Renglon + 1
            
        Muestra.Row = Renglon
                        
        Muestra.Col = 1
        Muestra.Text = Pusing("######", Str$(corte))
                        
        Muestra.Col = 2
        Muestra.Text = Fecha
                
        Muestra.Col = 3
        Muestra.Text = Cliente
                        
        Muestra.Col = 4
        Muestra.Text = ""
                        
        Muestra.Col = 5
        Muestra.Text = ""
                        
        Muestra.Col = 6
        Muestra.Text = "DEVOL"
        
        Muestra.Col = 7
        Muestra.Text = Pusing("###,###,###.##", Str$(Importe))
                        
        
        If Estado = "X" Then
            Muestra.Col = 8
            Muestra.Text = "Autorizad"
                Else
            Muestra.Col = 8
            Muestra.Text = ""
        End If
        
        If Impresa = "X" Then
            Muestra.Col = 9
            Muestra.Text = "SI"
                Else
            Muestra.Col = 9
            Muestra.Text = ""
        End If
        
        Muestra.Col = 10
        Muestra.Text = ""
                        
    End If
    
    rstPedidoDevol.Close
    
    End If
    
    Rem
    Rem MUESTRAS PARA CLIENTES
    Rem
    Rem
    Rem Pasa = 0
    Rem
    Rem Sql1 = "Select *"
    Rem Sql2 = " FROM Muestra"
    Rem Sql3 = " Where Autoriza = " + "'" + "X" + "'"
    Rem Sql4 = " Order by Pedido, Codigo"
    Rem spMuestra = Sql1 + Sql2 + Sql3 + Sql4
    Rem Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstMuestra.RecordCount > 0 Then
    Rem     With rstMuestra
    Rem         .MoveFirst
    Rem         If .NoMatch = False Then
    Rem             Do
    Rem
    Rem                 If Pasa = 0 Then
    Rem                     corte = rstMuestra!Pedido
    Rem                     Fecha = rstMuestra!Fecha
    Rem                     Cliente = rstMuestra!Cliente
    Rem                     Razon = rstMuestra!Razon
    Rem                     FEntrega = rstMuestra!Fecha
    Rem                     Pasa = 1
    Rem                 End If
    Rem
    Rem                 If corte <> rstMuestra!Pedido Then
    Rem
    Rem                     Renglon = Renglon + 1
    Rem
    Rem                     Muestra.Row = Renglon
    Rem
    Rem                     Muestra.Col = 1
    Rem                     Muestra.Text = Pusing("######", Str$(corte))
    Rem
    Rem                     Muestra.Col = 2
    Rem                     Muestra.Text = Fecha
    Rem
    Rem                     Muestra.Col = 3
    Rem                     Muestra.Text = Cliente
    Rem
    Rem                     Muestra.Col = 4
    Rem                     Muestra.Text = Razon
    Rem
    Rem                     Muestra.Col = 5
    Rem                     Muestra.Text = FEntrega
    Rem
    Rem                     Muestra.Col = 6
    Rem                     Muestra.Text = "MUESTRA"
    Rem
    Rem                     Muestra.Col = 7
    Rem                     Muestra.Text = ""
    Rem
    Rem                     Muestra.Col = 8
    Rem                     Muestra.Text = ""
    Rem
    Rem                     Muestra.Col = 9
    Rem                     Muestra.Text = ""
    Rem
    Rem                     Muestra.Col = 10
    Rem                     Muestra.Text = ""
    Rem
    Rem                     corte = rstMuestra!Pedido
    Rem                     Fecha = rstMuestra!Fecha
    Rem                     Cliente = rstMuestra!Cliente
    Rem                     Razon = rstMuestra!Razon
    Rem                     FEntrega = rstMuestra!Fecha
    Rem                     Pasa = 1
    Rem
    Rem                 End If
    Rem
    Rem                 .MoveNext
    Rem
    Rem                 If .EOF = True Then
    Rem                     Exit Do
    Rem                 End If
    Rem
    Rem             Loop
    Rem         End If
    Rem
    Rem     End With
    Rem
    Rem     If Pasa <> 0 Then
    Rem
    Rem         Renglon = Renglon + 1
    Rem
    Rem         Muestra.Row = Renglon
    Rem
    Rem         Muestra.Col = 1
    Rem         Muestra.Text = Pusing("######", Str$(corte))
    Rem
    Rem         Muestra.Col = 2
    Rem         Muestra.Text = Fecha
    Rem
    Rem         Muestra.Col = 3
    Rem         Muestra.Text = Cliente
    Rem
    Rem         Muestra.Col = 4
    Rem         Muestra.Text = Razon
    Rem
    Rem         Muestra.Col = 5
    Rem         Muestra.Text = FEntrega
    Rem
    Rem         Muestra.Col = 6
    Rem         Muestra.Text = "MUESTRA"
    Rem
    Rem         Muestra.Col = 7
    Rem         Muestra.Text = ""
    Rem
    Rem         Muestra.Col = 8
    Rem         Muestra.Text = ""
    Rem
    Rem         Muestra.Col = 9
    Rem         Muestra.Text = ""
    Rem
    Rem         Muestra.Col = 10
    Rem         Muestra.Text = ""
    Rem
    Rem     End If
    Rem
    Rem     rstMuestra.Close
    Rem
    Rem End If
    
    For Dada = 1 To Renglon
    
        Muestra.Row = Dada
                        
        Muestra.Col = 3
        WCliente = Muestra.Text
    
        spCliente = "ConsultaCliente " + "'" + WCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Muestra.Col = 4
            Muestra.Text = rstCliente!Razon
            rstCliente.Close
        End If
        
    Next Dada
    
    TotalPedidos = Renglon
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    
    Muestra.TopRow = 1
    
    Muestra.SetFocus

End Sub

Private Sub Limpia_Vector()
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
    Muestra.Col = 7
    Muestra.Text = "Importe"
    
    Muestra.Col = 8
    Muestra.Text = "Estado"
    
    Muestra.Col = 9
    Muestra.Text = "Impresa"
    
End Sub

Private Sub Muestra_DblClick()

    Muestra.Col = 6
    If Muestra.Text = "DEVOL" Then
        Muestra.Col = 1
        WXPed = Muestra.Text
        PrgPeddev.Show
            Else
        If Muestra.Text = "MUESTRA" Then
            Rem Muestra.Col = 1
            Rem WXPed = Muestra.Text
            Rem PrgMuestraAutoriza.Show
            Muestra.Col = 1
            WXPed = Muestra.Text
            PrgPed.Show
                Else
            Muestra.Col = 1
            WXPed = Muestra.Text
            PrgPed.Show
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


Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFecha.Text, Auxi)
        If Auxi = "S" Then
            HastaFecha.Text = DesdeFecha.Text
            Call Proceso_Click
                Else
            DesdeFecha.SetFocus
        End If
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFecha.Text, Auxi)
        If Auxi = "S" Then
            Call Proceso_Click
                Else
            HastaFecha.SetFocus
        End If
    End If
End Sub

Private Sub VtoII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(VtoII.Text, Auxi)
        If Auxi = "S" Then
            Problema.SetFocus
        End If
    End If
End Sub


Sub Ingresa_clave()

    WClave.Text = ""
    Clave1.Visible = True
    WClave.SetFocus
    
End Sub

Private Sub CancelaGraba_Click()

    Clave1.Visible = False

End Sub





Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WGraba = "N"
        WClave.Text = UCase(WClave.Text)
        If WClave.Text = "BMW" Then
            WGraba = "S"
            Clave1.Visible = False
            Call Graba_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Archivo de Materias Primas")
            WClave.SetFocus
        End If
    End If

End Sub




Private Sub Calcula_FecEntrega()

    Rem 1 - DOMINGO
    Rem 2 - LUNES
    Rem 3 - MARTES
    Rem 4 - MIERCOLES
    Rem 5 - JUEVES
    Rem 6 - VIERNES
    Rem 7 - SABADO
    
    XFec1 = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    strDia = Format$(XFec1, "dddd")
    BDia = Format(XFec1, "w")
    Select Case BDia
        Case 2, 3, 4
            SumaDia = 2
        Case 5, 6
            SumaDia = 4
        Case 7
            SumaDia = 3
        Case 1
            SumaDia = 2
        Case Else
    End Select
    SumaDia = SumaDia + 1
    Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
    ZZFecEntrega = XFec2

End Sub

Private Sub Calcula_Feriado()

    Erase DiaFeriado
    TotalFeriado = 0
    
    spFeriado = "ListaFeriado"
    Set rstFeriado = db.OpenRecordset(spFeriado, dbOpenSnapshot, dbSQLPassThrough)
    If rstFeriado.RecordCount > 0 Then
        With rstFeriado
            .MoveFirst
            Do
                If .EOF = False Then
                    TotalFeriado = TotalFeriado + 1
                    DiaFeriado(TotalFeriado) = rstFeriado!Fecha
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstFeriado.Close
    End If
    
    Do
    
        Feriado = "N"
        For Ciclo = 1 To TotalFeriado
            If DiaFeriado(Ciclo) = ZZFecEntrega Then
                Feriado = "S"
                Exit For
            End If
        Next Ciclo
                
        Rem 1 - DOMINGO
        Rem 2 - LUNES
        Rem 3 - MARTES
        Rem 4 - MIERCOLES
        Rem 5 - JUEVES
        Rem 6 - VIERNES
        Rem 7 - SABADO
        XFec1 = ZZFecEntrega
        strDia = Format$(XFec1, "dddd")
        BDia = Format(XFec1, "w")
        If BDia = 1 Or BDia = 7 Then
            Feriado = "S"
        End If
        
        If Feriado = "S" Then
            SumaDia = 2
            Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
            ZZFecEntrega = XFec2
                Else
            Exit Do
        End If
        
    Loop

End Sub


