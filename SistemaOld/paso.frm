VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPaso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Tareas"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11760
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8205
   ScaleWidth      =   11760
   Visible         =   0   'False
   Begin VB.ComboBox TipoSolicitud 
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
      Left            =   9120
      TabIndex        =   45
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   3
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   2760
      Width           =   375
   End
   Begin VB.Frame Clave 
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
      Left            =   3000
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Cancelagraba 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Ingrese su Password"
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame PantallaObservaciones 
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
      Height          =   3135
      Left            =   7320
      TabIndex        =   27
      Top             =   4680
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Respuesta7 
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
         MaxLength       =   50
         TabIndex        =   41
         Top             =   2640
         Width           =   4095
      End
      Begin VB.TextBox Respuesta6 
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
         MaxLength       =   50
         TabIndex        =   40
         Top             =   2280
         Width           =   4095
      End
      Begin VB.TextBox Respuesta5 
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
         MaxLength       =   50
         TabIndex        =   39
         Top             =   1920
         Width           =   4095
      End
      Begin VB.TextBox Respuesta4 
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
         MaxLength       =   50
         TabIndex        =   38
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox Respuesta3 
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
         MaxLength       =   50
         TabIndex        =   37
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox Respuesta2 
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
         MaxLength       =   50
         TabIndex        =   36
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox Respuesta1 
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
         MaxLength       =   50
         TabIndex        =   35
         Top             =   480
         Width           =   4095
      End
   End
   Begin VB.Frame PantallaEstado 
      Caption         =   "Estado de la Solicitud de Tarea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   7320
      TabIndex        =   26
      Top             =   1200
      Width           =   4335
      Begin VB.TextBox Estado6 
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
         MaxLength       =   50
         TabIndex        =   34
         Top             =   2760
         Width           =   4095
      End
      Begin VB.TextBox Estado5 
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
         MaxLength       =   50
         TabIndex        =   33
         Top             =   2400
         Width           =   4095
      End
      Begin VB.TextBox Estado4 
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
         MaxLength       =   50
         TabIndex        =   32
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox Estado3 
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
         MaxLength       =   50
         TabIndex        =   31
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox Estado2 
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
         MaxLength       =   50
         TabIndex        =   30
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox Estado1 
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
         MaxLength       =   50
         TabIndex        =   29
         Top             =   960
         Width           =   4095
      End
      Begin VB.ComboBox Estado 
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "ESTADO"
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
         Height          =   315
         Left            =   120
         TabIndex        =   42
         Top             =   390
         Width           =   1455
      End
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2640
      Width           =   375
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
      Left            =   1560
      TabIndex        =   19
      Top             =   2160
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   480
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   390
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
      Left            =   960
      TabIndex        =   17
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Solicitante 
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
      MaxLength       =   30
      TabIndex        =   12
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox Planta 
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
      Left            =   6840
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   10
      Top             =   120
      Width           =   2895
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
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   8
      Text            =   " "
      Top             =   840
      Width           =   7935
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4920
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ImpreInsumos.rpt"
      Destination     =   3
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   17
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5400
      TabIndex        =   7
      Top             =   6960
      Width           =   975
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Top             =   120
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
   Begin VB.TextBox Solicitud 
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
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3600
      TabIndex        =   2
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1800
      TabIndex        =   1
      Top             =   6960
      Width           =   975
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4800
      TabIndex        =   22
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
      Height          =   5535
      Left            =   120
      TabIndex        =   23
      Top             =   1200
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9763
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox Entrega 
      Height          =   285
      Left            =   6840
      TabIndex        =   24
      Top             =   480
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
   Begin VB.Label Label6 
      Caption         =   "Tipo "
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
      Left            =   8520
      TabIndex        =   44
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "F.Entrega"
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
      Left            =   5760
      TabIndex        =   25
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Solicitante"
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
      TabIndex        =   11
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Planta"
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
      Left            =   5760
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
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
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
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
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro de Tarea"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgPaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Cantidad As Single
Private XCantidad As String
Dim rstMantenimiento As Recordset
Dim spMantenimiento As String
Dim XParam As String
Private Auxi As String
Private WAuxi As String
Private WGraba As String
Private WSolicitud As String
Dim ZZOperador As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstAuxiliar
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    PrgPaso.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Graba_Click()

    On Error GoTo WError
    
    If Trim(Observaciones.Text) = "" Then
        m$ = "No se ha informado observaciones (Este campo es de ingreso obligatorio)"
        a% = MsgBox(m$, 0, "Solicitud de Compra de Tareas")
        Exit Sub
    End If
    
    If Val(Solicitud.Text) = 0 Then
    
        XEmpresa = WEmpresa
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        ZSuma = 0
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Mantenimiento"
        ZSql = ZSql + " Where Renglon = 1"
        ZSql = ZSql + " and Estado = 2"
        ZSql = ZSql + " and Operador = " + "'" + Str$(WOperador) + "'"
        ZSql = ZSql + " Order by Solicitud"
        spMantenimiento = ZSql
        Set rstMantenimiento = db.OpenRecordset(spMantenimiento, dbOpenSnapshot, dbSQLPassThrough)
        If rstMantenimiento.RecordCount > 0 Then
            With rstMantenimiento
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        ZSuma = ZSuma + 1
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
                End If
            End With
            rstMantenimiento.Close
        End If
    
        Call Conecta_Empresa
        
        If ZSuma >= 8 Then
            m$ = "No puede generar ninguna soliciutd adicional " + Chr$(13) + "ya que posee mas de 5 solicitudes con compra realizada"
            a% = MsgBox(m$, 0, "Solicitud de Compra de Tareas")
            Exit Sub
        End If
        
    End If
    
    
    
    
    
    
    
    
    
    If WGraba <> "S" Then
        Call Ingresa_clave
            Else
            
        WGraba = ""
        
        XEmpresa = WEmpresa
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        ZZOperador = Str$(WOperador)
        
        If Val(Solicitud.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM Mantenimiento"
            Sql3 = " Where Mantenimiento.Solicitud = " + "'" + Solicitud.Text + "'"
            Sql4 = " Order by Clave"
            spMantenimiento = Sql1 + Sql2 + Sql3 + Sql4
            Set rstMantenimiento = db.OpenRecordset(spMantenimiento, dbOpenSnapshot, dbSQLPassThrough)
            If rstMantenimiento.RecordCount > 0 Then
                ZZOperador = IIf(IsNull(rstMantenimiento!Operador), "0", rstMantenimiento!Operador)
                rstMantenimiento.Close
                    Else
                m$ = "No se puede asignar manualmente el numero de Solicitud"
                a% = MsgBox(m$, 0, "Solicitud de Compra de Tareas")
                Call Conecta_Empresa
                Exit Sub
            End If
        End If
    
        If Val(Solicitud.Text) = 0 Then
            Sql1 = "Select Max(Solicitud), Solicitud"
            Sql2 = " FROM Mantenimiento"
            Sql3 = " Group By Solicitud"
            Sql4 = " Order By Solicitud"
            spMantenimiento = Sql1 + Sql2 + Sql3 + Sql4
            Set rstMantenimiento = db.OpenRecordset(spMantenimiento, dbOpenSnapshot, dbSQLPassThrough)
            If rstMantenimiento.RecordCount > 0 Then
                With rstMantenimiento
                    .MoveLast
                    Solicitud.Text = rstMantenimiento!Solicitud + 1
                End With
                rstMantenimiento.Close
            End If
        End If
        
        If Val(Solicitud.Text) = 0 Then
            Solicitud.Text = "1"
        End If
    
        Rem Borra la solicitud original
        
        Sql1 = "DELETE Mantenimiento"
        Sql2 = " Where Solicitud = " + "'" + Solicitud.Text + "'"
        spMantenimiento = Sql1 + Sql2
        Set rstMantenimiento = db.OpenRecordset(spMantenimiento, dbOpenSnapshot, dbSQLPassThrough)
        
        Renglon = 0
        
        For a = 1 To 100
        
            WVector1.Row = a
                    
            WVector1.Col = 1
            XCantidad = WVector1.Text
            
            WVector1.Col = 2
            XDescripcion = RTrim(WVector1.Text)
            
            WVector1.Col = 3
            XEstadoItem = WVector1.Text
                    
            If XDescripcion <> "" Or Val(Cantidad) <> 0 Then
            
                Renglon = Renglon + 1
            
                WSolicitud = Solicitud.Text
                WRenglon = Str$(Renglon)
                WFecha = Fecha.Text
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WPlanta = Planta.Text
                WSolicitante = Solicitante.Text
                WObservaciones = Observaciones.Text
                WEntrega = Entrega.Text
                WEntregaord = Right$(Entrega.Text, 4) + Mid$(Entrega.Text, 4, 2) + Left$(Entrega.Text, 2)
                WCantidad = XCantidad
                WDescripcion = XDescripcion
                WEstado = Str$(Estado.ListIndex)
                WEstado1 = Estado1.Text
                WEstado2 = Estado2.Text
                WEstado3 = Estado3.Text
                WEstado4 = Estado4.Text
                WEstado5 = Estado5.Text
                WEstado6 = Estado6.Text
                WRespuesta1 = Respuesta1.Text
                WRespuesta2 = Respuesta2.Text
                WRespuesta3 = Respuesta3.Text
                WRespuesta4 = Respuesta4.Text
                WRespuesta5 = Respuesta5.Text
                WRespuesta6 = Respuesta6.Text
                WRespuesta7 = Respuesta7.Text
                WEstadoItem = XEstadoItem
                WTipoSolicitud = Str$(TipoSolicitud.ListIndex)
                
                Auxi1 = WSolicitud
                Auxi = WRenglon
                Call Ceros(Auxi1, 6)
                Call Ceros(Auxi, 2)
                WClave = Auxi1 + Auxi
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO Mantenimiento ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Solicitud ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "OrdFecha ,"
                ZSql = ZSql + "Planta ,"
                ZSql = ZSql + "Solicitante ,"
                ZSql = ZSql + "Observaciones ,"
                ZSql = ZSql + "Entrega ,"
                ZSql = ZSql + "OrdEntrega ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "Descripcion ,"
                ZSql = ZSql + "Estado ,"
                ZSql = ZSql + "Estado1 ,"
                ZSql = ZSql + "Estado2 ,"
                ZSql = ZSql + "Estado3 ,"
                ZSql = ZSql + "Estado4 ,"
                ZSql = ZSql + "Estado5 ,"
                ZSql = ZSql + "Estado6 ,"
                ZSql = ZSql + "Respuesta1 ,"
                ZSql = ZSql + "Respuesta2 ,"
                ZSql = ZSql + "Respuesta3 ,"
                ZSql = ZSql + "Respuesta4 ,"
                ZSql = ZSql + "Respuesta5 ,"
                ZSql = ZSql + "Respuesta6 ,"
                ZSql = ZSql + "Respuesta7 ,"
                ZSql = ZSql + "EstadoItem ,"
                ZSql = ZSql + "Operador ,"
                ZSql = ZSql + "TipoSolicitud )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + Solicitud.Text + "',"
                ZSql = ZSql + "'" + WRenglon + "',"
                ZSql = ZSql + "'" + Fecha.Text + "',"
                ZSql = ZSql + "'" + WOrdFecha + "',"
                ZSql = ZSql + "'" + Planta.Text + "',"
                ZSql = ZSql + "'" + Solicitante.Text + "',"
                ZSql = ZSql + "'" + Observaciones.Text + "',"
                ZSql = ZSql + "'" + Entrega.Text + "',"
                ZSql = ZSql + "'" + WOrdEntrega + "',"
                ZSql = ZSql + "'" + WCantidad + "',"
                ZSql = ZSql + "'" + WDescripcion + "',"
                ZSql = ZSql + "'" + WEstado + "',"
                ZSql = ZSql + "'" + WEstado1 + "',"
                ZSql = ZSql + "'" + WEstado2 + "',"
                ZSql = ZSql + "'" + WEstado3 + "',"
                ZSql = ZSql + "'" + WEstado4 + "',"
                ZSql = ZSql + "'" + WEstado5 + "',"
                ZSql = ZSql + "'" + WEstado6 + "',"
                ZSql = ZSql + "'" + WRespuesta1 + "',"
                ZSql = ZSql + "'" + WRespuesta2 + "',"
                ZSql = ZSql + "'" + WRespuesta3 + "',"
                ZSql = ZSql + "'" + WRespuesta4 + "',"
                ZSql = ZSql + "'" + WRespuesta5 + "',"
                ZSql = ZSql + "'" + WRespuesta6 + "',"
                ZSql = ZSql + "'" + WRespuesta7 + "',"
                ZSql = ZSql + "'" + WEstadoItem + "',"
                ZSql = ZSql + "'" + ZZOperador + "',"
                ZSql = ZSql + "'" + WTipoSolicitud + "')"
        
                spMantenimiento = ZSql
                Set rstMantenimiento = db.OpenRecordset(spMantenimiento, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                        
        Next a
        
        Call Conecta_Empresa
                
        T$ = "Solicitud de Compra de Tareas"
        m$ = "Desea Imprimir la Solicitud de Tarea"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            Call Impresion
        End If
        
        Call Limpia_Click
        Solicitud.SetFocus
    
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub


Private Sub Limpia_Click()

    Solicitud.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Solicitante.Text = ""
    Observaciones.Text = ""
    Entrega.Text = "  /  /    "
    Select Case Val(WEmpresa)
        Case 1
            Planta.Text = "SI"
        Case 2
            Planta.Text = "PI"
        Case 3
            Planta.Text = "SII"
        Case 4
            Planta.Text = "PII"
        Case 5
            Planta.Text = "SIII"
        Case 6
            Planta.Text = "SIV"
        Case 7
            Planta.Text = "SV"
        Case 8
            Planta.Text = "PV"
        Case 9
            Planta.Text = "PVI"
        Case 10
            Planta.Text = "SVI"
        Case 11
            Planta.Text = "SVII"
        Case Else
            Planta.Text = "SI"
    End Select
    
    Estado.ListIndex = 0
    TipoSolicitud.ListIndex = 0
    
    Estado1.Text = ""
    Estado2.Text = ""
    Estado3.Text = ""
    Estado4.Text = ""
    Estado5.Text = ""
    Estado6.Text = ""
    
    Respuesta1.Text = ""
    Respuesta2.Text = ""
    Respuesta3.Text = ""
    Respuesta4.Text = ""
    Respuesta5.Text = ""
    Respuesta6.Text = ""
    Respuesta7.Text = ""
    
    Call Limpia_Vector
    Renglon = 0

    Solicitud.SetFocus

End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    Estado.Clear
    
    Estado.AddItem ""
    Estado.AddItem "En Proceso"
    Estado.AddItem ""
    Estado.AddItem ""
    Estado.AddItem ""
    Estado.AddItem ""
    Estado.AddItem ""
    Estado.AddItem "Pedido Cumplido"
    
    TipoSolicitud.Clear
    
    TipoSolicitud.AddItem "a"
    TipoSolicitud.AddItem "b"
    TipoSolicitud.AddItem "c"
    
    Solicitud.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Observaciones.Text = ""
    Solicitante.Text = ""
    Entrega.Text = "  /  /    "
    Select Case Val(WEmpresa)
        Case 1
            Planta.Text = "SI"
        Case 2
            Planta.Text = "PI"
        Case 3
            Planta.Text = "SII"
        Case 4
            Planta.Text = "PII"
        Case 5
            Planta.Text = "SIII"
        Case 6
            Planta.Text = "SIV"
        Case 7
            Planta.Text = "SV"
        Case 8
            Planta.Text = "PV"
        Case 9
            Planta.Text = "PVI"
        Case 10
            Planta.Text = "SVI"
        Case 11
            Planta.Text = "SVII"
        Case Else
            Planta.Text = "SI"
    End Select
    
    Estado.ListIndex = 0
    TipoSolicitud.ListIndex = 0
    
    Estado1.Text = ""
    Estado2.Text = ""
    Estado3.Text = ""
    Estado4.Text = ""
    Estado5.Text = ""
    Estado6.Text = ""
    
    Respuesta1.Text = ""
    Respuesta2.Text = ""
    Respuesta3.Text = ""
    Respuesta4.Text = ""
    Respuesta5.Text = ""
    Respuesta6.Text = ""
    Respuesta7.Text = ""

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgPaso.Caption = "Ingreso de Solicitud de Tareas :  " + !Nombre
        End If
    End With
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    
    Renglon = 0
    
    XEmpresa = WEmpresa
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Sql1 = "Select *"
    Sql2 = " FROM Mantenimiento"
    Sql3 = " Where Mantenimiento.Solicitud = " + "'" + Solicitud.Text + "'"
    Sql4 = " Order by CLave"
    spMantenimiento = Sql1 + Sql2 + Sql3 + Sql4
    Set rstMantenimiento = db.OpenRecordset(spMantenimiento, dbOpenSnapshot, dbSQLPassThrough)
    If rstMantenimiento.RecordCount > 0 Then
            
        With rstMantenimiento
            .MoveFirst
            Do
                If .EOF = False Then
            
                    Renglon = Renglon + 1
                    WVector1.Row = Renglon
                    
                    If rstMantenimiento!Cantidad <> 0 Then
                        WVector1.Col = 1
                        WVector1.Text = rstMantenimiento!Cantidad
                            Else
                        WVector1.Col = 1
                        WVector1.Text = ""
                    End If
                
                    WVector1.Col = 2
                    WVector1.Text = rstMantenimiento!Descripcion
                    
                    WVector1.Col = 3
                    WVector1.Text = IIf(IsNull(rstMantenimiento!EstadoItem), "", rstMantenimiento!EstadoItem)
            
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMantenimiento.Close
    End If
    
    Call Conecta_Empresa
    
End Sub

Private Sub Solicitud_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi = Solicitud.Text
        Call Ceros(Auxi, 6)
        WClave = Auxi + "01"
            
        Entra = "N"
        
        XEmpresa = WEmpresa
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        Sql1 = "Select *"
        Sql2 = " FROM Mantenimiento"
        Sql3 = " Where Mantenimiento.Solicitud = " + "'" + Solicitud.Text + "'"
        Sql4 = " Order by Clave"
        spMantenimiento = Sql1 + Sql2 + Sql3 + Sql4
        Set rstMantenimiento = db.OpenRecordset(spMantenimiento, dbOpenSnapshot, dbSQLPassThrough)
        If rstMantenimiento.RecordCount > 0 Then
            Fecha.Text = rstMantenimiento!Fecha
            Entrega.Text = rstMantenimiento!Entrega
            Observaciones.Text = RTrim(rstMantenimiento!Observaciones)
            Planta.Text = RTrim(rstMantenimiento!Planta)
            Solicitante.Text = RTrim(rstMantenimiento!Solicitante)
            Estado.ListIndex = rstMantenimiento!Estado
            TipoSolicitud.ListIndex = rstMantenimiento!TipoSolicitud
            Estado1.Text = RTrim(rstMantenimiento!Estado1)
            Estado2.Text = RTrim(rstMantenimiento!Estado2)
            Estado3.Text = RTrim(rstMantenimiento!Estado3)
            Estado4.Text = RTrim(rstMantenimiento!Estado4)
            Estado5.Text = RTrim(rstMantenimiento!Estado5)
            Estado6.Text = RTrim(rstMantenimiento!Estado6)
            Respuesta1.Text = RTrim(rstMantenimiento!Respuesta1)
            Respuesta2.Text = RTrim(rstMantenimiento!Respuesta2)
            Respuesta3.Text = RTrim(rstMantenimiento!Respuesta3)
            Respuesta4.Text = RTrim(rstMantenimiento!Respuesta4)
            Respuesta5.Text = RTrim(rstMantenimiento!Respuesta5)
            Respuesta6.Text = RTrim(rstMantenimiento!Respuesta6)
            Respuesta7.Text = RTrim(rstMantenimiento!Respuesta7)
            rstMantenimiento.Close
            Entra = "S"
                Else
            Rem WSolicitud = Solicitud.Text
            Rem Call Limpia_Click
            Rem Solicitud.Text = WSolicitud
            If Val(Solicitud.Text) = 0 Then
                Fecha.SetFocus
                    Else
                Solicitud.SetFocus
            End If
        End If
        
        Call Conecta_Empresa
        
        If Entra = "S" Then
            Call Proceso_Click
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        End If
        
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Solicitante.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Solicitante_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Entrega.SetFocus
    End If
End Sub

Private Sub Entrega_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Entrega.Text, Auxi)
        If Auxi = "S" Then
            TipoSolicitud.SetFocus
                Else
            Entrega.SetFocus
        End If
    End If
End Sub

Private Sub TipoSolicitud_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
End Sub

Private Sub Observaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
End Sub

Private Sub Estado1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Estado2.SetFocus
    End If
End Sub

Private Sub Estado2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Estado3.SetFocus
    End If
End Sub

Private Sub Estado3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Estado4.SetFocus
    End If
End Sub

Private Sub Estado4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Estado5.SetFocus
    End If
End Sub

Private Sub Estado5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Estado6.SetFocus
    End If
End Sub

Private Sub Estado6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Estado1.SetFocus
    End If
End Sub

Private Sub Respuesta1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Respuesta2.SetFocus
    End If
End Sub

Private Sub Respuesta2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Respuesta3.SetFocus
    End If
End Sub

Private Sub Respuesta3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Respuesta4.SetFocus
    End If
End Sub

Private Sub Respuesta4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Respuesta5.SetFocus
    End If
End Sub

Private Sub Respuesta5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Respuesta6.SetFocus
    End If
End Sub

Private Sub Respuesta6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Respuesta7.SetFocus
    End If
End Sub

Private Sub Respuesta7_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Respuesta1.SetFocus
    End If
End Sub

Sub Impresion()

    XEmpresa = WEmpresa
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    Listado.WindowTitle = "Emision de Solicitud de Tareas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Matenimiento.Solicitud} in " + Solicitud.Text + " to " + Solicitud.Text
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = ""
                            
    Rem Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    Listado.Action = 1
    
    Call Conecta_Empresa
 
End Sub
 
Sub Ingresa_clave()

    WClave.Text = ""
    Clave.Visible = True
    WClave.SetFocus
    
End Sub

Private Sub CancelaGraba_Click()

    Clave.Visible = False
    Solicitud.SetFocus

End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WGraba = "N"
        If WClave.Text = "SOL" Then
            WGraba = "S"
            Clave.Visible = False
            Call Graba_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Solicitud de Tareas")
            WClave.SetFocus
        End If
    End If
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
            Rem Select Case KeyAscii
            Rem     Case 0 To Asc(" ")
            Rem         WTexto1.Text = WVector1.Text
            Rem         WTexto1.SelStart = Len(WTexto1.Text)
            Rem     Case Else
            Rem         WTexto1.Text = Chr$(KeyAscii)
            Rem         WTexto1.SelStart = 1
            Rem End Select
            WTexto1.Text = WVector1.Text
            WTexto1.SelStart = 0
            WTexto1.SelLength = Len(WTexto1.Text)
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Rem Select Case KeyAscii
            Rem     Case 0 To Asc(" ")
            Rem         WTexto2.Text = WVector1.Text
            Rem         Rem WTexto2.SelStart = Len(WTexto2.Text)
            Rem         WTexto2.SelStart = 0
            Rem     Case Else
            Rem         WTexto2.Text = Chr$(KeyAscii)
            Rem         WTexto2.SelStart = 1
            Rem End Select
            WTexto2.Text = WVector1.Text
            WTexto2.SelStart = 0
            WTexto2.SelLength = Len(WTexto2.Text)
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

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Case 112, 113, 114, 115, 121, 122, 116, 117, 123
            Rem WVector1.SetFocus
            Rem WFuncion = KeyCode
            Rem Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            If WPideAyuda <> "S" Then
                Call StartEdit
            End If
            WPideAyuda = ""

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

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Case 112, 113, 114, 115, 121, 122, 116, 117, 123
            Rem WVector1.SetFocus
            Rem WFuncion = KeyCode
            Rem Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            If WPideAyuda <> "S" Then
                Call StartEdit
            End If
            WPideAyuda = ""
    
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
            
        Case 112, 113, 114, 115, 121, 122, 116, 117, 123
            Rem WVector1.SetFocus
            Rem  WFuncion = KeyCode
            Rem Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            If WPideAyuda <> "S" Then
                Call StartEdit
            End If
            WPideAyuda = ""

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
        Case 3
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            If Val(WVector1.Text) = 0 Then
                WVector1.Text = ""
            End If
        Case Else
    End Select
    
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector1.Rows - 1
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        WVector1.Col = 2
        WAuxi2 = WVector1.Text
        If WAuxi1 <> "" Or WAuxi2 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For Da = 1 To WVector1.Cols - 1
            WVector1.Col = Da
            WVector1.Text = WBorra(Ciclo, Da)
        Next Da
    Next Ciclo
    
    End If
    
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
    WVector1.Cols = 3
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
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 7000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Ok"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 2
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
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
    Rem modificar el tama�o de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub


