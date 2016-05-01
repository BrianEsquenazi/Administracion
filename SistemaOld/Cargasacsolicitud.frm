VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCargaSacSolicitud 
   Caption         =   "Solicitud de Carga de SAC"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   11775
   Begin VB.TextBox Referencia 
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
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   30
      Text            =   " "
      Top             =   1200
      Width           =   9615
   End
   Begin VB.TextBox Numero 
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
      Left            =   5880
      MaxLength       =   6
      TabIndex        =   2
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Centro 
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
      Left            =   8280
      MaxLength       =   6
      TabIndex        =   19
      Text            =   " "
      Top             =   120
      Width           =   855
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
      Left            =   8880
      TabIndex        =   18
      Top             =   480
      Width           =   2775
   End
   Begin VB.ComboBox Origen 
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
      TabIndex        =   17
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Ano 
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
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   1
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Tipo 
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
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox Opcion 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   1440
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox Pantalla 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      ItemData        =   "Cargasacsolicitud.frx":0000
      Left            =   360
      List            =   "Cargasacsolicitud.frx":0007
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.TextBox Ayuda 
      BackColor       =   &H00FFFFC0&
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
      Left            =   360
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.TextBox Titulo 
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
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   5
      Text            =   " "
      Top             =   1560
      Width           =   9615
   End
   Begin VB.TextBox ResponsableDestino 
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
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   4
      Text            =   " "
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox ResponsableEmisor 
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
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   3
      Text            =   " "
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox IngresoNoCon 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   2280
      Width           =   9615
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11640
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   7800
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   1200
      TabIndex        =   20
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
   Begin VB.Label Label5 
      Caption         =   "Referencia"
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
      TabIndex        =   31
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Año"
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
      Left            =   3480
      TabIndex        =   29
      Top             =   120
      Width           =   735
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
      Left            =   120
      TabIndex        =   28
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Centro"
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
      Left            =   7200
      TabIndex        =   27
      Top             =   120
      Width           =   855
   End
   Begin VB.Label DesCentro 
      BackColor       =   &H00C0C000&
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
      Height          =   255
      Left            =   9240
      TabIndex        =   26
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Estado"
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
      Left            =   7920
      TabIndex        =   25
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Origen"
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
      Left            =   3960
      TabIndex        =   24
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Numero"
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
      Left            =   5040
      TabIndex        =   23
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Tipo"
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
      TabIndex        =   22
      Top             =   120
      Width           =   735
   End
   Begin VB.Label DesTipo 
      BackColor       =   &H00C0C000&
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
      Height          =   255
      Left            =   2040
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Titulo"
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
      TabIndex        =   16
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label DesResponsableDestino 
      BackColor       =   &H00C0C000&
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
      Height          =   255
      Left            =   7440
      TabIndex        =   15
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label DesResponsableEmisor 
      BackColor       =   &H00C0C000&
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
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion de la No Conformidad"
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
      Height          =   300
      Left            =   1200
      TabIndex        =   13
      Top             =   1920
      Width           =   9615
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   4800
      MouseIcon       =   "Cargasacsolicitud.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "Cargasacsolicitud.frx":031F
      ToolTipText     =   "Consulta de Datos"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Label Label12 
      Caption         =   "Resp. Inv."
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
      Left            =   5400
      TabIndex        =   12
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Emisor"
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
      Top             =   840
      Width           =   975
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   6720
      MouseIcon       =   "Cargasacsolicitud.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "Cargasacsolicitud.frx":0E6B
      ToolTipText     =   "Salida"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   8760
      MouseIcon       =   "Cargasacsolicitud.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "Cargasacsolicitud.frx":19B7
      ToolTipText     =   "Elimina el Registro"
      Top             =   7440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   3840
      MouseIcon       =   "Cargasacsolicitud.frx":21F9
      MousePointer    =   99  'Custom
      Picture         =   "Cargasacsolicitud.frx":2503
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   5760
      MouseIcon       =   "Cargasacsolicitud.frx":2D45
      MousePointer    =   99  'Custom
      Picture         =   "Cargasacsolicitud.frx":304F
      ToolTipText     =   "Limpia la pantalla"
      Top             =   7440
      Width           =   480
   End
End
Attribute VB_Name = "PrgCargaSacSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTipoSac As Recordset
Dim spTipoSac As String
Dim rstCargaSacSolicitud As Recordset
Dim spCargaSacSolicitud As String
Dim rstCentroSac As Recordset
Dim spCentroSac As String
Dim rstResponsableSac As Recordset
Dim spResponsableSac As String

Dim XParam As String
Dim ZZLugar As Integer

Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
Dim MSubject As String
Dim MBody As String
Dim AllPath As String


Sub Imprime_Descripcion()
    
    Sql1 = "Select *"
    Sql2 = " FROM TipoSac"
    Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
    spTipoSac = Sql1 + Sql2 + Sql3
    Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoSac.RecordCount > 0 Then
        DesTipo.Caption = Trim(rstTipoSac!Descripcion)
        rstTipoSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM CentroSac"
    Sql3 = " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
    spCentroSac = Sql1 + Sql2 + Sql3
    Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCentroSac.RecordCount > 0 Then
        DesCentro.Caption = Trim(rstCentroSac!Descripcion)
        rstCentroSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableEmisor.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsableEmisor.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableDestino.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsableDestino.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If

End Sub

Sub Verifica_datos()
End Sub

Sub Imprime_Datos()

    On Error GoTo WError

    ZTipo = Tipo.Text
    ZAno = Ano.Text
    ZNumero = Numero.Text
    
    Call CmdLimpiar_Click
    
    ZExiste = "N"
    
    Tipo.Text = ZTipo
    Ano.Text = ZAno
    Numero.Text = ZNumero
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacSolicitudSolicitud"
    ZSql = ZSql + " Where CargaSacSolicitud.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacSolicitud.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacSolicitud.Numero = " + "'" + Numero.Text + "'"
    spCargaSacSolicitud = ZSql
    Set rstCargaSacSolicitud = db.OpenRecordset(spCargaSacSolicitud, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacSolicitud.RecordCount > 0 Then
    
        Centro.Text = rstCargaSacSolicitud!Centro
        Fecha.Text = rstCargaSacSolicitud!Fecha
        Origen.ListIndex = rstCargaSacSolicitud!Origen
        Estado.ListIndex = rstCargaSacSolicitud!Estado
        ResponsableEmisor.Text = rstCargaSacSolicitud!ResponsableEmisor
        ResponsableDestino.Text = rstCargaSacSolicitud!ResponsableDestino
        Referencia.Text = Trim(rstCargaSacSolicitud!Referencia)
        Titulo.Text = Trim(rstCargaSacSolicitud!Titulo)
        IngresoNoCon.Text = rstCargaSacSolicitud!Causa
        
        rstCargaSacSolicitud.Close
    End If
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub cmdAdd_Click()

    If Val(Tipo.Text) <> 0 And Val(Ano.Text) <> 0 And Val(Numero.Text) = 0 Then
    
        If Val(Ano.Text) < 2000 Or Val(Ano.Text) > 2017 Then
            m$ = "Error en carga de año"
            A% = MsgBox(m$, 0, "Archivo de Carga de Solicitudes")
            Exit Sub
        End If
    
        Numero.Text = "1"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacSolicitud"
        ZSql = ZSql + " Where CargaSacSolicitud.Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and CargaSacSolicitud.Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " Order by CargaSacSolicitud.Numero"
        spCargaSacSolicitud = ZSql
        Set rstCargaSacSolicitud = db.OpenRecordset(spCargaSacSolicitud, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacSolicitud.RecordCount > 0 Then
            rstCargaSacSolicitud.MoveLast
            ZUltimo = IIf(IsNull(rstCargaSacSolicitud!Numero), "0", rstCargaSacSolicitud!Numero)
            Numero.Text = ZUltimo + 1
            rstCargaSacSolicitud.Close
        End If
        
        m$ = "El numero de Solicitud asignado es " + Numero.Text
        A% = MsgBox(m$, 0, "Archivo de Carga de Solicitudes")
        
    End If

    If Tipo.Text <> "" And Ano.Text <> "" And Numero.Text <> "" Then
    
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        
        Auxi3 = Tipo.Text
        Auxi1 = Ano.Text
        Auxi2 = Numero.Text
        Call Ceros(Auxi3, 4)
        Call Ceros(Auxi1, 4)
        Call Ceros(Auxi2, 6)
        WClave = Auxi3 + Auxi1 + Auxi2
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacSolicitud"
        ZSql = ZSql + " Where CargaSacSolicitud.Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and CargaSacSolicitud.Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and CargaSacSolicitud.Numero = " + "'" + Numero.Text + "'"
        spCargaSacSolicitud = ZSql
        Set rstCargaSacSolicitud = db.OpenRecordset(spCargaSacSolicitud, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacSolicitud.RecordCount > 0 Then
        
            rstCargaSacSolicitud.Close
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaSacSolicitud SET "
            ZSql = ZSql + " Centro = " + "'" + Centro.Text + "',"
            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
            ZSql = ZSql + " OrdFecha = " + "'" + WOrdFecha + "',"
            ZSql = ZSql + " Origen = " + "'" + Str$(Origen.ListIndex) + "',"
            ZSql = ZSql + " Estado = " + "'" + Str$(Estado.ListIndex) + "',"
            ZSql = ZSql + " ResponsableEmisor = " + "'" + ResponsableEmisor.Text + "',"
            ZSql = ZSql + " ResponsableDestino = " + "'" + ResponsableDestino.Text + "',"
            ZSql = ZSql + " Referencia = " + "'" + Referencia.Text + "',"
            ZSql = ZSql + " Titulo = " + "'" + Titulo.Text + "',"
            ZSql = ZSql + " IngresoNoCon = " + "'" + IngresoNoCon.Text + "'"
            ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
            ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
            spCargaSacSolicitud = ZSql
            Set rstCargaSacSolicitud = db.OpenRecordset(spCargaSacSolicitud, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaSacSolicitud ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Centro ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Origen ,"
            ZSql = ZSql + "Estado ,"
            ZSql = ZSql + "ResponsableEmisor ,"
            ZSql = ZSql + "ResponsableDestino ,"
            ZSql = ZSql + "Referencia ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "IngresoNoCon )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Tipo.Text + "',"
            ZSql = ZSql + "'" + Ano.Text + "',"
            ZSql = ZSql + "'" + Numero.Text + "',"
            ZSql = ZSql + "'" + Centro.Text + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + WOrdFecha + "',"
            ZSql = ZSql + "'" + Str$(Origen.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado.ListIndex) + "',"
            ZSql = ZSql + "'" + ResponsableEmisor.Text + "',"
            ZSql = ZSql + "'" + ResponsableDestino.Text + "',"
            ZSql = ZSql + "'" + Referencia.Text + "',"
            ZSql = ZSql + "'" + Titulo.Text + "',"
            ZSql = ZSql + "'" + IngresoNoCon.Text + "')"
            
            spCargaSacSolicitud = ZSql
            Set rstCargaSacSolicitud = db.OpenRecordset(spCargaSacSolicitud, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        If Estado.ListIndex = 0 Then
        
            T$ = "Carga de Solicitud de " + DesTipo.Caption
            m$ = "Desea enviar el aviso al Responsable de Calidad"
            ZRespuesta% = MsgBox(m$, 32 + 4, T$)
            If ZRespuesta% = 6 Then
        
                ZZEmail = "ebiglieri@surfactan.com.ar; calidad@surfactan.com.ar"
            
                sTo = ZZEmail
                sCC = ""
                sBCC = ""
                sSubject = "Pedido de " + DesTipo.Caption
                sBody = "Se solicita emitir " + _
                        DesTipo.Caption + _
                        " Numero : " + Numero.Text + _
                        " Referencia : " + Referencia.Text + _
                        " Titulo : " + Titulo.Text
    
                ret = Shell("Start.exe " _
                        & "mailto:" & """" & sTo & """" _
                        & "?Subject=" & """" & sSubject & """" _
                        & "&cc=" & """" & sCC & """" _
                        & "&bcc=" & """" & sBCC & """" _
                        & "&Body=" & """" & sBody & """" _
                        & "&File=" & """" & "c:\autoexec.bat" & """" _
                        , 0)
            End If
            
        End If
        
        If Estado.ListIndex = 1 Then
    
            Numero.Text = "1"
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaSac"
            ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
            ZSql = ZSql + " Order by CargaSac.Numero"
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaSac.RecordCount > 0 Then
                rstCargaSac.MoveLast
                ZUltimo = IIf(IsNull(rstCargaSac!Numero), "0", rstCargaSac!Numero)
                Numero.Text = ZUltimo + 1
                rstCargaSac.Close
            End If
            
            m$ = "El numero de Solicitud asignado es " + Numero.Text
            A% = MsgBox(m$, 0, "Archivo de Carga de Solicitudes")
        
            WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        
            Auxi3 = Tipo.Text
            Auxi1 = Ano.Text
            Auxi2 = Numero.Text
            Call Ceros(Auxi3, 4)
            Call Ceros(Auxi1, 4)
            Call Ceros(Auxi2, 6)
            WClave = Auxi3 + Auxi1 + Auxi2
            
            Auxi = ""
    
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaSac ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Centro ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Origen ,"
            ZSql = ZSql + "Estado ,"
            ZSql = ZSql + "ResponsableEmisor ,"
            ZSql = ZSql + "ResponsableDestino ,"
            ZSql = ZSql + "Referencia ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "IngresoNoCon ,"
            ZSql = ZSql + "IngresoCausa )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Tipo.Text + "',"
            ZSql = ZSql + "'" + Ano.Text + "',"
            ZSql = ZSql + "'" + Numero.Text + "',"
            ZSql = ZSql + "'" + Centro.Text + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + WOrdFecha + "',"
            ZSql = ZSql + "'" + Str$(Origen.ListIndex) + "',"
            ZSql = ZSql + "'" + "1" + "',"
            ZSql = ZSql + "'" + ResponsableEmisor.Text + "',"
            ZSql = ZSql + "'" + ResponsableDestino.Text + "',"
            ZSql = ZSql + "'" + Referencia.Text + "',"
            ZSql = ZSql + "'" + Titulo.Text + "',"
            ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
            ZSql = ZSql + "'" + "." + "')"
            
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
        
        
            T$ = "Carga de " + DesTipo.Caption
            m$ = "Desea enviar el aviso al Responsable del Area"
            ZRespuesta% = MsgBox(m$, 32 + 4, T$)
            If ZRespuesta% = 6 Then
        
                ZZResponsable = 0
        
                Sql1 = "Select *"
                Sql2 = " FROM CentroSac"
                Sql3 = " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
                spCentroSac = Sql1 + Sql2 + Sql3
                Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstCentroSac.RecordCount > 0 Then
                    ZZResponsable = rstCentroSac!Responsable
                    rstCentroSac.Close
                End If
                
                If ZZResponsable <> 0 Then
                
                    ZZEmail = ""
                    
                    Sql1 = "Select *"
                    Sql2 = " FROM ResponsableSac"
                    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Str$(ZZResponsable) + "'"
                    spResponsableSac = Sql1 + Sql2 + Sql3
                    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
                    If rstResponsableSac.RecordCount > 0 Then
                        ZZEmail = Trim(rstResponsableSac!Email)
                        rstResponsableSac.Close
                    End If
                    
                    If ZZEmail <> "" Then
                
                        sTo = ZZEmail
                        sCC = ""
                        sBCC = ""
                        sSubject = "Pededo de " + DesTipo.Caption
                        sBody = "Se solicida emitir " + DesTipo.Caption + ". Referencia : " + Referencia.Text + ". Titulo : " + Titulo.Text
        
                        ret = Shell("Start.exe " _
                                & "mailto:" & """" & sTo & """" _
                                & "?Subject=" & """" & sSubject & """" _
                                & "&cc=" & """" & sCC & """" _
                                & "&bcc=" & """" & sBCC & """" _
                                & "&Body=" & """" & sBody & """" _
                                & "&File=" & """" & "c:\autoexec.bat" & """" _
                                , 0)
            
                    End If
                End If
            End If
        
        
            T$ = "Carga de " + DesTipo.Caption
            m$ = "Desea enviar el aviso al Responsable de Investigacion"
            ZRespuesta% = MsgBox(m$, 32 + 4, T$)
            If ZRespuesta% = 6 Then
        
                ZZResponsable = Val(ResponsableDestino.Text)
        
                If ZZResponsable <> 0 Then
            
                    ZZEmail = ""
                
                    Sql1 = "Select *"
                    Sql2 = " FROM ResponsableSac"
                    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Str$(ZZResponsable) + "'"
                    spResponsableSac = Sql1 + Sql2 + Sql3
                    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
                    If rstResponsableSac.RecordCount > 0 Then
                        ZZEmail = Trim(rstResponsableSac!Email)
                        rstResponsableSac.Close
                    End If
                    
                    If ZZEmail <> "" Then
                
                        sTo = ZZEmail
                        sCC = ""
                        sBCC = ""
                        sSubject = "Carga de " + DesTipo.Caption
                        sBody = "Se ingreso una " + DesTipo.Caption + " : " + Ano.Text + "/" + Numero.Text + " para determinar las acciones correctivas correspondientes    Referencia : " + Referencia.Text
        
                        ret = Shell("Start.exe " _
                                & "mailto:" & """" & sTo & """" _
                                & "?Subject=" & """" & sSubject & """" _
                                & "&cc=" & """" & sCC & """" _
                                & "&bcc=" & """" & sBCC & """" _
                                & "&Body=" & """" & sBody & """" _
                                & "&File=" & """" & "c:\autoexec.bat" & """" _
                                , 0)
            
                    End If
                End If
            End If
        
        End If
        
        Call CmdLimpiar_Click
        Tipo.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()

    If Tipo.Text <> "" And Ano.Text <> "" And Numero.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacSolicitud"
        ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
        spCargaSacSolicitud = ZSql
        Set rstCargaSacSolicitud = db.OpenRecordset(spCargaSacSolicitud, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacSolicitud.RecordCount > 0 Then
        
            rstCargaSacSolicitud.Close
            
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
            
                ZSql = ""
                Sql1 = ZSql + "DELETE CargaSacSolicitud"
                ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
                ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
                ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
                spCargaSacSolicitud = Sql1 + Sql2
                Set rstCargaSacSolicitud = db.OpenRecordset(spCargaSacSolicitud, dbOpenSnapshot, dbSQLPassThrough)
                
                Call CmdLimpiar_Click
                
            End If
        End If
        
    End If
    
    Tipo.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()
    
    Tipo.Text = "1"
    DesTipo.Caption = "SAC"
    Ano.Text = "2010"
    Numero.Text = ""
    Centro.Text = ""
    DesCentro.Caption = ""
    Fecha.Text = "  /  /    "
    ResponsableEmisor.Text = ""
    ResponsableDestino.Text = ""
    DesResponsableEmisor.Caption = ""
    DesResponsableDestino.Caption = ""
    Referencia.Text = ""
    Titulo.Text = ""
    IngresoNoCon.Text = ""
    
    Origen.ListIndex = 0
    Estado.ListIndex = 0
    
    Tipo.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgCargaSacSolicitud.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()

    Tipo.Text = "1"
    DesTipo.Caption = "SAC"
    Ano.Text = "2010"
    Numero.Text = ""
    Centro.Text = ""
    DesCentro.Caption = ""
    Fecha.Text = "  /  /    "
    ResponsableEmisor.Text = ""
    ResponsableDestino.Text = ""
    Referencia.Text = ""
    Titulo.Text = ""
    IngresoNoCon.Text = ""
    
    Estado.Clear
    
    Estado.AddItem ""
    Estado.AddItem "APROBADO"
    Estado.AddItem "RECHAZADO"
    
    Estado.ListIndex = 0
    
    Origen.Clear
    
    Origen.AddItem ""
    Origen.AddItem "Auditoria"
    Origen.AddItem "Reclamo"
    Origen.AddItem "I. No Conformidad"
    Origen.AddItem "Proceso/Sist"
    Origen.AddItem "Otro"
    
    Origen.ListIndex = 0
    
End Sub

Private Sub Tipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM TipoSac"
        Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
        spTipoSac = Sql1 + Sql2 + Sql3
        Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoSac.RecordCount > 0 Then
            DesTipo.Caption = Trim(rstTipoSac!Descripcion)
            rstTipoSac.Close
            Ano.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Tipo.Text = ""
        DesTipo.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Numero.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Numero.Text <> "" Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaSacSolicitud"
            ZSql = ZSql + " Where CargaSacSolicitud.Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and CargaSacSolicitud.Ano = " + "'" + Ano.Text + "'"
            ZSql = ZSql + " and CargaSacSolicitud.Numero = " + "'" + Numero.Text + "'"
            spCargaSacSolicitud = ZSql
            Set rstCargaSacSolicitud = db.OpenRecordset(spCargaSacSolicitud, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaSacSolicitud.RecordCount > 0 Then
                
                Centro.Text = rstCargaSacSolicitud!Centro
                Fecha.Text = rstCargaSacSolicitud!Fecha
                Origen.ListIndex = rstCargaSacSolicitud!Origen
                Estado.ListIndex = rstCargaSacSolicitud!Estado
                ResponsableEmisor.Text = rstCargaSacSolicitud!ResponsableEmisor
                ResponsableDestino.Text = rstCargaSacSolicitud!ResponsableDestino
                Referencia.Text = rstCargaSacSolicitud!Referencia
                Titulo.Text = rstCargaSacSolicitud!Titulo
                IngresoNoCon.Text = rstCargaSacSolicitud!IngresoNoCon
                rstCargaSacSolicitud.Close
                
                Call Imprime_Descripcion
                Centro.SetFocus
                
                    Else
                    
                WTipo = Tipo.Text
                WAno = Ano.Text
                WNumero = Numero.Text
                CmdLimpiar_Click
                Ano.Text = WAno
                Numero.Text = WNumero
                Tipo.Text = WTipo
                Sql1 = "Select *"
                Sql2 = " FROM TipoSac"
                Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
                spTipoSac = Sql1 + Sql2 + Sql3
                Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstTipoSac.RecordCount > 0 Then
                    DesTipo.Caption = Trim(rstTipoSac!Descripcion)
                    rstTipoSac.Close
                    Ano.SetFocus
                End If
                Centro.SetFocus
                
            End If
            
        End If
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Centro_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM CentroSac"
        Sql3 = " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
        spCentroSac = Sql1 + Sql2 + Sql3
        Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstCentroSac.RecordCount > 0 Then
            DesCentro.Caption = Trim(rstCentroSac!Descripcion)
            rstCentroSac.Close
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Centro.Text = ""
        DesCentro.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Origen.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Origen_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Estado.SetFocus
    End If
End Sub

Private Sub Estado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ResponsableEmisor.SetFocus
    End If
End Sub

Private Sub ResponsableEmisor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableEmisor.Text + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            DesResponsableEmisor.Caption = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
            ResponsableDestino.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        ResponsableEmisor.Text = ""
        DesResponsableEmisor.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ResponsableDestino_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableDestino.Text + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            DesResponsableDestino.Caption = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
            Referencia.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        ResponsableDestino.Text = ""
        DesResponsableDestino.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Referencia_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Titulo.SetFocus
    End If
    If KeyAscii = 27 Then
        Referencia.Text = ""
    End If
End Sub

Private Sub Titulo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Centro.SetFocus
    End If
    If KeyAscii = 27 Then
        Titulo.Text = ""
    End If
End Sub

Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False

     Opcion.Clear

     Opcion.AddItem "Tipo"
     Opcion.AddItem "Centro"

     Opcion.Visible = True
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    Ayuda.Text = ""
    Ayuda.Visible = True
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM tiposac"
            Sql3 = " Order by tiposac.Codigo"
            spTipoSac = Sql1 + Sql2 + Sql3
            Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoSac.RecordCount > 0 Then
                With rstTipoSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstTipoSac!Codigo) + " " + rstTipoSac!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstTipoSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTipoSac.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM CentroSac"
            Sql3 = " Order by CentroSac.Codigo"
            spCentroSac = Sql1 + Sql2 + Sql3
            Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCentroSac.RecordCount > 0 Then
                With rstCentroSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstCentroSac!Codigo) + " " + rstCentroSac!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCentroSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCentroSac.Close
            End If
        
        Case 2
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Order by ResponsableSac.Codigo"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                With rstResponsableSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstResponsableSac!Codigo) + " " + rstResponsableSac!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstResponsableSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstResponsableSac.Close
            End If
        
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Tipo.Text = WIndice.List(Indice)
            Call Tipo_Keypress(13)
            
        Case 1
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Centro.Text = WIndice.List(Indice)
            Call Centro_Keypress(13)
            
        Case 2
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Select Case ZZLugar
                Case 1
                    ResponsableEmisor.Text = WIndice.List(Indice)
                    Call ResponsableEmisor_Keypress(13)
                Case Else
                    ResponsableDestino.Text = WIndice.List(Indice)
                    Call ResponsableDestino_Keypress(13)
            End Select
            
        Case Else
    End Select
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    LugarAyuda = 0
    WIndice.Clear
    Pantalla.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM TipoSac"
            Sql3 = " Order by TipoSac.Codigo"
            spTipoSac = Sql1 + Sql2 + Sql3
            Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoSac.RecordCount > 0 Then
                With rstTipoSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstTipoSac!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstTipoSac!Descripcion, aa, WEspacios) Then
                                    IngresaItem = Str$(rstTipoSac!Codigo) + " " + rstTipoSac!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstTipoSac!Codigo
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
                rstTipoSac.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM CentroSac"
            Sql3 = " Order by CentroSac.Codigo"
            spCentroSac = Sql1 + Sql2 + Sql3
            Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCentroSac.RecordCount > 0 Then
                With rstCentroSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstCentroSac!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstCentroSac!Descripcion, aa, WEspacios) Then
                                    IngresaItem = Str$(rstCentroSac!Codigo) + " " + rstCentroSac!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstCentroSac!Codigo
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
                rstCentroSac.Close
            End If
            

        Case 2
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Order by ResponsableSac.Codigo"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                With rstResponsableSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstResponsableSac!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstResponsableSac!Descripcion, aa, WEspacios) Then
                                    IngresaItem = Str$(rstResponsableSac!Codigo) + " " + rstResponsableSac!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstResponsableSac!Codigo
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
                rstResponsableSac.Close
            End If
                
        Case Else
    End Select
    
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub


Private Sub Tipo_DblClick()

    Opcion.Clear
    Opcion.AddItem ""
    Opcion.AddItem ""
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

End Sub

Private Sub Centro_DblClick()

    Opcion.Clear
    Opcion.AddItem ""
    Opcion.AddItem ""
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub ResponsableEmisor_DblClick()

    ZZLugar = 1

    Opcion.Clear
    Opcion.AddItem ""
    Opcion.AddItem ""
    Opcion.AddItem ""
    Rem Opcion.Visible = True
    Opcion.ListIndex = 2
    
    Rem Call Opcion_Click

End Sub

Private Sub ResponsableDestino_DblClick()

    ZZLugar = 2

    Opcion.Clear
    Opcion.AddItem ""
    Opcion.AddItem ""
    Opcion.AddItem ""
    Rem Opcion.Visible = True
    Opcion.ListIndex = 2
    
    Rem Call Opcion_Click

End Sub

