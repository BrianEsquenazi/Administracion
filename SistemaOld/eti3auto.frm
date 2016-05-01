VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEti3Auto 
   Caption         =   "Impresion de Etiquetas"
   ClientHeight    =   6825
   ClientLeft      =   735
   ClientTop       =   645
   ClientWidth     =   10425
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6825
   ScaleWidth      =   10425
   Begin VB.Frame PantaDirEntrega 
      Caption         =   "Seleccion de Lugar de Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   26
      Top             =   4080
      Visible         =   0   'False
      Width           =   9375
      Begin VB.ListBox ListaDirEntrega 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   9015
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   9855
      Begin VB.ComboBox TipoProceso 
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
         Left            =   3720
         TabIndex        =   28
         Text            =   " "
         Top             =   480
         Width           =   3735
      End
      Begin VB.CommandButton Limpia 
         Caption         =   "  Limpia Pantalla"
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
         Left            =   8520
         TabIndex        =   25
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Baja 
         Caption         =   "  Limpia Etiquetas"
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
         Left            =   8520
         TabIndex        =   24
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Tara 
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
         Left            =   5640
         MaxLength       =   6
         TabIndex        =   23
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ComboBox Tipo 
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
         Left            =   4320
         TabIndex        =   21
         Text            =   " "
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox Descripcion 
         Enabled         =   0   'False
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
         MaxLength       =   25
         TabIndex        =   19
         Text            =   " "
         Top             =   1920
         Width           =   5775
      End
      Begin VB.TextBox Etiquetas 
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
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   16
         Text            =   " "
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Cantidad 
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
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   15
         Text            =   " "
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Lote 
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
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   1
         Text            =   "  "
         Top             =   480
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Terminado 
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   327680
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   960
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
         Left            =   8520
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
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
         Left            =   8520
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Tara "
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
         Left            =   5040
         TabIndex        =   22
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label DesProducto 
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
         Left            =   4200
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label Label6 
         Caption         =   "Descripcion"
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
         TabIndex        =   18
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label DesCliente 
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
         Left            =   3720
         TabIndex        =   17
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad de Etiquetas"
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
         TabIndex        =   13
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
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
         TabIndex        =   12
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label1 
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
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Lote"
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
         TabIndex        =   10
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Producto Terminado"
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
         TabIndex        =   9
         Top             =   1440
         Width           =   2055
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9600
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "eti1.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8640
      TabIndex        =   5
      Top             =   4800
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
      Height          =   2460
      ItemData        =   "eti3auto.frx":0000
      Left            =   240
      List            =   "eti3auto.frx":0007
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   7320
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   7200
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgEti3Auto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WLote As String
Private WCantidad As String
Private WImpreadi As String
Private WClase As String
Private WRiesgo As String
Private WIntervencion As String
Private WNaciones As String
Private WEmbalaje As String
Private WDescriUno As String
Private WTipoeti As String
Private WObservaciones As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim XParam As String
Dim Da As Integer
Dim WTerminado(100) As String
Dim LugarTerminado As Integer
Dim WConservacion As String
Dim WElaboracion As String
Dim WVencimento As String
Dim WVida As Single
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer
Dim DiaFeriado(100) As String
Dim XMes As String
Dim XAno As String
Dim WMes As Single
Dim WAno As Single
Dim WDirentrega As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
Dim WImpreVto As Integer
Dim ZZImpreVtoTermi As Integer

Dim ZFechaVto As String
Dim ZVto As String
Dim CargaEmpresa(12, 2) As String

Private Sub Acepta_Click()

    On Error GoTo WError
    
    spHoja = "ListaHoja " + "'" + Lote.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        If rstHoja!Real = 0 And Cliente.Text <> "" And rstHoja!Producto = Terminado.Text Then
            If Val(XEmpresa) = 1 Then
                rstHoja.Close
                m$ = "El lote informado no esta aprobado por laboratorio"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Exit Sub
            End If
            WMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
            If WMarcaVencida = "S" Then
                rstHoja.Close
                m$ = "El lote informado se encuentra vencido"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Exit Sub
            End If
        End If
        rstHoja.Close
    End If
    
    If Left$(Terminado.Text, 2) <> "PT" Then
        m$ = "Solo se puede emitir etiqietas a productos PT"
        G% = MsgBox(m$, 0, "Impresion de Etiquetas")
        Terminado.Text = "  -     -   "
        Lote.SetFocus
        Exit Sub
    End If
    
    WMarcaVencida = "S"
    WEntra = "N"
    
    XParam = "'" + Lote.Text + "','" _
            + Terminado.Text + "'"
    spHoja = "ListaHojaProducto " + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        WEntra = "S"
        WMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
        rstHoja.Close
    End If

    If WEntra = "N" Then
        XParam = "'" + Terminado.Text + "','" _
                + Lote.Text + "'"
        spMovguia = "ListaMovguiaLote1 " + XParam
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
            WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
            rstMovguia.Close
        End If
    End If
    
    If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
        m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
             "Por favor comuniquese con el laboratorio para su revalida"
        G% = MsgBox(m$, 0, "Impresion de Etiquetas")
        Terminado.Text = "  -     -   "
        Lote.SetFocus
        Exit Sub
    End If
    
    Rem Listado.DataFiles(0) = WEmpresa + "coti.mdb"
    Rem Listado.DataFiles(1) = ""
    Rem Listado.DataFiles(0) = WEmpresa + "VENT.mdb"
    Rem Listado.DataFiles(2) = WEmpresa + "admi.mdb"
    
    Salida = "N"
    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Impresion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    If Salida <> "S" Then
            
    XEmpresa = Wempresa
        
    Select Case Val(XEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2, 4, 8, 9
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    
    WVida = 0
    WLinea = 0
    ZZImpreVtoTermi = 0
                
    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        DesProducto.Caption = rstTerminado!Descripcion
        WLinea = rstTerminado!Linea
        ZZImpreVtoTermi = IIf(IsNull(rstTerminado!ImpreVto), "0", rstTerminado!ImpreVto)
        If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 9 Then
            Descripcion.Text = rstTerminado!Descripcion
        End If
        
        WImpreadi = ""
        WClase = ""
        WRiesgo = ""
        WIntervencion = ""
        WNaciones = ""
        WEmbalaje = ""
        wdescriOnu = ""
        
        WImpreadi = IIf(IsNull(rstTerminado!Impreadi), "", rstTerminado!Impreadi)
        WClase = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
        WRiesgo = IIf(IsNull(rstTerminado!Riesgo), "", rstTerminado!Riesgo)
        WIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
        WNaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
        WEmbalaje = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
        wdescriOnu = IIf(IsNull(rstTerminado!Descrionu), "", rstTerminado!Descrionu)
        
        WTipoeti = IIf(IsNull(rstTerminado!TipoEti), "", rstTerminado!TipoEti)
        WObservaciones = IIf(IsNull(rstTerminado!observaciones), "", rstTerminado!observaciones)
        WObservaciones = ""
        WVida = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
        WConservacion = IIf(IsNull(rstTerminado!Conservacion), "", rstTerminado!Conservacion)
        WConservacion = RTrim(WConservacion)
        WConservacionII = IIf(IsNull(rstTerminado!ConservacionII), "", rstTerminado!ConservacionII)
        WConservacionII = RTrim(WConservacionII)
        rstTerminado.Close
    End If
    
    If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 9 Then
        WVida = 0
    End If
        
    spPrecios = "ConsultaPrecios " + "'" + Cliente.Text + Terminado.Text + "'"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        Descripcion.Text = Left$(rstPrecios!Descripcion, 25)
        rstPrecios.Close
    End If
                
    spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then
        DesCliente.Caption = rstClientes!Razon
        rstClientes.Close
    End If
                
    Select Case Val(XEmpresa)
        Case 1
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            Wempresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            Wempresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            Wempresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            Wempresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            Wempresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            Wempresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 11
            Wempresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    
    spHoja = "ListaHoja " + "'" + Lote.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        WMes = Val(Mid$(rstHoja!Fecha, 4, 2))
        WAno = Val(Right$(rstHoja!Fecha, 4))
        
        ZZRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
        ZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
        ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
            
        If Val(ZZRevalida) <> 0 Then
            WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
            WAno = Val(Right$(ZZFechaRevalida, 4))
            WVida = Val(ZZMesesRevalida)
        End If
        
        For Ciclo = 1 To WVida
            WMes = WMes + 1
            If WMes > 12 Then
                WAno = WAno + 1
                WMes = 1
            End If
        Next Ciclo
        WElaboracion = rstHoja!Fecha
        Rem XFec1 = WElaboracion
        Rem SumaDia = WVida + 1
        Rem Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
        If WVida <> 0 Then
            XMes = Str$(WMes)
            XAno = Str$(WAno)
            Call Ceros(XMes, 2)
            Call Ceros(XAno, 4)
            Wvencimiento = "01/" + XMes + "/" + XAno
        End If
        rstHoja.Close
    End If
    
    
    
    ZZRenglon = 0
    ZZTipo = ""
    ZZTerminado = ""
    ZZArticulo = ""
    ZZCantidad = 0
    ZZCantidadLote = 0
    ZZLote = ""
    spHoja = "ListaHoja " + "'" + Lote.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
                    ZZRenglon = ZZRenglon + 1
                    ZZTipo = rstHoja!Tipo
                    ZZTerminado = rstHoja!Terminado
                    ZZArticulo = rstHoja!Articulo
                    ZZCantidad = rstHoja!Cantidad
                    ZZCantidadLote = rstHoja!Canti1
                    ZZLote = rstHoja!lote1
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If
    
    TipoPro = "PT"
    XCodigo = Val(Mid$(Terminado.Text, 4, 5))
    If Left$(Terminado.Text, 2) <> "PT" Then
        Select Case Left$(Terminado.Text, 2)
            Case "DY", "DS"
                TipoPro = "CO"
            Case "QC"
                TipoPro = "FA"
            Case Else
                TipoPro = "PT"
        End Select
            Else
        If XCodigo >= 0 And XCodigo <= 999 Then
            TipoPro = "CO"
                Else
            If XCodigo >= 11000 And XCodigo <= 12999 Then
                TipoPro = "CO"
                    Else
                If XCodigo >= 25000 And XCodigo <= 25999 Then
                    TipoPro = "FA"
                        Else
                    If XCodigo >= 2300 And XCodigo <= 2399 Then
                        TipoPro = "BI"
                            Else
                        TipoPro = "PT"
                    End If
                End If
            End If
        End If
    End If
    
    If Left$(Terminado.Text, 2) = "YQ" Then
        TipoPro = "PT"
    End If
    If Left$(Terminado.Text, 2) = "YH" Then
        TipoPro = "PT"
    End If
    If Left$(Terminado.Text, 2) = "YP" Then
        TipoPro = "PT"
    End If
    If Left$(Terminado.Text, 2) = "YF" Then
        TipoPro = "FA"
    End If
    
    XCodigo = Val(Mid$(Terminado.Text, 4, 5))
    If (XCodigo >= 25000 And XCodigo <= 25999) Or WLinea = 10 Or WLinea = 20 Or WLinea = 22 Or WLinea = 24 Or WLinea = 25 Or WLinea = 26 Or WLinea = 27 Or WLinea = 28 Or WLinea = 29 Or WLinea = 30 Then
        If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
            TipoPro = "FA"
        End If
    End If
    
    Rem If Tipopro <> "FA" Then
    Rem     If ZZRenglon = 1 And ZZCantidad = ZZCantidadLote And ZZTipo = "M" Then
    Rem
    Rem         Select Case Val(WEmpresa)
    Rem             Case 1, 3, 5, 6, 7, 10, 11
    Rem                 CargaEmpresa(1, 1) = "0001"
    Rem                 CargaEmpresa(1, 2) = "Empresa01"
    Rem                 CargaEmpresa(2, 1) = "0003"
    Rem                 CargaEmpresa(2, 2) = "Empresa03"
    Rem                 CargaEmpresa(3, 1) = "0005"
    Rem                 CargaEmpresa(3, 2) = "Empresa05"
    Rem                 CargaEmpresa(4, 1) = "0006"
    Rem                 CargaEmpresa(4, 2) = "Empresa06"
    Rem                 CargaEmpresa(5, 1) = "0007"
    Rem                 CargaEmpresa(5, 2) = "Empresa07"
    Rem                 CargaEmpresa(6, 1) = "0010"
    Rem                 CargaEmpresa(6, 2) = "Empresa10"
    Rem                 CargaEmpresa(7, 1) = "0011"
    Rem                 CargaEmpresa(7, 2) = "Empresa11"
    Rem                 ZHasta = 7
    Rem             Case Else
    Rem                 CargaEmpresa(1, 1) = "0002"
    Rem                 CargaEmpresa(1, 2) = "Empresa02"
    Rem                 CargaEmpresa(2, 1) = "0004"
    Rem                 CargaEmpresa(2, 2) = "Empresa04"
    Rem                 CargaEmpresa(3, 1) = "0008"
    Rem                 CargaEmpresa(3, 2) = "Empresa08"
    Rem                 CargaEmpresa(4, 1) = "0009"
    Rem                 CargaEmpresa(4, 2) = "Empresa09"
    Rem                 ZHasta = 4
    Rem         End Select
    Rem
    Rem         ZVto = ""
    Rem         ZLaudo = ZZLote
    Rem         ZArticulo = ZZArticulo
    Rem         ZFecha = ""
    Rem         ZFechaVto = ""
    Rem
    Rem         For ZCiclo = 1 To ZHasta
    Rem
    Rem             WEmpresa = CargaEmpresa(ZCiclo, 1)
    Rem             txtOdbc = CargaEmpresa(ZCiclo, 2)
    Rem             strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Rem             Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Rem
    Rem             ZSql = ""
    Rem             ZSql = ZSql + "Select *"
    Rem             ZSql = ZSql + " FROM Laudo"
    Rem             ZSql = ZSql + " Where Laudo = " + "'" + ZLaudo + "'"
    Rem             ZSql = ZSql + " and Articulo = " + "'" + ZArticulo + "'"
    Rem             spLaudo = ZSql
    Rem             Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    Rem             If rstLaudo.RecordCount > 0 Then
    Rem                 ZFecha = rstLaudo!Fecha
     Rem                ZFechaVto = IIf(IsNull(rstLaudo!FechaVencimiento), "", rstLaudo!FechaVencimiento)
    Rem                 rstLaudo.Close
    Rem                 Exit For
    Rem             End If
    Rem
    Rem         Next ZCiclo
    Rem
    Rem         Call Conecta_Empresa
    Rem
    Rem         ZVto = ""
    Rem         ZOrdFecha = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
    Rem         If ZFechaVto <> "" And ZFechaVto <> "  /  /    " And ZFechaVto <> "00/00/0000" Then
    Rem             Call Valida_fecha(ZFechaVto, Auxi)
    Rem             If Auxi = "S" Then
    Rem                 ZVto = ZFechaVto
    Rem             End If
    Rem         End If
    Rem
    Rem         If ZVto = "" Then
    Rem
    Rem             ZMeses = 0
    Rem             ZSql = ""
    Rem             ZSql = ZSql + "Select *"
    Rem             ZSql = ZSql + " FROM Articulo"
    Rem             ZSql = ZSql + " Where Codigo = " + "'" + ZArticulo + "'"
    Rem             spArticulo = ZSql
    Rem             Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    Rem             If rstArticulo.RecordCount > 0 Then
    Rem                 ZMeses = rstArticulo!Meses
    Rem                 rstArticulo.Close
    Rem             End If
    Rem
    Rem             WMes = Val(Mid$(ZFecha, 4, 2))
    Rem             WAno = Val(Right$(ZFecha, 4))
    Rem             For ZCiclo = 1 To ZMeses
    Rem                 WMes = WMes + 1
    Rem                 If WMes > 12 Then
    Rem                     WAno = WAno + 1
    Rem                     WMes = 1
    Rem                 End If
    Rem             Next ZCiclo
    Rem
    Rem             XMes = Str$(WMes)
    Rem             XAno = Str$(WAno)
    Rem             Call Ceros(XMes, 2)
    Rem             Call Ceros(XAno, 4)
    Rem             If Val(Left$(ZFecha, 2)) <= 30 Then
    Rem                 If Val(XMes) = 2 And Val(Left$(ZFecha, 2)) > 28 Then
    Rem                     ZVto = "28/" + XMes + "/" + XAno
    Rem                         Else
    Rem                     ZVto = Left$(ZFecha, 3) + XMes + "/" + XAno
    Rem                 End If
    Rem                     Else
    Rem                 If Val(XMes) = 2 Then
    Rem                     ZVto = "28/" + XMes + "/" + XAno
    Rem                         Else
    Rem                     ZVto = "30/" + XMes + "/" + XAno
    Rem                 End If
    Rem             End If
    Rem
    Rem         End If
    Rem
    Rem         Wvencimiento = ZVto
    Rem
    Rem     End If
    Rem End If
    
    
    
    
    
    
    
    
    
    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                     Exit Do
                End If
            Loop
        End If
    End With
    
    WTara = Val(Tara.Text)
    WNeto = Val(Cantidad.Text)
    
    If WTara = 0 Then
        WBruto = 0
            Else
        WBruto = WTara + WNeto
    End If
    
    WRazon = ""
    Rem WDirEntrega = ""
            
    XEmpresa = Wempresa
    Select Case Val(XEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2, 4, 8, 9
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
    End Select
            
    WImpreVto = 0
    spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then
        WRazon = rstClientes!Razon
        WImpreVto = IIf(IsNull(rstClientes!ImpreVto), "0", rstClientes!ImpreVto)
        Rem WDirEntrega = rstClientes!DirEntrega
        rstClientes.Close
    End If
    
    ZVencimiento = Wvencimiento
    If (XCodigo >= 25000 And XCodigo <= 25999) Or WLinea = 10 Or WLinea = 20 Or WLinea = 22 Or WLinea = 24 Or WLinea = 25 Or WLinea = 26 Or WLinea = 27 Or WLinea = 28 Or WLinea = 29 Or WLinea = 30 Then
        Rem no hago nada
            Else
        If ZZImpreVtoTermi = 0 Then
            If WImpreVto = 0 Then
                Rem ZVencimiento = ""
            End If
        End If
    End If
            
    Select Case Val(XEmpresa)
        Case 1
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            Wempresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            Wempresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            Wempresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            Wempresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            Wempresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            Wempresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 11
            Wempresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    
    If TipoPro <> "FA" Then
        Da = 0
        If Len(Descripcion.Text) > 16 Then
            For Da = 17 To 1 Step -1
                If Mid$(Descripcion.Text, Da, 1) = Space$(1) Then
                    ZZNombre = Mid$(Descripcion.Text, 1, Da)
                    ZZNombreII = Mid$(Descripcion.Text, Da + 1, 100)
                    Exit For
                End If
            Next Da
                Else
            ZZNombre = Descripcion.Text
            ZZNombreII = ""
        End If
        If TipoProceso.ListIndex > 0 Then
            ZZNombre = Descripcion.Text
            ZZNombreII = ""
        End If
            Else
        ZZNombre = Descripcion.Text
        ZZNombreII = ""
    End If
    
    If Tipo.ListIndex = 2 Then
        If Len(Descripcion.Text) > 20 Then
            For Da = 21 To 1 Step -1
                If Mid$(Descripcion.Text, Da, 1) = Space$(1) Then
                    ZZNombre = Mid$(Descripcion.Text, 1, Da)
                    ZZNombreII = Mid$(Descripcion.Text, Da + 1, 100)
                    Exit For
                End If
            Next Da
                Else
            ZZNombre = Descripcion.Text
            ZZNombreII = ""
        End If
    End If
    
    Da = 0
    ZRazon = ""
    ZRazonII = ""
    If Len(WRazon) > 27 Then
        For Da = 27 To 1 Step -1
            If Mid$(WRazon, Da, 1) = Space$(1) Then
                ZRazon = Mid$(WRazon, 1, Da)
                ZRazonII = Mid$(WRazon, Da + 1, 100)
                Exit For
            End If
        Next Da
            Else
        ZRazon = WRazon
        ZRazonII = ""
    End If
    
    If Val(WProv) = 24 Then
        If Idioma.ListIndex = 0 Then
            WObservaciones = "Hecho en Argentina"
                Else
            WObservaciones = "Made in Argentina"
        End If
    End If
        
    
    
    With rstEtiqueta
        For Da = 1 To Val(Etiquetas)
            .Index = "Codigo"
            .AddNew
            !Codigo = Da
            WLote = Lote.Text
            Call Ceros(WLote, 6)
            WCantidad = Cantidad.Text
            Call Ceros(WCantidad, 4)
            !Terminado = Terminado.Text
            !Lote = Val(Lote.Text)
            !Cliente = Cliente.Text
            !Cantidad = Val(Cantidad.Text)
            !Nombre = Left$(ZZNombre, 30)
            !NombreII = Left$(ZZNombreII, 30)
            !Impre1 = Mid$(Terminado.Text, 4, 5) + Right$(Terminado.Text, 3) + " " + WLote
            !Razon = ZRazon
            !DirEntrega = ZRazonII
            !Clase = WRiesgo
            !Intervencion = WIntervencion
            !Naciones = WNaciones
            !Embalaje = WEmbalaje
            !Descrionu = wdescriOnu
            !Bruto = WBruto
            !Tara = WTara
            !Neto = WNeto
            !observaciones = Left$(WObservaciones, 20)
            !Elaboracion = Right$(WElaboracion, 7)
            !Vencimiento = Right$(ZVencimiento, 7)
            !Conservacion = WConservacion
            !ConservacionII = WConservacionII
            !TipoPro = ""
            If Trim(Cliente.Text) = "" Then
                !TipoPro = Left$(Terminado.Text, 2)
            End If
            
            !NombreFarmaI = "MANTENER EN ENVASE ORIGINAL CERRADO ENTRE 5 Y 35Cº"
            !NombreFarmaII = ""
            
            ZFazon = "N"
            Select Case Val(WLinea)
                Case 3, 4, 5, 7, 8, 9, 11, 12, 14, 19, 22
                    ZFazon = "N"
                Case 6, 16, 17
                    ZFazon = "N"
                Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                    ZFazon = "N"
                Case Else
                    ZFazon = "S"
            End Select
            If TipoPro = "CO" Then
                !NombreFarmaI = ""
                !NombreFarmaII = ""
            End If
            If TipoPro = "FA" Then
                !NombreFarmaI = ""
                !NombreFarmaII = ""
            End If
            If ZFazon = "S" Then
                !NombreFarmaI = ""
                !NombreFarmaII = ""
            End If
            
            If Tipo.ListIndex = 2 Then
                !Impre1 = Mid$(Terminado.Text, 4, 5) + Right$(Terminado.Text, 3)
                !Impre2 = WLote
            End If
            
            
            .Update
        Next Da
    End With

    Listado.WindowTitle = "Emision de Etiquetas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Da = 0
    Rem If Len(Descripcion.Text) > 16 Then
    Rem     For Da = 25 To 1 Step -1
    Rem        If Mid$(Descripcion.Text, Da, 1) <> Space$(1) Then
    Rem             Exit For
    Rem         End If
    Rem     Next Da
    Rem End If
    Da = Len(Trim(Descripcion.Text))
    
    If Tipo.ListIndex = 0 Then
        Select Case Val(XEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                If TipoProceso.ListIndex = 0 Then
                
                    If WImpreadi = "S" Then
                        m$ = " Coloque la etiqueta de producto Peligroso Clase Nro.: " + WClase
                        G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                    End If
                    Listado.ReportFileName = "WEti1Nuevo.rpt"
                    
                        Else
                        
                    If WImpreadi <> "S" Then
                        If Da > 20 Then
                            Listado.ReportFileName = "eti10.rpt"
                                Else
                            Listado.ReportFileName = "eti1.rpt"
                        End If
                            Else
                        m$ = " Coloque la etiqueta que en su margen tengo el Numero " + WTipoeti
                        G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                        If Da > 20 Then
                            Listado.ReportFileName = "eti110.rpt"
                                Else
                            Listado.ReportFileName = "eti101.rpt"
                        End If
                    End If
                    
                End If
                
            Case Else
                If WImpreadi <> "S" Then
                    If Da > 20 Then
                        Listado.ReportFileName = "eti10Pellital.rpt"
                            Else
                        Listado.ReportFileName = "eti1Pellital.rpt"
                    End If
                        Else
                    m$ = " Coloque la etiqueta que en su margen tengo el Numero " + WTipoeti
                    G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                    If Da > 20 Then
                        Listado.ReportFileName = "eti110Pellital.rpt"
                            Else
                        Listado.ReportFileName = "eti101Pellital.rpt"
                    End If
                End If
                
        End Select
        
            Else
            
        If TipoProceso.ListIndex = 0 Then
        
            If WImpreadi = "S" Then
                m$ = " Coloque la etiqueta de producto Peligroso Clase Nro.: " + WClase
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            End If
            If Da > 23 Then
                Listado.ReportFileName = "weti20Nuevo.rpt"
                    Else
                If Da > 16 Then
                    Listado.ReportFileName = "weti30Nuevo.rpt"
                        Else
                    Listado.ReportFileName = "weti2Nuevo.rpt"
                End If
            End If
            
                Else
                
            If WImpreadi <> "S" Then
                If Da > 20 Then
                    Listado.ReportFileName = "weti20.rpt"
                        Else
                    Listado.ReportFileName = "weti2.rpt"
                End If
                    Else
                Rem m$ = "Producto Peligrosos no se pueden imprimir en etiquetas Chicas"
                Rem G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                If Da > 20 Then
                    Listado.ReportFileName = "weti20.rpt"
                        Else
                    Listado.ReportFileName = "weti2.rpt"
                End If
            End If
            
        End If
        
    End If
    
    Rem If WVida <> 0 Then
    Rem     WFechaActual = "01" + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Rem     WFechaActualOrd = Right$(WFechaActual, 4) + Mid$(WFechaActual, 4, 2) + Left$(WFechaActual, 2)
    Rem
    Rem     WFechaVencimiento = "01" + Mid$(WVencimiento, 3, 10)
    Rem     WFechaVencimientoOrd = Right$(WFechaVencimiento, 4) + Mid$(WFechaVencimiento, 4, 2) + Left$(WFechaVencimiento, 2)
    Rem
    Rem     Pasa = "S"
    Rem     If WFechaActualOrd >= WFechaVencimientoOrd Then
    Rem         Pasa = "N"
    Rem             Else
    Rem         Meses = 0
    Rem         WMes = Val(Mid$(WFechaActual, 4, 2))
    Rem         WAno = Val(Right$(WFechaActual, 4))
    Rem         Do
    Rem             Meses = Meses + 1
    Rem             WMes = WMes + 1
    Rem             If WMes > 12 Then
    Rem                 WAno = WAno + 1
    Rem                 WMes = 1
    Rem             End If
    Rem             XMes = Str$(WMes)
    Rem             XAno = Str$(WAno)
    Rem             Call Ceros(XMes, 2)
    Rem             Call Ceros(XAno, 4)
    Rem             WCompara = "01/" + XMes + "/" + XAno
    Rem             If WCompara = WFechaVencimiento Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem
    Rem         ZMeses = Int(WVida / 2)
    Rem         If ZMeses > 12 Then
    Rem             ZMeses = 12
    Rem         End If
    Rem
    Rem         If Meses <= ZMeses Then
    Rem             Pasa = "N"
    Rem         End If
    Rem
    Rem     End If
    Rem
    Rem     If Pasa = "N" Then
    Rem         m$ = "EL Producto tiene menos de un año de vida util"
    Rem         G% = MsgBox(m$, 0, "Impresion de Etiquetas")
    Rem         Da = 0
    Rem         With rstEtiqueta
    Rem             .Index = "Codigo"
    Rem             .Seek ">=", Da
    Rem             If .NoMatch = False Then
    Rem                 Do
    Rem                     .Delete
    Rem                     .MoveNext
    Rem                     If .EOF = True Then
    Rem                         Exit Do
    Rem                     End If
    Rem                 Loop
    Rem             End If
    Rem         End With
    Rem         Exit Sub
    Rem     End If
    Rem
    Rem End If
    
    XCodigo = Val(Mid$(Terminado.Text, 4, 5))
    If (XCodigo >= 25000 And XCodigo <= 25999) Or WLinea = 10 Or WLinea = 20 Or WLinea = 22 Or WLinea = 24 Or WLinea = 25 Or WLinea = 26 Or WLinea = 27 Or WLinea = 28 Or WLinea = 29 Or WLinea = 30 Then
        If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
            m$ = "Coloque la etiqueta correspondirentes a los productos de Farma"
            G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            If Da > 20 Then
                Listado.ReportFileName = "EtiFarmaII.rpt"
                    Else
                Listado.ReportFileName = "EtiFarma.rpt"
            End If
        End If
    End If
    
    If Tipo.ListIndex = 2 Then
        If Trim(WClase) <> "" Then
            m$ = " Coloque la etiqueta de producto Peligroso Clase Nro.: " + WClase
            G% = MsgBox(m$, 0, "Impresion de Etiquetas")
        End If
        Listado.ReportFileName = "EtiquetaInteriorQuimicos.rpt"
    End If
    
    
    Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
   
    Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
    Rem Listado.Connect = Connect()
    
    If Val(Wempresa) = 2 Or Val(Wempresa) = 4 Then
        Listado.ReportFileName = "WEtiBlanco.rpt"
    End If
    
    If Tipo.ListIndex = 2 Then
        Listado.ReportFileName = "EtiquetaInteriorQuimicos.rpt"
    End If
    
    Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
    Listado.DataFiles(1) = ""
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    Listado.PrinterCopies = 1
    Listado.Action = 1
    
    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Cancela_Click()
    With rstEmpresa
        .Close
    End With
    PrgEti3Auto.Hide
    Unload Me
    PrgHoja.Show
End Sub

Private Sub Baja_Click()
    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
End Sub


Sub Form_Load()

    XEmpresa = Wempresa
    
    TipoProceso.Clear
    
    TipoProceso.AddItem "Etiqueta Nueva"
    TipoProceso.AddItem "Etiqueta Anterior"
    
    TipoProceso.ListIndex = 0
    
    Tipo.Clear
    
    Tipo.AddItem "Grande"
    Tipo.AddItem "Chica"
    Tipo.AddItem "Etiqueta Autoadhesivas"
    
    Select Case Val(XEmpresa)
        Case 5
            Tipo.ListIndex = 1
        Case Else
            Tipo.ListIndex = 0
    End Select

    Cliente.Text = ""
    Terminado.Text = "  -     -   "
    Lote.Text = ""
    Descripcion.Text = ""
    Cantidad.Text = ""
    Etiquetas.Text = ""
    Tara.Text = ""
    
    DesCliente.Caption = ""
    DesProducto.Caption = ""
    
    Lote.Text = PLote
    Call Lote_keypress(13)
    
End Sub


Sub Limpia_Click()

    Tipo.Clear
    
    Tipo.AddItem "Grande"
    Tipo.AddItem "Chica"
    
    Select Case Val(XEmpresa)
        Case 5
            Tipo.ListIndex = 1
        Case Else
            Tipo.ListIndex = 0
    End Select

    Cliente.Text = ""
    Terminado.Text = "  -     -   "
    Lote.Text = ""
    Descripcion.Text = ""
    Cantidad.Text = ""
    Etiquetas.Text = ""
    Tara.Text = ""
    
    DesCliente.Caption = ""
    DesProducto.Caption = ""
    
    Lote.SetFocus
    
End Sub


Private Sub Lote_keypress(KeyAscii As Integer)

    On Error GoTo WError

    If KeyAscii = 13 Then
    
    LugarTerminado = 0
    
    Terminado.Text = "  -     -   "
    Ingresa = "N"
    
    spHoja = "ListaHoja " + "'" + Lote.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        LugarTerminado = LugarTerminado + 1
        WTerminado(LugarTerminado) = UCase(rstHoja!Producto)
        rstHoja.Close
    End If
        
    spMovguia = "ListaMovguiaLoteSolo " + "'" + Lote.Text + "'"
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
        With rstMovguia
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaTerminado = "S"
                    For Ciclo = 1 To LugarTerminado
                        If WTerminado(Ciclo) = rstMovguia!Terminado Then
                            IngresaTerminado = "N"
                            Exit For
                        End If
                    Next Ciclo
                    If IngresaTerminado = "S" Then
                        LugarTerminado = LugarTerminado + 1
                        WTerminado(LugarTerminado) = rstMovguia!Terminado
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovguia.Close
    End If
    
    If LugarTerminado = 1 Then
        Terminado.Text = WTerminado(1)
        Ingresa = "S"
    End If
        
    If LugarTerminado > 1 Then
        Call Elije_Lote
    End If
        
    If Ingresa = "S" Then
        If Left$(Terminado.Text, 2) <> "PT" Then
            m$ = "Solo se puede emitir etiqietas a productos PT"
            G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            Terminado.Text = "  -     -   "
            Lote.SetFocus
            Exit Sub
        End If
    End If
        
    If Ingresa = "N" Then
        Lote.SetFocus
            Else
        Call Ejecuta_Lote
    End If
    
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Ejecuta_Lote()

    On Error GoTo WError

    XEmpresa = Wempresa
    Select Case Val(XEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2, 4, 8, 9
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
                
    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        DesProducto.Caption = rstTerminado!Descripcion
        If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 9 Then
            Descripcion.Text = rstTerminado!Descripcion
        End If
        
        WImpreadi = ""
        WClase = ""
        WRiesgo = ""
        WIntervencion = ""
        WNaciones = ""
        WEmbalaje = ""
        wdescriOnu = ""
        
        WImpreadi = IIf(IsNull(rstTerminado!Impreadi), "", rstTerminado!Impreadi)
        WClase = rstTerminado!Clase
        WRiesgo = IIf(IsNull(rstTerminado!Riesgo), "", rstTerminado!Riesgo)
        WIntervencion = rstTerminado!Intervencion
        WNaciones = rstTerminado!Naciones
        WEmbalaje = rstTerminado!Embalaje
        wdescriOnu = IIf(IsNull(rstTerminado!Descrionu), "", rstTerminado!Descrionu)
        
        WTipoeti = IIf(IsNull(rstTerminado!TipoEti), "", rstTerminado!TipoEti)
        WObservaciones = IIf(IsNull(rstTerminado!observaciones), "", rstTerminado!observaciones)
        WObservaciones = ""
        rstTerminado.Close
    End If
                
    spPrecios = "ConsultaPrecios " + "'" + Cliente.Text + Terminado.Text + "'"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        Descripcion.Text = Left$(rstPrecios!Descripcion, 25)
        rstPrecios.Close
    End If
                
    Select Case Val(XEmpresa)
        Case 1
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            Wempresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            Wempresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            Wempresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            Wempresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            Wempresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            Wempresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 11
            Wempresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
                
    Cliente.SetFocus
    
    Exit Sub

WError:

    Resume Next
    
End Sub


Private Sub Elije_Lote()

    Pantalla.Clear
    
    For Ciclo = 1 To LugarTerminado
        spTerminado = "ConsultaTerminado " + "'" + WTerminado(Ciclo) + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WDesTerminado = rstTerminado!Descripcion
            rstTerminado.Close
        End If
    
        Pantalla.AddItem WTerminado(Ciclo) + "   " + WDesTerminado
        
    Next Ciclo
    
    Pantalla.Visible = True
    
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Terminado.Text = Left$(Pantalla.Text, 12)
    Call Ejecuta_Lote
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Cliente.Text <> "" Then
        
            XEmpresa = Wempresa
            Select Case Val(XEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2, 4, 8, 9
                    Wempresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        
            spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                Cliente.Text = rstClientes!Cliente
                DesCliente.Caption = rstClientes!Razon
                
                Erase ZDirEntrega
                
                ZDirEntrega(1) = rstClientes!DirEntrega
                ZDirEntrega(2) = Trim(IIf(IsNull(rstClientes!DirEntregaII), "", rstClientes!DirEntregaII))
                ZDirEntrega(3) = Trim(IIf(IsNull(rstClientes!DirEntregaIII), "", rstClientes!DirEntregaIII))
                ZDirEntrega(4) = Trim(IIf(IsNull(rstClientes!DirEntregaIV), "", rstClientes!DirEntregaIV))
                ZDirEntrega(5) = Trim(IIf(IsNull(rstClientes!DirEntregaV), "", rstClientes!DirEntregaV))
                
                WDirentrega = ""
                CantiLugarEntrega = 0
                For CicloDirEntrega = 1 To 5
                    If ZDirEntrega(CicloDirEntrega) <> "" Then
                        WDirentrega = ZDirEntrega(CicloDirEntrega)
                        ZLugarDirEntrega = CicloDirEntrega
                        CantiLugarEntrega = CantiLugarEntrega + 1
                    End If
                Next CicloDirEntrega
                
                If CantiLugarEntrega > 1 Then
                    ListaDirEntrega.Clear
                    For CicloDirEntrega = 1 To 5
                        If ZDirEntrega(CicloDirEntrega) <> "" Then
                            ListaDirEntrega.AddItem ZDirEntrega(CicloDirEntrega)
                        End If
                    Next CicloDirEntrega
                    PantaDirEntrega.Top = 840
                    PantaDirEntrega.Visible = True
                    ListaDirEntrega.SetFocus
                        Else
                    ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
                End If
                
                rstClientes.Close
                
                spPrecios = "ConsultaPrecios " + "'" + Cliente.Text + Terminado.Text + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Descripcion.Text = Left$(rstPrecios!Descripcion, 25)
                    rstPrecios.Close
                    Cantidad.SetFocus
                        Else
                    If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 9 Then
                        Descripcion.Text = DesProducto.Caption
                    End If
                    Cantidad.SetFocus
                End If
            End If
            
            Select Case Val(XEmpresa)
                Case 1
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2
                    Wempresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 3
                    Wempresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 4
                    Wempresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 5
                    Wempresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 6
                    Wempresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 7
                    Wempresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 8
                    Wempresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 9
                    Wempresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 10
                    Wempresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 11
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
            
                Else
                
            If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 9 Then
                Descripcion.Text = DesProducto.Caption
            End If
            Cantidad.SetFocus
            
        End If
    End If
End Sub

Private Sub ListaDirEntrega_Click()
    ZLugarDirEntrega = ListaDirEntrega.ListIndex + 1
    WDirentrega = ZDirEntrega(ZLugarDirEntrega)
    ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
    PantaDirEntrega.Visible = False
    Cantidad.SetFocus
End Sub


Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Lote.SetFocus
    End If
End Sub

Private Sub Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Lote_keypress(13)
        Etiquetas.SetFocus
    End If
End Sub

Private Sub Etiquetas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tara.SetFocus
    End If
End Sub

Private Sub Tara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Etiqueta
End Sub

