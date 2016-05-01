VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMovguia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de E/S de Materia Prima y Productos"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   11835
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11835
   Visible         =   0   'False
   Begin VB.Frame PantaOrden 
      Height          =   1455
      Left            =   3120
      TabIndex        =   38
      Top             =   1800
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox ZSaldo 
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
         Left            =   2760
         MaxLength       =   20
         TabIndex        =   45
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox ZDescontar 
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
         Left            =   2760
         MaxLength       =   20
         TabIndex        =   43
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text2 
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
         Left            =   2760
         MaxLength       =   20
         TabIndex        =   41
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox ZOrden 
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
         Left            =   2760
         MaxLength       =   20
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Cantidad a Descontar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label16 
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label15 
         Caption         =   "Orden de Fabricacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame IngresoTransito 
      Height          =   1095
      Left            =   3120
      TabIndex        =   35
      Top             =   3120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox WTransito 
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
         MaxLength       =   20
         TabIndex        =   37
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Ingreso el Numero de Transito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.ComboBox Destino 
      Height          =   315
      Left            =   8880
      TabIndex        =   32
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   30
      Text            =   " "
      Top             =   480
      Width           =   5055
   End
   Begin VB.ComboBox Tipomov 
      Height          =   315
      Left            =   8880
      TabIndex        =   28
      Text            =   " "
      Top             =   120
      Width           =   2415
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11040
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impreord.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   500
      Left            =   2520
      TabIndex        =   16
      Top             =   6480
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   4680
      TabIndex        =   15
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   5160
      TabIndex        =   14
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Codigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   500
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      Height          =   500
      Left            =   1320
      TabIndex        =   10
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   500
      Left            =   2520
      TabIndex        =   8
      Top             =   5880
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   11175
      Begin VB.TextBox WLote 
         Height          =   285
         Left            =   9960
         MaxLength       =   20
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox WMovi 
         Height          =   285
         Left            =   8760
         MaxLength       =   1
         TabIndex        =   20
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox WCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   19
         Text            =   " "
         Top             =   600
         Width           =   1335
      End
      Begin MSMask.MaskEdBox WTerminado 
         Height          =   285
         Left            =   840
         TabIndex        =   18
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox WTipo 
         Height          =   285
         Left            =   360
         MaxLength       =   1
         TabIndex        =   17
         Text            =   " "
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote"
         Height          =   255
         Left            =   9960
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E/S"
         Height          =   255
         Left            =   8760
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   7440
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   3840
         TabIndex        =   24
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Materia Prima"
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto Terminado"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M/T"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label WDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   3840
         TabIndex        =   6
         Top             =   600
         Width           =   3615
      End
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   500
      Left            =   120
      TabIndex        =   4
      Top             =   6480
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3735
      Left            =   240
      OleObjectBlob   =   "movguia.frx":0000
      TabIndex        =   3
      Top             =   960
      Width           =   11295
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10560
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   2205
      ItemData        =   "movguia.frx":0A12
      Left            =   3840
      List            =   "movguia.frx":0A19
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Destno"
      Height          =   255
      Left            =   7080
      TabIndex        =   31
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Movimiento"
      Height          =   285
      Left            =   7080
      TabIndex        =   27
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro Movimiento"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgMovguia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 10 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Tipo As String
Private Articulo As String
Private Terminado As String
Private WTipomov As String
Private WDestino As String
Private Auxiliar(100, 10) As String
Private XAuxiliar(100, 10) As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstCargaProyeccion As Recordset
Dim spCargaProyeccion As String
Dim XParam As String
Dim WSalida As String
Dim WControl(100, 2) As String
Dim WVector(100) As String
Dim ZCodigo As String
Dim QSaldo As Double

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    DBGrid1.Col = 3
    DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
    DBGrid1.Col = 5
    DBGrid1.Text = ""
    
    DBGrid1.Col = 6
    DBGrid1.Text = ""
    
    DBGrid1.Col = 7
    DBGrid1.Text = ""
    
    DBGrid1.Col = 8
    DBGrid1.Text = ""
    
    DBGrid1.Col = 9
    DBGrid1.Text = ""
    
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WMovi.Text = ""
    WLote.Text = ""
    WTransito.Text = ""
    ZOrden.Text = ""
    ZSaldo.Text = ""
    ZDescontar.Text = ""
    
    WLinea.Text = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click
    PrgMovguia.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima"

     Opcion.Visible = True
     
 End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    Rem OPEN_FILE_Movguia
    Rem OPEN_FILE_TERMINADO
    Rem OPEN_FILE_Articulo
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
            spTerminado = "ListaTerminado"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstTerminado.RecordCount > 0 Then
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstTerminado!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTerminado.Close
            End If
            
        Case 1
            spArticulo = "ListaArticulo"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstArticulo.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 1 Then
        WLinea.Text = DBGrid1.Row + 1
        WTipo.Text = DBGrid1.Text
            Else
        WTipo.Text = ""
        WLinea.Text = ""
    End If

    DBGrid1.Col = 1
    If Len(DBGrid1.Text) = 12 Then
        WTerminado.Text = DBGrid1.Text
            Else
        WTerminado.Text = "  -     -   "
    End If

    DBGrid1.Col = 2
    If Len(DBGrid1.Text) = 10 Then
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -   -   "
    End If
    
    DBGrid1.Col = 3
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 4
    WCantidad.Text = DBGrid1.Text
    
    DBGrid1.Col = 5
    WMovi.Text = DBGrid1.Text
    
    DBGrid1.Col = 6
    WLote.Text = DBGrid1.Text
    
    DBGrid1.Col = 7
    WTransito.Text = DBGrid1.Text
    
    DBGrid1.Col = 8
    ZOrden.Text = DBGrid1.Text
    
    DBGrid1.Col = 9
    ZDescontar.Text = DBGrid1.Text
    
    WTipo.SetFocus

End Sub

Private Sub Graba_Click()

    Call Valida_fecha(Fecha.Text, Auxi)
    If Auxi <> "S" Then
        m$ = "La fecha de la guia de traslado es incorrecta"
        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
        Exit Sub
    End If
    
    WTipomov = Str$(Tipomov.ListIndex)
    Call Ceros(WTipomov, 2)
    WDestino = Str$(Destino.ListIndex)
    Call Ceros(WDestino, 2)
    
    If Val(WTipomov) <> 0 Then
        Exit Sub
    End If
    
    If Val(WDestino) = 0 Then
        Exit Sub
    End If
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            If Val(WEmpresa) = Val(WDestino) Or Val(WDestino) = 2 Or Val(WDestino) = 4 Or Val(WDestino) = 8 Or Val(WDestino) = 9 Then
                Exit Sub
            End If
        Case Else
            If Val(WEmpresa) = Val(WDestino) Or Val(WDestino) = 1 Or Val(WDestino) = 3 Or Val(WDestino) = 5 Or Val(WDestino) = 6 Or Val(WDestino) = 7 Or Val(WDestino) = 10 Or Val(WDestino) = 11 Then
                Exit Sub
            End If
    End Select
    
    WTipomov = Str$(Tipomov.ListIndex)
    Call Ceros(WTipomov, 2)
    
    Select Case Val(WEmpresa)
        Case 1
            Codigo.Text = Str$(Val(Codigo.Text) + 300000)
        Case 3
            Codigo.Text = Str$(Val(Codigo.Text) + 400000)
        Case 5
            Codigo.Text = Str$(Val(Codigo.Text) + 500000)
        Case 7
            Codigo.Text = Str$(Val(Codigo.Text) + 600000)
        Case 10
            Codigo.Text = Str$(Val(Codigo.Text) + 700000)
        Case 11
            Codigo.Text = Str$(Val(Codigo.Text) + 800000)
        Case Else
    End Select
    
    XParam = "'" + WTipomov + "','" _
                 + Codigo.Text + "'"
    spMovguia = "ListaMovguia " + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
        rstMovguia.Close
        Exit Sub
    End If
    
    Rem *************** BUSCAR SALDO LOTE
    
    
    
    
    
    
    
    
    
    WSalida = "S"
    
    Rem
    Rem VERIFICA SI HAY CONEXCION CON LA OTRA OPLANTA
    Rem
    
    On Error GoTo Control_Error
    XEmpresa = WEmpresa
        
    Select Case Val(WDestino)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 11
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    
    Select Case Val(XEmpresa)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 11
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    
    If WSalida = "N" Then Exit Sub
    
    On Error GoTo 0

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    XRenglon = 0
    Erase XAuxiliar
    DBGrid1.Refresh
                
    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Tipo = DBGrid1.Text
                                       
            DBGrid1.Col = 1
            Terminado = DBGrid1.Text
                    
            DBGrid1.Col = 2
            Articulo = DBGrid1.Text
                    
            DBGrid1.Col = 4
            Cantidad = DBGrid1.Text
                    
            DBGrid1.Col = 5
            Movi = DBGrid1.Text
            
            DBGrid1.Col = 6
            Lote = DBGrid1.Text
            
            DBGrid1.Col = 7
            Transito = DBGrid1.Text
            
            DBGrid1.Col = 8
            Orden = DBGrid1.Text
            
            DBGrid1.Col = 9
            Descontar = DBGrid1.Text
                    
            If Tipo <> "" Then
            
                XRenglon = XRenglon + 1
                
                XAuxiliar(XRenglon, 1) = Tipo
                XAuxiliar(XRenglon, 2) = Terminado
                XAuxiliar(XRenglon, 3) = Articulo
                XAuxiliar(XRenglon, 4) = Cantidad
                XAuxiliar(XRenglon, 5) = Movi
                XAuxiliar(XRenglon, 6) = Lote
                XAuxiliar(XRenglon, 7) = Transito
                XAuxiliar(XRenglon, 8) = Orden
                XAuxiliar(XRenglon, 9) = Descontar

            End If
                
        Next iRow
            
    Next a
    
    Pasa = "S"
    
    For Ciclo1 = 1 To XRenglon
        For Ciclo2 = Ciclo1 + 1 To XRenglon
            If XAuxiliar(Ciclo1, 1) = XAuxiliar(Ciclo2, 1) Then
                If XAuxiliar(Ciclo1, 2) = XAuxiliar(Ciclo2, 2) Then
                    If XAuxiliar(Ciclo1, 3) = XAuxiliar(Ciclo2, 3) Then
                        If XAuxiliar(Ciclo1, 6) = XAuxiliar(Ciclo2, 6) Then
                            Pasa = "N"
                        End If
                    End If
                End If
            End If
        Next Ciclo2
    Next Ciclo1
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    If Pasa = "N" Then
        m$ = "Productos repetidos en la carga de la guia"
        G% = MsgBox(m$, 0, "Guias de Traslado Internos")
        Exit Sub
    End If
    
    
    
    If Destino.ListIndex = 10 Or Destino.ListIndex = 11 Then
    
    
        For a = 1 To 99
            
            ZZTipo = XAuxiliar(a, 1)
            ZZCodigo = XAuxiliar(a, 3)
            ZZCodSedronar = ""
            
            If UCase(ZZTipo) = "M" Then
                If Trim(ZZCodigo) <> "" Then
                
                    spArticulo = "ConsultaArticulo " + "'" + ZZCodigo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        ZZCodSedronar = rstArticulo!CodSedronar
                        rstArticulo.Close
                    End If
                    
                    If Trim(ZZCodSedronar) <> "" Then
                        m$ = "No se puede efectuar el envio la materia prima " + ZZCodigo + " por tener que informarse luego al sedronar"
                        AAAa% = MsgBox(m$, 0, "Carga de Gastos de Importacion")
                        Exit Sub
                    End If
            
                End If
            End If
            
        Next a
    
    End If
    
    
    
    Renglon = 0
    Erase Auxiliar
    DBGrid1.Refresh
    Erase WVector
                
    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Tipo = DBGrid1.Text
                                       
            DBGrid1.Col = 1
            Terminado = DBGrid1.Text
                    
            DBGrid1.Col = 2
            Articulo = DBGrid1.Text
                    
            DBGrid1.Col = 4
            Cantidad = DBGrid1.Text
                    
            DBGrid1.Col = 5
            Movi = DBGrid1.Text
            
            DBGrid1.Col = 6
            Lote = DBGrid1.Text
            PartiOri = ""
            
            DBGrid1.Col = 7
            Transito = DBGrid1.Text
            
            DBGrid1.Col = 8
            Orden = DBGrid1.Text
            
            DBGrid1.Col = 9
            Descontar = DBGrid1.Text
            
            If Tipo = "M" Then
            
            ZArticulo = Left$(Articulo, 2)
            If ZArticulo = "DS" Then
                ZArticuloII = Mid$(Articulo, 4, 3)
                If Val(ZArticuloII) < 100 Then
                    ZArticulo = ""
                End If
            End If
            
            If ZArticulo = "DY" Or ZArticulo = "DS" Then
            
                PartiOri = Lote
                Lote = ""
                WEntra = "N"
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Laudo"
                ZSql = ZSql + " Where Laudo.Articulo = " + "'" + Articulo + "'"
                ZSql = ZSql + " and Laudo.PartiOri = " + "'" + PartiOri + "'"
                ZSql = ZSql + " Order by Laudo.FechaOrd, Laudo.Laudo"
                spLaudo = ZSql
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    With rstLaudo
                        .MoveFirst
                        Lote = IIf(IsNull(rstLaudo!Laudo), "", rstLaudo!Laudo)
                        WEntra = "S"
                        rstLaudo.Close
                    End With
                End If
     
                If WEntra = "N" Then
                
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Guia"
                    ZSql = ZSql + " Where Guia.Articulo = " + "'" + Articulo + "'"
                    ZSql = ZSql + " and Guia.PartiOri = " + "'" + PartiOri + "'"
                    ZSql = ZSql + " Order by Guia.FechaOrd, Guia.Codigo"
                    spMovguia = ZSql
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        With rstMovguia
                            .MoveFirst
                            Lote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                            WEntra = "S"
                            rstMovguia.Close
                        End With
                    End If
                    
                End If
            
            End If
            
            End If
                    
            If Tipo <> "" Then
                    
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Codigo.Text
                Call Ceros(Auxi1, 6)
                
                WTipomov = Str$(Tipomov.ListIndex)
                Call Ceros(WTipomov, 2)
                WDestino = Str$(Destino.ListIndex)
                Call Ceros(WDestino, 2)
                WCodigo = Codigo.Text
                WRenglon = Str$(Renglon)
                WFecha = Fecha.Text
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WTipo = Tipo
                WArticulo = Articulo
                WTerminado = Terminado
                WCantidad = Cantidad
                WMovi = Movi
                WObservaciones = Observaciones.Text
                WClave = Trim(Str$(Val(WTipomov))) + Auxi1 + Auxi
                WDate = Date$
                WMarca = ""
                
                Rem uso la partida
                WPartida = Lote
                WLote = ""
                WSaldo = "0"
                WPartiOri = PartiOri
                WTransito = Transito
                WOrden = Orden
                WDescontar = Descontar
                
                WVector(Renglon) = Lote
                
                Auxiliar(Renglon, 1) = WTipo
                Auxiliar(Renglon, 2) = WTerminado
                Auxiliar(Renglon, 3) = WArticulo
                Auxiliar(Renglon, 4) = WCantidad
                Auxiliar(Renglon, 5) = WMovi
                Auxiliar(Renglon, 6) = WPartida
                Auxiliar(Renglon, 7) = WPartiOri
                Auxiliar(Renglon, 8) = WTransito
                Auxiliar(Renglon, 9) = WOrden
                Auxiliar(Renglon, 10) = WDescontar
           
                 Rem dada
                 Rem by nan buso marcavencida
                 Rem dada
                 Rem XParam = "'" + WPartida + "','" _
                 rem             + Terminado + "'"
                 Rem spHoja = "ListaHojaProducto " + XParam
                 Rem Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                 Rem If rstHoja.RecordCount > 0 Then
                 Rem    Rem WWMarcavencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
                 Rem    rstHoja.Close
                 Rem        Else
                 Rem    XParam = "'" + Terminado + "','" _
                 rem             + Lote + "'"
                 Rem    spMovguia = "ListaMovguiaLote1 " + XParam
                 Rem    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                 Rem    If rstMovguia.RecordCount > 0 Then
                 Rem        Rem WWMarcavencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                 Rem        rstMovguia.Close
                 Rem    End If
                 Rem End If
                 
                 Rem WWMarcavencida = ""
           
                Rem termino de buscar
                Rem aca inserto
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO Guia ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "TipoMov ,"
                ZSql = ZSql + "Codigo ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Tipo ,"
                ZSql = ZSql + "Articulo ,"
                ZSql = ZSql + "Terminado ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "FechaOrd ,"
                ZSql = ZSql + "Movi,"
                ZSql = ZSql + "Observaciones,"
                ZSql = ZSql + "Marca,"
                ZSql = ZSql + "Destino,"
                ZSql = ZSql + "Lote,"
                ZSql = ZSql + "Saldo,"
                ZSql = ZSql + "Partida,"
                ZSql = ZSql + "PartiOri,"
                ZSql = ZSql + "Transito,"
                ZSql = ZSql + "Orden,"
                ZSql = ZSql + "Descontar )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + WTipomov + "',"
                ZSql = ZSql + "'" + WCodigo + "',"
                ZSql = ZSql + "'" + WRenglon + "',"
                ZSql = ZSql + "'" + WFecha + "',"
                ZSql = ZSql + "'" + WTipo + "',"
                ZSql = ZSql + "'" + WArticulo + "',"
                ZSql = ZSql + "'" + WTerminado + "',"
                ZSql = ZSql + "'" + WCantidad + "',"
                ZSql = ZSql + "'" + WFechaord + "',"
                ZSql = ZSql + "'" + WMovi + "',"
                ZSql = ZSql + "'" + WObservaciones + "',"
                ZSql = ZSql + "'" + WMarca + "',"
                ZSql = ZSql + "'" + WDestino + "',"
                ZSql = ZSql + "'" + WLote + "',"
                ZSql = ZSql + "'" + WSaldo + "',"
                ZSql = ZSql + "'" + WPartida + "',"
                ZSql = ZSql + "'" + WPartiOri + "',"
                ZSql = ZSql + "'" + WTransito + "',"
                ZSql = ZSql + "'" + WOrden + "',"
                ZSql = ZSql + "'" + WDescontar + "')"
                
                spMovguia = ZSql
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                
        Next iRow
            
    Next a
                
    For Da = 1 To Renglon
    
        Tipo = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        Articulo = Auxiliar(Da, 3)
        Cantidad = Auxiliar(Da, 4)
        Movi = Auxiliar(Da, 5)
        Lote = Auxiliar(Da, 6)
        PartiOri = Auxiliar(Da, 7)
        Transito = Trim(Auxiliar(Da, 8))
        Orden = Trim(Auxiliar(Da, 9))
        Descontar = Trim(Auxiliar(Da, 10))
        
        If Transito <> "" Then
        
            Sql1 = "UPDATE Laudo SET "
            Sql2 = " SaldoTransito = SaldoTransito - " + "'" + Str$(Val(Cantidad)) + "'"
            Sql3 = " Where Laudo.Articulo = " + "'" + Articulo + "'"
            Sql4 = " and Laudo.Laudo = " + "'" + Lote + "'"
            Sql5 = " and Laudo.Transito = " + "'" + Transito + "'"
            spLaudo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
        If Val(Descontar) <> 0 Then
        
            Sql1 = "UPDATE CargaProyeccion SET "
            Sql2 = " Entregado = Entregado + " + "'" + Str$(Val(Descontar)) + "'"
            Sql3 = " Where CargaProyeccion.Articulo = " + "'" + Terminado + "'"
            Sql4 = " and CargaProyeccion.Orden = " + "'" + Orden + "'"
            spCargaProyeccion = Sql1 + Sql2 + Sql3 + Sql4
            Set rstCargaProyeccion = db.OpenRecordset(spCargaProyeccion, dbOpenSnapshot, dbSQLPassThrough)
            
            Sql1 = "UPDATE CargaProyeccion SET "
            Sql2 = " Saldo = Cantidad - Entregado"
            Sql3 = " Where CargaProyeccion.Articulo = " + "'" + Terminado + "'"
            Sql4 = " and CargaProyeccion.Orden = " + "'" + Orden + "'"
            spCargaProyeccion = Sql1 + Sql2 + Sql3 + Sql4
            Set rstCargaProyeccion = db.OpenRecordset(spCargaProyeccion, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        Select Case Tipo
            Case "M"
                WControla = 0
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
        
                    WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                    If WControla = 2 Then
                        WControla = 0
                    End If
                    WCodigo = Articulo
                    If Movi = "E" Then
                        WEntradas = Str$(rstArticulo!Entradas + Val(Cantidad))
                        WSalidas = Str$(rstArticulo!Salidas)
                            Else
                        WSalidas = Str$(rstArticulo!Salidas + Val(Cantidad))
                        WEntradas = Str$(rstArticulo!Entradas)
                    End If
                    rstArticulo.Close
                    WDate = Date$
                
                    XParam = "'" + WCodigo + "','" _
                            + WEntradas + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
                    spArticulo = "ModificaArticuloMovimientos " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If WControla = 0 And Val(Lote) <> 0 Then
                        XParam = "'" + Lote + "','" _
                                    + Articulo + "'"
                        spLaudo = "ListaLaudoArticulo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            WClave = rstLaudo!Clave
                            If Movi = "E" Then
                                WSaldo = Str$(rstLaudo!Saldo + Val(Cantidad))
                                    Else
                                WSaldo = Str$(rstLaudo!Saldo - Val(Cantidad))
                            End If
                            WDate = Date$
                            rstLaudo.Close
                            
                            XParam = "'" + WClave + "','" _
                                + WDate + "','" _
                                + WSaldo + "'"
                            spLaudo = "ModificaLaudoSaldo " + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                                
                            XParam = "'" + Articulo + "','" _
                                    + Lote + "'"
                            spMovguia = "ListaMovguiaLote " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WClave = rstMovguia!Clave
                                If Movi = "E" Then
                                    WSaldo = Str$(rstMovguia!Saldo + Val(Cantidad))
                                        Else
                                    WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad))
                                End If
                                WDate = Date$
                                rstMovguia.Close
                            
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spMovguia = "ModificaMovguiaSaldo " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
                        End If
                    End If
                End If
                
            Case "T"
                WControla = 0
                spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
        
                    WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                    If WControla = 2 Then
                        WControla = 0
                    End If
                    WCodigo = Terminado
                    If Movi = "E" Then
                        WEntradas = Str$(rstTerminado!Entradas + Val(Cantidad))
                        WSalidas = Str$(rstTerminado!Salidas)
                            Else
                        WSalidas = Str$(rstTerminado!Salidas + Val(Cantidad))
                        WEntradas = Str$(rstTerminado!Entradas)
                    End If
                    WDate = Date$
                    rstTerminado.Close
                
                    XParam = "'" + WCodigo + "','" _
                            + WEntradas + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
                    spTerminado = "ModificaTerminadoMovimientos " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Rem aca puedo tomar marca vencido
                    If WControla = 0 And Val(Lote) <> 0 Then
                        XParam = "'" + Lote + "','" _
                                    + Terminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                            WClave = rstHoja!Clave
                            If Movi = "E" Then
                                WSaldo = Str$(rstHoja!Saldo + Val(Cantidad))
                                    Else
                                WSaldo = Str$(rstHoja!Saldo - Val(Cantidad))
                            End If
                            WDate = Date$
                            rstHoja.Close
                            
                            XParam = "'" + WClave + "','" _
                                + WDate + "','" _
                                + WSaldo + "'"
                            spHoja = "ModificahojaSaldo " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
            
                            XParam = "'" + Terminado + "','" _
                                    + Lote + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WClave = rstMovguia!Clave
                                If Movi = "E" Then
                                    WSaldo = Str$(rstMovguia!Saldo + Val(Cantidad))
                                        Else
                                    WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad))
                                End If
                                WDate = Date$
                                rstMovguia.Close
                            
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spMovguia = "ModificaMovguiaSaldo " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
                        End If
                    End If
                    
                End If
            
            Case Else
        End Select
        
    Next Da
    
    
    Rem If Val(WEmpresa) = 3 Then
    Rem     If Val(WDestino) = 5 Or Val(WDestino) = 6 Then
    Rem         WPasa = "S"
    Rem     End If
    Rem End If
    
    Rem If Val(WEmpresa) = 5 Then
    Rem     If Val(WDestino) = 3 Or Val(WDestino) = 6 Then
    Rem         WPasa = "S"
    Rem     End If
    Rem End If
    
    Rem  If Val(WEmpresa) = 6 Then
    Rem     If Val(WDestino) = 3 Or Val(WDestino) = 5 Then
    Rem         WPasa = "S"
    Rem     End If
    Rem End If
    
    Rem If Val(WEmpresa) = 7 Then
    Rem     If Val(WDestino) = 7 Then
    Rem         WPasa = "S"
    Rem     End If
    Rem End If
    
    Rem dada
    Rem dada
    Rem dada
    Rem dada
    Rem dada
    
    WPasa = "S"
    
    If WPasa = "S" Then
    
        XEmpresa = WEmpresa
        Select Case Val(WDestino)
            Case 1
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                WEmpresa = "0002"
                txtOdbc = "Empresa02"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 5
                WEmpresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 6
                WEmpresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 7
                WEmpresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 8
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 9
                WEmpresa = "0009"
                txtOdbc = "Empresa09"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 10
                WEmpresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 11
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
    
        Renglon = 0
        Erase Auxiliar
        DBGrid1.Refresh
                
        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Tipo = DBGrid1.Text
                                       
                DBGrid1.Col = 1
                Terminado = DBGrid1.Text
                    
                DBGrid1.Col = 2
                Articulo = DBGrid1.Text
                    
                DBGrid1.Col = 4
                Cantidad = DBGrid1.Text
                    
                DBGrid1.Col = 5
                Movi = DBGrid1.Text
                
                DBGrid1.Col = 6
                Lote = DBGrid1.Text
                PartiOri = ""
                
                DBGrid1.Col = 7
                Transito = DBGrid1.Text
                
                If Tipo <> "" Then
                 
                    Renglon = Renglon + 1
                    Auxi = Str$(Renglon)
                    Call Ceros(Auxi, 2)
                    
                    If Tipo = "M" Then
                    
                        ZArticulo = Left$(Articulo, 2)
                        If ZArticulo = "DS" Then
                            ZArticuloII = Mid$(Articulo, 4, 3)
                            If Val(ZArticuloII) < 100 Then
                                ZArticulo = ""
                            End If
                        End If
                    
                        If ZArticulo = "DY" Or ZArticulo = "DS" Then
                            PartiOri = Lote
                            Lote = WVector(Renglon)
                            WEntra = "N"
                        End If
                        
                    End If
                        
                    Auxi1 = Codigo.Text
                    Call Ceros(Auxi1, 6)
                
                    WTipomov = Str$(Val(XEmpresa))
                    Call Ceros(WTipomov, 2)
                    WDestino = "0"
                    Call Ceros(WDestino, 2)
                    WCodigo = Codigo.Text
                    WRenglon = Str$(Renglon)
                    WFecha = Fecha.Text
                    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    WTipo = Tipo
                    WArticulo = Articulo
                    WTerminado = Terminado
                    WCantidad = Cantidad
                    If Movi = "E" Then
                        WMovi = "S"
                            Else
                        WMovi = "E"
                    End If
                    WObservaciones = Observaciones.Text
                    WClave = Trim(Str$(Val(WTipomov))) + Auxi1 + Auxi
                    WDate = Date$
                    WMarca = ""
                    WLote = Lote
                    WSaldo = Cantidad
                    WPartida = ""
                    WPartiOri = PartiOri
                    WTransito = Transito
                    
                    spMovguia = "ConsultaMovguia " + "'" + WClave + "'"
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        rstMovguia.Close
                            Else
                        Select Case WTipo
                            Case "M"
                                WEntra = "N"
        
                                XParam = "'" + WLote + "','" _
                                         + WArticulo + "'"
                                spLaudo = "ListaLaudoArticulo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    WEntra = "S"
                                    XClave = rstLaudo!Clave
                                    WSaldo = Str$(rstLaudo!Saldo + Val(WCantidad))
                                    WDate = Date$
                                    rstLaudo.Close
                            
                                    XParam = "'" + XClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "'"
                                    spLaudo = "ModificaLaudoSaldo " + XParam
                                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                                If WEntra = "N" Then
                                
                                    XParam = "'" + WArticulo + "','" _
                                            + WLote + "'"
                                    spMovguia = "ListaMovguiaLote " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WEntra = "S"
                                        XClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo + Val(WCantidad))
                                        WDate = Date$
                                        rstMovguia.Close
                            
                                        XParam = "'" + XClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                        spMovguia = "ModificaMovguiaSaldo " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    End If
                                    
                                End If
                                
                            Case "T"
                                WEntra = "N"
            
                                XParam = "'" + WLote + "','" _
                                        + WTerminado + "'"
                                spHoja = "ListaHojaProducto " + XParam
                                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                If rstHoja.RecordCount > 0 Then
                                    WEntra = "S"
                                    XClave = rstHoja!Clave
                                    WSaldo = Str$(rstHoja!Saldo + Val(WCantidad))
                                    WDate = Date$
                                    rstHoja.Close
                            
                                    XParam = "'" + XClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "'"
                                    spHoja = "ModificaHojaSaldo " + XParam
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                                If WEntra = "N" Then
                                    XParam = "'" + WTerminado + "','" _
                                            + WLote + "'"
                                    spMovguia = "ListaMovguiaLote1 " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WEntra = "S"
                                        XClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo + Val(WCantidad))
                                        WDate = Date$
                                        rstMovguia.Close
                            
                                        XParam = "'" + XClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                        spMovguia = "ModificaMovguiaSaldo " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    End If
                                End If
                            Case Else
                        End Select
                        
                        If WEntra = "S" Then
                            WSaldo = "0"
                        End If
                        
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Guia ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "TipoMov ,"
                        ZSql = ZSql + "Codigo ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "Tipo ,"
                        ZSql = ZSql + "Articulo ,"
                        ZSql = ZSql + "Terminado ,"
                        ZSql = ZSql + "Cantidad ,"
                        ZSql = ZSql + "FechaOrd ,"
                        ZSql = ZSql + "Movi,"
                        ZSql = ZSql + "Observaciones,"
                        ZSql = ZSql + "Marca,"
                        ZSql = ZSql + "Destino,"
                        ZSql = ZSql + "Lote,"
                        ZSql = ZSql + "Saldo,"
                        ZSql = ZSql + "Partida,"
                        ZSql = ZSql + "PartiOri,"
                        ZSql = ZSql + "Transito)"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + WClave + "',"
                        ZSql = ZSql + "'" + WTipomov + "',"
                        ZSql = ZSql + "'" + WCodigo + "',"
                        ZSql = ZSql + "'" + WRenglon + "',"
                        ZSql = ZSql + "'" + WFecha + "',"
                        ZSql = ZSql + "'" + WTipo + "',"
                        ZSql = ZSql + "'" + WArticulo + "',"
                        ZSql = ZSql + "'" + WTerminado + "',"
                        ZSql = ZSql + "'" + WCantidad + "',"
                        ZSql = ZSql + "'" + WFechaord + "',"
                        ZSql = ZSql + "'" + WMovi + "',"
                        ZSql = ZSql + "'" + WObservaciones + "',"
                        ZSql = ZSql + "'" + WMarca + "',"
                        ZSql = ZSql + "'" + WDestino + "',"
                        ZSql = ZSql + "'" + WLote + "',"
                        ZSql = ZSql + "'" + WSaldo + "',"
                        ZSql = ZSql + "'" + WPartida + "',"
                        ZSql = ZSql + "'" + WPartiOri + "',"
                        ZSql = ZSql + "'" + WTransito + "')"
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        
                        Select Case WTipo
                            Case "M"
                                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstArticulo.RecordCount > 0 Then
                                    WCodigo = WArticulo
                                    If WMovi = "E" Then
                                        WEntradas = Str$(rstArticulo!Entradas + Val(WCantidad))
                                        WSalidas = Str$(rstArticulo!Salidas)
                                                Else
                                        WSalidas = Str$(rstArticulo!Salidas + Val(WCantidad))
                                        WEntradas = Str$(rstArticulo!Entradas)
                                    End If
                                    WDate = Date$
                    
                                    XParam = "'" + WCodigo + "','" _
                                            + WEntradas + "','" _
                                            + WSalidas + "','" _
                                            + WDate + "'"
                                           
                                    spArticulo = "ModificaArticuloMovimientos " + XParam
                                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                            Case "T"
                                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
        
                                    WCodigo = WTerminado
                                    If WMovi = "E" Then
                                        WEntradas = Str$(rstTerminado!Entradas + Val(WCantidad))
                                        WSalidas = Str$(rstTerminado!Salidas)
                                            Else
                                        WSalidas = Str$(rstTerminado!Salidas + Val(WCantidad))
                                        WEntradas = Str$(rstTerminado!Entradas)
                                    End If
                                    WDate = Date$
                
                                    XParam = "'" + WCodigo + "','" _
                                            + WEntradas + "','" _
                                            + WSalidas + "','" _
                                            + WDate + "'"
                                           
                                    spTerminado = "ModificaTerminadoMovimientos " + XParam
                                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                End If
            
                            Case Else
                        End Select
                    End If
                
                End If
                    
            Next iRow
            
        Next a
    
        Select Case Val(XEmpresa)
            Case 1
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                WEmpresa = "0002"
                txtOdbc = "Empresa02"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 5
                WEmpresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 6
                WEmpresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 7
                WEmpresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 8
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 9
                WEmpresa = "0009"
                txtOdbc = "Empresa09"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 10
                WEmpresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 11
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
    
    End If
    
    T$ = "Guias de Traslado Interno"
    m$ = "Desea Imprimir la guia de traslado interno"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Call Impresion
    End If
        
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Codigo.SetFocus
    
    Exit Sub
    
Control_Error:
    MsgBox Err.Description
    WSalida = "N"
    Resume Next
    
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WMovi.Text = ""
    WLote.Text = ""
    WTransito.Text = ""
    ZOrden.Text = ""
    ZSaldo.Text = ""
    ZDescontar.Text = ""
    
    WTipo.SetFocus
    
End Sub

Private Sub Limpia_Click()

    WLinea.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WMovi.Text = ""
    WLote.Text = ""
    WTransito.Text = ""
    ZOrden.Text = ""
    ZSaldo.Text = ""
    ZDescontar.Text = ""

    Codigo.Text = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 6
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    Rem With rstMovguia
    Rem     .Index = "Clave"
    Rem     Claveven$ = "99999999"
    Rem     .Seek "<=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Codigo.Text = !Codigo + 1
    Rem             Else
    Rem         Codigo.Text = ""
    Rem     End If
    Rem End With
    
    Tipomov.ListIndex = 0
    Codigo.Text = ""
    WTipomov = Str$(Tipomov.ListIndex)
    Call Ceros(WTipomov, 2)
    Destino.ListIndex = 0
    
    Rem spMovguia = "ListaMovguiaNumero " + "'" + WTipomov + "'"
    Rem Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstMovguia.RecordCount > 0 Then
    Rem     With rstMovguia
    Rem         .MoveLast
    Rem         Do
    Rem             Codigo.Text = rstMovguia!Codigo + 1
    Rem             If Val(Codigo.Text) > 900000 Then
    Rem                 .MovePrevious
    Rem                     Else
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem     End With
    Rem     rstMovguia.Close
    Rem         Else
    Rem     Codigo.Text = "1"
    Rem End If
    
    DBGrid1.FirstRow = 0
    Renglon = 0

    Graba.Enabled = True


    Codigo.SetFocus

End Sub

Private Sub WTipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WTipo.Text = "M" Or WTipo.Text = "T" Then
            If WTipo.Text = "M" Then
                WArticulo.SetFocus
                    Else
                WTerminado.SetFocus
            End If
                Else
            WTipo.SetFocus
        End If
    End If
End Sub

Private Sub WTerminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WDescripcion.Caption = rstTerminado!Descripcion
            rstTerminado.Close
            WCantidad.SetFocus
                Else
            WTerminado.SetFocus
        End If
    End If
End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
                WDescripcion.Caption = rstArticulo!Descripcion
                WCantidad.SetFocus
                    Else
                WArticulo.SetFocus
        End If
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        Select Case Tipomov.ListIndex
            Case 0
                WMovi.Text = "S"
            Case Else
                WMovi.Text = "E"
        End Select
        WLote.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If WTipo.Text = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            WArticulo.Text = UCase(WArticulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                If WControla = 2 Then
                    WControla = 0
                End If
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
            
                ZArticulo = Left$(WArticulo.Text, 2)
                If ZArticulo = "DS" Then
                    ZArticuloII = Mid$(WArticulo.Text, 4, 3)
                    If Val(ZArticuloII) < 100 Then
                        ZArticulo = ""
                    End If
                End If
            
                If ZArticulo = "DY" Or ZArticulo = "DS" Then
                
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Laudo"
                    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArticulo.Text + "'"
                    ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote.Text + "'"
                    ZSql = ZSql + " Order by Laudo.FechaOrd, Laudo.Laudo"
                    spLaudo = ZSql
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            WCanti = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                            WEntra = "S"
                            rstLaudo.Close
                        End With
                    End If
                    
                    If WEntra = "N" Then
                    
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Guia"
                        ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArticulo.Text + "'"
                        ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote.Text + "'"
                        ZSql = ZSql + " Order by Guia.FechaOrd, Guia.Codigo"
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
                                WCanti = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                WEntra = "S"
                                rstMovguia.Close
                            End With
                        End If
                        
                    End If
                    
                        Else
                
                    WCanti = 0
                    XParam = "'" + WLote.Text + "','" _
                            + WArticulo.Text + "'"
                    spLaudo = "ListaLaudoArticulo " + XParam
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        WCanti = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                        WEntra = "S"
                        rstLaudo.Close
                    End If
                
                    If WEntra = "N" Then
                        XParam = "'" + WArticulo.Text + "','" _
                                     + WLote.Text + "'"
                        spMovguia = "ListaMovguiaLote " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            WCanti = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                            WEntra = "S"
                            rstMovguia.Close
                        End If
                    End If
                    
                End If
                
                    Else
                    
                WCanti = Val(WCantidad.Text)
                WEntra = "S"
                
            End If
            
            Rem dada
            Rem dada
            
            If WEntra = "S" Then
            
                If ZArticulo <> "DY" And ZArticulo <> "DS" Then
                
                    XRenglon = 0
                    Erase XAuxiliar
                
                    For a = 0 To 3
        
                        Suma = a * 10
                        DBGrid1.FirstRow = Suma
            
                        For iRow = 0 To 9
                
                            WRow = iRow
                            DBGrid1.Row = WRow
                    
                            DBGrid1.Col = 0
                            Tipo = DBGrid1.Text
                                       
                            DBGrid1.Col = 1
                            Terminado = DBGrid1.Text
                    
                            DBGrid1.Col = 2
                            Articulo = DBGrid1.Text
                        
                            DBGrid1.Col = 4
                            Cantidad = DBGrid1.Text
                    
                            DBGrid1.Col = 5
                            Movi = DBGrid1.Text
            
                            DBGrid1.Col = 6
                            Lote = DBGrid1.Text
            
                            DBGrid1.Col = 7
                            Transito = DBGrid1.Text
            
                            DBGrid1.Col = 8
                            Orden = DBGrid1.Text
            
                            DBGrid1.Col = 9
                            Descontar = DBGrid1.Text
                    
                            If Tipo <> "" Then
            
                                XRenglon = XRenglon + 1
                
                                XAuxiliar(XRenglon, 1) = Tipo
                                XAuxiliar(XRenglon, 2) = Terminado
                                XAuxiliar(XRenglon, 3) = Articulo
                                XAuxiliar(XRenglon, 4) = Cantidad
                                XAuxiliar(XRenglon, 5) = Movi
                                XAuxiliar(XRenglon, 6) = Lote
                                XAuxiliar(XRenglon, 7) = Transito
                                XAuxiliar(XRenglon, 8) = Orden
                                XAuxiliar(XRenglon, 9) = Descontar

                            End If
                
                        Next iRow
            
                    Next a
                
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Laudo"
                    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArticulo.Text + "'"
                    ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    spLaudo = ZSql
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            If .NoMatch = False Then
                                Do
                                    If .EOF = True Then
                                        Exit Do
                                    End If
                               
                                    WMarcaVencida = IIf(IsNull(rstLaudo!MarcaVencida), "", rstLaudo!MarcaVencida)
                                    QSaldo = rstLaudo!Saldo
                                    Call Redondeo(QSaldo)
                                    If QSaldo <> 0 And Trim(WMarcaVencida) = "" Then
                                        If rstLaudo!Articulo = Articulo Then
                                                
                                            WLaudo = rstLaudo!Laudo
                                            ZEntra = "S"
                                                    
                                            If WLaudo >= 190000 And WLaudo <= 194999 Then
                                                ZEntra = "N"
                                            End If
                                            If WLaudo >= 990000 And WLaudo <= 994999 Then
                                                ZEntra = "N"
                                            End If
                                            If WLaudo >= 290000 And WLaudo <= 294999 Then
                                                ZEntra = "N"
                                            End If
                                            If WLaudo >= 390000 And WLaudo <= 394999 Then
                                                ZEntra = "N"
                                            End If
                                            If WLaudo >= 490000 And WLaudo <= 494999 Then
                                                ZEntra = "N"
                                            End If
                                            If WLaudo >= 590000 And WLaudo <= 594999 Then
                                                ZEntra = "N"
                                            End If
                                            If WLaudo >= 690000 And WLaudo <= 694999 Then
                                                ZEntra = "N"
                                            End If
                                            If WLaudo >= 790000 And WLaudo <= 794999 Then
                                                ZEntra = "N"
                                            End If
                                            If WLaudo >= 890000 And WLaudo <= 894999 Then
                                                ZEntra = "N"
                                            End If
                                                    
                                            If ZEntra = "S" Then
                                            
                                                ZCompara = 1
                                                Do
                                                    If WLaudo <> XAuxiliar(ZCompara, 6) Then
                                                        ZEntra = "N"
                                                    End If
                                                    Exit Do
                                                Loop
                                                
                                            End If
                                                    
                                        End If
                                    End If
                               
                                    .MoveNext
                               
                                    If .EOF = True Then
                                        Exit Do
                                    End If
                               
                                Loop
                            End If
                        End With
                        rstLaudo.Close
                    End If
                    
        
                        
                    Rem ZSql = ""
                    Rem ZSql = ZSql + "Select *"
                    Rem ZSql = ZSql + " FROM Guia"
                    Rem ZSql = ZSql + " Where Guia.Articulo = " + "'" + Articulo + "'"
                    Rem ZSql = ZSql + " Order by Guia.Codigo"
                    Rem spMovguia = ZSql
                    Rem Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    Rem If rstMovguia.RecordCount > 0 Then
                    Rem
                    Rem      With rstMovguia
                    Rem
                    Rem         .MoveFirst
                    Rem
                    Rem         If .NoMatch = False Then
                    Rem             Do
                    Rem
                    Rem                 If .EOF = True Then
                    Rem                     Exit Do
                    Rem                 End If
                    Rem
                    Rem                 WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                    Rem                 QSaldo = rstMovguia!Saldo
                    Rem                 Call Redondeo(QSaldo)
                    Rem                 If QSaldo <> 0 And Trim(WMarcaVencida) = "" Then
                    Rem                     If rstMovguia!Articulo = Articulo Then
                    Rem
                    Rem                         WLaudo = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                    Rem                         ZEntra = "S"
                    Rem
                    Rem                         If WLaudo >= 190000 And WLaudo <= 194999 Then
                    Rem                             ZEntra = "N"
                    Rem                         End If
                    Rem                         If WLaudo >= 990000 And WLaudo <= 994999 Then
                    Rem                             ZEntra = "N"
                    Rem                         End If
                    Rem                         If WLaudo >= 290000 And WLaudo <= 294999 Then
                    Rem                             ZEntra = "N"
                    Rem                         End If
                    Rem                         If WLaudo >= 390000 And WLaudo <= 394999 Then
                    Rem                             ZEntra = "N"
                    Rem                         End If
                    Rem                         If WLaudo >= 490000 And WLaudo <= 494999 Then
                    Rem                             ZEntra = "N"
                    Rem                         End If
                    Rem                         If WLaudo >= 590000 And WLaudo <= 594999 Then
                    Rem                             ZEntra = "N"
                    Rem                         End If
                    Rem                         If WLaudo >= 690000 And WLaudo <= 694999 Then
                    Rem                             ZEntra = "N"
                    Rem                         End If
                    Rem                         If WLaudo >= 790000 And WLaudo <= 794999 Then
                    Rem                             ZEntra = "N"
                    Rem                         End If
                    Rem                         If WLaudo >= 890000 And WLaudo <= 894999 Then
                    Rem                             ZEntra = "N"
                    Rem                         End If
                    Rem
                    Rem                         If ZEntra = "S" Then
                    Rem
                    Rem                         End If
                    Rem
                    Rem                     End If
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
                    Rem     End With
                    Rem     rstMovguia.Close
                    Rem End If
                    
                End If
                
            End If
            
            
            
            If WEntra = "S" Then
                If WCanti >= Val(WCantidad.Text) Or WMovi.Text = "E" Then
                    If Left$(WArticulo.Text, 2) = "DY" Or Left$(WArticulo.Text, 2) = "DS" Then
                        IngresoTransito.Visible = True
                        WTransito.SetFocus
                            Else
                        Call Alta_Vector
                        Call Ingresa_Click
                        WTipo.SetFocus
                    End If
                        Else
                    m$ = WArticulo.Text + " Stock Insufucuente. Cantidad:" + Str$(WCanti)
                    G% = MsgBox(m$, 0, "Guias de Traslado Internos")
                End If
                    Else
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + WLote.Text + " inexistente"
                G% = MsgBox(m$, 0, "Guias de Traslado Internos")
            End If
            
            
            
                Else
                
                
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                If WControla = 2 Then
                    WControla = 0
                End If
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote.Text + "','" _
                        + WTerminado.Text + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WCanti = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WTerminado.Text + "','" _
                            + WLote.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WCanti = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra = "S" Then
                If WCanti >= Val(WCantidad.Text) Or WMovi.Text = "E" Then
                    If Val(WEmpresa) = 4 Then
                        PantaOrden.Visible = True
                        ZOrden.SetFocus
                            Else
                        Call Alta_Vector
                        Call Ingresa_Click
                        WTipo.SetFocus
                    End If
                        Else
                    m$ = WTerminado.Text + " Stock Insufucuente. Cantidad:" + Str$(WCanti)
                    G% = MsgBox(m$, 0, "Guias de Traslado Internos")
                End If
                    Else
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
                G% = MsgBox(m$, 0, "Guias de Traslado Internos")
            End If
            
        End If
    
    End If
    
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub WTransito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        WTransito.Text = Trim(WTransito.Text)
        If WTransito.Text = "" Then
            IngresoTransito.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WTipo.SetFocus
            Exit Sub
        End If
        
        WArticulo = WArticulo.Text
        WLote = WLote.Text
        WTransito = WTransito.Text
        WSaldo = 0
        WCantidad = Val(WCantidad.Text)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArticulo + "'"
        ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote + "'"
        ZSql = ZSql + " and Laudo.Transito = " + "'" + WTransito + "'"
        ZSql = ZSql + " Order by Laudo.FechaOrd, Laudo.Laudo"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            With rstLaudo
                .MoveFirst
                WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                rstLaudo.Close
            End With
        End If
        
        If WSaldo >= Val(WCantidad) Then
            IngresoTransito.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WTipo.SetFocus
        End If

    End If
End Sub

Private Sub ZOrden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        If Val(ZOrden.Text) = 0 Then
            ZDescontar.Text = ""
            PantaOrden.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WTipo.SetFocus
            Exit Sub
        End If
        
        ZSaldo.Text = ""
        
        Sql1 = "Select *"
        Sql2 = " FROM CargaProyeccion"
        Sql3 = " Where CargaProyeccion.Articulo = " + "'" + WTerminado.Text + "'"
        Sql4 = " and CargaProyeccion.Orden = " + "'" + ZOrden.Text + "'"
        spCargaProyeccion = Sql1 + Sql2 + Sql3 + Sql4
        Set rstCargaProyeccion = db.OpenRecordset(spCargaProyeccion, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaProyeccion.RecordCount > 0 Then
            ZSaldo.Text = IIf(IsNull(rstCargaProyeccion!Saldo), "0", rstCargaProyeccion!Saldo)
            rstCargaProyeccion.Close
        End If
        
        If Val(ZSaldo.Text) = 0 Then
            ZOrden.SetFocus
                Else
            ZDescontar.SetFocus
        End If

    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ZDescontar_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PantaOrden.Visible = False
        Call Alta_Vector
        Call Ingresa_Click
        WTipo.SetFocus
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
        
            spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WTipo.Text = "T"
                WTerminado.Text = Claveven$
                WDescripcion.Caption = rstTerminado!Descripcion
                    
                DBGrid1.Col = 0
                DBGrid1.Text = "T"
                DBGrid1.Col = 1
                DBGrid1.Text = rstTerminado!Codigo
                DBGrid1.Col = 3
                DBGrid1.Text = rstTerminado!Descripcion
                rstTerminado.Close
                    
                Call Alta_Vector
                WLinea.Text = WAnterior + 1
                If Val(WLinea.Text) > 0 Then
                    DBGrid1.Row = Val(WLinea.Text) - 1
                End If
                    
                Call DBGrid1.SetFocus
                WCantidad.SetFocus
                    
            End If
            
        Case 1
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
        
            spArticulo = "ConsultaArticulo " + "'" + Claveven$ + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WTipo.Text = "M"
                WArticulo.Text = rstArticulo!Codigo
                WDescripcion.Caption = rstArticulo!Descripcion
                    
                DBGrid1.Col = 0
                DBGrid1.Text = "M"
                DBGrid1.Col = 2
                DBGrid1.Text = rstArticulo!Codigo
                DBGrid1.Col = 3
                DBGrid1.Text = rstArticulo!Descripcion
                rstArticulo.Close
                    
                Call Alta_Vector
                WLinea.Text = WAnterior + 1
                If Val(WLinea.Text) > 0 Then
                    DBGrid1.Row = Val(WLinea.Text) - 1
                End If
                                        
                Call DBGrid1.SetFocus
                WCantidad.SetFocus
                    
            End If
            
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4, 5
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 40 Then
                            DBGrid1.Row = DBGrid1.Row + 1
                            DBGrid1.Col = 0
                            KeyCode = 0
                        End If
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

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 9, 0 To 40)

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
For i = 0 To 9
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Tipo"
             DBGrid1.Columns(newcnt).Width = 400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Prod.Terminado"
             DBGrid1.Columns(newcnt).Width = 1500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Materia Prima"
             DBGrid1.Columns(newcnt).Width = 1500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3620
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Movimiento"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
         Case 6
             DBGrid1.Columns(newcnt).Caption = "Lote"
             DBGrid1.Columns(newcnt).Width = 1100
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
         Case 7
             DBGrid1.Columns(newcnt).Caption = ""
             DBGrid1.Columns(newcnt).Width = 100
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
         Case 8
             DBGrid1.Columns(newcnt).Caption = ""
             DBGrid1.Columns(newcnt).Width = 100
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
         Case 9
             DBGrid1.Columns(newcnt).Caption = ""
             DBGrid1.Columns(newcnt).Width = 100
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
         Case Else
     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    Codigo.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
 
    Rem With rstMovguia
    Rem     .Index = "Clave"
    Rem    Claveven$ = "99999999"
    Rem    .Seek "<=", Claveven$
    Rem    If .NoMatch = False Then
    Rem        Codigo.Text = !Codigo + 1
    Rem            Else
    Rem        Codigo.Text = ""
    Rem    End If
    Rem End With
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgMovguia.Caption = "Listado de E/S de Materia Prima y Productos :  " + !Nombre
        End If
    End With
    
    Tipomov.Clear
    
    Tipomov.AddItem "Emision de Guia de Traslado interno"
    Tipomov.AddItem "Recepcion de Surfactan"
    Tipomov.AddItem "Recepcion de Pellital"
    Tipomov.AddItem "Recepcion de Surfactan II"
    Tipomov.AddItem "Recepcion de Pellital II"
    Tipomov.AddItem "Recepcion de Surfactan III"
    Tipomov.AddItem "Recepcion de Surfactan IV"
    Tipomov.AddItem "Recepcion de Surfactan V"
    Tipomov.AddItem "Recepcion de Pellital V"
    Tipomov.AddItem "Recepcion de Pellital IV"
    Tipomov.AddItem "Recepcion de Surfactan VI"
    Tipomov.AddItem "Recepcion de Surfactan VII"
    Tipomov.AddItem "Otro"
    
    Tipomov.ListIndex = 0
    
    Destino.Clear
    
    Destino.AddItem ""
    Destino.AddItem "Envio hacia Surfactan"
    Destino.AddItem "Envio hacia Pellital"
    Destino.AddItem "Envio hacia Surfactan II"
    Destino.AddItem "Envio hacia Pellital II"
    Destino.AddItem "Envio hacia Surfactan III"
    Destino.AddItem "Envio hacia Surfactan IV"
    Destino.AddItem "Envio hacia Surfactan V"
    Destino.AddItem "Envio hacia Pelitall V"
    Destino.AddItem "Envio hacia Pelitall IV"
    Destino.AddItem "Envio hacia Surfactan VI"
    Destino.AddItem "Envio hacia Surfactan VII"
    Destino.AddItem "Otro"
    
    Destino.ListIndex = 0
    
    WTipomov = Str$(Tipomov.ListIndex)
    Call Ceros(WTipomov, 2)
    
    Codigo.Text = ""
    
    Rem spMovguia = "ListaMovguiaNumero " + "'" + WTipomov + "'"
    Rem Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstMovguia.RecordCount > 0 Then
    Rem     With rstMovguia
    Rem         .MoveLast
    Rem         Do
    Rem             Codigo.Text = rstMovguia!Codigo + 1
    Rem             If Val(Codigo.Text) > 900000 Then
    Rem                 .MovePrevious
    Rem                     Else
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem     End With
    Rem     rstMovguia.Close
    Rem         Else
    Rem     Codigo.Text = "1"
    Rem End If
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Graba.Enabled = True
    
    Codigo.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 9
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Erase Auxiliar
    Renglon = 0
    
    Select Case Val(WEmpresa)
        Case 1
            ZCodigo = Str$(Val(Codigo.Text) + 300000)
        Case 3
            ZCodigo = Str$(Val(Codigo.Text) + 400000)
        Case 5
            ZCodigo = Str$(Val(Codigo.Text) + 500000)
        Case 7
            ZCodigo = Str$(Val(Codigo.Text) + 600000)
        Case 10
            ZCodigo = Str$(Val(Codigo.Text) + 700000)
        Case 11
            ZCodigo = Str$(Val(Codigo.Text) + 800000)
        Case Else
    End Select
    
    WTipomov = Str$(Tipomov.ListIndex)
    Call Ceros(WTipomov, 2)
    
    XParam = "'" + WTipomov + "','" _
                + ZCodigo + "'"
    spMovguia = "ListaMovguia " + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)

    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstMovguia!Tipo
                
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstMovguia!Terminado
                    Auxi1 = rstMovguia!Terminado
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = rstMovguia!Articulo
                    Auxi2 = rstMovguia!Articulo
                
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", (rstMovguia!Cantidad))
                
                    DBGrid1.Col = 5
                    DBGrid1.Text = rstMovguia!Movi
                    
                    DBGrid1.Col = 6
                    DBGrid1.Text = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                    
                    DBGrid1.Col = 7
                    DBGrid1.Text = IIf(IsNull(rstMovguia!Transito), "", rstMovguia!Transito)
                    
                    DBGrid1.Col = 8
                    DBGrid1.Text = IIf(IsNull(rstMovguia!Orden), "", rstMovguia!Orden)
                    
                    DBGrid1.Col = 9
                    DBGrid1.Text = IIf(IsNull(rstMovguia!Descontar), "", rstMovguia!Descontar)
                    
                    Tipomov.ListIndex = Val(rstMovguia!Tipomov)
                    Destino.ListIndex = Val(rstMovguia!Destino)
                    Observaciones.Text = rstMovguia!Observaciones
                    Fecha.Text = rstMovguia!Fecha
                    
                    Auxiliar(Renglon, 1) = Auxi1
                    Auxiliar(Renglon, 2) = Auxi2
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovguia.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0

    For Da = 1 To WRenglon
    
        Auxi1 = Auxiliar(Da, 1)
        Auxi2 = Auxiliar(Da, 2)
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
    
        spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DBGrid1.Col = 3
            DBGrid1.Text = rstTerminado!Descripcion
            rstTerminado.Close
        End If
        
        spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DBGrid1.Col = 3
            DBGrid1.Text = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
    Next Da

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
    
    WTipo.SetFocus

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
            DBGrid1.Text = WTipo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WTerminado.Text
            
            DBGrid1.Col = 2
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 3
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 4
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
                
            DBGrid1.Col = 5
            DBGrid1.Text = WMovi.Text
            
            DBGrid1.Col = 6
            DBGrid1.Text = WLote.Text
            
            DBGrid1.Col = 7
            DBGrid1.Text = WTransito.Text
            
            DBGrid1.Col = 8
            DBGrid1.Text = ZOrden.Text
            
            DBGrid1.Col = 9
            DBGrid1.Text = ZDescontar.Text
            
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WTipo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WTerminado.Text
            
            DBGrid1.Col = 2
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 3
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 4
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
                
            DBGrid1.Col = 5
            DBGrid1.Text = WMovi.Text
            
            DBGrid1.Col = 6
            DBGrid1.Text = WLote.Text
            
            DBGrid1.Col = 7
            DBGrid1.Text = WTransito.Text
            
            DBGrid1.Col = 8
            DBGrid1.Text = ZOrden.Text
            
            DBGrid1.Col = 9
            DBGrid1.Text = ZDescontar.Text
            
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WTipomov = Str$(Tipomov.ListIndex)
        Call Ceros(WTipomov, 2)
    
        ZCodigo = Codigo.Text
        Select Case Val(WEmpresa)
            Case 1
                ZCodigo = Str$(Val(Codigo.Text) + "300000")
            Case 3
                ZCodigo = Str$(Val(Codigo.Text) + 400000)
            Case 5
                ZCodigo = Str$(Val(Codigo.Text) + 500000)
            Case 7
                ZCodigo = Str$(Val(Codigo.Text) + 600000)
            Case 10
                ZCodigo = Str$(Val(Codigo.Text) + 700000)
            Case 11
                ZCodigo = Str$(Val(Codigo.Text) + 800000)
            Case Else
        End Select
    
        XParam = "'" + WTipomov + "','" _
                + ZCodigo + "'"
        spMovguia = "ListaMovguia " + XParam
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    
        If rstMovguia.RecordCount > 0 Then
            Fecha.Text = rstMovguia!Fecha
            rstMovguia.Close
            Graba.Enabled = False
            Call Proceso_Click
                Else
            WCodigo = Codigo.Text
            Call Limpia_Click
            Codigo.Text = WCodigo
            Graba.Enabled = True
            Fecha.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Observaciones.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WTipo.SetFocus
    End If
End Sub

Sub Impresion()

        If Val(WEmpresa) = 1 Then
            Rem Open "DADA.TXT" For Output As #1
            Open "lpt1" For Output As #1
                Else
            Rem Open "DADA.TXT" For Output As #1
            
            Open "lpt1" For Output As #1
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
        End If
  
        Rem  #1, 255

        For FF = 1 To 2

        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "2" + Chr$(72)
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        Print #1, ""
        Print #1, Tab(48); "GUIA DE TRASLADO INTERNO"
        Print #1, ""
        Print #1, ""
        Print #1, Tab(53); Fecha.Text
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        
        Select Case Val(WDestino)
            Case 1, 6, 10, 11
                Print #1, Tab(7); "Surfactan"
                Print #1, Tab(7); "Malvinas Argentinas 4589"
                Print #1, Tab(7); "Victoria"
                Print #1, Tab(7); "Pcia. Bs.As."
                Print #1, ""
                Print #1, Tab(7); "Inscripto";
                Print #1, Tab(48); "30-54916508-3"
                Print #1, ""
                Print #1, Tab(30); "Direccion Entrega : Malvinas Argentinas 4589";
                Print #1, ""
            Case 3
                Print #1, Tab(7); "Surfactan"
                Print #1, Tab(7); "Malvinas Argentinas 4589"
                Print #1, Tab(7); "Victoria"
                Print #1, Tab(7); "Pcia. Bs.As."
                Print #1, ""
                Print #1, Tab(7); "Inscripto";
                Print #1, Tab(48); "30-54916508-3"
                Print #1, ""
                Print #1, Tab(30); "Direccion Entrega  : Uruguay 2671";
                Print #1, ""
            Case 5
                Print #1, Tab(7); "Surfactan"
                Print #1, Tab(7); "Malvinas Argentinas 4589"
                Print #1, Tab(7); "Victoria"
                Print #1, Tab(7); "Pcia. Bs.As."
                Print #1, ""
                Print #1, Tab(7); "Inscripto";
                Print #1, Tab(48); "30-54916508-3"
                Print #1, ""
                Print #1, Tab(30); "Direccion Entrega : Kenedy 2689 Esq. Entre Rios";
                Print #1, ""
            Case 7
                Print #1, Tab(7); "Surfactan"
                Print #1, Tab(7); "Malvinas Argentinas 4589"
                Print #1, Tab(7); "Victoria"
                Print #1, Tab(7); "Pcia. Bs.As."
                Print #1, ""
                Print #1, Tab(7); "Inscripto";
                Print #1, Tab(48); "30-54916508-3"
                Print #1, ""
                Print #1, Tab(30); "Direccion Entrega : Tucuman 3275";
                Print #1, ""
            Case 2
                Print #1, Tab(7); "Pellital"
                Print #1, Tab(7); "Tucuman 3275"
                Print #1, Tab(7); "Victoria"
                Print #1, Tab(7); "Pcia. Bs.As."
                Print #1, ""
                Print #1, Tab(7); "Inscripto";
                Print #1, Tab(48); "30-61052459-8"
                Print #1, ""
                Print #1, Tab(30); "Direccion Entrega : Malvinas Argentinas 4589";
                Print #1, ""
            Case 4
                Print #1, Tab(7); "Pellital"
                Print #1, Tab(7); "Tucuman 3275"
                Print #1, Tab(7); "Victoria"
                Print #1, Tab(7); "Pcia. Bs.As."
                Print #1, ""
                Print #1, Tab(7); "Inscripto";
                Print #1, Tab(48); "30-61052459-8"
                Print #1, ""
                Print #1, Tab(30); "Direccion Entrega : Uruguay 2671";
                Print #1, ""
            Case 8
                Print #1, Tab(7); "Pellital"
                Print #1, Tab(7); "Tucuman 3275"
                Print #1, Tab(7); "Victoria"
                Print #1, Tab(7); "Pcia. Bs.As."
                Print #1, ""
                Print #1, Tab(7); "Inscripto";
                Print #1, Tab(48); "30-61052459-8"
                Print #1, ""
                Print #1, Tab(30); "Direccion Entrega : Tucuman 3275";
                Print #1, ""
            Case Else
        End Select
                
        If FF = 1 Then
            Print #1, Tab(60); "ORIGINAL"
                Else
            Print #1, Tab(60); "DUPLICADO"
        End If
        Print #1, ""
        
        Impre = 0

        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 3
                Descri = DBGrid1.Text
            
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                
                If Cantidad <> 0 Then
                    Print #1, Tab(14); Left$(Descri, 40);
                    Print #1, Tab(58); Alinea("#####.##", Str$(Cantidad));
                    Print #1, " Kg";
                    Print #1, Tab(71); "Netos"
                    Impre = Impre + 1
                End If
                    
            Next iRow
            
        Next a
        
        Rem For aa = Impre To 22
        Rem         Print #1, ""
        Rem Next aa
        
        Rem Print #1, ""
        Rem Print #1, ""
        Rem Print #1, ""
        
        Rem For Da = 1 To 9
        Rem         Print #1, ""
        Rem Next Da
        
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        
        Rem For xda = 2 To 4
        Rem         Print #1, ""
        Rem         Print #1, ""
        Rem Next xda
        
        Rem Print #1, ""
        Rem Select Case XX
        Rem         Case 1
        Rem                 Print #1, Tab(10); "ORIGINAL";
        Rem         Case 2
        Rem                 Print #1, Tab(10); "DUPLICADO";
        Rem         Case 3
        Rem                 Print #1, Tab(10); "TRIPLICADO";
        Rem         Case Else
        Rem End Select
        Rem Print #1, Tab(10); "Nro. Control : "; Codigo.Text
        Rem Print #1, Chr$(12)

        Next FF

        Close #1

End Sub

