VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCargaZona 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Mercaderia de Zona Franca"
   ClientHeight    =   8190
   ClientLeft      =   0
   ClientTop       =   525
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   8190
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin MSFlexGridLib.MSFlexGrid WVector2 
      Height          =   615
      Left            =   9360
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      _Version        =   327680
      BackColor       =   12648384
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   20
      Text            =   " "
      Top             =   480
      Width           =   1095
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
      Index           =   5
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   " "
      Top             =   2400
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
      Index           =   4
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   " "
      Top             =   2400
      Width           =   375
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
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   15
      Top             =   840
      Width           =   5295
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   1095
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   " "
      Top             =   2400
      Width           =   375
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
      Left            =   1800
      TabIndex        =   10
      Top             =   1800
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      Top             =   2400
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
      Left            =   1200
      TabIndex        =   8
      Top             =   1800
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2400
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
      Index           =   2
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2400
      Width           =   375
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
      Top             =   5040
      Visible         =   0   'False
      Width           =   6855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10560
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "PedidoFranca.rpt"
      PrintFileType   =   17
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
      Top             =   5520
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   6240
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
      ItemData        =   "CargaZona.frx":0000
      Left            =   120
      List            =   "CargaZona.frx":0007
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   1800
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
      Height          =   3735
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   6588
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   4800
      TabIndex        =   16
      Top             =   120
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
   Begin VB.Label Label2 
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   240
      TabIndex        =   22
      Top             =   480
      Width           =   1575
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
      Left            =   3000
      TabIndex        =   21
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   120
      Width           =   1215
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
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   1335
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   10560
      MouseIcon       =   "CargaZona.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "CargaZona.frx":031F
      ToolTipText     =   "Salida"
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   10560
      MouseIcon       =   "CargaZona.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "CargaZona.frx":0E6B
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   10560
      MouseIcon       =   "CargaZona.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "CargaZona.frx":19B7
      ToolTipText     =   "Consulta de Datos"
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   10560
      MouseIcon       =   "CargaZona.frx":21F9
      MousePointer    =   99  'Custom
      Picture         =   "CargaZona.frx":2503
      ToolTipText     =   "Limpia la pantalla"
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Solicitud"
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
Attribute VB_Name = "PrgCargaZona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstCargaZona As Recordset
Dim spCargaZona As String

Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Cantidad As Double
Dim XPaso As String
Dim Renglon As Integer

Dim EmailAddress As String
Dim WEmail(100) As String
Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
Dim MSubject As String
Dim MBody As String
Dim AllPath As String

Dim ZSaldo As Double

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String


Private Sub Consulta_Click()

     Opcion.Clear
     Opcion.AddItem "Materia Prima"
     Opcion.AddItem "Clientes"
     
     Opcion.Visible = True
     
     Rem Opcion.ListIndex = 0
     Rem Call Opcion_Click
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Ayuda.Visible = True
    Ayuda.Text = ""
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Order by Codigo"
            spArticulo = Sql1 + Sql2 + Sql3
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
        Case 1
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
            
    Ayuda.SetFocus
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub cmdClose1_Click()

    Call Limpia_Click
    PrgCargaZona.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()

    Sql1 = "DELETE CargaZona"
    Sql2 = " Where Solicitud = " + "'" + Solicitud.Text + "'"
    spCargaZona = Sql1 + Sql2
    Set rstCargaZona = db.OpenRecordset(spCargaZona, dbOpenSnapshot, dbSQLPassThrough)
    
    ZRazon = DesCliente.Caption
    
    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
        
        WVector1.Row = iRow
            
        WVector1.Col = 1
        PArticulo = WVector1.Text
        
        WVector1.Col = 3
        XCantidad = WVector1.Text
        
        If PArticulo <> "" Or XCantidad <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    WRenglon = 0
    For iRow = 1 To HastaRenglon
        
        WVector1.Row = iRow
            
        WVector1.Col = 1
        ZArticulo = WVector1.Text
        
        WVector1.Col = 2
        ZDescriArticulo = WVector1.Text
        
        WVector1.Col = 3
        ZCantidad = WVector1.Text
        
        WVector1.Col = 4
        ZPartidaOri = WVector1.Text
        
        WVector1.Col = 5
        ZTransito = WVector1.Text
        
        ZEntregado = "0"
        
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
        
        Auxi1 = Str$(Solicitud.Text)
        Call Ceros(Auxi1, 6)
        
        WClave = Auxi1 + Auxi
        ZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            
        Sql1 = "INSERT INTO CargaZona ("
        Sql2 = "Clave ,"
        Sql3 = "Solicitud ,"
        Sql4 = "Renglon ,"
        Sql5 = "Cliente ,"
        Sql6 = "Fecha ,"
        Sql7 = "OrdFecha ,"
        Sql8 = "Observaciones ,"
        Sql9 = "Articulo ,"
        Sql10 = "Cantidad ,"
        Sql11 = "Entregado ,"
        Sql12 = "Partida ,"
        Sql13 = "PartiOri ,"
        Sql14 = "Transito ,"
        Sql15 = "DescriArticulo ,"
        Sql16 = "Razon )"
        Sql17 = "Values ("
        Sql18 = "'" + WClave + "',"
        Sql19 = "'" + Solicitud.Text + "',"
        Sql20 = "'" + Str$(WRenglon) + "',"
        Sql21 = "'" + Cliente.Text + "',"
        Sql22 = "'" + Fecha.Text + "',"
        Sql23 = "'" + ZOrdFecha + "',"
        Sql24 = "'" + Observaciones.Text + "',"
        Sql25 = "'" + ZArticulo + "',"
        Sql26 = "'" + ZCantidad + "',"
        Sql27 = "'" + ZEntregado + "',"
        Sql28 = "'" + ZPartida + "',"
        Sql29 = "'" + ZPartidaOri + "',"
        Sql30 = "'" + ZTransito + "',"
        Sql31 = "'" + ZDescriArticulo + "',"
        Sql32 = "'" + ZRazon + "')"
            
        spCargaZona = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 _
                         + Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 _
                         + Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + Sql31 + Sql32
        Set rstCargaZona = db.OpenRecordset(spCargaZona, dbOpenSnapshot, dbSQLPassThrough)
            
    Next iRow
    
    
    T$ = "Solicitud de Fabricacion"
    m$ = "Desea imprimir la Solicitud"
    Respuesta% = MsgBox(m$, 4, T$)
    If Respuesta% = 6 Then
    
        Listado.GroupSelectionFormula = "{CargaZona.Solicitud} in " + Solicitud.Text + " to " + Solicitud.Text
        Listado.Destination = 1
        Rem Listado.Destination = 0
    
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
    
        Listado.SQLQuery = "SELECT CargaZona.Clave, CargaZona.Solicitud, CargaZona.Cliente, CargaZona.Fecha, CargaZona.Observaciones, CargaZona.Articulo, CargaZona.Cantidad, CargaZona.PartiOri, CargaZona.Transito, CargaZona.DescriArticulo, CargaZona.Razon " _
                        + "From " _
                        + DSQ + ".dbo.CargaZona CargaZona " _
                        + "Where " _
                        + "CargaZona.Solicitud >= " + Solicitud.Text + " AND " _
                        + "CargaZona.Solicitud <= " + Solicitud.Text
                        
        Rem Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
        Listado.Connect = Connect()
        Listado.Action = 1
        
    End If
    
    If MiRuta = "" Then
        MiRuta = CurDir + "\"
        Rem MiRutaII = Left$(CurDir, 1)
    End If
    
    T$ = "Solicitud de Fabricacion"
    m$ = "Desea enviar la Solicitud via email"
    Respuesta% = MsgBox(m$, 4, T$)
    If Respuesta% = 6 Then
    
        Listado.Destination = 3
        Listado.PrintFileType = crptWinWord

        Listado.EMailToList = "alejandro-perera@comex.com.ar"
        Listado.EMailCCList = "biglesias@surfactan.com.ar; jlgdeposito@comex.com.ar"
        Listado.EMailSubject = "Pedido de Zona Franca"
        Listado.EMailMessage = "Se adjunta un archivo con la mercaderia que se debe gestionar de la zona franca."
    
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
    
        Listado.SQLQuery = "SELECT CargaZona.Clave, CargaZona.Solicitud, CargaZona.Cliente, CargaZona.Fecha, CargaZona.Observaciones, CargaZona.Articulo, CargaZona.Cantidad, CargaZona.PartiOri, CargaZona.Transito, CargaZona.DescriArticulo, CargaZona.Razon " _
                        + "From " _
                        + DSQ + ".dbo.CargaZona CargaZona " _
                        + "Where " _
                        + "CargaZona.Solicitud >= " + Solicitud.Text + " AND " _
                        + "CargaZona.Solicitud <= " + Solicitud.Text
                        
        Rem Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
        Listado.Connect = Connect()
        
        Listado.Action = 1
        
        Rem ChDrive MiRutaII
        ChDir MiRuta
            
    End If
    
    Call Limpia_Click

    WVector1.Col = 1
    WVector1.Row = 1
        
    Solicitud.SetFocus
        
End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector

    Solicitud.Text = ""
    Observaciones.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = "  /  /    "
    
    Renglon = 0
    Graba.Enabled = True
    
    Sql1 = "Select Max(Solicitud) as [SolicitudMayor]"
    Sql2 = " FROM CargaZona"
    spCargaZona = Sql1 + Sql2
    Set rstCargaZona = db.OpenRecordset(spCargaZona, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaZona.RecordCount > 0 Then
        rstCargaZona.MoveLast
        ZSolicitudMayor = IIf(IsNull(rstCargaZona!SolicitudMayor), "0", rstCargaZona!SolicitudMayor)
        Solicitud.Text = Str$(ZSolicitudMayor + 1)
        rstCargaZona.Close
            Else
        Solicitud.Text = "1"
    End If
    
    WVector1.Col = 1
    WVector1.Row = 1
    
    Solicitud.SetFocus

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WPArticulo = WIndice.List(Indice)
            
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Where Articulo.Codigo = " + "'" + WPArticulo + "'"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 1
                WVector1.Text = Trim(rstArticulo!Codigo)
                WVector1.Col = 2
                If WVector1.Text = "" Then
                    WVector1.Text = Trim(rstArticulo!Descripcion)
                End If
                WVector1.Col = 3
                rstArticulo.Close
                Call StartEdit
            End If
            Rem Ayuda.Visible = False
            
        Case 1
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spCliente = "ConsultaCliente " + "'" + Claveven$ + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = Claveven$
                DesCliente.Caption = rstCliente!Razon
                rstCliente.Close
            End If
            Cliente.SetFocus
        
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    WVector1.Col = 1
    WVector1.Row = 1

    Solicitud.Text = ""
    Observaciones.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = "  /  /    "
    
    Sql1 = "Select Max(Solicitud) as [SolicitudMayor]"
    Sql2 = " FROM CargaZona"
    spCargaZona = Sql1 + Sql2
    Set rstCargaZona = db.OpenRecordset(spCargaZona, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaZona.RecordCount > 0 Then
        rstCargaZona.MoveLast
        ZSolicitudMayor = IIf(IsNull(rstCargaZona!SolicitudMayor), "0", rstCargaZona!SolicitudMayor)
        Solicitud.Text = Str$(ZSolicitudMayor + 1)
        rstCargaZona.Close
            Else
        Solicitud.Text = "1"
    End If
    
    Renglon = 0
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    WRenglon = 0
    
    Sql1 = "Select CargaZona.Clave, CargaZona.Solicitud, CargaZona.Articulo, CargaZona.Cantidad, CargaZona.PartiOri, CargaZona.Transito"
    Sql2 = " FROM CargaZona"
    Sql3 = " Where CargaZona.Solicitud = " + "'" + Solicitud.Text + "'"
    Sql4 = " Order by CargaZona.Clave"
    
    spCargaZona = Sql1 + Sql2 + Sql3 + Sql4
    Set rstCargaZona = db.OpenRecordset(spCargaZona, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaZona.RecordCount > 0 Then
        With rstCargaZona
            .MoveFirst
            Do
                If .EOF = False Then
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = Trim(rstCargaZona!Articulo)
                    
                    WVector1.Col = 2
                    WVector1.Text = ""
            
                    WVector1.Col = 3
                    WVector1.Text = rstCargaZona!Cantidad
                    WVector1.Text = Pusing("######.##", WVector1.Text)
                    
                    WVector1.Col = 4
                    WVector1.Text = rstCargaZona!PartiOri
                    
                    WVector1.Col = 5
                    WVector1.Text = rstCargaZona!Transito
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaZona.Close
    End If
    
    For Ciclo = 1 To WRenglon
        Sql1 = "Select *"
        Sql2 = " FROM Articulo"
        Sql3 = " Where Articulo.Codigo = " + "'" + WVector1.TextMatrix(Ciclo, 1) + "'"
        spArticulo = Sql1 + Sql2 + Sql3
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WVector1.TextMatrix(Ciclo, 2) = Trim(rstArticulo!Descripcion)
            rstArticulo.Close
        End If
    Next Ciclo
    
    Graba.Enabled = True

End Sub

Private Sub Solicitud_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM CargaZona"
        Sql3 = " Where CargaZona.Solicitud = " + "'" + Solicitud.Text + "'"
        spCargaZona = Sql1 + Sql2 + Sql3
        Set rstCargaZona = db.OpenRecordset(spCargaZona, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaZona.RecordCount > 0 Then
            Fecha.Text = rstCargaZona!Fecha
            Observaciones.Text = rstCargaZona!Observaciones
            Cliente.Text = rstCargaZona!Cliente
            rstCargaZona.Close
            Call Proceso_Click
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                Else
            Graba.Enabled = True
            WSolicitud = Solicitud.Text
            Call Limpia_Click
            Solicitud.Text = WSolicitud
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Solicitud.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Cliente.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        If Cliente.Text <> "" Then
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                rstCliente.Close
                Observaciones.SetFocus
                    Else
                Cliente.Text = Claveven$
                Cliente.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
        DesCliente.Caption = ""
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    Busqueda = Left$(Ayuda.Text, WEspacios)
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Order by Codigo"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Da = Len(rstArticulo!Descripcion) - WEspacios
                            For aa = 1 To Da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstArticulo!Descripcion, aa, WEspacios) Then
                                    IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstArticulo!Codigo
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
                rstArticulo.Close
            End If
            
        Case 1
            spCliente = "ListaClienteConsulta"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            Da = Len(rstCliente!Razon) - WEspacios
                
                            For aa = 1 To Da
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
            
        Case Else
    End Select
            
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
        Case 5
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
            WVector1.Text = UCase(WVector1.Text)
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Where Articulo.Codigo = " + "'" + WVector1.Text + "'"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 2
                If WVector1.Text = "" Then
                    WVector1.Text = Trim(rstArticulo!Descripcion)
                End If
                rstArticulo.Close
                    Else
                WControl = "N"
            End If
        Case 3
            Call ficha_Mp
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub ficha_Mp()
    
    Call Limpia_Vector2II
    WArticulo = Left$(WVector1.TextMatrix(WVector1.Row, 1), 3) + Right$(WVector1.TextMatrix(WVector1.Row, 1), 7)
    
    XRenglon = 0
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
        With rstLaudo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                ZSaldoTransito = IIf(IsNull(rstLaudo!SaldoTransito), "0", rstLaudo!SaldoTransito)
                If rstLaudo!Marca = "X" And rstLaudo!Saldo = 0 And ZSaldoTransito <> 0 Then
                
                        Else
                    
                    If rstLaudo!Articulo = WArticulo Then
                    
                        ZArticulo = rstLaudo!Articulo
                        ZCantidad = rstLaudo!Liberada
                        ZFecha = rstLaudo!Fecha
                        ZLaudo = rstLaudo!Laudo
                        ZOrden = rstLaudo!Orden
                        Zdevuelta = IIf(IsNull(rstLaudo!devuelta), "0", rstLaudo!devuelta)
                        ZRechazo = IIf(IsNull(rstLaudo!Rechazo), "0", rstLaudo!Rechazo)
                        ZSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                        ZLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                        ZPArtiOri = IIf(IsNull(rstLaudo!PartiOri), "0", rstLaudo!PartiOri)
                        ZTransito = IIf(IsNull(rstLaudo!Transito), "", rstLaudo!Transito)
                        ZSaldoTransito = IIf(IsNull(rstLaudo!SaldoTransito), "0", rstLaudo!SaldoTransito)
                        Call Redondeo(ZSaldo)
                        
                        If ZLiberada <> 0 And ZSaldo <> 0 Then
                        
                            XRenglon = XRenglon + 1
                            WVector2.Row = XRenglon
                
                            WVector2.Col = 1
                            WVector2.Text = "Laudo"
                        
                            WVector2.Col = 2
                            WVector2.Text = ZLaudo
                                               
                            WVector2.Col = 3
                            WVector2.Text = ZFecha
                        
                            WVector2.Col = 4
                            WVector2.Text = ZOrden
                        
                            WVector2.Col = 5
                            WVector2.Text = ZCantidad
                
                            WVector2.Col = 6
                            WVector2.Text = ZSaldo
                            
                            WVector2.Col = 7
                            WVector2.Text = ZPArtiOri
                
                            WVector2.Col = 8
                            WVector2.Text = ZLaudo
                            
                            WVector2.Col = 9
                            WVector2.Text = ""
                            
                            WVector2.Col = 10
                            WVector2.Text = ZTransito
                            
                            WVector2.Col = 11
                            WVector2.Text = ZSaldoTransito
                            
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
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovguia = "ListaMovguiaArticuloDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then

        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                        
                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArticulo Then
                    
                        ZArticulo = rstMovguia!Articulo
                        ZCantidad = rstMovguia!Cantidad
                        ZFecha = rstMovguia!Fecha
                        ZCodigo = rstMovguia!Codigo
                        ZMovi = rstMovguia!Movi
                        WDestino = rstMovguia!Destino
                        ZTipomov = rstMovguia!Tipomov
                        ZPartida = IIf(IsNull(rstMovguia!Lote), "o", rstMovguia!Lote)
                        ZPartidaOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                        ZFecha = rstMovguia!Fecha
                        If Val(ZCodigo) > 900000 Then
                            WWTipo = "Prestamo"
                            ZCodigo = ZCodigo - 900000
                                Else
                            WWTipo = "Guia In"
                        End If
                        ZSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(ZSaldo)
                                
                        If rstMovguia!Movi = "E" And ZSaldo <> 0 Then
                            
                            XRenglon = XRenglon + 1
                            WVector2.Row = XRenglon
                
                            WVector2.Col = 1
                            WVector2.Text = WWTipo
                        
                            WVector2.Col = 2
                            WVector2.Text = ZCodigo
                                               
                            WVector2.Col = 3
                            WVector2.Text = ZFecha
                        
                            WVector2.Col = 4
                            WVector2.Text = ""
                        
                            WVector2.Col = 5
                            WVector2.Text = ZCantidad
                
                            WVector2.Col = 6
                            WVector2.Text = ZSaldo
                
                            WVector2.Col = 7
                            WVector2.Text = ZPartidaOri
                            
                            WVector2.Col = 8
                            WVector2.Text = ZPartida
                            
                            WVector2.Col = 9
                            WVector2.Text = ""
                            
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
        rstMovguia.Close
    End If
    
    For Ciclo = 1 To XRenglon
    
        XParam = "'" + WVector2.TextMatrix(Ciclo, 4) + "','" _
                 + WArticulo + "'"
        spInforme = "ListaInformeOrdenArticulo " + XParam
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            WEnvase = Str$(rstInforme!Envase)
            rstInforme.Close
        End If
        
        spEnvase = "ConsultaEnvases " + "'" + WEnvase + "'"
        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvase.RecordCount > 0 Then
            WVector2.TextMatrix(Ciclo, 9) = rstEnvase!Abreviatura
            rstEnvase.Close
        End If
        
        WLote = WVector2.TextMatrix(Ciclo, 7)
        WTermi = "DY-00" + Mid$(WArticulo, 4, 7)
        
        XParam = "'" + WLote + "'"
        spLaudo = "ListaLaudoPartiOri " + XParam
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
        
            WLote = IIf(IsNull(rstLaudo!Laudo), "", rstLaudo!Laudo)
            rstLaudo.Close
            
                Else
                
            Rem WLote = WVector2.TextMatrix(Ciclo, 7)
            Rem WTermi = "DW-00" + Mid$(WArticulo, 4, 7)
        
            Rem XParam = "'" + WLote + "'"
            Rem spLaudo = "ListaLaudoPartiOri " + XParam
            Rem Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstLaudo.RecordCount > 0 Then
            Rem     WLote = IIf(IsNull(rstLaudo!Laudo), "", rstLaudo!Laudo)
            Rem     rstLaudo.Close
            Rem End If
                
        End If
        
    Next Ciclo
    
    WVector2.Col = 1
    WVector2.Row = 1
    
    WVector2.TopRow = 1
    
End Sub

Private Sub Limpia_Vector2II()

    WVector2.Height = 4095
    WVector2.Left = 120
    WVector2.Top = 1350
    WVector2.Width = 12000

    WVector2.Clear
    WVector2.Font.Bold = True
    
    WVector2.FixedCols = 1
    WVector2.Cols = 12
    WVector2.FixedRows = 1
    WVector2.Rows = 5001
    
    WVector2.ColWidth(0) = 200
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Tipo"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector2.Text = "Numero"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                WVector2.Text = "Fecha"
                WVector2.ColWidth(Ciclo) = 1200
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                WVector2.Text = "Orden"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WVector2.Text = "Cantidad"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector2.Text = "Saldo"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector2.Text = "Part.Orig"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                WVector2.Text = "Partida"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 9
                WVector2.Text = "Envase"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 10
                WVector2.Text = "Nro. Transito"
                WVector2.ColWidth(Ciclo) = 2000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 11
                WVector2.Text = "Stock"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Visible = True
    
    WVector2.Col = 1
    WVector2.Row = 1
    
End Sub


Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
    
    RenglonAuxiliar = WVector1.Row

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
        
        WVector1.Row = iRow
            
        WVector1.Col = 1
        PArticulo = WVector1.Text
        
        WVector1.Col = 3
        XCantidad = WVector1.Text
            
        If PArticulo <> "" Or XCantidad <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        If Ciclo <> RenglonAuxiliar Then
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

Private Sub WTexto2_DblClick()

    Opcion.Clear
    
     Opcion.AddItem "Productos Articulos"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos Articulos a Utilizar"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click
    
    
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
    WVector1.Cols = 6
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
                WVector1.Text = "P.Articulo"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 2500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 70
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "######.##"
            Case 4
                WVector1.Text = "Partida"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 20
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Nro.Transito"
                WVector1.ColWidth(Ciclo) = 2600
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 20
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

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub

Private Sub WVector2_Click()
    WVector2.Visible = False
    WPartida = Trim(WVector2.TextMatrix(WVector2.Row, 7))
    WTransito = Trim(WVector2.TextMatrix(WVector2.Row, 10))
    WVector1.Col = 4
    WVector1.Text = WPartida
    WVector1.Col = 5
    WVector1.Text = WTransito
    WVector1.Col = 4
    WVector1.Text = WPartida
    WVector1.Row = WVector1.Row + 1
    WVector1.Col = 1
    Call StartEdit
End Sub
