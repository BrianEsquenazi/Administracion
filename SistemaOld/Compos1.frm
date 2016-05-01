VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCompos1 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Composicion de Productos Terminados"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1935
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox TipoListado 
         Height          =   315
         Left            =   1800
         TabIndex        =   14
         Top             =   1080
         Width           =   2415
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1800
         TabIndex        =   12
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Listado "
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Producto"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Producto"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wcompos1.rpt"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "Compos1.frx":0000
      Left            =   120
      List            =   "Compos1.frx":0007
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   5760
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgCompos1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Producto As String
Private Costo As Double
Private Costo1 As Double
Private Costo2 As Double
Private WCosto1 As String
Private WCosto2 As String
Private Auxiliar(100, 7) As String
Private XVector(20000, 6) As String
Private lista(20000) As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim XParam As String

Private Sub Acepta_Click()

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    Erase lista
    Renglon = 0
    
    XParam = "'" + Desde.Text + "','" _
                + Hasta.Text + "'"
                                         
    Set rstTerminado = db.OpenRecordset("ListaTerminadoDesdeHasta " + XParam, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then

    With rstTerminado
    
        .MoveFirst
        If .NoMatch = False Then
        
            Do
            
                Renglon = Renglon + 1
                    
                lista(Renglon) = rstTerminado!Codigo
                    
                .MoveNext
                   
                If .EOF = True Then
                    Exit Do
                End If
                        
            Loop
            
        End If
            
    End With
    rstTerminado.Close
    
    End If
    
    Total = Renglon
    
    Erase XVector
    Renglon = 0
    
    For Cicla = 1 To Total
    
        WCodigo = lista(Cicla)
        WHoja = 0
    
        spHoja = "ConsultaHojaEspecial " + "'" + WCodigo + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            With rstHoja
                .MoveFirst
                Do
                    If .EOF = False Then
                        WHoja = rstHoja!Hoja
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstHoja.Close
        End If
        
        spHoja = "ListaHoja " + "'" + WHoja + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            With rstHoja
                .MoveFirst
                Do
                    If .EOF = False Then
            
                        Renglon = Renglon + 1
                    
                        XVector(Renglon, 1) = rstHoja!Tipo
                        XVector(Renglon, 2) = rstHoja!Articulo
                        XVector(Renglon, 3) = rstHoja!Terminado
                        XVector(Renglon, 4) = rstHoja!Cantidad
                        XVector(Renglon, 5) = rstHoja!Clave
                        XVector(Renglon, 6) = WCodigo
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstHoja.Close
        End If
        
    Next Cicla
        
    For Da = 1 To Renglon
    
        Tipo = XVector(Da, 1)
        Articulo1 = XVector(Da, 2)
        Articulo2 = XVector(Da, 3)
        Cantidad = Val(Vector(Da, 4))
        Clave = XVector(Da, 5)
        Terminado = XVector(Da, 6)
        
        DescriTerminado = ""
        DescriArticulo1 = ""
        DescriArticulo2 = ""
        
        Select Case Tipo
            Case "T"
                Producto = Articulo2
                Call Calcula_Costo(Producto, Costo)
                spTerminado = "ConsultaTerminado " + "'" + Articulo2 + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                        DescriArticulo2 = Left$(rstTerminado!Descripcion, 30)
                        rstTerminado.Close
                End If
                
                
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + Articulo1 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                        If TipoListado.ListIndex = 0 Then
                            Costo = rstArticulo!Costo2
                                Else
                            Costo = rstArticulo!Costo1
                        End If
                        DescriArticulo1 = Left$(rstArticulo!Descripcion, 30)
                        rstArticulo.Close
                End If
            Case Else
        End Select
        
        spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DescriTerminado = Left$(rstTerminado!Descripcion, 30)
            rstTerminado.Close
        End If
        
        Costo1 = Costo
        Call Redondeo(Costo1)
        WCosto1 = Costo1
        Costo2 = Costo * Cantidad
        Call Redondeo(Costo2)
        WCosto2 = Costo2
        WCosto1 = Pusing("###,###.##", WCosto1)
        WCosto2 = Pusing("###,###.##", WCosto2)
            
        XParam = "'" + Clave + "','" _
                    + WCosto1 + "','" _
                    + WCosto2 + "','" _
                    + DescriTerminado + "','" _
                    + DescriArticulo1 + "','" _
                    + DescriArticulo2 + "'"
                                           
                     dada
                                           
        spComposicion = "ModificaComposicionCosto " + XParam
        Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Da

    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            If TipoListado.ListIndex = 0 Then
                !Varios = "(Costo Standard)"
                    Else
                !Varios = "(Costo Ultima Compra)"
            End If
            .Update
        End If
    End With

    Listado.WindowTitle = "Listado de Composicion de Productos Terminados"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Composicion.terminado} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.Connect = Connect()
    
    Listado.SQLQuery = "SELECT Composicion.Clave , Composicion.Terminado, Composicion.Tipo, Composicion.Articulo1, Composicion.Articulo2, Composicion.Cantidad, Composicion.Costo1, Composicion.Costo2, Composicion.DescriTerminado, Composicion.DescriArticulo1, Composicion.DescriArticulo2 " + _
                        "From " + DSQ + ".dbo.Composicion Composicion " + _
                        "Where Composicion.Terminado >= '" + Desde.Text + "' AND Composicion.Terminado <= '" + Hasta.Text + "'"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgCompos.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.Text = Desde.Text
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Auxiliar
    OPEN_FILE_Empresa
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgCompos.Caption = "Listado de composicion de Productos Terminados :  " + !Nombre
        End If
    End With
    
    
    TipoListado.Clear
    
    TipoListado.AddItem "Costo Standard"
    TipoListado.AddItem "Costo Ultima Compra"
    
    TipoListado.ListIndex = 0

    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub
Private Sub Consulta_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    XIndice = 0
    
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
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub


Private Sub Pantalla_Click()

    Pantalla.Visible = False
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                    Desde.Text = rstTerminado!Codigo
                    Hasta.Text = rstTerminado!Codigo
                    rstTerminado.Close
            End If
            Desde.SetFocus
    End Select
    
End Sub

Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim Vector(100, 2) As String
    Erase Auxiliar
    Renglon = 0
    
    Vector(1, 1) = Producto
    Vector(1, 2) = "1"
    Costo = 0
    Lugar = 1
    Cicla = 0
    
    Do
        Cicla = Cicla + 1
        If Vector(Cicla, 1) <> "" Then
    
            spComposicion = "ConsultaComposicionProducto " + "'" + Vector(Cicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Tipo = rstComposicion!Tipo
                        Articulo1 = rstComposicion!Articulo1
                        Articulo2 = rstComposicion!Articulo2
                        Cantidad = rstComposicion!Cantidad
                        
                        Select Case Tipo
                            Case "T"
                                If Producto <> Articulo2 Then
                                    Lugar = Lugar + 1
                                    Vector(Lugar, 1) = Articulo2
                                    Vector(Lugar, 2) = Str$(Cantidad * Val(Vector(Cicla, 2)))
                                End If
                            Case "M"
                                Renglon = Renglon + 1
                                Auxiliar(Renglon, 1) = Articulo1
                                Auxiliar(Renglon, 2) = Cantidad
                                Auxiliar(Renglon, 3) = Vector(Cicla, 2)
                            Case Else
                        End Select
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstComposicion.Close
            End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For Da = 1 To Renglon
        Articulo = Auxiliar(Da, 1)
        Cantidad = Val(Auxiliar(Da, 2))
        WVector = Auxiliar(Da, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            If TipoListado.ListIndex = 0 Then
                WCosto = (Cantidad * rstArticulo!Costo2 * Val(WVector))
                Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(WVector))
                    Else
                WCosto = (Cantidad * rstArticulo!Costo1 * Val(WVector))
                Costo = Costo + (Cantidad * rstArticulo!Costo1 * Val(WVector))
            End If
            rstArticulo.Close
        End If
    Next Da
    
End Sub


