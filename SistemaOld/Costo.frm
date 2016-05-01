VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCosto 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Costos Historicos de Productos"
   ClientHeight    =   4125
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4125
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox TipoListado 
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
         Left            =   2160
         TabIndex        =   12
         Top             =   2640
         Width           =   2655
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   2040
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
         Left            =   2160
         TabIndex        =   9
         Top             =   1680
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   2160
         TabIndex        =   0
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   327680
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
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
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
         Left            =   3600
         TabIndex        =   6
         Top             =   3240
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
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
         Left            =   1680
         TabIndex        =   5
         Top             =   3240
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
         Left            =   4320
         TabIndex        =   4
         Top             =   600
         Width           =   1215
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
         Left            =   4320
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2160
         TabIndex        =   13
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   327680
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
      Begin VB.Label Label5 
         Caption         =   "Desde Producto"
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
         Left            =   480
         TabIndex        =   14
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label2 
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
         Left            =   480
         TabIndex        =   11
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Producto"
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
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7800
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wcosto.rpt"
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
End
Attribute VB_Name = "PrgCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XParam As String
Private Producto As String
Private Costo As Double
Private Auxiliar(100, 7) As String
Private Vector(10000) As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstComposicion As Recordset
Dim spComposicion As String

Private Sub Acepta_Click()

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Hoja SET "
    ZSql = ZSql + " Hoja.PorceDife = (Hoja.Real - Hoja.Teorico) / (Hoja.Teorico/100)" + ","
    ZSql = ZSql + " Hoja.ImpreReal = Hoja.Real"
    ZSql = ZSql + " Where Hoja.Teorico <> 0 and Hoja.Real <> 0"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Hoja SET "
    ZSql = ZSql + " Hoja.PorceDife = (Hoja.Realant - Hoja.Teorico) / (Hoja.Teorico/100)" + ","
    ZSql = ZSql + " Hoja.ImpreReal = Hoja.Realant"
    ZSql = ZSql + " Where Hoja.Teorico <> 0 and Hoja.Realant <> 0"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    If TipoListado.ListIndex <> 2 Then
    
        Erase Vector
        Renglon = 0
        
        spTerminado = "ListaTerminado"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            With rstTerminado
                .MoveFirst
            
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        If rstTerminado!Codigo >= Desde.Text And rstTerminado!Codigo <= Hasta.Text Then
                            Renglon = Renglon + 1
                            Vector(Renglon) = rstTerminado!Codigo
                        End If
                
                        .MoveNext
                
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
            
            End With
            rstTerminado.Close
        End If
        
        ZRenglon = Renglon
    
        For ZDa = 1 To ZRenglon
            Call Calcula_Costo(Vector(ZDa), Costo)
            WCosto = Str$(Costo)
        
            XParam = "'" + Vector(ZDa) + "','" _
                         + WCosto + "'"
            spTerminado = "ModificaTerminadoCosto " + XParam
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        Next ZDa
        
    End If
    
    WDesdeFecha = Right$(DesdeFecha.Text, 4) + Mid$(DesdeFecha.Text, 4, 2) + Left$(DesdeFecha.Text, 2)
    WHastaFecha = Right$(HastaFecha.Text, 4) + Mid$(HastaFecha.Text, 4, 2) + Left$(HastaFecha.Text, 2)
    
    Listado.WindowTitle = "Listado de Costos Historicos de Productos Terminados"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.Connect = Connect()
    
    Select Case TipoListado.ListIndex
        Case 0
            Listado.GroupSelectionFormula = "{Hoja.FechaIngOrd} in " + Chr$(34) + WDesdeFecha + Chr$(34) + " to " + Chr$(34) + WHastaFecha + Chr$(34) + " and {Hoja.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
            Listado.SelectionFormula = "{Hoja.FechaIngOrd} in " + Chr$(34) + WDesdeFecha + Chr$(34) + " to " + Chr$(34) + WHastaFecha + Chr$(34) + " and {Hoja.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
            Listado.SQLQuery = "SELECT Hoja.Clave, Hoja.Hoja, Hoja.Producto, Hoja.Cantidad, Hoja.Tipo, Hoja.Articulo, Hoja.Terminado, Hoja.Teorico, Hoja.Real, Hoja.FechaIng, Hoja.FechaIngOrd, Hoja.Costo1, Hoja.Costo2, Hoja.PorceDife, Hoja.ImpreReal, " _
                        + "Terminado.Costo " _
                        + "From " _
                        + DSQ + ".dbo.Hoja Hoja, " _
                        + DSQ + ".dbo.Terminado Terminado " _
                        + "Where " _
                        + "Hoja.Producto = Terminado.Codigo AND " _
                        + "Hoja.Producto >= '" + Desde.Text + "' AND " _
                        + "Hoja.Producto <= '" + Hasta.Text + "' AND " _
                        + "Hoja.FechaIngOrd >= '" + WDesdeFecha + "' AND " _
                        + "Hoja.FechaIngOrd <= '" + WHastaFecha + "'"
            
            Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
            Listado.ReportFileName = "WCosto.rpt"
            Listado.Action = 1
            
        Case 1
            Listado.GroupSelectionFormula = "{Hoja.FechaIngOrd} in " + Chr$(34) + WDesdeFecha + Chr$(34) + " to " + Chr$(34) + WHastaFecha + Chr$(34) + " and {Hoja.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
            Listado.SelectionFormula = "{Hoja.FechaIngOrd} in " + Chr$(34) + WDesdeFecha + Chr$(34) + " to " + Chr$(34) + WHastaFecha + Chr$(34) + " and {Hoja.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
            Listado.SQLQuery = "SELECT Hoja.Clave, Hoja.Hoja, Hoja.Producto, Hoja.Cantidad, Hoja.Tipo, Hoja.Articulo, Hoja.Terminado, Hoja.Teorico, Hoja.Real, Hoja.FechaIng, Hoja.FechaIngOrd, Hoja.Costo1, Hoja.Costo2, Hoja.Realant, Hoja.PorceDife, Hoja.MotivoDesvio, Hoja.ObservaDesvio, Hoja.ImpreReal, " _
                        + "Terminado.Costo " _
                        + "From " _
                        + DSQ + ".dbo.Hoja Hoja, " _
                        + DSQ + ".dbo.Terminado Terminado " _
                        + "Where " _
                        + "Hoja.Producto = Terminado.Codigo AND " _
                        + "Hoja.Producto >= '" + Desde.Text + "' AND " _
                        + "Hoja.Producto <= '" + Hasta.Text + "' AND " _
                        + "Hoja.FechaIngOrd >= '" + WDesdeFecha + "' AND " _
                        + "Hoja.FechaIngOrd <= '" + WHastaFecha + "' AND " _
                        + "(Hoja.PorceDife <= -3 OR Hoja.PorceDife >= 3)"
            
            Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
            Listado.ReportFileName = "WCostoII.rpt"
            Listado.Action = 1
            
        Case Else
            Listado.GroupSelectionFormula = "{Hoja.FechaIngOrd} in " + Chr$(34) + WDesdeFecha + Chr$(34) + " to " + Chr$(34) + WHastaFecha + Chr$(34) + " and {Hoja.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
            Listado.SelectionFormula = "{Hoja.FechaIngOrd} in " + Chr$(34) + WDesdeFecha + Chr$(34) + " to " + Chr$(34) + WHastaFecha + Chr$(34) + " and {Hoja.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
            Listado.SQLQuery = "SELECT Hoja.Clave, Hoja.Hoja, Hoja.Producto, Hoja.Cantidad, Hoja.Tipo, Hoja.Articulo, Hoja.Terminado, Hoja.Teorico, Hoja.Real, Hoja.FechaIng, Hoja.FechaIngOrd, Hoja.Costo1, Hoja.Costo2, Hoja.PorceDife, Hoja.MotivoDesvio, Hoja.ObservaDesvio, Hoja.ImpreReal, " _
                        + "Terminado.Costo " _
                        + "From " _
                        + DSQ + ".dbo.Hoja Hoja, " _
                        + DSQ + ".dbo.Terminado Terminado " _
                        + "Where " _
                        + "Hoja.Producto = Terminado.Codigo AND " _
                        + "Hoja.Producto >= '" + Desde.Text + "' AND " _
                        + "Hoja.Producto <= '" + Hasta.Text + "' AND " _
                        + "Hoja.FechaIngOrd >= '" + WDesdeFecha + "' AND " _
                        + "Hoja.FechaIngOrd <= '" + WHastaFecha + "' AND " _
                        + "(Hoja.PorceDife <= -3 OR Hoja.PorceDife >= 3)"
            
            Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
            Listado.ReportFileName = "WCostoIII.rpt"
            Listado.Action = 1
            
    End Select
            
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgCosto.Hide
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

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        DesdeFecha.SetFocus
    End If
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.SetFocus
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()

    TipoListado.Clear
    
    TipoListado.AddItem "Completo"
    TipoListado.AddItem "Desvio 3%"
    TipoListado.AddItem "Desvio 3% Resumido"
    
    TipoListado.ListIndex = 0

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgCosto.Caption = "Listado de composicion de Productos Terminados :  " + !Nombre
        End If
    End With

    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
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
    
            Entra = "S"
            
            spComposicion = "ConsultaComposicionProducto " + "'" + Vector(Cicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Entra = "N"
                        
                        Tipo = rstComposicion!Tipo
                        Articulo1 = rstComposicion!Articulo1
                        Articulo2 = rstComposicion!Articulo2
                        Cantidad = rstComposicion!Cantidad
                        
                        Rem If Left$(Articulo1, 2) = "DW" Then
                        Rem     Tipo = "T"
                        Rem     Articulo2 = Left$(Articulo1, 3) + "00" + Right$(Articulo1, 7)
                        Rem End If
                        
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
                                Auxiliar(Renglon, 2) = Str$(Cantidad)
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
            
            Rem If Entra = "S" And Left$(Vector(Cicla, 1), 2) = "DW" Then
            Rem     Renglon = Renglon + 1
            Rem     Auxiliar(Renglon, 1) = Left$(Vector(Cicla, 1), 3) + Right$(Vector(Cicla, 1), 7)
            Rem     Auxiliar(Renglon, 2) = "1"
            Rem     Auxiliar(Renglon, 3) = Vector(Cicla, 2)
            Rem End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For Da = 1 To Renglon
        Articulo = Auxiliar(Da, 1)
        Cantidad = Val(Auxiliar(Da, 2))
        XVector = Auxiliar(Da, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCosto = (Cantidad * rstArticulo!Costo2 * Val(XVector))
            Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(XVector))
            rstArticulo.Close
        End If
    Next Da
    
End Sub


