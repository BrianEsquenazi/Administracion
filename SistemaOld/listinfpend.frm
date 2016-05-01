VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form prglistinfpend 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Informe de Recepcion Pendientes de Aprobacion"
   ClientHeight    =   2400
   ClientLeft      =   1875
   ClientTop       =   2520
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   2400
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   5535
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1800
         TabIndex        =   9
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
         Left            =   2640
         TabIndex        =   8
         Top             =   1440
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
         Left            =   1080
         TabIndex        =   7
         Top             =   1440
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
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   360
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
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6840
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wlistinfpend.rpt"
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
      Left            =   6600
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "prglistinfpend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WTerminado As String
Private WInicial As Double
Private WEntrada As Double
Private WSalida As Double
Private WTipo As Integer
Private WNumero As String
Private Impre1 As String
Private Impre2 As String
Private WFecha As String
Dim Vector(10000, 4) As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstArticulo As Recordset
Dim spArticulo As String

Private Sub Acepta_Click()

    Erase Vector
    Renglon = 0
    
    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    spInforme = "ModificaInformeProcesoSaldo"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
    spArticulo = "ModificaArticuloLaboratorio0"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    XParam = "'" + "20020101" + "'"
    spInforme = "ModificaInformeProceso0 " + XParam
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
    XParam = "'" + WDesde + "','" _
                 + WHasta + "'"
    spInforme = "ListaInformeDesdeHastaFecha" + XParam
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        With rstInforme
            .MoveFirst
            Do
                If .EOF = False Then
                    If !FechaOrd > "20020101" Then
                        Renglon = Renglon + 1
                        Vector(Renglon, 1) = rstInforme!Clave
                        Vector(Renglon, 2) = rstInforme!Informe
                        Vector(Renglon, 3) = rstInforme!Articulo
                        Vector(Renglon, 4) = rstInforme!Cantidad
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstInforme.Close
    End If
    
    For Ciclo = 1 To Renglon
    
        WClave = Vector(Ciclo, 1)
        WInforme = Vector(Ciclo, 2)
        WArticulo = Vector(Ciclo, 3)
        WCantidad = Val(Vector(Ciclo, 4))
        WResta = 0
        
        XParam = "'" + WInforme + "','" _
                 + WArticulo + "'"
        spLaudo = "ListaLaudoInforme " + XParam
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
    
            With rstLaudo
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        WLiberada = rstLaudo!Liberada
                        WDevuelta = rstLaudo!devuelta
                        WSuma = WLiberada + WDevuelta
                        
                        WLiberadaAnt = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)
                        WDevueltaAnt = IIf(IsNull(rstLaudo!devueltaant), "0", rstLaudo!devueltaant)
                        WSumaAnt = WLiberadaAnt + WDevueltaAnt
                        
                        If WSumaAnt <> 0 Then
                            WResta = WResta + WSumaAnt
                                Else
                            WResta = WResta + WSuma
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
        
        XParam = "'" + WClave + "','" _
                 + Str$(WResta) + "'"
        spInforme = "ModificaInformeProceso " + XParam
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        
        WDife = WCantidad - WResta
        If WDife > 0 Then
            XParam = "'" + WArticulo + "','" _
                         + Str$(WDife) + "'"
            spArticulo = "ModificaArticuloLaboratorio " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Next Ciclo
    
    spInforme = "ModificaInformeProcesoDife"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)

    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    Listado.WindowTitle = "Listado de Informe de Recepcion Pendientes de Aprobacion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Informe.fechaord} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Informe.Informe, Informe.Fecha, Informe.Proveedor, Informe.Orden, Informe.Articulo, Informe.Cantidad, Informe.Fechaord, Informe.CantidadLaudo, Informe.Dife, Informe.Certificado1, Informe.Estado1, " _
                        + "Articulo.Descripcion, " _
                        + "Proveedor.Nombre " _
                        + "From " _
                        + DSQ + ".dbo.Informe Informe, " _
                        + DSQ + ".dbo.Articulo Articulo, " _
                        + DSQ + ".dbo.Proveedor Proveedor " _
                        + "Where " _
                        + "Informe.Articulo = Articulo.Codigo AND " _
                        + "Informe.Proveedor = Proveedor.Proveedor AND " _
                        + "Informe.Fechaord >= '" + WDesde + "' AND Informe.Fechaord <= '" + WHasta + "' AND " _
                        + "Informe.Dife <> 0."
                        
    Listado.DataFiles(3) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    
    DbsEmpresa.Close
    
    Desde.SetFocus
    prglistinfpend.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
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

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub


Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            prglistinfpend.Caption = "Listado de Informe de Recepcion Pendientes de Aprobacion :  " + !Nombre
        End If
    End With
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

