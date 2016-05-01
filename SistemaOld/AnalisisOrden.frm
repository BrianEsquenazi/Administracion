VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAnalisisOrden 
   AutoRedraw      =   -1  'True
   Caption         =   "Analisis de Cumplimiento de Ordenes de Compra"
   ClientHeight    =   3495
   ClientLeft      =   2025
   ClientTop       =   1050
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3495
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2280
         TabIndex        =   9
         Top             =   1200
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
         Left            =   2280
         TabIndex        =   0
         Top             =   720
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
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   2160
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
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   2160
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
         Left            =   4200
         TabIndex        =   6
         Top             =   600
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
         Left            =   4200
         TabIndex        =   5
         Top             =   1080
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
         Left            =   240
         TabIndex        =   4
         Top             =   1200
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
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7080
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wlistinf.rpt"
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
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgAnalisisOrden"
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
Dim WVector(10000, 10) As String
Dim WDevuelta As String
Dim WLiberada As String
Dim WPartida1 As String
Dim WPartida2 As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstLaudo As Recordset
Dim spLaudo As String

Private Sub Acepta_Click()

    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    Sql1 = "UPDATE Orden SET "
    Sql2 = " Suma1 = 0,"
    Sql3 = " Suma2 = 0,"
    Sql4 = " Dias = 0"
    spOrden = Sql1 + Sql2 + Sql3 + Sql4
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem dada dada
    
    Erase WVector
    
    Sql1 = "Select Clave, Orden, Articulo, Cantidad, FechaOrd, Fecha"
    Sql2 = " FROM Orden"
    Sql3 = " Where Orden.FechaOrd >= " + "'" + WDesde + "'"
    Sql4 = " and Orden.FechaOrd <= " + "'" + WHasta + "'"
    spOrden = Sql1 + Sql2 + Sql3 + Sql4
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    WVector(Renglon, 1) = rstOrden!Clave
                    WVector(Renglon, 2) = rstOrden!Orden
                    WVector(Renglon, 3) = rstOrden!Articulo
                    WVector(Renglon, 4) = rstOrden!Cantidad
                    WVector(Renglon, 5) = rstOrden!Fecha
                    WVector(Renglon, 6) = rstOrden!FechaOrd
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If

Stop

    Listado.WindowTitle = "Analisis de Cumplimiento de Ordenes de Compra"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
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
    
    Listado.SQLQuery = "SELECT Informe.Informe, Informe.Fecha, Informe.Remito, Informe.Proveedor, Informe.Orden, Informe.Articulo, Informe.Cantidad, Informe.Fechaord, Informe.FechaOrden, Informe.Difefecha, " _
                    + "Articulo.Descripcion, " _
                    + "Proveedor.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Informe Informe, " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Proveedor Proveedor " _
                    + "Where " _
                    + "Informe.Articulo = Articulo.Codigo AND Informe.Proveedor = Proveedor.Proveedor AND " _
                    + "Informe.Proveedor >= '" + DesdeProv.Text + "' AND Informe.Proveedor <= '" + HastaProv.Text + "' AND " _
                    + "Informe.Fechaord >= '" + WDesde + "' AND Informe.Fechaord <= '" + WHasta + "'"
                        
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
    PrgAnalisisOrden.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgAnalisisOrden.Caption = "Analisis de Cumplimiento de Ordenes de Compra :  " + !Nombre
        End If
    End With
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

