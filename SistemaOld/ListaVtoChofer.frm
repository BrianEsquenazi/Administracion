VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaVtoChofer 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Vencimientos de Choferes"
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
      Left            =   6600
      Top             =   720
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
      Left            =   6360
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaVtoChofer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vector(10000, 6) As String
Dim rstChofer As Recordset
Dim spChofer As String

Private Sub Acepta_Click()

    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    Listado.WindowTitle = "Listado de Vencimiento de Choferes"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    ZZTitulo = "desde el " + Desde.Text + " al " + Hasta.Text
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Chofer SET "
    ZSql = ZSql + " Lista = " + "'" + "" + "',"
    ZSql = ZSql + " ListaI = " + "'" + "" + "',"
    ZSql = ZSql + " ListaII = " + "'" + "" + "',"
    ZSql = ZSql + " ListaIII = " + "'" + "" + "',"
    ZSql = ZSql + " Titulo = " + "'" + ZZTitulo + "'"
    spChofer = ZSql
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    
    Erase Vector
    ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Chofer"
    spChofer = ZSql
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    If rstChofer.RecordCount > 0 Then
    
        With rstChofer
            .MoveFirst
            Do
            
                WFechaVtoI = IIf(IsNull(rstChofer!FechaVtoI), "00/00/0000", rstChofer!FechaVtoI)
                WFechaVtoII = IIf(IsNull(rstChofer!FechaVtoII), "00/00/0000", rstChofer!FechaVtoII)
                WFechaVtoIII = IIf(IsNull(rstChofer!FechaVtoIII), "00/00/0000", rstChofer!FechaVtoIII)
                
                WOrdFechaVtoI = Right$(WFechaVtoI, 4) + Mid$(WFechaVtoI, 4, 2) + Left$(WFechaVtoI, 2)
                WOrdFechaVtoII = Right$(WFechaVtoII, 4) + Mid$(WFechaVtoII, 4, 2) + Left$(WFechaVtoII, 2)
                WOrdFechaVtoIII = Right$(WFechaVtoIII, 4) + Mid$(WFechaVtoIII, 4, 2) + Left$(WFechaVtoIII, 2)
                
                ZEntra = "N"
                
                ZEntraI = "N"
                ZEntraII = "N"
                ZEntraIII = "N"
                
                If WOrdFechaVtoI >= WDesde And WOrdFechaVtoI <= WHasta Then
                    ZEntra = "S"
                    ZEntraI = "S"
                End If
                If WOrdFechaVtoII >= WDesde And WOrdFechaVtoII <= WHasta Then
                    ZEntra = "S"
                    ZEntraII = "S"
                End If
                If WOrdFechaVtoIII >= WDesde And WOrdFechaVtoIII <= WHasta Then
                    ZEntra = "S"
                    ZEntraIII = "S"
                End If
                
                If ZEntra = "S" Then
                    ZLugar = ZLugar + 1
                    Vector(ZLugar, 1) = rstChofer!Codigo
                    Vector(ZLugar, 2) = ZEntraI
                    Vector(ZLugar, 3) = ZEntraII
                    Vector(ZLugar, 4) = ZEntraIII
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        
        rstChofer.Close
        
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZZCodigo = Vector(Ciclo, 1)
        ZZListaI = Vector(Ciclo, 2)
        ZZListaII = Vector(Ciclo, 3)
        ZZListaIII = Vector(Ciclo, 4)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Chofer SET "
        ZSql = ZSql + " Lista = " + "'" + "S" + "',"
        ZSql = ZSql + " ListaI = " + "'" + ZZListaI + "',"
        ZSql = ZSql + " ListaII = " + "'" + ZZListaII + "',"
        ZSql = ZSql + " ListaIII = " + "'" + ZZListaIII + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ZZCodigo + "'"
        spChofer = ZSql
        Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    Listado.GroupSelectionFormula = "{Chofer.Lista} = " + Chr$(34) + "S" + Chr$(34)
    Listado.SelectionFormula = "{Chofer.Lista} = " + Chr$(34) + "S" + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.ReportFileName = "ListaVtoChofer.rpt"
    
    Listado.SQLQuery = "SELECT Chofer.Codigo, Chofer.Descripcion, Chofer.FechaVtoI, Chofer.FechaVtoII, Chofer.FechaVtoIII, Chofer.Proveedor, Chofer.Lista, Chofer.Titulo, Chofer.ListaI, Chofer.ListaII, Chofer.ListaIII, " _
            + "Proveedor.Nombre " _
            + "From " _
            + DSQ + ".dbo.Chofer Chofer, " _
            + DSQ + ".dbo.Proveedor Proveedor " _
            + "Where " _
            + "Chofer.Proveedor = Proveedor.Proveedor AND " _
            + "Chofer.Lista = 'S'"
                        
    Listado.Connect = Connect()
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    PrgListaVtoChofer.Hide
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

Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

