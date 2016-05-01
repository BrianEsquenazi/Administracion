VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAju 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Muestras a Clientes"
   ClientHeight    =   8325
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11790
   LinkTopic       =   "Form2"
   ScaleHeight     =   8325
   ScaleWidth      =   11790
   Begin VB.TextBox AnoII 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10680
      MaxLength       =   4
      TabIndex        =   39
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Ayuda 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   38
      Text            =   " "
      Top             =   840
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Frame PantaRemito 
      Height          =   7095
      Left            =   4440
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox DesClienteRemito 
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
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   25
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton ConfirmaRemito 
         Caption         =   "Confirma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   20
         Top             =   6360
         Width           =   1455
      End
      Begin VB.CommandButton CancelaRemito 
         Caption         =   "Cancela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   19
         Top             =   6360
         Width           =   1455
      End
      Begin VB.TextBox NumeroRemito 
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
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin MSMask.MaskEdBox FechaRemito 
         Height          =   285
         Left            =   1800
         TabIndex        =   21
         Top             =   720
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
      Begin MSFlexGridLib.MSFlexGrid VectorRemito 
         Height          =   4695
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   8281
         _Version        =   327680
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente"
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
         TabIndex        =   24
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Nro. Remito"
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
         TabIndex        =   23
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha "
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
         TabIndex        =   22
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.TextBox Ano 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10680
      MaxLength       =   4
      TabIndex        =   37
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame PantaCantiEtiqueta 
      Height          =   1455
      Left            =   4320
      TabIndex        =   33
      Top             =   2760
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox CantiEtiqueta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   35
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Cantidad"
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
         Left            =   240
         TabIndex        =   34
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame PantaEtiqueta 
      Height          =   6735
      Left            =   1080
      TabIndex        =   27
      Top             =   1440
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton CancelaEtiqueta 
         Caption         =   "Cancela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   29
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton ConfirmaEtiqueta 
         Caption         =   "Confirma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   28
         Top             =   6240
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid VectorEtiqueta 
         Height          =   4815
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   8493
         _Version        =   327680
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "IMPRESION DE ETIQUETAS DE MUESTRA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.CommandButton Etiqueta 
      Caption         =   "Etiquetas (F6)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   5400
      TabIndex        =   31
      Top             =   0
      Width           =   1215
   End
   Begin Crystal.CrystalReport ListaRemito 
      Left            =   9960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      CopiesToPrinter =   2
   End
   Begin VB.CommandButton Remito 
      Caption         =   "Remito (F8)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   6720
      TabIndex        =   16
      Top             =   0
      Width           =   1215
   End
   Begin VB.Frame PantaExporta 
      Height          =   4695
      Left            =   2280
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton ConfirmaExporta 
         Caption         =   "Confirma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   15
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton CancelaExporta 
         Caption         =   "Cancela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   14
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox NombreExporta 
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Top             =   360
         Width           =   2655
      End
      Begin VB.DirListBox Dir1 
         Height          =   2340
         Left            =   2760
         TabIndex        =   11
         Top             =   840
         Width           =   2655
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   600
         TabIndex        =   10
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
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
         Left            =   600
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Exportaii 
      Caption         =   "Exportacion (F5)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   4080
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin Crystal.CrystalReport ListaGRilla 
      Left            =   9720
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "muestra.rpt"
   End
   Begin VB.ListBox Lista 
      Height          =   645
      Left            =   2640
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox Pantalla 
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
      Height          =   5715
      ItemData        =   "aju.frx":0000
      Left            =   3480
      List            =   "aju.frx":0007
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   4815
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   7335
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12938
      _Version        =   327680
      BackColor       =   16777215
      ForeColor       =   4210752
      FocusRect       =   2
      GridLines       =   0
   End
   Begin VB.CommandButton Labora 
      Caption         =   "    Actualiza Laboratorio (F4)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   2760
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Modifica 
      Caption         =   "Modifica / Baja Muestra   (F3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   1440
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Alta 
      Caption         =   "Alta (F1)"
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
      Height          =   780
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Impresion 
      Caption         =   "Impresion (F9)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   8040
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Fin (F10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   9360
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin Crystal.CrystalReport ListaEtiqueta 
      Left            =   10080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Año"
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
      Left            =   10680
      TabIndex        =   36
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "PrgAju"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstMuestra As Recordset
Dim spMuestra As String
Dim rstMuestraImpre As Recordset
Dim spMuestraImpre As String
Dim rstImpreEtiqueta As Recordset
Dim spImpreEtiqueta As String
Dim XParam As String
Dim Auxiliar(20000)
Dim XEmpresa As String
Dim WFecha As String
Dim WFecha2 As String
Dim SeparaFecha As Integer
Dim SumaDia As Integer
Dim SumaMes As Integer
Dim WDia As String
Dim WMes As String
Dim WCod As String
Dim ColumnaOpcion As Integer
Dim Seleccion As String
Dim WPasa(20000) As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Dim LugarRemito As Integer
Dim WBorra(1000, 10) As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
Dim ZImpreEti(100, 20) As String
Dim ZAno As String
Dim ZAnoII As String

Private Sub Alta_Click()
    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    Muestra.Visible = False
    WMuestra = 0
    PrgMuestraNueva.Show
End Sub


Private Sub CancelaExporta_Click()
    
    PantaExporta.Visible = False
    
End Sub

Private Sub Command1_Click()
    Sql1 = "UPDATE Muestra SET "
    Sql2 = " Pedido = Codigo"
    spMuestra = Sql1 + Sql2
    Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
End Sub

Private Sub ConfirmaExporta_Click()

    If NombreExporta.Text = "" Then
        m$ = "Se debe informar un nombre de archivo"
        A% = MsgBox(m$, 0, "Exportacion de Muestras")
        Exit Sub
    End If

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            WDesEmpresa = !Nombre
        End If
    End With

    spMuestraImpre = "BorrarMuestraImpre "
    Set rstMuestraImpre = db.OpenRecordset(spMuestraImpre, dbOpenSnapshot, dbSQLPassThrough)
    
    rowini = Muestra.Row
    RowFin = Muestra.RowSel
    
    For Ciclo = rowini To RowFin
    
        ZNumero = Str$(Ciclo)
        ZPedido = Left$(Muestra.TextMatrix(Ciclo, 1), 6)
        ZFecha = Left$(Muestra.TextMatrix(Ciclo, 2), 10)
        ZCodigo = Left$(Muestra.TextMatrix(Ciclo, 3), 15)
        ZDescripcion = Left$(Muestra.TextMatrix(Ciclo, 4), 50)
        ZCantidad = Left$(Muestra.TextMatrix(Ciclo, 5), 10)
        ZDescriCliente = Left$(Muestra.TextMatrix(Ciclo, 6), 50)
        ZCliente = Left$(Muestra.TextMatrix(Ciclo, 7), 50)
        ZObservaciones = Left$(Muestra.TextMatrix(Ciclo, 8), 50)
        ZFecha2 = Left$(Muestra.TextMatrix(Ciclo, 9), 10)
        ZRemito = Left$(Muestra.TextMatrix(Ciclo, 10), 10)
        ZHojaRuta = Left$(Muestra.TextMatrix(Ciclo, 11), 10)
        ZCodigo2 = Left$(Muestra.TextMatrix(Ciclo, 12), 15)
        ZDescripcion2 = Left$(Muestra.TextMatrix(Ciclo, 13), 50)
        ZLote = Left$(Muestra.TextMatrix(Ciclo, 14), 10)
        ZObservaciones2 = Left$(Muestra.TextMatrix(Ciclo, 15), 50)
        ZCantidad2 = Left$(Muestra.TextMatrix(Ciclo, 16), 10)
        ZActualiza = Left$(Muestra.TextMatrix(Ciclo, 17), 1)
        
        Sql1 = "INSERT INTO MuestraImpre ("
        Sql2 = "Numero ,"
        Sql3 = "Fecha ,"
        Sql4 = "Codigo ,"
        Sql5 = "Descripcion ,"
        Sql6 = "Cantidad ,"
        Sql7 = "DescriCliente ,"
        Sql8 = "Cliente ,"
        Sql9 = "Observaciones ,"
        Sql10 = "Fecha2 ,"
        Sql11 = "Codigo2 ,"
        Sql12 = "Descripcion2 ,"
        Sql13 = "Lote ,"
        Sql14 = "Observaciones2 ,"
        Sql15 = "Cantidad2 ,"
        Sql16 = "Actualiza ,"
        Sql17 = "DesEmpresa) "
        Sql18 = "Values ("
        Sql19 = "'" + ZNumero + "',"
        Sql20 = "'" + ZFecha + "',"
        Sql21 = "'" + ZCodigo + "',"
        Sql22 = "'" + ZDescripcion + "',"
        Sql23 = "'" + ZCantidad + "',"
        Sql24 = "'" + ZDescriCliente + "',"
        Sql25 = "'" + ZCliente + "',"
        Sql26 = "'" + ZObservaciones + "',"
        Sql27 = "'" + ZFecha2 + "',"
        Sql28 = "'" + ZCodigo2 + "',"
        Sql29 = "'" + ZDescripcion2 + "',"
        Sql30 = "'" + ZLote + "',"
        Sql31 = "'" + ZObservaciones2 + "',"
        Sql32 = "'" + ZCantidad2 + "',"
        Sql33 = "'" + ZActualiza + "',"
        Sql34 = "'" + WDesEmpresa + "')"
       
        spMuestraImpre = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                     Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                     Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                     Sql31 + Sql32 + Sql33 + Sql34
        Set rstMuestraImpre = db.OpenRecordset(spMuestraImpre, dbOpenSnapshot, dbSQLPassThrough)
    Next Ciclo
    
    DoEvents

    ListaGRilla.Destination = 2
    ListaGRilla.PrintFileType = crptExcel50
    ListaGRilla.PrintFileName = Dir1.Path + "\" + NombreExporta.Text + ".xls"
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    ListaGRilla.SQLQuery = "SELECT MuestraImpre.Numero, MuestraImpre.Fecha, MuestraImpre.Codigo, MuestraImpre.Descripcion, MuestraImpre.Cantidad, MuestraImpre.DescriCLiente, MuestraImpre.Cliente, MuestraImpre.Observaciones, MuestraImpre.Fecha2, MuestraImpre.Codigo2, MuestraImpre.Descripcion2, MuestraImpre.Lote, MuestraImpre.Observaciones2, MuestraImpre.Cantidad2 " _
                    + "From " _
                    + DSQ + ".dbo.MuestraImpre MuestraImpre " _
                    + "Where " _
                    + "MuestraImpre.Numero >= 0 AND " _
                    + "MuestraImpre.Numero <= 999999 " _
                    + "Order By MuestraImpre.Numero ASC"
    ListaGRilla.Connect = Connect()
    ListaGRilla.Action = 1
    
    PantaExporta.Visible = False
    
End Sub

Private Sub Exportaii_Click()

    NombreExporta.Text = ""
    Drive1.Drive = "C:"
    Dir1.Path = Drive1.Drive    ' Establece la ruta del directorio.
    PantaExporta.Visible = True

End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive    ' Establece la ruta del directorio.
End Sub

Private Sub Form_Load()
    Provincia(0) = "Capital Federal"
    Provincia(1) = "Buenos Aires"
    Provincia(2) = "Catamarca"
    Provincia(3) = "Cordoba"
    Provincia(4) = "Corrientes"
    Provincia(5) = "Chaco"
    Provincia(6) = "Chubut"
    Provincia(7) = "Entre Rios"
    Provincia(8) = "Formosa"
    Provincia(9) = "Jujuy"
    Provincia(10) = "La Pampa"
    Provincia(11) = "La Rioja"
    Provincia(12) = "Mendoza"
    Provincia(13) = "Misiones"
    Provincia(14) = "Neuquen"
    Provincia(15) = "Rio Negro"
    Provincia(16) = "Salta"
    Provincia(17) = "San Juan"
    Provincia(18) = "San Luis"
    Provincia(19) = "Santa Cruz"
    Provincia(20) = "Santa Fe"
    Provincia(21) = "Santiago del Estero"
    Provincia(22) = "Tucuman"
    Provincia(23) = "Tierra del Fuego"
    Provincia(24) = "Exterior"
    Provincia(25) = ""
    
    Iva(1) = "Inscripto"
    Iva(2) = "No Inscripto"
    Iva(3) = "Inscripto"
    Iva(4) = "Inscripto"
    Iva(5) = "Inscripto"
    Iva(6) = "Inscripto"
End Sub

Private Sub Labora_Click()
    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    Muestra.Visible = False
    Fila = Muestra.Row
    WMuestra = Auxiliar(Fila)
    If Val(WMuestra) <> 0 Then
        PrgMuestraLaboNuevo.Show
    End If
End Sub

Private Sub Modifica_Click()
    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    Muestra.Visible = False
    Fila = Muestra.Row
    WMuestra = Auxiliar(Fila)
    If Val(WMuestra) <> 0 Then
        PrgMuestraNueva.Show
    End If
End Sub

Private Sub Baja_Click()
    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    Fila = Muestra.Row
    WMuestra = Auxiliar(Fila)
    If Val(WMuestra) <> 0 Then
        T$ = "Borrar Registro"
        m$ = "Desea Borrar la muestra"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            spMuestra = "BorrarMuestra " + "'" + WMuestra + "'"
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenDynaset, dbSQLPassThrough)
            Call Proceso_Click
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    PrgAju.Hide
    Unload Me
    Close
    End
End Sub

Private Sub Impresion_Click()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            WDesEmpresa = !Nombre
        End If
    End With

    spMuestraImpre = "BorrarMuestraImpre "
    Set rstMuestraImpre = db.OpenRecordset(spMuestraImpre, dbOpenSnapshot, dbSQLPassThrough)
    
    rowini = Muestra.Row
    RowFin = Muestra.RowSel
    
    For Ciclo = rowini To RowFin
        ZNumero = Str$(Ciclo)
        ZPedido = Left$(Muestra.TextMatrix(Ciclo, 1), 6)
        ZFecha = Left$(Muestra.TextMatrix(Ciclo, 2), 10)
        ZCodigo = Left$(Muestra.TextMatrix(Ciclo, 3), 15)
        ZDescripcion = Left$(Muestra.TextMatrix(Ciclo, 4), 50)
        ZCantidad = Left$(Muestra.TextMatrix(Ciclo, 5), 10)
        ZDescriCliente = Left$(Muestra.TextMatrix(Ciclo, 6), 50)
        ZCliente = Left$(Muestra.TextMatrix(Ciclo, 7), 50)
        ZObservaciones = Left$(Muestra.TextMatrix(Ciclo, 8), 50)
        ZFecha2 = Left$(Muestra.TextMatrix(Ciclo, 9), 10)
        ZRemito = Left$(Muestra.TextMatrix(Ciclo, 10), 10)
        ZHojaRuta = Left$(Muestra.TextMatrix(Ciclo, 11), 10)
        ZCodigo2 = Left$(Muestra.TextMatrix(Ciclo, 12), 15)
        ZDescripcion2 = Left$(Muestra.TextMatrix(Ciclo, 13), 50)
        ZLote = Left$(Muestra.TextMatrix(Ciclo, 14), 10)
        ZObservaciones2 = Left$(Muestra.TextMatrix(Ciclo, 15), 50)
        ZCantidad2 = Left$(Muestra.TextMatrix(Ciclo, 16), 10)
        ZActualiza = Left$(Muestra.TextMatrix(Ciclo, 17), 1)
        
        Sql1 = "INSERT INTO MuestraImpre ("
        Sql2 = "Numero ,"
        Sql3 = "Fecha ,"
        Sql4 = "Codigo ,"
        Sql5 = "Descripcion ,"
        Sql6 = "Cantidad ,"
        Sql7 = "DescriCliente ,"
        Sql8 = "Cliente ,"
        Sql9 = "Observaciones ,"
        Sql10 = "Fecha2 ,"
        Sql11 = "Codigo2 ,"
        Sql12 = "Descripcion2 ,"
        Sql13 = "Lote ,"
        Sql14 = "Observaciones2 ,"
        Sql15 = "Cantidad2 ,"
        Sql16 = "Actualiza ,"
        Sql17 = "DesEmpresa) "
        Sql18 = "Values ("
        Sql19 = "'" + ZNumero + "',"
        Sql20 = "'" + ZFecha + "',"
        Sql21 = "'" + ZCodigo + "',"
        Sql22 = "'" + ZDescripcion + "',"
        Sql23 = "'" + ZCantidad + "',"
        Sql24 = "'" + ZDescriCliente + "',"
        Sql25 = "'" + ZCliente + "',"
        Sql26 = "'" + ZObservaciones + "',"
        Sql27 = "'" + ZFecha2 + "',"
        Sql28 = "'" + ZCodigo2 + "',"
        Sql29 = "'" + ZDescripcion2 + "',"
        Sql30 = "'" + ZLote + "',"
        Sql31 = "'" + ZObservaciones2 + "',"
        Sql32 = "'" + ZCantidad2 + "',"
        Sql33 = "'" + ZActualiza + "',"
        Sql34 = "'" + WDesEmpresa + "')"
       
        spMuestraImpre = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                     Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                     Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                     Sql31 + Sql32 + Sql33 + Sql34
        Set rstMuestraImpre = db.OpenRecordset(spMuestraImpre, dbOpenSnapshot, dbSQLPassThrough)
    Next Ciclo

    ListaGRilla.Destination = 1
    Rem ListaGRilla.Destination = 0
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    ListaGRilla.SQLQuery = "SELECT MuestraImpre.Numero, MuestraImpre.Fecha, MuestraImpre.Codigo, MuestraImpre.Descripcion, MuestraImpre.Cantidad, MuestraImpre.DescriCLiente, MuestraImpre.Cliente, MuestraImpre.Observaciones, MuestraImpre.Fecha2, MuestraImpre.Codigo2, MuestraImpre.Descripcion2, MuestraImpre.Lote, MuestraImpre.Observaciones2, MuestraImpre.Cantidad2 " _
                    + "From " _
                    + DSQ + ".dbo.MuestraImpre MuestraImpre " _
                    + "Where " _
                    + "MuestraImpre.Numero >= 0 AND " _
                    + "MuestraImpre.Numero <= 999999 " _
                    + "Order By MuestraImpre.Numero ASC"
    ListaGRilla.Connect = Connect()
    ListaGRilla.Action = 1
    
End Sub

Private Sub Proceso_Click()

    If Val(Ano.Text) = 0 Or Val(AnoII.Text) = 0 Then
        Exit Sub
    End If

    DesClienteRemito.Text = ""
    Call Limpia_Vector
    
    ZFecDesde = Ano.Text + "0101"
    ZFecHasta = AnoII.Text + "1231"
 
    Select Case ColumnaOpcion
        Case 0, 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Muestra"
            ZSql = ZSql + " Where Muestra.OrdFecha >= " + "'" + ZFecDesde + "'"
            ZSql = ZSql + " and Muestra.OrdFecha <= " + "'" + ZFecHasta + "'"
            ZSql = ZSql + " Order by Muestra.Codigo"
            spMuestra = ZSql
            
        Case 2
            spMuestra = "ListaMuestraFechaSolo " + "'" + Seleccion + "'"
            
        Case 3
            If Left(Seleccion, 2) = "PT" Then
                spTerminado = "ConsultaTerminado " + "'" + Seleccion + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    spMuestra = "ListaMuestraProductoSolo " + "'" + Seleccion + "'"
                    rstTerminado.Close
                        Else
                    spMuestra = "ListaMuestraEnsayoSolo " + "'" + Seleccion + "'"
                End If
                    Else
                If Mid(Seleccion, 3, 1) = "-" And Seleccion <> "  -   -   " Then
                    spArticulo = "ConsultaArticulo " + "'" + Seleccion + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        spMuestra = "ListaMuestraArticuloSolo " + "'" + Seleccion + "'"
                        rstArticulo.Close
                            Else
                        spMuestra = "ListaMuestraEnsayoSolo " + "'" + Seleccion + "'"
                    End If
                        Else
                    spMuestra = "ListaMuestraEnsayoSolo " + "'" + Seleccion + "'"
                End If
            End If
            
        Case 4
            spMuestra = "ListaMuestraNombreSolo " + "'" + Seleccion + "'"
            
        Case 5
            spMuestra = "ListaMuestraCantidadSolo " + "'" + Seleccion + "'"
            
        Case 6
            spMuestra = "ListaMuestraDescriClienteSolo " + "'" + Seleccion + "'"
            
        Case 7
            spMuestra = "ListaMuestraClienteSolo " + "'" + Seleccion + "'"
            
        Case 8
            spMuestra = "ListaMuestraObservacionesSolo " + "'" + Seleccion + "'"
            
        Case Else
    End Select
            
    Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
    If rstMuestra.RecordCount > 0 Then
        With rstMuestra
    
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    If rstMuestra!OrdFecha >= ZFecDesde And rstMuestra!OrdFecha <= ZFecHasta Then
                
                        WLugar = WLugar + 1
                        Auxiliar(WLugar) = Str$(rstMuestra!Codigo)
                        
                        Rem Auxiliar(WLugar, 2) = rstMuestra!Fecha
                        Rem Auxiliar(WLugar, 3) = rstMuestra!Ensayo
                        Rem Auxiliar(WLugar, 4) = rstMuestra!Producto
                        Rem Auxiliar(WLugar, 5) = rstMuestra!Articulo
                        Rem Auxiliar(WLugar, 6) = rstMuestra!Cantidad
                        Rem Auxiliar(WLugar, 7) = rstMuestra!DescriCliente
                        Rem Auxiliar(WLugar, 8) = rstMuestra!Cliente
                        Rem Auxiliar(WLugar, 17) = rstMuestra!Observaciones
                        Rem Auxiliar(WLugar, 9) = IIf(IsNull(rstMuestra!Producto2), "", rstMuestra!Producto2)
                        Rem Auxiliar(WLugar, 10) = IIf(IsNull(rstMuestra!Articulo2), "", rstMuestra!Articulo2)
                        Rem Auxiliar(WLugar, 11) = IIf(IsNull(rstMuestra!ensayo2), "", rstMuestra!ensayo2)
                        Rem Auxiliar(WLugar, 12) = IIf(IsNull(rstMuestra!fecha2), "", rstMuestra!fecha2)
                        Rem Auxiliar(WLugar, 13) = IIf(IsNull(rstMuestra!Cantidad2), "", rstMuestra!Cantidad2)
                        Rem Auxiliar(WLugar, 14) = IIf(IsNull(rstMuestra!lote2), "", rstMuestra!lote2)
                        Rem Auxiliar(WLugar, 15) = IIf(IsNull(rstMuestra!Observaciones2), "", rstMuestra!Observaciones2)
                        Rem Auxiliar(WLugar, 16) = IIf(IsNull(rstMuestra!Stock2), "", rstMuestra!Stock2)
                        Rem Auxiliar(WLugar, 18) = IIf(IsNull(rstMuestra!Razon), "", rstMuestra!Razon)
                        Rem Auxiliar(WLugar, 19) = IIf(IsNull(rstMuestra!Nombre), "", rstMuestra!Nombre)
                        Rem Auxiliar(WLugar, 20) = IIf(IsNull(rstMuestra!Nombre2), "", rstMuestra!Nombre2)
                        
                        Rem Muestra.Row = WLugar
                        
                        Rem Muestra.Col = 1
                        Muestra.TextMatrix(WLugar, 1) = Left$(IIf(IsNull(rstMuestra!pedido), "", rstMuestra!pedido), 6)
                        
                        Muestra.TextMatrix(WLugar, 2) = Left$(rstMuestra!Fecha, 5) + "/" + Mid$(rstMuestra!Fecha, 9, 2)
                        aa = rstMuestra!OrdFecha
            
                        Espa1 = Len(rstMuestra!Ensayo)
                        Espa2 = Len(rstMuestra!Producto)
                        Espa3 = Len(rstMuestra!Articulo)
            
            
                        If rstMuestra!Ensayo <> "" And rstMuestra!Ensayo <> Space$(Espa1) Then
                            Rem Muestra.Col = 3
                            Muestra.TextMatrix(WLugar, 3) = rstMuestra!Ensayo
                        End If
                
                        If rstMuestra!Producto <> "" And rstMuestra!Producto <> "  -     -   " And rstMuestra!Producto <> Space$(Espa2) Then
                            Rem Muestra.Col = 3
                            Muestra.TextMatrix(WLugar, 3) = rstMuestra!Producto
                        End If
                        
                        If rstMuestra!Articulo <> "" And rstMuestra!Articulo <> "  -   -   " And rstMuestra!Articulo <> Space$(Espa3) Then
                            Rem Muestra.Col = 3
                            Muestra.TextMatrix(WLugar, 3) = rstMuestra!Articulo
                        End If
            
                        Rem Muestra.Col = 4
                        ZNombre = IIf(IsNull(rstMuestra!Nombre), "", rstMuestra!Nombre)
                        Muestra.TextMatrix(WLugar, 4) = ZNombre
                        
                        Rem Muestra.Col = 5
                        Muestra.TextMatrix(WLugar, 5) = rstMuestra!Cantidad
            
                        Rem Muestra.Col = 6
                        Muestra.TextMatrix(WLugar, 6) = rstMuestra!descricliente
            
                        Rem Muestra.Col = 7
                        ZRazon = IIf(IsNull(rstMuestra!Razon), "", rstMuestra!Razon)
                        Muestra.TextMatrix(WLugar, 7) = ZRazon
            
                        Rem Muestra.Col = 8
                        Muestra.TextMatrix(WLugar, 8) = rstMuestra!Observaciones
            
                        Rem Muestra.Col = 9
                        Muestra.TextMatrix(WLugar, 9) = Left$(IIf(IsNull(rstMuestra!fecha2), "", rstMuestra!fecha2), 5)
            
                        ZRemito = IIf(IsNull(rstMuestra!Remito), "", rstMuestra!Remito)
                        If ZRemito <> "" Then
                            Rem Muestra.Col = 10
                            Muestra.TextMatrix(WLugar, 10) = ZRemito
                        End If
                        
                        ZHojaRuta = IIf(IsNull(rstMuestra!HojaRuta), "", rstMuestra!HojaRuta)
                        If ZHojaRuta <> "" Then
                            Rem Muestra.Col = 11
                            Muestra.TextMatrix(WLugar, 11) = ZHojaRuta
                        End If
                        
                        ZEnsayo2 = IIf(IsNull(rstMuestra!ensayo2), "", rstMuestra!ensayo2)
                        If ZEnsayo2 <> "" And ZEnsayo2 <> Space$(15) Then
                            Rem Muestra.Col = 12
                            Muestra.TextMatrix(WLugar, 12) = ZEnsayo2
                        End If
                        
                        ZArticulo2 = IIf(IsNull(rstMuestra!Articulo2), "", rstMuestra!Articulo2)
                        If ZArticulo2 <> "" And ZArticulo2 <> "  -   -   " Then
                            Rem Muestra.Col = 12
                            Muestra.TextMatrix(WLugar, 12) = ZArticulo2
                        End If
                
                        ZProducto2 = IIf(IsNull(rstMuestra!Producto2), "", rstMuestra!Producto2)
                        If ZProducto2 <> "" And ZProducto2 <> "  -     -   " Then
                            Rem Muestra.Col = 12
                            Muestra.TextMatrix(WLugar, 12) = ZProducto2
                        End If
            
                        Rem Muestra.Col = 13
                        ZNombre2 = IIf(IsNull(rstMuestra!Nombre2), "", rstMuestra!Nombre2)
                        Muestra.TextMatrix(WLugar, 13) = ZNombre2
            
                        Rem Muestra.Col = 14
                        Muestra.TextMatrix(WLugar, 14) = IIf(IsNull(rstMuestra!lote2), "", rstMuestra!lote2)
            
                        Rem Muestra.Col = 15
                        Muestra.TextMatrix(WLugar, 15) = IIf(IsNull(rstMuestra!Observaciones2), "", rstMuestra!Observaciones2)
            
                        Rem Muestra.Col = 16
                        Muestra.TextMatrix(WLugar, 16) = IIf(IsNull(rstMuestra!Cantidad2), "", rstMuestra!Cantidad2)
            
                        WStock2 = IIf(IsNull(rstMuestra!Stock2), "", rstMuestra!Stock2)
                        If Val(WStock2) = 1 Then
                            Rem Muestra.Col = 17
                            Muestra.TextMatrix(WLugar, 17) = "          S"
                                Else
                            Rem Muestra.Col = 17
                            Muestra.TextMatrix(WLugar, 17) = ""
                        End If
                        
                        WOrdenTrabajo = IIf(IsNull(rstMuestra!OrdenTrabajo), "", rstMuestra!OrdenTrabajo)
                        Muestra.TextMatrix(WLugar, 18) = WOrdenTrabajo
                    
                    End If
                    
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
        
        End With
        rstMuestra.Close
    End If
    
    Muestra.Visible = True
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    If WPosi1 <> 0 And WPosi2 <> 0 And WPosi3 <> 0 Then
        Muestra.TopRow = WPosi1
        Muestra.Col = WPosi3
        Muestra.Row = WPosi2
            Else
        If WLugar > 20 Then
            Muestra.TopRow = WLugar - 20
                Else
            Muestra.TopRow = 1
        End If
        Muestra.Col = 1
        Muestra.Row = WLugar
    End If
    
    Muestra.SetFocus
    
End Sub

Private Sub Muestra_Click()
    Ayuda.Visible = True
    Ayuda.Text = ""
    ColumnaOpcion = Muestra.Col
End Sub

Private Sub Muestra_DblClick()

    If Val(Ano.Text) = 0 Or Val(AnoII.Text) = 0 Then
        Exit Sub
    End If
    
    Ayuda.Visible = True
    Ayuda.Text = ""
    ColumnaOpcion = Muestra.Col
    
    WPosi1 = 1
    WPosi2 = 1
    WPosi3 = 1
    
    ZFecDesde = Ano.Text + "0101"
    ZFecHasta = AnoII.Text + "1231"
    
    pantalla.Clear
    Select Case ColumnaOpcion
        Case 2
            Pasa = 0
            Corte = ""
            
            ZSql = ""
            ZSql = ZSql + "Select Muestra.fecha, Muestra.OrdFecha"
            ZSql = ZSql + " FROM Muestra"
            ZSql = ZSql + " Where Muestra.OrdFecha >= " + "'" + ZFecDesde + "'"
            ZSql = ZSql + " and Muestra.OrdFecha <= " + "'" + ZFecHasta + "'"
            ZSql = ZSql + " Order by Muestra.OrdFecha"
            spMuestra = ZSql
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            If rstMuestra.RecordCount > 0 Then
                With rstMuestra
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Pasa = 0 Then
                                pantalla.AddItem ""
                                Pasa = 1
                                Corte = rstMuestra!Fecha
                            End If
                            If Corte <> rstMuestra!Fecha Then
                                pantalla.AddItem Corte
                                Corte = rstMuestra!Fecha
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                pantalla.AddItem Corte
                rstMuestra.Close
            End If
            pantalla.Visible = True
            
        Case 3
            Lista.Clear
            
            Pasa = 0
            Corte = ""
            
            ZSql = ""
            ZSql = ZSql + "Select Muestra.fecha, Muestra.OrdFecha, Muestra.Producto, Muestra.Articulo, Muestra.Ensayo"
            ZSql = ZSql + " FROM Muestra"
            ZSql = ZSql + " Where Muestra.OrdFecha >= " + "'" + ZFecDesde + "'"
            ZSql = ZSql + " and Muestra.OrdFecha <= " + "'" + ZFecHasta + "'"
            ZSql = ZSql + " Order by Muestra.Articulo, Muestra.Producto, Muestra.Ensayo"
            
            spMuestra = ZSql
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            If rstMuestra.RecordCount > 0 Then
                With rstMuestra
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            WAgrega = ""
                            If Left(rstMuestra!Producto, 2) = "PT" Then
                                WAgrega = rstMuestra!Producto
                                    Else
                                If Mid(rstMuestra!Articulo, 3, 1) = "-" And rstMuestra!Articulo <> "  -   -   " Then
                                    WAgrega = rstMuestra!Articulo
                                        Else
                                    If rstMuestra!Ensayo <> "" And rstMuestra!Ensayo <> Space(15) Then
                                        WAgrega = rstMuestra!Ensayo
                                    End If
                                End If
                            End If
                            If WAgrega <> "" Then
                                If Pasa = 0 Then
                                    Pasa = 1
                                    Corte = WAgrega
                                End If
                                If Corte <> WAgrega Then
                                    Lista.AddItem Corte
                                    Corte = WAgrega
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstMuestra.Close
                Lista.AddItem Corte
            End If
            
            Erase WPasa
            Hasta = Lista.ListCount
            For Ciclo = 0 To Hasta - 1
                Lista.ListIndex = Ciclo
                WPasa(Ciclo + 1) = Lista.Text
            Next Ciclo
            
            pantalla.AddItem ""
            
            Pasa = 0
            Corte = ""
            For Ciclo = 1 To Hasta
                WAgrega = WPasa(Ciclo)
                If WAgrega <> "" Then
                    If Pasa = 0 Then
                        Pasa = 1
                        Corte = WAgrega
                    End If
                    If Corte <> WAgrega Then
                        pantalla.AddItem Corte
                        Corte = WAgrega
                    End If
                End If
            Next Ciclo
            pantalla.AddItem Corte
            
            pantalla.Visible = True
            
        Case 4
            Pasa = 0
            Corte = ""
            
            ZSql = ""
            ZSql = ZSql + "Select Muestra.fecha, Muestra.OrdFecha, Muestra.Nombre"
            ZSql = ZSql + " FROM Muestra"
            ZSql = ZSql + " Where Muestra.OrdFecha >= " + "'" + ZFecDesde + "'"
            ZSql = ZSql + " and Muestra.OrdFecha <= " + "'" + ZFecHasta + "'"
            ZSql = ZSql + " Order by Muestra.Nombre"
            
            spMuestra = ZSql
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            If rstMuestra.RecordCount > 0 Then
                With rstMuestra
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Pasa = 0 Then
                                pantalla.AddItem ""
                                Pasa = 1
                                Corte = rstMuestra!Nombre
                            End If
                            If Corte <> rstMuestra!Nombre Then
                                pantalla.AddItem Corte
                                Corte = rstMuestra!Nombre
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstMuestra.Close
            End If
                
            pantalla.AddItem Corte
            pantalla.Visible = True
            
        Case 5
            Pasa = 0
            Corte = ""
            
            ZSql = ""
            ZSql = ZSql + "Select Muestra.fecha, Muestra.OrdFecha, Muestra.cantidad"
            ZSql = ZSql + " FROM Muestra"
            ZSql = ZSql + " Where Muestra.OrdFecha >= " + "'" + ZFecDesde + "'"
            ZSql = ZSql + " and Muestra.OrdFecha <= " + "'" + ZFecHasta + "'"
            ZSql = ZSql + " Order by Muestra.Cantidad"
            
            spMuestra = ZSql
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            If rstMuestra.RecordCount > 0 Then
                With rstMuestra
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Pasa = 0 Then
                                pantalla.AddItem ""
                                Pasa = 1
                                Corte = rstMuestra!Cantidad
                            End If
                            If Corte <> rstMuestra!Cantidad Then
                                pantalla.AddItem Corte
                                Corte = rstMuestra!Cantidad
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                pantalla.AddItem Corte
                rstMuestra.Close
            End If
            pantalla.Visible = True
            
        Case 6
            ZFecDesde = Ano.Text + "0101"
            ZFecHasta = AnoII.Text + "1231"
        
            Pasa = 0
            Corte = ""
            
            ZSql = ""
            ZSql = ZSql + "Select Muestra.DescriCliente, Muestra.OrdFecha"
            ZSql = ZSql + " FROM Muestra"
            ZSql = ZSql + " Where Muestra.OrdFecha >= " + "'" + ZFecDesde + "'"
            ZSql = ZSql + " and Muestra.OrdFecha <= " + "'" + ZFecHasta + "'"
            ZSql = ZSql + " Order by Muestra.DescriCliente"
            spMuestra = ZSql
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            If rstMuestra.RecordCount > 0 Then
                With rstMuestra
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            If rstMuestra!OrdFecha >= ZFecDesde And rstMuestra!OrdFecha <= ZFecHasta Then
                        
                                If Pasa = 0 Then
                                    pantalla.AddItem ""
                                    Pasa = 1
                                    Corte = rstMuestra!descricliente
                                End If
                                If Corte <> rstMuestra!descricliente Then
                                    pantalla.AddItem Corte
                                    Corte = rstMuestra!descricliente
                                End If
                            
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                pantalla.AddItem Corte
                rstMuestra.Close
            End If
            
            pantalla.Visible = True
            
        Case 7
            Ayuda.Visible = True
            Ayuda.Text = ""
            Pasa = 0
            Corte = ""
            
            ZFecDesde = Ano.Text + "0101"
            ZFecHasta = AnoII.Text + "1231"
        
            Pasa = 0
            Corte = ""
            
            ZSql = ""
            ZSql = ZSql + "Select Muestra.Razon, Muestra.OrdFecha, Muestra.Cliente"
            ZSql = ZSql + " FROM Muestra"
            ZSql = ZSql + " Where Muestra.OrdFecha >= " + "'" + ZFecDesde + "'"
            ZSql = ZSql + " and Muestra.OrdFecha <= " + "'" + ZFecHasta + "'"
            ZSql = ZSql + " Order by Muestra.Razon"
            spMuestra = ZSql
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            If rstMuestra.RecordCount > 0 Then
                With rstMuestra
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If rstMuestra!Cliente <> "" Then
                                If Pasa = 0 Then
                                    pantalla.AddItem ""
                                    Pasa = 1
                                    Corte = rstMuestra!Razon
                                End If
                                If Corte <> rstMuestra!Razon Then
                                    pantalla.AddItem Corte
                                    Corte = rstMuestra!Razon
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                pantalla.AddItem Corte
                rstMuestra.Close
            End If
            pantalla.Visible = True
            
        Case 8
            Pasa = 0
            Corte = ""
            spMuestra = "ListaMuestraObservaciones"
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            With rstMuestra
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            pantalla.AddItem ""
                            Pasa = 1
                            Corte = rstMuestra!Observaciones
                        End If
                        If Corte <> rstMuestra!Observaciones Then
                            pantalla.AddItem Corte
                            Corte = rstMuestra!Observaciones
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            pantalla.AddItem Corte
            rstMuestra.Close
            pantalla.Visible = True
            
        Case Else
        
    End Select
    
    Ayuda.SetFocus
            
    Rem Muestra.Col = 10
    Rem Muestra.Col = 1
    Rem WXSol = Muestra.Text
    Rem PrgSol.Show
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    If Val(ZAno) = 0 Then
        ZAno = Right$(Date$, 4)
    End If
    If Val(ZAnoII) = 0 Then
        ZAnoII = Right$(Date$, 4)
    End If
    Ano.Text = ZAno
    AnoII.Text = ZAnoII
    Call Proceso_Click
End Sub

Private Sub Muestra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Alta_Click
        Case 113
            Call Baja_Click
        Case 114
            Call Modifica_Click
        Case 115
            Call Labora_Click
        Case 120
            Call Impresion_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector()

    Muestra.Clear

    Rem ponga la muestra en negritas
    Rem Muestra.Font.Bold = True

    ' Establesco loa Valores de la muestra
    
    Muestra.FixedCols = 1
    Muestra.Cols = 19
    Muestra.FixedRows = 1
    Muestra.Rows = 18000
    
    Muestra.ColWidth(0) = 200
    Muestra.Row = 0
    
    For Ciclo = 1 To Muestra.Cols - 1
        Muestra.Col = Ciclo
        Select Case Ciclo
            Case 1
                Muestra.Text = "Pedido"
                Muestra.ColWidth(Ciclo) = 650
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                Muestra.Text = "Fecha"
                Muestra.ColWidth(Ciclo) = 850
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                Muestra.Text = "Codigo"
                Muestra.ColWidth(Ciclo) = 1200
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                Muestra.Text = "Descripcion"
                Muestra.ColWidth(Ciclo) = 1300
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                Muestra.Text = "Cantidad"
                Muestra.ColWidth(Ciclo) = 800
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 6
                Muestra.Text = "Nombre para el Cliente"
                Muestra.ColWidth(Ciclo) = 1500
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 7
                Muestra.Text = "Cliente"
                Muestra.ColWidth(Ciclo) = 1250
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                Muestra.Text = "Observaciones"
                Muestra.ColWidth(Ciclo) = 1600
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 9
                Muestra.Text = "Fec.OK"
                Muestra.ColWidth(Ciclo) = 750
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 10
                Muestra.Text = "Remito"
                Muestra.ColWidth(Ciclo) = 750
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 11
                Muestra.Text = "H.Ruta"
                Muestra.ColWidth(Ciclo) = 750
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 12
                Muestra.Text = "Codigo Conf."
                Muestra.ColWidth(Ciclo) = 1350
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 13
                Muestra.Text = "Descripcion"
                Muestra.ColWidth(Ciclo) = 2000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 14
                Muestra.Text = "Lote"
                Muestra.ColWidth(Ciclo) = 1200
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 15
                Muestra.Text = "Observaciones"
                Muestra.ColWidth(Ciclo) = 2300
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 16
                Muestra.Text = "Cantidad"
                Muestra.ColWidth(Ciclo) = 1000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 17
                Muestra.Text = "Actualiza Stock"
                Muestra.ColWidth(Ciclo) = 1600
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 18
                Muestra.Text = "O.Trabajo"
                Muestra.ColWidth(Ciclo) = 1600
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Muestra.AllowUserResizing = flexResizeBoth
    
    Muestra.Col = 1
    Muestra.Row = 1
    
End Sub

Private Sub Limpia_VectorII()

    VectorRemito.Clear

    Rem ponga la muestra en negritas
    Rem Muestra.Font.Bold = True

    ' Establesco loa Valores de la muestra
    
    VectorRemito.FixedCols = 1
    VectorRemito.Cols = 8
    VectorRemito.FixedRows = 1
    VectorRemito.Rows = 100
    
    VectorRemito.ColWidth(0) = 200
    VectorRemito.Row = 0
    
    For Ciclo = 1 To VectorRemito.Cols - 1
        VectorRemito.Col = Ciclo
        Select Case Ciclo
            Case 1
                VectorRemito.Text = "Descripcion"
                VectorRemito.ColWidth(Ciclo) = 2500
                VectorRemito.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                VectorRemito.Text = "Cantidad"
                VectorRemito.ColWidth(Ciclo) = 900
                VectorRemito.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                VectorRemito.Text = "Muestra"
                VectorRemito.ColWidth(Ciclo) = 900
                VectorRemito.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                VectorRemito.Text = "Partida"
                VectorRemito.ColWidth(Ciclo) = 900
                VectorRemito.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                VectorRemito.Text = "Pedido"
                VectorRemito.ColWidth(Ciclo) = 900
                VectorRemito.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                VectorRemito.Text = "Codigo"
                VectorRemito.ColWidth(Ciclo) = 10
                VectorRemito.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                VectorRemito.Text = "Codigo"
                VectorRemito.ColWidth(Ciclo) = 10
                VectorRemito.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    VectorRemito.AllowUserResizing = flexResizeBoth
    
    VectorRemito.Col = 1
    VectorRemito.Row = 1
    
End Sub


Private Sub Pantalla_Click()
    If pantalla.ListIndex <> 0 Then
        Seleccion = pantalla.Text
            Else
        Seleccion = ""
        ColumnaOpcion = 0
    End If
    pantalla.Visible = False
    Ayuda.Visible = False
    Call Proceso_Click
End Sub

Private Sub Remito_Click()

    rowini = Muestra.Row
    RowFin = Muestra.RowSel
    
    Pasa = 0
    
    For Ciclo = rowini To RowFin
        
        ZNumero = Str$(Ciclo)
        ZPedido = Left$(Muestra.TextMatrix(Ciclo, 1), 6)
        ZFecha = Left$(Muestra.TextMatrix(Ciclo, 2), 10)
        ZCodigo = Left$(Muestra.TextMatrix(Ciclo, 3), 15)
        ZDescripcion = Left$(Muestra.TextMatrix(Ciclo, 4), 50)
        ZCantidad = Left$(Muestra.TextMatrix(Ciclo, 5), 10)
        ZDescriCliente = Left$(Muestra.TextMatrix(Ciclo, 6), 50)
        ZCliente = Left$(Muestra.TextMatrix(Ciclo, 7), 50)
        ZObservaciones = Left$(Muestra.TextMatrix(Ciclo, 8), 50)
        ZFecha2 = Left$(Muestra.TextMatrix(Ciclo, 8), 10)
        ZRemito = Left$(Muestra.TextMatrix(Ciclo, 10), 10)
        ZHojaRuta = Left$(Muestra.TextMatrix(Ciclo, 11), 10)
        ZCodigo2 = Left$(Muestra.TextMatrix(Ciclo, 12), 15)
        ZDescripcion2 = Left$(Muestra.TextMatrix(Ciclo, 13), 50)
        ZLote = Left$(Muestra.TextMatrix(Ciclo, 14), 10)
        ZObservaciones2 = Left$(Muestra.TextMatrix(Ciclo, 15), 50)
        ZCantidad2 = Left$(Muestra.TextMatrix(Ciclo, 16), 10)
        ZActualiza = Left$(Muestra.TextMatrix(Ciclo, 17), 1)
        
        WMuestra = Auxiliar(Ciclo)
        
        If Pasa = 0 Then
            Pasa = 1
            WCliente = ZCliente
        End If
        
        If WCliente <> ZCliente Then
            m$ = "Se ha seleccionado muestras de distintos clientes"
            A% = MsgBox(m$, 0, "Impresion de Remitos")
            Exit Sub
        End If
       
        If DesClienteRemito.Text <> "" Then
            If DesClienteRemito.Text <> ZCliente Then
                m$ = "Se ha seleccionado muestras de distintos clientes"
                A% = MsgBox(m$, 0, "Impresion de Remitos")
                Exit Sub
            End If
        End If
        
        WRemito = ""
        WEnsayo = ""
        WEnsayoII = ""
        Sql1 = "Select *"
        Sql2 = " FROM Muestra"
        Sql3 = " WHERE Codigo = " + "'" + WMuestra + "'"
        spMuestra = Sql1 + Sql2 + Sql3
        Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
        If rstMuestra.RecordCount > 0 Then
            WRemito = IIf(IsNull(rstMuestra!Remito), "0", rstMuestra!Remito)
            WEnsayo = IIf(IsNull(rstMuestra!Ensayo), "", rstMuestra!Ensayo)
            WEnsayoII = IIf(IsNull(rstMuestra!ensayo2), "", rstMuestra!ensayo2)
            rstMuestra.Close
        End If
        
        If Left$(ZCodigo, 2) = "ML" And Val(Mid$(ZCodigo, 4, 3)) >= 100 Then
            If Trim(WEnsayo) = "" And Trim(WEnsayoII) = "" Then
                m$ = "Se debe informar numero de ensayo"
               A% = MsgBox(m$, 0, "Impresion de Remitos")
               Exit Sub
            End If
        End If
        
        
       If Val(WRemito) <> 0 Then
           m$ = "Ya se ha emitido el remito corresondiente"
          A% = MsgBox(m$, 0, "Impresion de Remitos")
          Exit Sub
        End If
        
    Next Ciclo
    
    If PantaRemito.Visible = False Then
    
        Call Limpia_VectorII
    
        NumeroRemito.Text = ""
        FechaRemito.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        DesClienteRemito.Text = ""
    
        Sql1 = "Select Max(Remito) as [RemitoMayor]"
        Sql2 = " FROM Muestra"
        spMuestra = Sql1 + Sql2
        Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
        If rstMuestra.RecordCount > 0 Then
            rstMuestra.MoveLast
            WRemitoMayor = IIf(IsNull(rstMuestra!RemitoMayor), "0", rstMuestra!RemitoMayor)
            WRemito = Mid$(Str$(WRemitoMayor + 1), 2, 8)
            rstMuestra.Close
                Else
            WRemito = "1"
        End If
    
        NumeroRemito.Text = WRemito
        PantaRemito.Visible = True
        NumeroRemito.SetFocus
        
        LugarRemito = 0
        
    End If
    
    For CiclaRemito = 1 To 99
        If VectorRemito.TextMatrix(CiclaRemito, 1) = "" Or VectorRemito.TextMatrix(CiclaRemito, 2) = "" Then
            LugarRemito = CiclaRemito - 1
            Exit For
        End If
    Next CiclaRemito
    
    For Ciclo = rowini To RowFin
        
        ZNumero = Str$(Ciclo)
        ZPedido = Left$(Muestra.TextMatrix(Ciclo, 1), 6)
        ZFecha = Left$(Muestra.TextMatrix(Ciclo, 2), 10)
        ZCodigo = Left$(Muestra.TextMatrix(Ciclo, 3), 15)
        ZDescripcion = Left$(Muestra.TextMatrix(Ciclo, 4), 50)
        ZCantidad = Left$(Muestra.TextMatrix(Ciclo, 5), 10)
        ZDescriCliente = Left$(Muestra.TextMatrix(Ciclo, 6), 50)
        ZCliente = Left$(Muestra.TextMatrix(Ciclo, 7), 50)
        ZObservaciones = Left$(Muestra.TextMatrix(Ciclo, 8), 50)
        ZFecha2 = Left$(Muestra.TextMatrix(Ciclo, 8), 10)
        Rem Zxx = Left$(Muestra.TextMatrix(Ciclo, 9), 10)
        ZRemito = Left$(Muestra.TextMatrix(Ciclo, 10), 10)
        ZHojaRuta = Left$(Muestra.TextMatrix(Ciclo, 11), 10)
        ZCodigo2 = Left$(Muestra.TextMatrix(Ciclo, 12), 15)
        ZDescripcion2 = Left$(Muestra.TextMatrix(Ciclo, 13), 50)
        ZLote = Left$(Muestra.TextMatrix(Ciclo, 14), 10)
        ZObservaciones2 = Left$(Muestra.TextMatrix(Ciclo, 15), 50)
        ZCantidad2 = Left$(Muestra.TextMatrix(Ciclo, 16), 10)
        ZActualiza = Left$(Muestra.TextMatrix(Ciclo, 17), 1)

        WMuestra = Auxiliar(Ciclo)
        
        DesClienteRemito.Text = ZCliente
        
        LugarRemito = LugarRemito + 1
        
        VectorRemito.TextMatrix(LugarRemito, 1) = ZDescriCliente
        If Val(ZCantidad2) <> 0 Then
            VectorRemito.TextMatrix(LugarRemito, 2) = Str$(Val(ZCantidad2))
                Else
            VectorRemito.TextMatrix(LugarRemito, 2) = Str$(Val(ZCantidad))
        End If
        VectorRemito.TextMatrix(LugarRemito, 3) = WMuestra
        VectorRemito.TextMatrix(LugarRemito, 4) = ""
        VectorRemito.TextMatrix(LugarRemito, 5) = ZPedido
        
        If Len(ZCodigo) = 10 Then
            ZZCodigo = Left$(ZCodigo, 3) + "00" + Right$(ZCodigo, 7)
                Else
            ZZCodigo = ZCodigo
        End If
        VectorRemito.TextMatrix(LugarRemito, 6) = ZZCodigo
        VectorRemito.TextMatrix(LugarRemito, 7) = ZCodigo
        
        Sql1 = "Select *"
        Sql2 = " FROM Pedido"
        Sql3 = " WHERE Pedido = " + "'" + ZPedido + "'"
        Sql4 = " and Terminado = " + "'" + ZZCodigo + "'"
        spPedido = Sql1 + Sql2 + Sql3 + Sql4
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            VectorRemito.TextMatrix(LugarRemito, 4) = rstPedido!lote1
            rstPedido.Close
        End If
        
        If Val(VectorRemito.TextMatrix(LugarRemito, 4)) = 0 Then
            Sale = "N"
            If Left$(ZCodigo, 2) = "ML" Then
                Rem lo dejo pasar
                Rem es una muestra de laboratorio
                Sale = "S"
                    Else
                If Left$(ZCodigo, 2) = "YQ" Or Left$(ZCodigo, 2) = "YF" Then
                    If Val(ZCantidad) <= 20 Then
                        Rem lo dejo pasar
                        Rem es un ensayo con menos de 10 Kgs.
                        Sale = "S"
                    End If
                End If
            End If
            If Sale = "N" Then
                m$ = "No se ha actualizado el pedido con las partidas correspondientes"
                A% = MsgBox(m$, 0, "Impresion de Remitos")
                Call Limpia_VectorII
                NumeroRemito.Text = ""
                FechaRemito.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                DesClienteRemito.Text = ""
                PantaRemito.Visible = False
                Exit Sub
            End If
        End If
        
    Next Ciclo
            
End Sub

Private Sub CancelaRemito_Click()
    Call Limpia_VectorII
    NumeroRemito.Text = ""
    FechaRemito.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    DesClienteRemito.Text = ""
    PantaRemito.Visible = False
End Sub

Private Sub ConfirmaRemito_Click()

    For Ciclo = 1 To 99
        
        If Trim(VectorRemito.TextMatrix(Ciclo, 1)) = "" Or Trim(VectorRemito.TextMatrix(Ciclo, 2)) = "" Then
            Exit For
        End If
        
        ZDescripcion = VectorRemito.TextMatrix(Ciclo, 1)
        ZCantidad = VectorRemito.TextMatrix(Ciclo, 2)
        ZMuestra = VectorRemito.TextMatrix(Ciclo, 3)
        ZPartida = VectorRemito.TextMatrix(Ciclo, 4)
        ZPedido = VectorRemito.TextMatrix(Ciclo, 5)
        ZCodigo = VectorRemito.TextMatrix(Ciclo, 7)
        
        Rem If Trim(ZPartida) = "" Then
        Rem     m$ = "El pedido no esta actualizado"
        Rem     A% = MsgBox(m$, 0, "Emision de Remitos")
        Rem     Exit Sub
        Rem End If
            
    Next Ciclo



    WMuestra = VectorRemito.TextMatrix(1, 3)

    Sql1 = "Select *"
    Sql2 = " FROM Muestra"
    Sql3 = " WHERE Codigo = " + "'" + WMuestra + "'"
    spMuestra = Sql1 + Sql2 + Sql3
    Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
    If rstMuestra.RecordCount > 0 Then
        WCliente = rstMuestra!Cliente
        WRazon = rstMuestra!Razon
        rstMuestra.Close
    End If
    
    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = rstCliente!Razon
        WPago1 = rstCliente!Pago1
        WPago2 = rstCliente!Pago2
        WVendedor = rstCliente!Vendedor
        WProv = rstCliente!Provincia
        WRubro = rstCliente!Rubro
        WCodIva = rstCliente!Iva
        WCodIb = rstCliente!Ib
        WRazon = rstCliente!Razon
        WDireccion = rstCliente!Direccion
        WLocalidad = rstCliente!Localidad
        WPostal = rstCliente!Postal
        WCuit = rstCliente!Cuit
        WDirentrega = rstCliente!DirEntrega
        rstCliente.Close
    End If

    Rem If Val(WEmpresa) = 1 Then
    Rem     Rem Open "DADA.TXT" For Output As #1
    Rem    Rem Open "lpt1" For Output As #1
    Rem     Open "DADA.TXT" For Output As #1
    Rem         Else
    Rem     If Val(WEmpresa) <> 9 And Val(WEmpresa) <> 10 Then
    Rem         Rem Open "DADA.TXT" For Output As #1
    Rem         Open "lpt1" For Output As #1
    Rem         Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
    Rem         Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
    Rem             Else
    Rem         Open "DADA.TXT" For Output As #1
    Rem     End If
    Rem End If
    Rem
    Rem For FF = 1 To 2
    Rem
    Rem     Print #1, Chr$(27) + Chr$(40) + "19U"
    Rem     Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "2" + Chr$(72)
    Rem     Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
    Rem     Print #1, ""
    Rem     Print #1, ""
    Rem     Print #1, ""
    Rem     Print #1, ""
    Rem     Print #1, Tab(53); FechaRemito.Text
    Rem     Print #1, ""
    Rem     Print #1, ""
    Rem     Print #1, ""
    Rem     Print #1, ""
    Rem     Print #1, ""
    Rem     Print #1, ""
    Rem     Print #1, Tab(7); WRazon
    Rem     Print #1, Tab(7); Left$(WDireccion, 33)
    Rem     Print #1, Tab(7); Left$(WLocalidad, 33);
    Rem     Print #1, Tab(44); ;
    Rem     Print #1, Tab(57); WCliente;
    Rem     Print #1, Tab(68); ""
    Rem     Print #1, Tab(7); Provincia(Val(WProv)); "("; WPostal; ")"
    Rem     Print #1, ""
    Rem     Print #1, Tab(7); Iva(Val(WCodIva));
    Rem     Print #1, Tab(48); WCuit
    Rem     Print #1, ""
    Rem     Print #1, Tab(30); WDirentrega;
    Rem     Print #1, ""
    Rem     If FF = 1 Then
    Rem         Print #1, Tab(60); "ORIGINAL"
    Rem             Else
    Rem         Print #1, Tab(60); "DUPLICADO"
    Rem     End If
    Rem     Print #1, ""
    Rem
    Rem     Impre = 0
    Rem
    Rem     For Ciclo = 1 To 99
    Rem
    Rem         If Trim(VectorRemito.TextMatrix(Ciclo, 1)) = "" Or Trim(VectorRemito.TextMatrix(Ciclo, 2)) = "" Then
    Rem             Exit For
    Rem         End If
    Rem
    Rem         ZDescripcion = VectorRemito.TextMatrix(Ciclo, 1)
    Rem         ZCantidad = VectorRemito.TextMatrix(Ciclo, 2)
    Rem         ZMuestra = VectorRemito.TextMatrix(Ciclo, 3)
    Rem
    Rem         WMuestra = ZMuestra
    Rem         Descri = ZDescripcion
    Rem         Cantidad = Val(ZCantidad)
    Rem
    Rem         If Cantidad <> 0 Then
    Rem             Print #1, Tab(14); Left$(Descri, 40);
    Rem             Print #1, Tab(58); Alinea("#####.##", Str$(Cantidad));
    Rem             Print #1, " Kg";
    Rem             Print #1, Tab(71); "Netos"
    Rem             Impre = Impre + 1
    Rem         End If
    Rem
    Rem     Next Ciclo
    Rem
    Rem     If FF = 2 Then
    Rem
    Rem         If Val(WEmpresa) = 4 Or Val(WEmpresa) = 8 Then
    Rem             For aa = Impre To 10
    Rem                 Impre = Impre + 1
    Rem                 Print #1, ""
    Rem             Next aa
    Rem                 Else
    Rem             For aa = Impre To 12
    Rem                 Impre = Impre + 1
    Rem                 Print #1, ""
    Rem             Next aa
    Rem         End If
    Rem
    Rem         Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "15" + Chr$(72);
    Rem         Print #1, "  -----------------------------------------------------------------------------------------------------------------"
    Rem         Print #1, "  |   ETIQUETADO          | Si/No | TRANSPORTE                    | Si/No | SI HAY SUSTANCIAS PELIGROSAS  | Si/No |"
    Rem         Print #1, "  -----------------------------------------------------------------------------------------------------------------"
    Rem         Print #1, "  | Cliente               |       | Conductor                     |       | Ficha de Intervencion         |       |"
    Rem         Print #1, "  | Nombre                |       | H.de Ruta/Guia Traslasdo      |       | Rotulos Externos              |       |"
    Rem         Print #1, "  | Codigo                |       | Remitos                       |       |----------------------------------------"
    Rem         Print #1, "  | Partida               |       | Facturas                      |       | VERIFICO :                            |"
    Rem         Print #1, "  | Neto                  |       | Certificado de Analisis       |       |---------------------------------------|"
    Rem         Print #1, "  | Vencimiento           |       | Hoja de Seguridad             |       | ENTREGA ENVASES               | Si/No |"
    Rem         Print #1, "  | Etiq./Irradiacion     |       | Certificado de Irradicacion   |       | MOTIVO :                              |"
    Rem         Print #1, "  |                       |       | Van Muestras                  |       |                                       |"
    Rem         Print #1, "  ------------------------------------------------------------------------|                                       |"
    Rem         Print #1, "  | CONTROL TRANSPORTISTA | Entro |               | Salio |               | Firma                                 |"
    Rem         Impre = Impre + 12
    Rem
    Rem     End If
    Rem
    Rem     Select Case Val(WEmpresa)
    Rem         Case 4, 8
    Rem             If FF = 1 Then
    Rem                 For aa = Impre To 17
    Rem                     Print #1, ""
    Rem                 Next aa
    Rem
    Rem                 Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)
    Rem
    Rem                 Print #1, Tab(3); "Pellital S.A. no se responsabiliza por los daños que pudiera causar la aplicación inadecuada de estos productos,"
    Rem                 Print #1, Tab(3); "el reuso de envases o la mala disposición final de los residuos generados a partir de los mismos."
    Rem                 Print #1, Tab(3); "Los residuos generados a partir de los productos remitidos con  este comprobante y que presenten riesgos para"
    Rem                 Print #1, Tab(3); "la salud o para el medio ambiente, deberán ser destruidos y dispuestos según lo establecen las reglamentaciones "
    Rem                 Print #1, Tab(3); "vigentes del ámbito municipal, provincial y nacional"
    Rem                 Print #1, ""
    Rem             End If
    Rem
    Rem             Print #1, ""
    Rem             Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
    Rem             Print #1, ""
    Rem             Print #1, Chr$(12)
    Rem
    Rem         Case Else
    Rem             If FF = 1 Then
    Rem                 For aa = Impre To 19
    Rem                     Print #1, ""
    Rem                 Next aa
    Rem
    Rem                 Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)
    Rem
    Rem                 Print #1, Tab(3); "Surfactan S.A. no se responsabiliza por los daños que pudiera causar la aplicación inadecuada de estos productos,"
    Rem                 Print #1, Tab(3); "el reuso de envases o la mala disposición final de los residuos generados a partir de los mismos."
    Rem                 Print #1, Tab(3); "Los residuos generados a partir de los productos remitidos con  este comprobante y que presenten riesgos para"
    Rem                 Print #1, Tab(3); "la salud o para el medio ambiente, deberán ser destruidos y dispuestos según lo establecen las reglamentaciones "
    Rem                 Print #1, Tab(3); "vigentes del ámbito municipal, provincial y nacional"
    Rem                 Print #1, ""
    Rem             End If
    Rem
    Rem             Print #1, ""
    Rem             Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
    Rem             Print #1, ""
    Rem             Print #1, Chr$(12)
    Rem
    Rem     End Select
    Rem
    Rem Next FF
    Rem
    Rem Close #1

    For Ciclo = 1 To 99
        
        If Trim(VectorRemito.TextMatrix(Ciclo, 1)) = "" Or Trim(VectorRemito.TextMatrix(Ciclo, 2)) = "" Then
            Exit For
        End If
            
        ZDescripcion = VectorRemito.TextMatrix(Ciclo, 1)
        ZCantidad = VectorRemito.TextMatrix(Ciclo, 2)
        ZMuestra = VectorRemito.TextMatrix(Ciclo, 3)
        ZPartida = VectorRemito.TextMatrix(Ciclo, 4)
        ZPedido = VectorRemito.TextMatrix(Ciclo, 5)
        ZZCodigo = VectorRemito.TextMatrix(Ciclo, 6)
        ZCodigo = VectorRemito.TextMatrix(Ciclo, 7)
        
        If ZZCodigo > "ML-00100-100" And ZZCodigo < "ML-99999-100" Then
            ZZCodigo = "ML-00008-100"
        End If
            
        WMuestra = ZMuestra
        Descri = ZDescripcion
        Cantidad = Val(ZCantidad)
        
        Sql1 = "UPDATE Muestra SET "
        Sql2 = " Remito = " + "'" + NumeroRemito.Text + "'"
        Sql3 = " Where Codigo = " + "'" + WMuestra + "'"
        spMuestra = Sql1 + Sql2 + Sql3
        
        
        
        Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
        
        Sql1 = "Select *"
        Sql2 = " FROM Pedido"
        Sql3 = " WHERE Pedido = " + "'" + ZPedido + "'"
        Sql4 = " and Terminado = " + "'" + ZZCodigo + "'"
        spPedido = Sql1 + Sql2 + Sql3 + Sql4
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            ZTipoPedido = rstPedido!TipoPedido
            Select Case rstPedido!TipoPedido
                Case 1
                    WTipoPedido = "CO"
                Case 3
                    WTipoPedido = "BI"
                Case 4
                    WTipoPedido = "FA"
                Case 5
                    WTipoPedido = "PG"
                Case Else
                    WTipoPedido = "PT"
            End Select
            rstPedido.Close
        End If
        
        If Val(ZPartida) <> 0 Then
        
        If Len(ZCodigo) = 10 Then
            XTipoproDy = "M"
                Else
            XTipoproDy = "T"
        End If
        
        If XTipoproDy = "M" Then
                    
            Select Case WTipoPedido
                Case "PG", "CO"
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case "FA"
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    Wempresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
                    
            spArticulo = "ConsultaArticulo " + "'" + ZCodigo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WCodigo = ZCodigo
                WPedido = Str$(rstArticulo!Venta - Val(ZCantidad))
                WSalidas = Str$(rstArticulo!Salidas + Val(ZCantidad))
                WDate = Date$
                rstArticulo.Close
                XParam = "'" + WCodigo + "','" _
                            + WPedido + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                spArticulo = "ModificaArticuloFacturas " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
                        
            XParam = "'" + ZPartida + "','" _
                         + ZCodigo + "'"
            spLaudo = "ListaLaudoArticulo " + XParam
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                WClave = rstLaudo!Clave
                WSaldo = Str$(rstLaudo!Saldo - Val(ZCantidad))
                WDate = Date$
                rstLaudo.Close
                            
                XParam = "'" + WClave + "','" _
                             + WDate + "','" _
                             + WSaldo + "'"
                spLaudo = "ModificaLaudoSaldo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                    Else
                                
                XParam = "'" + ZCodigo + "','" _
                            + ZPartida + "'"
                spMovguia = "ListaMovguiaLote " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    WClave = rstMovguia!Clave
                    WSaldo = Str$(rstMovguia!Saldo - Val(ZCantidad))
                    WDate = Date$
                    rstMovguia.Close
                        
                    XParam = "'" + WClave + "','" _
                                 + WDate + "','" _
                                 + WSaldo + "'"
                    spMovguia = "ModificaMovguiaSaldo " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                End If
                            
            End If
            
            
            
            WMovlab = ""
            
            spMovlab = "ListamovlabNumero"
            Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovlab.RecordCount > 0 Then
                With rstMovlab
                    .MoveLast
                    WMovlab = Str$(rstMovlab!Codigo + 1)
                End With
                rstMovlab.Close
            End If
        
            Renglon = 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = WMovlab
            Call Ceros(Auxi1, 6)
                
            WCodigo = WMovlab
            WRenglon = Str$(Renglon)
            WFecha = FechaRemito.Text
            WFechaOrd = Right$(FechaRemito.Text, 4) + Mid$(FechaRemito.Text, 4, 2) + Left$(FechaRemito.Text, 2)
            WTipo = "M"
            WArticulo = ZCodigo
            WTerminado = "  -     -   "
            WCantidad = ZCantidad
            WMovi = "S"
            WTipoMov = "1"
            Wobservaciones = ""
            Wobservaciones = Left$("Muestra a " + DesClienteRemito.Text, 50)
            WClave = Auxi1 + Auxi
            WDate = Date$
            WMarca = ""
            WLote = ZPartida
                
            XParam = "'" + WClave + "','" _
                         + WCodigo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WTipo + "','" _
                         + WArticulo + "','" _
                         + WTerminado + "','" _
                         + WCantidad + "','" _
                         + WFechaOrd + "','" _
                         + WMovi + "','" _
                         + WTipoMov + "','" _
                         + Wobservaciones + "','" _
                         + WDate + "','" _
                         + WMarca + "','" _
                         + WLote + "'"
                         
            spMovlab = "Altamovlab " + XParam
            Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
            
            
                        
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                Else
                            
            Select Case WTipoPedido
                Case "PG", "CO"
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case "FA"
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    Wempresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
                            
            ZClase = ""
            ZIntervencion = ""
            ZNaciones = ""
            ZImpre = ""
            spTerminado = "ConsultaTerminado " + "'" + ZCodigo + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WCodigo = ZCodigo
                WPedido = Str$(rstTerminado!pedido - Val(ZCantidad))
                WSalidas = Str$(rstTerminado!Salidas + Val(ZCantidad))
                WDate = Date$
                
                ZClase = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
                ZIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
                ZNaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
                ZClase = Trim(ZClase)
                ZIntervencion = Trim(ZIntervencion)
                ZNaciones = Trim(ZNaciones)
                ZImpre = ""
                If Trim(ZClase) <> "" Then
                    ZImpre = "Guia:" + ZIntervencion + " N.ONU:" + ZNaciones + " Clase:" + ZClase
                    ZImpre = Left$(ZImpre, 32)
                End If
                        
                rstTerminado.Close
                    
                XParam = "'" + WCodigo + "','" _
                             + WPedido + "','" _
                             + WSalidas + "','" _
                             + WDate + "'"
                                            
                spTerminado = "ModificaTerminadoFacturas " + XParam
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            End If
                        
            XParam = "'" + ZPartida + "','" _
                         + ZCodigo + "'"
            spHoja = "ListaHojaProducto " + XParam
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                                
                WClave = rstHoja!Clave
                WSaldo = Str$(rstHoja!Saldo - Val(ZCantidad))
                WDate = Date$
                rstHoja.Close
                                    
                XParam = "'" + WClave + "','" _
                             + WDate + "','" _
                             + WSaldo + "'"
                spHoja = "ModificaHojaSaldo " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                
                    Else
                                    
                XParam = "'" + ZCodigo + "','" _
                            + ZPartida + "'"
                spMovguia = "ListaMovguiaLote1 " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    WClave = rstMovguia!Clave
                    WSaldo = Str$(rstMovguia!Saldo - Val(ZCantidad))
                    WDate = Date$
                    rstMovguia.Close
                                    
                    XParam = "'" + WClave + "','" _
                                 + WDate + "','" _
                                 + WSaldo + "'"
                    spMovguia = "ModificaMovguiaSaldo " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
            End If
            
            WMovlab = ""
            
            spMovlab = "ListamovlabNumero"
            Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovlab.RecordCount > 0 Then
                With rstMovlab
                    .MoveLast
                    WMovlab = Str$(rstMovlab!Codigo + 1)
                End With
                rstMovlab.Close
            End If
        
            Renglon = 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = WMovlab
            Call Ceros(Auxi1, 6)
                
            WCodigo = WMovlab
            WRenglon = Str$(Renglon)
            WFecha = FechaRemito.Text
            WFechaOrd = Right$(FechaRemito.Text, 4) + Mid$(FechaRemito.Text, 4, 2) + Left$(FechaRemito.Text, 2)
            WTipo = "T"
            WArticulo = "  -   -   "
            WTerminado = ZCodigo
            WCantidad = ZCantidad
            WMovi = "S"
            WTipoMov = "1"
            Wobservaciones = ""
            Wobservaciones = Left$("Muestra a " + DesClienteRemito.Text, 50)
            WClave = Auxi1 + Auxi
            WDate = Date$
            WMarca = ""
            WLote = ZPartida
                
            XParam = "'" + WClave + "','" _
                         + WCodigo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WTipo + "','" _
                         + WArticulo + "','" _
                         + WTerminado + "','" _
                         + WCantidad + "','" _
                         + WFechaOrd + "','" _
                         + WMovi + "','" _
                         + WTipoMov + "','" _
                         + Wobservaciones + "','" _
                         + WDate + "','" _
                         + WMarca + "','" _
                         + WLote + "'"
                         
            spMovlab = "Altamovlab " + XParam
            Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
            
                        
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
        End If
        
        End If
        
        
        
        ZClase = ""
        ZIntervencion = ""
        ZNaciones = ""
        ZImpre = ""
        ZImpreII = ""
        
        spTerminado = "ConsultaTerminado " + "'" + ZCodigo + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            
            WCodigo = ZCodigo
            WPedido = Str$(rstTerminado!pedido - Val(ZCantidad))
            WSalidas = Str$(rstTerminado!Salidas + Val(ZCantidad))
            WDate = Date$

            ZClase = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
            ZIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
            ZNaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
            ZDescriOnu = IIf(IsNull(rstTerminado!DescriOnu), "", rstTerminado!DescriOnu)
            ZEmbalaje = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
            
            ZClase = Trim(ZClase)
            ZIntervencion = Trim(ZIntervencion)
            ZNaciones = Trim(ZNaciones)
            
            If Trim(ZClase) <> "" Then
                ZImpre = ZDescriOnu
                ZImpreII = "Clase:" + ZClase + " N.ONU:" + ZNaciones + " GRUPO DE EMBALAJE:" + ZEmbalaje
            End If
            
            rstTerminado.Close
                
        End If
        
        Rem ZImpre = "cuia 24234 clase 3534"
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Muestra SET "
        ZSql = ZSql + " Peligroso = " + "'" + ZImpre + "',"
        ZSql = ZSql + " PeligrosoII = " + "'" + ZImpreII + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WMuestra + "'"
        spMuestra = ZSql
        Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                            
                            
        Sql1 = "Select *"
        Sql2 = " FROM Pedido"
        Sql3 = " WHERE Pedido = " + "'" + ZPedido + "'"
        Sql4 = " and Terminado = " + "'" + ZZCodigo + "'"
        Sql5 = " and Facturado = 0"
        spPedido = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            WFacturado = Str$(rstPedido!Facturado + Val(ZCantidad))
            WClavePedido = rstPedido!Clave
            rstPedido.Close
            XParam = "'" + WClavePedido + "','" _
                         + WFacturado + "'"
                                           
            spPedido = "ModificaPedidoFacturas " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        ZSql = ""
        ZSql = ZSql & "UPDATE Pedido SET "
        ZSql = ZSql & "MarcaFactura = " + "'" + "0" + "'"
        ZSql = ZSql & " Where Pedido = " + "'" + ZPedido + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    ListaRemito.GroupSelectionFormula = "{Muestra.Remito} in " + NumeroRemito.Text + " to " + NumeroRemito.Text
    ListaRemito.Destination = 1
    Rem ListaRemito.Destination = 0
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    ListaRemito.SQLQuery = "SELECT Muestra.Cantidad, Muestra.Cliente, Muestra.Razon, Muestra.DescriCliente, Muestra.Remito, " _
                    + "Cliente.Direccion, Cliente.Localidad, Cliente.Cuit, Cliente.DirEntrega " _
                    + "From " _
                    + DSQ + ".dbo.Muestra Muestra, " _
                    + DSQ + ".dbo.Cliente Cliente " _
                    + "Where " _
                    + "Muestra.Cliente = Cliente.Cliente AND " _
                    + "Muestra.Remito >= " + NumeroRemito.Text + " AND " _
                    + "Muestra.Remito <= " + NumeroRemito.Text
                    
    ListaRemito.CopiesToPrinter = 2
    ListaRemito.Connect = Connect()
    
    If Trim(WCliente) = "" Then
        ListaRemito.ReportFileName = "remitoSc.rpt"
            Else
        ListaRemito.ReportFileName = "remito.rpt"
    End If
    ListaRemito.Action = 1
    
    PantaRemito.Visible = False
    
    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    
    Call Proceso_Click


        Sql2 = " FROM Pedido"
        Sql2 = " FROM Pedido"
        Sql2 = " FROM Pedido"
        Sql2 = " FROM Pedido"

        Sql4 = " and Terminado = " + "'" + ZZCodigo + "'"
        Sql4 = " and Terminado = " + "'" + ZZCodigo + "'"
        Sql4 = " and Terminado = " + "'" + ZZCodigo + "'"

End Sub

Private Sub VectorRemito_DblClick()

    If VectorRemito.Col = 0 Or VectorRemito.Col = 1 Then
    
        For Ciclo = 1 To VectorRemito.Cols - 1
            VectorRemito.Col = Ciclo
            VectorRemito.Text = ""
        Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To VectorRemito.Rows - 1
        VectorRemito.Row = Ciclo
        VectorRemito.Col = 1
        WAuxi1 = VectorRemito.Text
        VectorRemito.Col = 2
        WAuxi2 = VectorRemito.Text
        If WAuxi1 <> "" Or WAuxi2 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To VectorRemito.Cols - 1
                VectorRemito.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = VectorRemito.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_VectorII
    
    For Ciclo = 1 To EntraVector
        VectorRemito.Row = Ciclo
        For da = 1 To VectorRemito.Cols - 1
            VectorRemito.Col = da
            VectorRemito.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
End Sub





























Private Sub Etiqueta_Click()

    rowini = Muestra.Row
    RowFin = Muestra.RowSel
    
    Pasa = 0
    LugarEtiqueta = 0
    Call Limpia_VectorIII
    
    For Ciclo = rowini To RowFin
        
        ZNumero = Str$(Ciclo)
        ZPedido = Left$(Muestra.TextMatrix(Ciclo, 1), 6)
        ZFecha = Left$(Muestra.TextMatrix(Ciclo, 2), 10)
        ZCodigo = Left$(Muestra.TextMatrix(Ciclo, 3), 15)
        ZDescripcion = Left$(Muestra.TextMatrix(Ciclo, 4), 50)
        ZCantidad = Left$(Muestra.TextMatrix(Ciclo, 5), 10)
        ZDescriCliente = Left$(Muestra.TextMatrix(Ciclo, 6), 50)
        ZCliente = Left$(Muestra.TextMatrix(Ciclo, 7), 50)
        ZObservaciones = Left$(Muestra.TextMatrix(Ciclo, 8), 50)
        ZFecha2 = Left$(Muestra.TextMatrix(Ciclo, 8), 10)
        ZRemito = Left$(Muestra.TextMatrix(Ciclo, 10), 10)
        ZHojaRuta = Left$(Muestra.TextMatrix(Ciclo, 11), 10)
        ZCodigo2 = Left$(Muestra.TextMatrix(Ciclo, 12), 15)
        ZDescripcion2 = Left$(Muestra.TextMatrix(Ciclo, 13), 50)
        ZLote = Left$(Muestra.TextMatrix(Ciclo, 14), 10)
        ZObservaciones2 = Left$(Muestra.TextMatrix(Ciclo, 15), 50)
        ZCantidad2 = Left$(Muestra.TextMatrix(Ciclo, 16), 10)
        ZActualiza = Left$(Muestra.TextMatrix(Ciclo, 17), 1)


        WMuestra = Auxiliar(Ciclo)
        
        LugarEtiqueta = LugarEtiqueta + 1
        
        VectorEtiqueta.TextMatrix(LugarEtiqueta, 1) = ZDescriCliente
        VectorEtiqueta.TextMatrix(LugarEtiqueta, 2) = ""
        VectorEtiqueta.TextMatrix(LugarEtiqueta, 3) = WMuestra
        
    Next Ciclo
    
    PantaEtiqueta.Visible = True
            
End Sub

Private Sub CancelaEtiqueta_Click()
    Call Limpia_VectorIII
    PantaEtiqueta.Visible = False
End Sub

Private Sub ConfirmaEtiqueta_Click()

    Erase ZImpreEti
    LugarEti = 0
    CiclaLugar = 1

    For Ciclo = 1 To 99

        WMuestra = VectorEtiqueta.TextMatrix(Ciclo, 3)
        
        If Val(WMuestra) <> 0 Then

            Sql1 = "Select *"
            Sql2 = " FROM Muestra"
            Sql3 = " WHERE Codigo = " + "'" + WMuestra + "'"
            spMuestra = Sql1 + Sql2 + Sql3
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            If rstMuestra.RecordCount > 0 Then
                WCliente = rstMuestra!Cliente
                WRazon = rstMuestra!Razon
                ZZZCodigo = rstMuestra!Producto
                rstMuestra.Close
            End If
            
            
            ZClase = ""
            ZIntervencion = ""
            ZNaciones = ""
            ZImpre = ""
            spTerminado = "ConsultaTerminado " + "'" + ZZZCodigo + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                ZClase = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
                ZIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
                ZNaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
                ZClase = Trim(ZClase)
                ZIntervencion = Trim(ZIntervencion)
                ZNaciones = Trim(ZNaciones)
                ZImpre = ""
                If Trim(ZClase) <> "" Then
                    ZImpre = "Guia:" + ZIntervencion + " N.ONU:" + ZNaciones + " Clase:" + ZClase
                    ZImpre = Left$(ZImpre, 50)
                End If
                rstTerminado.Close
            End If
            
            Rem ZImpre = "clase 3534 iner 34534"
            
    
            ZDescripcion = VectorEtiqueta.TextMatrix(Ciclo, 1)
            ZCantidad = VectorEtiqueta.TextMatrix(Ciclo, 2)
            ZMuestra = VectorEtiqueta.TextMatrix(Ciclo, 3)
            
            WMuestra = ZMuestra
            Descri = ZDescripcion
            Cantidad = Val(ZCantidad)
            
            For CicloII = 1 To Cantidad
                If CiclaLugar = 1 Then
                
                    LugarEti = LugarEti + 1
                    CiclaLugar = 2
                    
                    If Val(Wempresa) = 1 Then
                        ZImpreEti(LugarEti, 1) = "SURFACTAN S.A."
                            Else
                        ZImpreEti(LugarEti, 1) = "PELITAL S.A."
                    End If
                    ZImpreEti(LugarEti, 2) = WRazon
                    ZImpreEti(LugarEti, 3) = ZDescripcion
                    ZImpreEti(LugarEti, 4) = "Fecha : " + Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                    If Val(Wempresa) = 1 Then
                        ZImpreEti(LugarEti, 5) = "MALVINAS ARGENTINAS 4589 (1644) VICTORIA"
                        ZImpreEti(LugarEti, 6) = "BS.AS. 4714-4097/4085 surfac@surfactan.com"
                            Else
                        ZImpreEti(LugarEti, 5) = "PELITAL S.A."
                        ZImpreEti(LugarEti, 6) = "PELITAL S.A."
                    
                    End If
                        Else
                        
                    CiclaLugar = 1
                        
                    If Val(Wempresa) = 1 Then
                        ZImpreEti(LugarEti, 7) = "SURFACTAN S.A."
                            Else
                        ZImpreEti(LugarEti, 7) = "PELITAL S.A."
                    End If
                    ZImpreEti(LugarEti, 8) = WRazon
                    ZImpreEti(LugarEti, 9) = ZDescripcion
                    ZImpreEti(LugarEti, 10) = "Fecha : " + Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                    If Val(Wempresa) = 1 Then
                        ZImpreEti(LugarEti, 11) = "MALVINAS ARGENTINAS 4589 (1644) VICTORIA"
                        ZImpreEti(LugarEti, 12) = "BS.AS. 4714-4097/4085 surfac@surfactran.com"
                            Else
                        ZImpreEti(LugarEti, 11) = "PELITAL S.A."
                        ZImpreEti(LugarEti, 12) = "PELITAL S.A."
                    End If
                    
                    ZImpreEti(LugarEti, 13) = ZImpre
                    ZImpreEti(LugarEti, 14) = ZImpre
                    
                End If
                
            Next CicloII
            
        End If
                
    Next Ciclo
    
    
    
    ZSql = "DELETE ImpreEtiqueta"
    spImpreEtiqueta = ZSql
    Set rstImpreEtiqueta = db.OpenRecordset(spImpreEtiqueta, dbOpenSnapshot, dbSQLPassThrough)
    
    SumaII = 0
    Corte = 1
    
    For Ciclo = 1 To LugarEti
    
        SumaII = SumaII + 1
        If SumaII > 4 Then
            SumaII = 1
            Corte = Corte + 1
        End If
        
        ZCiclo = Str$(Ciclo)
        ZCorte = Str$(Corte)
        ZEmpresa = Left$(ZImpreEti(Ciclo, 1), 30)
        ZCliente = Left$(ZImpreEti(Ciclo, 2), 35)
        ZDescripcion = Left$(ZImpreEti(Ciclo, 3), 35)
        ZFecha = ZImpreEti(Ciclo, 4)
        ZDireccionI = Left$(ZImpreEti(Ciclo, 5), 50)
        ZDireccionII = Left$(ZImpreEti(Ciclo, 6), 50)
        ZEmpresaII = Left$(ZImpreEti(Ciclo, 7), 30)
        ZClienteII = Left$(ZImpreEti(Ciclo, 8), 35)
        ZDescripcionII = Left$(ZImpreEti(Ciclo, 9), 35)
        ZFechaII = ZImpreEti(Ciclo, 10)
        ZDireccionIII = Left$(ZImpreEti(Ciclo, 11), 50)
        ZDireccionIV = Left$(ZImpreEti(Ciclo, 12), 50)
        ZPeligroso = Left$(ZImpreEti(Ciclo, 13), 50)
        ZPeligrosoII = Left$(ZImpreEti(Ciclo, 14), 50)
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreEtiqueta ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Corte ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "DireccionI ,"
        ZSql = ZSql + "DireccionII ,"
        ZSql = ZSql + "EmpresaII ,"
        ZSql = ZSql + "ClienteII ,"
        ZSql = ZSql + "DescripcionII ,"
        ZSql = ZSql + "FechaII ,"
        ZSql = ZSql + "DireccionIII ,"
        ZSql = ZSql + "DireccionIV ,"
        ZSql = ZSql + "Peligroso ,"
        ZSql = ZSql + "PeligrosoII )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZCodigo + "',"
        ZSql = ZSql + "'" + ZCorte + "',"
        ZSql = ZSql + "'" + ZEmpresa + "',"
        ZSql = ZSql + "'" + ZCliente + "',"
        ZSql = ZSql + "'" + ZDescripcion + "',"
        ZSql = ZSql + "'" + ZFecha + "',"
        ZSql = ZSql + "'" + ZDireccionI + "',"
        ZSql = ZSql + "'" + ZDireccionII + "',"
        ZSql = ZSql + "'" + ZEmpresaII + "',"
        ZSql = ZSql + "'" + ZClienteII + "',"
        ZSql = ZSql + "'" + ZDescripcionII + "',"
        ZSql = ZSql + "'" + ZFechaII + "',"
        ZSql = ZSql + "'" + ZDireccionIII + "',"
        ZSql = ZSql + "'" + ZDireccionIV + "',"
        ZSql = ZSql + "'" + ZPeligroso + "',"
        ZSql = ZSql + "'" + ZPeligrosoII + "')"
        
        spImpreEtiqueta = ZSql
        Set rstImpreEtiqueta = db.OpenRecordset(spImpreEtiqueta, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    Rem ListaEtiqueta.GroupSelectionFormula = "{ImpreEtiqueta.Codigo} in " + NumeroRemito.Text + " to " + NumeroRemito.Text
    
    ListaEtiqueta.Destination = 1
    Rem ListaRemito.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    ListaEtiqueta.SQLQuery = "SELECT ImpreEtiqueta.Codigo, ImpreEtiqueta.Corte, ImpreEtiqueta.Empresa, ImpreEtiqueta.Cliente, ImpreEtiqueta.Descripcion, ImpreEtiqueta.Fecha, ImpreEtiqueta.DireccionI, ImpreEtiqueta.DireccionII, ImpreEtiqueta.EmpresaII, ImpreEtiqueta.ClienteII, ImpreEtiqueta.DescripcionII, ImpreEtiqueta.FechaII, ImpreEtiqueta.DireccionIII, ImpreEtiqueta.DireccionIV, ImpreEtiqueta.Peligroso, ImpreEtiqueta.PeligrosoII " _
                    + "From " _
                    + DSQ + ".dbo.ImpreEtiqueta ImpreEtiqueta " _
                    + "Where " _
                    + "ImpreEtiqueta.Codigo >= 0 AND " _
                    + "ImpreEtiqueta.Codigo <= 999999"
                    
    ListaEtiqueta.Connect = Connect()
    ListaEtiqueta.ReportFileName = "impreetiqueta.rpt"
    ListaEtiqueta.Action = 1
    
    PantaEtiqueta.Visible = False
    
End Sub

Private Sub VectorEtiqueta_DblClick()
    CantiEtiqueta.Text = VectorEtiqueta.TextMatrix(VectorEtiqueta.Row, 2)
    PantaCantiEtiqueta.Visible = True
    CantiEtiqueta.SetFocus
End Sub

Private Sub CantiEtiqueta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        VectorEtiqueta.TextMatrix(VectorEtiqueta.Row, 2) = CantiEtiqueta.Text
        PantaCantiEtiqueta.Visible = False
    End If
    If KeyAscii = 27 Then
        CantiEtiqueta.Text = ""
    End If
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Ano.Text) <> 0 Then
            Ayuda.Visible = False
            ZAno = Ano.Text
            ColumnaOpcion = 0
            Call Proceso_Click
        End If
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
End Sub

Private Sub AnoII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(AnoII.Text) <> 0 Then
            Ayuda.Visible = False
            ZAnoII = AnoII.Text
            ColumnaOpcion = 0
            Call Proceso_Click
        End If
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
End Sub

Private Sub Limpia_VectorIII()

    VectorEtiqueta.Clear

    Rem ponga la muestra en negritas
    Rem Muestra.Font.Bold = True

    ' Establesco loa Valores de la muestra
    
    VectorEtiqueta.FixedCols = 1
    VectorEtiqueta.Cols = 4
    VectorEtiqueta.FixedRows = 1
    VectorEtiqueta.Rows = 100
    
    VectorEtiqueta.ColWidth(0) = 200
    VectorEtiqueta.Row = 0
    
    For Ciclo = 1 To VectorEtiqueta.Cols - 1
        VectorEtiqueta.Col = Ciclo
        Select Case Ciclo
            Case 1
                VectorEtiqueta.Text = "Descripcion"
                VectorEtiqueta.ColWidth(Ciclo) = 3000
                VectorEtiqueta.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                VectorEtiqueta.Text = "Cantidad"
                VectorEtiqueta.ColWidth(Ciclo) = 1200
                VectorEtiqueta.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                VectorEtiqueta.Text = "Muestra"
                VectorEtiqueta.ColWidth(Ciclo) = 1200
                VectorEtiqueta.ColAlignment(Ciclo) = flexAlignRightCenter
             Case Else
        End Select
    Next Ciclo
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    VectorEtiqueta.AllowUserResizing = flexResizeBoth
    
    VectorEtiqueta.Col = 1
    VectorEtiqueta.Row = 1
    
End Sub

Rem DADA

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    
    If Val(Ano.Text) = 0 Or Val(AnoII.Text) = 0 Then
        Exit Sub
    End If
    
    ZFecDesde = Ano.Text + "0101"
    ZFecHasta = AnoII.Text + "1231"
    
    If KeyAscii = 13 Then
        Ayuda = UCase(Ayuda)
 
        Tipo = Left$(Ayuda, 2)
 
 
        Ayuda = UCase(Ayuda)
        pantalla.Clear
        ColumnaOpcion = Muestra.Col
        Pasa = 0
        Corte = ""
    
    
        Select Case ColumnaOpcion
            Case 2
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Muestra"
                ZSql = ZSql + " Where Muestra.fecha LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
                ZSql = ZSql + " Order by Muestra.FECHA"
                spMuestra = ZSql
                Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                If rstMuestra.RecordCount > 0 Then
                    With rstMuestra
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                If rstMuestra!OrdFecha >= ZFecDesde And rstMuestra!OrdFecha <= ZFecHasta Then
                                    If Pasa = 0 Then
                                        pantalla.AddItem ""
                                        Pasa = 1
                                        Corte = rstMuestra!Fecha
                                    End If
                                    If Corte <> rstMuestra!Fecha Then
                                        pantalla.AddItem Corte
                                        Corte = rstMuestra!Fecha
                                    End If
                                End If
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstMuestra.Close
                End If
                pantalla.AddItem Corte
                pantalla.Visible = True
            
            Case 3
                If Tipo = "PT" Then
                
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Muestra"
                    ZSql = ZSql + " Where Muestra.producto LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
                    ZSql = ZSql + " Order by Muestra.PRODUCTO"
                    spMuestra = ZSql
                    Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMuestra.RecordCount > 0 Then
                        With rstMuestra
                            .MoveFirst
                            Do
                                If .EOF = False Then
                                    If rstMuestra!OrdFecha >= ZFecDesde And rstMuestra!OrdFecha <= ZFecHasta Then
                                        If Pasa = 0 Then
                                            pantalla.AddItem ""
                                            Pasa = 1
                                            Corte = rstMuestra!Producto
                                        End If
                                        If Corte <> rstMuestra!Producto Then
                                            pantalla.AddItem Corte
                                            Corte = rstMuestra!Producto
                                        End If
                                    End If
                                    .MoveNext
                                        Else
                                    Exit Do
                                End If
                            Loop
                        End With
                        rstMuestra.Close
                    End If
                    
                        Else
                
                    Rem BUSCO DY
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Muestra"
                    ZSql = ZSql + " Where Muestra.articulo LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
                    ZSql = ZSql + " Order by Muestra.ARTICULO"
                    spMuestra = ZSql
                    Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMuestra.RecordCount > 0 Then
                        With rstMuestra
                            .MoveFirst
                            Do
                                If .EOF = False Then
                                    If rstMuestra!OrdFecha >= ZFecDesde And rstMuestra!OrdFecha <= ZFecHasta Then
                                        If Pasa = 0 Then
                                            pantalla.AddItem ""
                                            Pasa = 1
                                            Corte = rstMuestra!Articulo
                                        End If
                                        If Corte <> rstMuestra!Articulo Then
                                            pantalla.AddItem Corte
                                            Corte = rstMuestra!Articulo
                                        End If
                                    End If
                                    .MoveNext
                                        Else
                                    Exit Do
                                End If
                            Loop
                        End With
                        rstMuestra.Close
                    End If
   
                End If
        
                pantalla.AddItem Corte
                pantalla.Visible = True
  
            Case 4
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Muestra"
                ZSql = ZSql + " Where Muestra.nombre LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
                ZSql = ZSql + " Order by Muestra.nombre"
                spMuestra = ZSql
                Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                If rstMuestra.RecordCount > 0 Then
                    With rstMuestra
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                If rstMuestra!OrdFecha >= ZFecDesde And rstMuestra!OrdFecha <= ZFecHasta Then
                                    If Pasa = 0 Then
                                        pantalla.AddItem ""
                                        Pasa = 1
                                        Corte = rstMuestra!Nombre
                                    End If
                                    If Corte <> rstMuestra!Nombre Then
                                        pantalla.AddItem Corte
                                        Corte = rstMuestra!Nombre
                                    End If
                                End If
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstMuestra.Close
                End If
                pantalla.AddItem Corte
                pantalla.Visible = True
   
            Case 6
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Muestra"
                ZSql = ZSql + " Where Muestra.descricliente LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
                ZSql = ZSql + " Order by Muestra.descricliente"
                spMuestra = ZSql
                Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                If rstMuestra.RecordCount > 0 Then
                    With rstMuestra
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                If rstMuestra!OrdFecha >= ZFecDesde And rstMuestra!OrdFecha <= ZFecHasta Then
                                    If Pasa = 0 Then
                                        pantalla.AddItem ""
                                        Pasa = 1
                                        Corte = rstMuestra!descricliente
                                    End If
                                    If Corte <> rstMuestra!descricliente Then
                                        pantalla.AddItem Corte
                                        Corte = rstMuestra!descricliente
                                    End If
                                End If
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstMuestra.Close
                End If
                pantalla.AddItem Corte
                pantalla.Visible = True
   
            Case 7
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Muestra"
                ZSql = ZSql + " Where Muestra.Razon LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
                ZSql = ZSql + " Order by Muestra.Razon"
                spMuestra = ZSql
                Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                If rstMuestra.RecordCount > 0 Then
                    With rstMuestra
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                If rstMuestra!OrdFecha >= ZFecDesde And rstMuestra!OrdFecha <= ZFecHasta Then
                                    If Pasa = 0 Then
                                        pantalla.AddItem ""
                                        Pasa = 1
                                        Corte = rstMuestra!Razon
                                    End If
                                    If Corte <> rstMuestra!Razon Then
                                        pantalla.AddItem Corte
                                        Corte = rstMuestra!Razon
                                    End If
                                End If
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstMuestra.Close
                End If
                pantalla.AddItem Corte
                pantalla.Visible = True
   

        End Select
    End If
End Sub

