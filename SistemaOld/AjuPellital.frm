VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgAju 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Muestras a Clientes"
   ClientHeight    =   8325
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11790
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   8325
   ScaleWidth      =   11790
   Begin Crystal.CrystalReport ListaRemito 
      Left            =   10800
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      CopiesToPrinter =   2
   End
   Begin VB.Frame PantaRemito 
      Height          =   7095
      Left            =   7440
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   4215
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
         Left            =   2160
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
         Left            =   360
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
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   3975
         _ExtentX        =   7011
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
      Height          =   540
      Left            =   6840
      TabIndex        =   16
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame PantaExporta 
      Height          =   4695
      Left            =   2040
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
      Height          =   540
      Left            =   5160
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin Crystal.CrystalReport ListaGRilla 
      Left            =   11280
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
      ItemData        =   "AjuPellital.frx":0000
      Left            =   3480
      List            =   "AjuPellital.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   4815
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   7455
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13150
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
      Height          =   540
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   1575
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
      Height          =   540
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Alta 
      Caption         =   "Alta (F1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
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
      Height          =   540
      Left            =   8520
      TabIndex        =   1
      Top             =   120
      Width           =   1575
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
      Height          =   540
      Left            =   10200
      TabIndex        =   0
      Top             =   120
      Width           =   1575
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
Dim XParam As String
Dim Auxiliar(10000)
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
Dim WPasa(10000) As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Dim LugarRemito As Integer
Dim WBorra(1000, 10) As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String

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
        .Seek "=", Val(WEmpresa)
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
        ZCodigo2 = Left$(Muestra.TextMatrix(Ciclo, 11), 15)
        ZDescripcion2 = Left$(Muestra.TextMatrix(Ciclo, 12), 50)
        ZLote = Left$(Muestra.TextMatrix(Ciclo, 13), 10)
        ZObservaciones2 = Left$(Muestra.TextMatrix(Ciclo, 14), 50)
        ZCantidad2 = Left$(Muestra.TextMatrix(Ciclo, 15), 10)
        ZActualiza = Left$(Muestra.TextMatrix(Ciclo, 16), 1)
        
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
        PrgMuestraLabo.Show
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
        .Seek "=", Val(WEmpresa)
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
        ZRemito = Left$(Muestra.TextMatrix(Ciclo, 1), 6)
        ZFecha = Left$(Muestra.TextMatrix(Ciclo, 2), 10)
        ZCodigo = Left$(Muestra.TextMatrix(Ciclo, 3), 15)
        ZDescripcion = Left$(Muestra.TextMatrix(Ciclo, 4), 50)
        ZCantidad = Left$(Muestra.TextMatrix(Ciclo, 5), 10)
        ZDescriCliente = Left$(Muestra.TextMatrix(Ciclo, 6), 50)
        ZCliente = Left$(Muestra.TextMatrix(Ciclo, 7), 50)
        ZObservaciones = Left$(Muestra.TextMatrix(Ciclo, 8), 50)
        ZFecha2 = Left$(Muestra.TextMatrix(Ciclo, 9), 10)
        ZRemito = Left$(Muestra.TextMatrix(Ciclo, 10), 10)
        ZCodigo2 = Left$(Muestra.TextMatrix(Ciclo, 11), 15)
        ZDescripcion2 = Left$(Muestra.TextMatrix(Ciclo, 12), 50)
        ZLote = Left$(Muestra.TextMatrix(Ciclo, 13), 10)
        ZObservaciones2 = Left$(Muestra.TextMatrix(Ciclo, 14), 50)
        ZCantidad2 = Left$(Muestra.TextMatrix(Ciclo, 15), 10)
        ZActualiza = Left$(Muestra.TextMatrix(Ciclo, 16), 1)
        
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

    DesClienteRemito.Text = ""
    Call Limpia_Vector
        
    Select Case ColumnaOpcion
        Case 0, 1
            spMuestra = "ListaMuestraTotal "
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
                    
                    Muestra.TextMatrix(WLugar, 2) = Left$(rstMuestra!Fecha, 5)
        
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
                    
                    ZEnsayo2 = IIf(IsNull(rstMuestra!ensayo2), "", rstMuestra!ensayo2)
                    If ZEnsayo2 <> "" And ZEnsayo2 <> Space$(15) Then
                        Rem Muestra.Col = 11
                        Muestra.TextMatrix(WLugar, 11) = ZEnsayo2
                    End If
                    
                    ZArticulo2 = IIf(IsNull(rstMuestra!Articulo2), "", rstMuestra!Articulo2)
                    If ZArticulo2 <> "" And ZArticulo2 <> "  -   -   " Then
                        Rem Muestra.Col = 11
                        Muestra.TextMatrix(WLugar, 11) = ZArticulo2
                    End If
            
                    ZProducto2 = IIf(IsNull(rstMuestra!Producto2), "", rstMuestra!Producto2)
                    If ZProducto2 <> "" And ZProducto2 <> "  -     -   " Then
                        Rem Muestra.Col = 11
                        Muestra.TextMatrix(WLugar, 11) = ZProducto2
                    End If
        
                    Rem Muestra.Col = 12
                    ZNombre2 = IIf(IsNull(rstMuestra!Nombre2), "", rstMuestra!Nombre2)
                    Muestra.TextMatrix(WLugar, 12) = ZNombre2
        
                    Rem Muestra.Col = 13
                    Muestra.TextMatrix(WLugar, 13) = IIf(IsNull(rstMuestra!lote2), "", rstMuestra!lote2)
        
                    Rem Muestra.Col = 14
                    Muestra.TextMatrix(WLugar, 14) = IIf(IsNull(rstMuestra!Observaciones2), "", rstMuestra!Observaciones2)
        
                    Rem Muestra.Col = 15
                    Muestra.TextMatrix(WLugar, 15) = IIf(IsNull(rstMuestra!Cantidad2), "", rstMuestra!Cantidad2)
        
                    WStock2 = IIf(IsNull(rstMuestra!Stock2), "", rstMuestra!Stock2)
                    If Val(WStock2) = 1 Then
                        Rem Muestra.Col = 16
                        Muestra.TextMatrix(WLugar, 16) = "          S"
                            Else
                        Rem Muestra.Col = 16
                        Muestra.TextMatrix(WLugar, 16) = ""
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

Private Sub Muestra_DblClick()

    ColumnaOpcion = Muestra.Col
    WPosi1 = 1
    WPosi2 = 1
    WPosi3 = 1
    
    pantalla.Clear
    Select Case ColumnaOpcion
        Case 2
        
            Pasa = 0
            corte = ""
            spMuestra = "ListaMuestraFecha"
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            With rstMuestra
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            pantalla.AddItem ""
                            Pasa = 1
                            corte = rstMuestra!Fecha
                        End If
                        If corte <> rstMuestra!Fecha Then
                            pantalla.AddItem corte
                            corte = rstMuestra!Fecha
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            pantalla.AddItem corte
            rstMuestra.Close
            pantalla.Visible = True
            
        Case 3
            Lista.Clear
            
            Pasa = 0
            corte = ""
            spMuestra = "ListaMuestraArticulo"
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
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
                            corte = WAgrega
                        End If
                        If corte <> WAgrega Then
                            Lista.AddItem corte
                            corte = WAgrega
                        End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            Lista.AddItem corte
            
            Erase WPasa
            Hasta = Lista.ListCount
            For Ciclo = 0 To Hasta - 1
                Lista.ListIndex = Ciclo
                WPasa(Ciclo + 1) = Lista.Text
            Next Ciclo
            
            pantalla.AddItem ""
            
            Pasa = 0
            corte = ""
            For Ciclo = 1 To Hasta
                WAgrega = WPasa(Ciclo)
                If WAgrega <> "" Then
                    If Pasa = 0 Then
                        Pasa = 1
                        corte = WAgrega
                    End If
                    If corte <> WAgrega Then
                        pantalla.AddItem corte
                        corte = WAgrega
                    End If
                End If
            Next Ciclo
            pantalla.AddItem corte
            
            rstMuestra.Close
            pantalla.Visible = True
            
        Case 4
            Pasa = 0
            corte = ""
            spMuestra = "ListaMuestraNombre"
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            With rstMuestra
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            pantalla.AddItem ""
                            Pasa = 1
                            corte = rstMuestra!Nombre
                        End If
                        If corte <> rstMuestra!Nombre Then
                            pantalla.AddItem corte
                            corte = rstMuestra!Nombre
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            pantalla.AddItem corte
            rstMuestra.Close
            pantalla.Visible = True
            
        Case 5
            Pasa = 0
            corte = ""
            spMuestra = "ListaMuestraCantidad"
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            With rstMuestra
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            pantalla.AddItem ""
                            Pasa = 1
                            corte = rstMuestra!Cantidad
                        End If
                        If corte <> rstMuestra!Cantidad Then
                            pantalla.AddItem corte
                            corte = rstMuestra!Cantidad
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            pantalla.AddItem corte
            rstMuestra.Close
            pantalla.Visible = True
            
        Case 6
            Pasa = 0
            corte = ""
            spMuestra = "ListaMuestraDescriCliente"
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            With rstMuestra
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            pantalla.AddItem ""
                            Pasa = 1
                            corte = rstMuestra!descricliente
                        End If
                        If corte <> rstMuestra!descricliente Then
                            pantalla.AddItem corte
                            corte = rstMuestra!descricliente
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            pantalla.AddItem corte
            rstMuestra.Close
            pantalla.Visible = True
            
            
        Case 7
            Pasa = 0
            corte = ""
            spMuestra = "ListaMuestraCliente"
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            With rstMuestra
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            pantalla.AddItem ""
                            Pasa = 1
                            corte = rstMuestra!Razon
                        End If
                        If corte <> rstMuestra!Razon Then
                            pantalla.AddItem corte
                            corte = rstMuestra!Razon
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            pantalla.AddItem corte
            rstMuestra.Close
            pantalla.Visible = True
            
        Case 8
            Pasa = 0
            corte = ""
            spMuestra = "ListaMuestraObservaciones"
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            With rstMuestra
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            pantalla.AddItem ""
                            Pasa = 1
                            corte = rstMuestra!Observaciones
                        End If
                        If corte <> rstMuestra!Observaciones Then
                            pantalla.AddItem corte
                            corte = rstMuestra!Observaciones
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            pantalla.AddItem corte
            rstMuestra.Close
            pantalla.Visible = True
            
        Case Else
        
    End Select
            
    
    Rem Muestra.Col = 10
    Rem Muestra.Col = 1
    Rem WXSol = Muestra.Text
    Rem PrgSol.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
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
    
    Rem dada
    
    Muestra.FixedCols = 1
    Muestra.Cols = 17
    Muestra.FixedRows = 1
    Muestra.Rows = 10000
    
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
                Muestra.ColWidth(Ciclo) = 650
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                Muestra.Text = "Codigo"
                Muestra.ColWidth(Ciclo) = 1200
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                Muestra.Text = "Descripcion"
                Muestra.ColWidth(Ciclo) = 1400
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                Muestra.Text = "Cantidad"
                Muestra.ColWidth(Ciclo) = 800
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 6
                Muestra.Text = "Nombre para el Cliente"
                Muestra.ColWidth(Ciclo) = 1850
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 7
                Muestra.Text = "Cliente"
                Muestra.ColWidth(Ciclo) = 1550
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
                Muestra.Text = "Codigo Conf."
                Muestra.ColWidth(Ciclo) = 1350
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 12
                Muestra.Text = "Descripcion"
                Muestra.ColWidth(Ciclo) = 2000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 13
                Muestra.Text = "Lote"
                Muestra.ColWidth(Ciclo) = 1200
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 14
                Muestra.Text = "Observaciones"
                Muestra.ColWidth(Ciclo) = 2300
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 15
                Muestra.Text = "Cantidad"
                Muestra.ColWidth(Ciclo) = 1000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 16
                Muestra.Text = "Actualiza Stock"
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
    VectorRemito.Cols = 4
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
                VectorRemito.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                VectorRemito.Text = "Pedido"
                VectorRemito.ColWidth(Ciclo) = 800
                VectorRemito.ColAlignment(Ciclo) = flexAlignLeftCenter
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
    Call Proceso_Click
End Sub

Private Sub Remito_Click()

    rowini = Muestra.Row
    RowFin = Muestra.RowSel
    
    Pasa = 0
    
    For Ciclo = rowini To RowFin
        
        ZNumero = Str$(Ciclo)
        ZRemito = Left$(Muestra.TextMatrix(Ciclo, 1), 6)
        ZFecha = Left$(Muestra.TextMatrix(Ciclo, 2), 10)
        ZCodigo = Left$(Muestra.TextMatrix(Ciclo, 3), 15)
        ZDescripcion = Left$(Muestra.TextMatrix(Ciclo, 4), 50)
        ZCantidad = Left$(Muestra.TextMatrix(Ciclo, 5), 10)
        ZDescriCliente = Left$(Muestra.TextMatrix(Ciclo, 6), 50)
        ZCliente = Left$(Muestra.TextMatrix(Ciclo, 7), 50)
        ZObservaciones = Left$(Muestra.TextMatrix(Ciclo, 8), 50)
        ZFecha2 = Left$(Muestra.TextMatrix(Ciclo, 8), 10)
        ZRemito = Left$(Muestra.TextMatrix(Ciclo, 10), 10)
        ZCodigo2 = Left$(Muestra.TextMatrix(Ciclo, 11), 15)
        ZDescripcion2 = Left$(Muestra.TextMatrix(Ciclo, 12), 50)
        ZLote = Left$(Muestra.TextMatrix(Ciclo, 13), 10)
        ZObservaciones2 = Left$(Muestra.TextMatrix(Ciclo, 14), 50)
        ZCantidad2 = Left$(Muestra.TextMatrix(Ciclo, 15), 10)
        ZActualiza = Left$(Muestra.TextMatrix(Ciclo, 16), 1)
        
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
        Sql1 = "Select *"
        Sql2 = " FROM Muestra"
        Sql3 = " WHERE Codigo = " + "'" + WMuestra + "'"
        spMuestra = Sql1 + Sql2 + Sql3
        Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
        If rstMuestra.RecordCount > 0 Then
            WRemito = IIf(IsNull(rstMuestra!Remito), "0", rstMuestra!Remito)
            rstMuestra.Close
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
        ZRemito = Left$(Muestra.TextMatrix(Ciclo, 1), 6)
        ZFecha = Left$(Muestra.TextMatrix(Ciclo, 2), 10)
        ZCodigo = Left$(Muestra.TextMatrix(Ciclo, 3), 15)
        ZDescripcion = Left$(Muestra.TextMatrix(Ciclo, 4), 50)
        ZCantidad = Left$(Muestra.TextMatrix(Ciclo, 5), 10)
        ZDescriCliente = Left$(Muestra.TextMatrix(Ciclo, 6), 50)
        ZCliente = Left$(Muestra.TextMatrix(Ciclo, 7), 50)
        ZObservaciones = Left$(Muestra.TextMatrix(Ciclo, 8), 50)
        ZFecha2 = Left$(Muestra.TextMatrix(Ciclo, 8), 10)
        ZRemito = Left$(Muestra.TextMatrix(Ciclo, 10), 10)
        ZCodigo2 = Left$(Muestra.TextMatrix(Ciclo, 11), 15)
        ZDescripcion2 = Left$(Muestra.TextMatrix(Ciclo, 12), 50)
        ZLote = Left$(Muestra.TextMatrix(Ciclo, 13), 10)
        ZObservaciones2 = Left$(Muestra.TextMatrix(Ciclo, 14), 50)
        ZCantidad2 = Left$(Muestra.TextMatrix(Ciclo, 15), 10)
        ZActualiza = Left$(Muestra.TextMatrix(Ciclo, 16), 1)

        WMuestra = Auxiliar(Ciclo)
        
        DesClienteRemito.Text = ZCliente
        
        LugarRemito = LugarRemito + 1
        
        VectorRemito.TextMatrix(LugarRemito, 1) = ZDescriCliente
        If Val(ZCantidad2) <> 0 Then
            VectorRemito.TextMatrix(LugarRemito, 2) = ZCantidad2
                Else
            VectorRemito.TextMatrix(LugarRemito, 2) = ZCantidad
        End If
        VectorRemito.TextMatrix(LugarRemito, 3) = WMuestra
        
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
            
        WMuestra = ZMuestra
        Descri = ZDescripcion
        Cantidad = Val(ZCantidad)
        
        Sql1 = "UPDATE Muestra SET "
        Sql2 = " Remito = " + "'" + NumeroRemito.Text + "',"
        Sql3 = " RazonRemito = " + "'" + DesClienteRemito.Text + "'"
        Sql4 = " Where Codigo = " + "'" + WMuestra + "'"
        spMuestra = Sql1 + Sql2 + Sql3 + Sql4
        Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                
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
    ListaRemito.ReportFileName = "remito.rpt"
    ListaRemito.Action = 1
    
    PantaRemito.Visible = False
    
    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    
    Call Proceso_Click

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



