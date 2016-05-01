VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgHomologaProve 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Homologacion de Muestras"
   ClientHeight    =   8325
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11790
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   8325
   ScaleWidth      =   11790
   Begin VB.ComboBox TipoMp 
      Height          =   315
      Left            =   8160
      TabIndex        =   16
      Top             =   360
      Width           =   2175
   End
   Begin Crystal.CrystalReport ListaRemito 
      Left            =   10440
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      CopiesToPrinter =   2
   End
   Begin VB.Frame PantaExporta 
      Height          =   4695
      Left            =   1440
      TabIndex        =   9
      Top             =   1800
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
      Left            =   11280
      Top             =   0
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
      ItemData        =   "homologaprove.frx":0000
      Left            =   3480
      List            =   "homologaprove.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   4815
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   7335
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12938
      _Version        =   393216
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
      Caption         =   "Modifica / Baja  (F3)"
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
      Width           =   1215
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
      Height          =   780
      Left            =   120
      TabIndex        =   2
      Top             =   0
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
      Left            =   5400
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
      Left            =   6720
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin Crystal.CrystalReport ListaEtiqueta 
      Left            =   10800
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
End
Attribute VB_Name = "PrgHomologaProve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer

Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstHomologa As Recordset
Dim spHomologa As String

Dim XParam As String
Dim Auxiliar(10000)

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
Dim WBorra(1000, 10) As String

Private Sub Alta_Click()
    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    Muestra.Visible = False
    WMuestra = 0
    PrgAltaHomologa.Show
End Sub


Private Sub CancelaExporta_Click()
    PantaExporta.Visible = False
End Sub

Private Sub ConfirmaExporta_Click()

    If NombreExporta.Text = "" Then
        m$ = "Se debe informar un nombre de archivo"
        a% = MsgBox(m$, 0, "Exportacion de Homologacion de Muestras")
        Exit Sub
    End If
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select

    ZSql = ""
    ZSql = ZSql + "UPDATE Homologa SET "
    ZSql = ZSql + "Marca =  " + "'" + "" + "'"
    spHomologa = ZSql
    Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
    
    RowIni = Muestra.Row
    Rowfin = Muestra.RowSel
    
    For Ciclo = RowIni To Rowfin
    
        ZNumero = Str$(Ciclo)
        ZRemito = Left$(Muestra.TextMatrix(Ciclo, 1), 6)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Homologa SET "
        ZSql = ZSql + "Marca =  " + "'" + "X" + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ZRemito + "'"
        spHomologa = ZSql
        Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo

    ListaGRilla.Destination = 2
    ListaGRilla.PrintFileType = crptExcel50
    ListaGRilla.PrintFileName = Dir1.Path + "\" + NombreExporta.Text + ".xls"
    
    ListaGRilla.ReportFileName = "ListaHomologaII.rpt"
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    ListaGRilla.SQLQuery = "SELECT Homologa.Codigo, Homologa.Material, Homologa.Fecha, Homologa.DesProveedor, Homologa.Precio, Homologa.Comentarios, Homologa.Resultado, Homologa.Observaciones, Homologa.Marca " _
            + "From " _
            + DSQ + ".dbo.Homologa Homologa " _
            + "Where " _
            + "Homologa.Marca = 'X'"
    
    ListaGRilla.GroupSelectionFormula = "{Homologa.Marca} = " + Chr$(34) + "X" + Chr$(34)
    ListaGRilla.SelectionFormula = "{Homologa.Marca} = " + Chr$(34) + "X" + Chr$(34)
    
    ListaGRilla.Connect = Connect()
    ListaGRilla.Action = 1
    
    PantaExporta.Visible = False
    
    Call Conecta_Empresa

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

Private Sub Labora_Click()
    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    Muestra.Visible = False
    Fila = Muestra.Row
    WMuestra = Auxiliar(Fila)
    If Val(WMuestra) <> 0 Then
        PrgActualizaHomologa.Show
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
        PrgAltaHomologa.Show
    End If
End Sub

Private Sub cmdClose_Click()
    PrgHomologaProve.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Impresion_Click()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select

    ZSql = ""
    ZSql = ZSql + "UPDATE Homologa SET "
    ZSql = ZSql + "Marca =  " + "'" + "" + "'"
    spHomologa = ZSql
    Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
    
    RowIni = Muestra.Row
    Rowfin = Muestra.RowSel
    
    For Ciclo = RowIni To Rowfin
    
        ZNumero = Str$(Ciclo)
        ZRemito = Left$(Muestra.TextMatrix(Ciclo, 1), 6)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Homologa SET "
        ZSql = ZSql + "Marca =  " + "'" + "X" + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ZRemito + "'"
        spHomologa = ZSql
        Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo

    ListaGRilla.ReportFileName = "ListaHomologa.rpt"

    ListaGRilla.Destination = 1
    Rem ListaGRilla.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    ListaGRilla.SQLQuery = "SELECT Homologa.Codigo, Homologa.Material, Homologa.Solicita, Homologa.Fecha, Homologa.DesProveedor, Homologa.EspecificacionesProve, Homologa.Certificado, Homologa.Precio, Homologa.Origen, Homologa.Ct, Homologa.Nombre, Homologa.Comentarios, Homologa.Entregado, Homologa.FechaII, Homologa.Unidad, Homologa.Resultado, Homologa.Observaciones, Homologa.Responsable, Homologa.CodigoMp, Homologa.ResultadoEntrega, Homologa.ComentariosII, Homologa.Marca " _
            + "From " _
            + DSQ + ".dbo.Homologa Homologa " _
            + "Where " _
            + "Homologa.Marca = 'X'"
    
    ListaGRilla.GroupSelectionFormula = "{Homologa.Marca} = " + Chr$(34) + "X" + Chr$(34)
    ListaGRilla.SelectionFormula = "{Homologa.Marca} = " + Chr$(34) + "X" + Chr$(34)
    
    ListaGRilla.Connect = Connect()
    ListaGRilla.Action = 1
    
    PantaExporta.Visible = False
    
    Call Conecta_Empresa
    
End Sub

Private Sub Proceso_Click()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select

    Call Limpia_Vector
        
    Select Case ColumnaOpcion
        Case 0, 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Homologa"
            ZSql = ZSql + " Order by Homologa.Codigo"
            spHomologa = ZSql
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Homologa"
            ZSql = ZSql + " Where Homologa.Material = " + "'" + Seleccion + "'"
            ZSql = ZSql + " Order by Homologa.Codigo"
            spHomologa = ZSql
        Case 5
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Homologa"
            ZSql = ZSql + " Where Homologa.DesProveedor = " + "'" + Seleccion + "'"
            ZSql = ZSql + " Order by Homologa.Codigo"
            spHomologa = ZSql
        Case Else
    End Select
            
    Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
    If rstHomologa.RecordCount > 0 Then
        With rstHomologa
    
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    ZTipo = IIf(IsNull(rstHomologa!TipoMp), "1", rstHomologa!TipoMp)
                    
                    If TipoMp.ListIndex = 0 Or ZTipo = TipoMp.ListIndex Then
                
                        WLugar = WLugar + 1
                        Auxiliar(WLugar) = Str$(rstHomologa!Codigo)
                        
                        Muestra.TextMatrix(WLugar, 1) = rstHomologa!Codigo
                        Muestra.TextMatrix(WLugar, 2) = rstHomologa!Material
                        Muestra.TextMatrix(WLugar, 3) = rstHomologa!Solicita
                        Muestra.TextMatrix(WLugar, 4) = rstHomologa!Fecha
                        Muestra.TextMatrix(WLugar, 5) = rstHomologa!DesProveedor
                        Muestra.TextMatrix(WLugar, 6) = rstHomologa!EspecificacionesProve
                        Muestra.TextMatrix(WLugar, 7) = rstHomologa!Certificado
                        Muestra.TextMatrix(WLugar, 8) = rstHomologa!Precio
                        Muestra.TextMatrix(WLugar, 9) = rstHomologa!Origen
                        Muestra.TextMatrix(WLugar, 10) = rstHomologa!Ct
                        Muestra.TextMatrix(WLugar, 11) = rstHomologa!Nombre
                        Muestra.TextMatrix(WLugar, 12) = rstHomologa!Comentarios
                        Muestra.TextMatrix(WLugar, 13) = rstHomologa!Entregado
                        Muestra.TextMatrix(WLugar, 14) = rstHomologa!FechaII
                        Muestra.TextMatrix(WLugar, 15) = rstHomologa!Unidad
                        Muestra.TextMatrix(WLugar, 16) = rstHomologa!Resultado
                        Muestra.TextMatrix(WLugar, 17) = rstHomologa!Observaciones
                        Muestra.TextMatrix(WLugar, 18) = rstHomologa!Responsable
                        Muestra.TextMatrix(WLugar, 19) = rstHomologa!codigomp
                        Muestra.TextMatrix(WLugar, 20) = rstHomologa!ResultadoEntrega
                        Muestra.TextMatrix(WLugar, 21) = rstHomologa!ComentariosII
                    
                    End If
                    
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
        
        End With
        rstHomologa.Close
    End If
    
    Call Conecta_Empresa
    
    Muestra.Visible = True
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    If WPosi1 <> 0 And WPosi2 <> 0 And WPosi3 <> 0 Then
        Muestra.TopRow = WPosi1
        Muestra.Col = WPosi3
        Muestra.Row = WPosi2
        WPosi1 = 0
        WPosi3 = 0
        WPosi2 = 0
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

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select

    ColumnaOpcion = Muestra.Col
    WPosi1 = 1
    WPosi2 = 1
    WPosi3 = 1
    
    Pantalla.Clear
    Select Case ColumnaOpcion
        Case 2
            Pasa = 0
            Corte = ""
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Homologa"
            ZSql = ZSql + " Order by Homologa.Material"
            spHomologa = ZSql
            Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
            If rstHomologa.RecordCount > 0 Then
            With rstHomologa
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            Pantalla.AddItem ""
                            Pasa = 1
                            Corte = rstHomologa!Material
                        End If
                        If Corte <> rstHomologa!Material Then
                            Pantalla.AddItem Corte
                            Corte = rstHomologa!Material
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            Pantalla.AddItem Corte
            rstHomologa.Close
            Pantalla.Visible = True
            End If
            
        Case 5
            Pasa = 0
            Corte = ""
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Homologa"
            ZSql = ZSql + " Order by Homologa.DesProveedor"
            
            spHomologa = ZSql
            Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
            If rstHomologa.RecordCount > 0 Then
            
            With rstHomologa
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            Pantalla.AddItem ""
                            Pasa = 1
                            Corte = rstHomologa!DesProveedor
                        End If
                        If Corte <> rstHomologa!DesProveedor Then
                            Pantalla.AddItem Corte
                            Corte = rstHomologa!DesProveedor
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            Pantalla.AddItem Corte
            rstHomologa.Close
            Pantalla.Visible = True
            
            End If
            
        Case Else
        
    End Select
    
    Call Conecta_Empresa
            
End Sub

Private Sub Form_Activate()

    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If

    OPEN_FILE_Empresa

    TipoMp.Clear
    
    TipoMp.AddItem "Total"
    TipoMp.AddItem "Materias Primas"
    TipoMp.AddItem "Envases"
    
    TipoMp.ListIndex = 0
    Rem Call Proceso_Click
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
    Muestra.Cols = 22
    Muestra.FixedRows = 1
    Muestra.Rows = 10000
    
    Muestra.ColWidth(0) = 200
    Muestra.Row = 0
    
    For Ciclo = 1 To Muestra.Cols - 1
        Muestra.Col = Ciclo
        Select Case Ciclo
            Case 1
                Muestra.Text = "Numero"
                Muestra.ColWidth(Ciclo) = 800
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                Muestra.Text = "Material"
                Muestra.ColWidth(Ciclo) = 2000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                Muestra.Text = "Solicita"
                Muestra.ColWidth(Ciclo) = 1000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                Muestra.Text = "Fecha"
                Muestra.ColWidth(Ciclo) = 1000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                Muestra.Text = "Proveedor"
                Muestra.ColWidth(Ciclo) = 2500
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 6
                Muestra.Text = "Especif.del Proveedor"
                Muestra.ColWidth(Ciclo) = 2000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 7
                Muestra.Text = "Cert.Analisis"
                Muestra.ColWidth(Ciclo) = 1000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                Muestra.Text = "Precio"
                Muestra.ColWidth(Ciclo) = 1000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 9
                Muestra.Text = "Origen"
                Muestra.ColWidth(Ciclo) = 1100
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 10
                Muestra.Text = "C / T"
                Muestra.ColWidth(Ciclo) = 1500
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 11
                Muestra.Text = "Denominacion Comercial"
                Muestra.ColWidth(Ciclo) = 2500
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 12
                Muestra.Text = "Comentarios de Compras"
                Muestra.ColWidth(Ciclo) = 2500
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 13
                Muestra.Text = "Entregado a "
                Muestra.ColWidth(Ciclo) = 1500
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 14
                Muestra.Text = "Fecha"
                Muestra.ColWidth(Ciclo) = 1000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 15
                Muestra.Text = "Negocio"
                Muestra.ColWidth(Ciclo) = 1000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 16
                Muestra.Text = "Resultado"
                Muestra.ColWidth(Ciclo) = 2500
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 17
                Muestra.Text = "Observaciones"
                Muestra.ColWidth(Ciclo) = 2500
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 18
                Muestra.Text = "Responsable"
                Muestra.ColWidth(Ciclo) = 1500
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 19
                Muestra.Text = "M.P."
                Muestra.ColWidth(Ciclo) = 1000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 20
                Muestra.Text = "Resultado 1ra Entrega"
                Muestra.ColWidth(Ciclo) = 2500
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 21
                Muestra.Text = "Comantarios Laboratorio"
                Muestra.ColWidth(Ciclo) = 2500
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Muestra.AllowUserResizing = flexResizeBoth
    
    Muestra.Col = 1
    Muestra.Row = 1
    
End Sub

Private Sub pantalla_Click()
    If Pantalla.ListIndex <> 0 Then
        Seleccion = Pantalla.Text
            Else
        Seleccion = ""
        ColumnaOpcion = 0
    End If
    Pantalla.Visible = False
    Call Proceso_Click
End Sub

Private Sub TipoMp_Click()
    Call Proceso_Click
End Sub

