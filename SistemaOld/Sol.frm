VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgSol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitudes de Compra"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11760
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8625
   ScaleWidth      =   11760
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Impresion de Solicitud de Orden de Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   17
      Top             =   7320
      Width           =   3495
   End
   Begin VB.CommandButton Impresion 
      Caption         =   "Impresion de Solicitud de Orden de Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   16
      Top             =   7320
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid Datos 
      Height          =   4815
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8493
      _Version        =   327680
      Rows            =   4000
      Cols            =   6
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Regreso a Consulta de Solicitudes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   14
      Top             =   6480
      Width           =   3495
   End
   Begin VB.TextBox Solicitante 
      Height          =   285
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   13
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox Planta 
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   11
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   4800
      TabIndex        =   9
      Top             =   6360
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   8
      Text            =   " "
      Top             =   840
      Width           =   9855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4080
      Top             =   8400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "orden.rpt"
      Destination     =   3
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   17
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   9120
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Solicitud 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   3
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   8160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1620
      ItemData        =   "Sol.frx":0000
      Left            =   4800
      List            =   "Sol.frx":0007
      TabIndex        =   0
      Top             =   6720
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Label9 
      Caption         =   "Solicitante"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Planta"
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Observaciones"
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro de Solicitud"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ver As String
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Cantidad As Single
Private XCantidad As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstSolic As Recordset
Dim spSolic As String
Dim XParam As String
Dim Vector(100, 2) As String
Private TipoConsulta As String
Private XVector(3, 4) As String
Private Auxi As String
Private WAuxi As String
Private WSaldo As Double
Private Desdelugar As Integer


Private Sub cmdClose_Click()

    With rstAuxiliar
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    PrgSol.Hide
    Unload Me
    
    PrgMIrasol.Show
    
End Sub

Private Sub Command1_Click()

 WEmpresa = ver
 XEmpresa = WEmpresa
 
 Call Conecta_Empresa
 
 Rem   WEmpresa = "0001"
 Rem   txtOdbc = "Empresa08"
  Rem  strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
 Rem   Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    Listado.ReportFileName = "ImpreInsumosII.rpt"
    
    Listado.WindowTitle = "Emision de Solicitud de Compras de Insumos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
  Rem  RowIni = Datos.Row
    RowIni = 1
    Rowfin = Datos.RowSel = 1
    
Rem    For Ciclo = RowIni To Rowfin
    
      Rem  WSolicitud = Datos.TextMatrix(Ciclo, 1)
      
     WSolicitud = Solicitud.Text
        Listado.GroupSelectionFormula = "{solic.Solicitud} in " + WSolicitud + " to " + WSolicitud
  Rem  Listado.Destination = 1
       Listado.Destination = 0
    Rem  Listado.SQLQuery = "SELECT solic.Solicitud, solic.Fecha, solic.Planta, solic.Solicitante, Solic.Observaciones, solic.Entrega, solic.Cantidad, solic.Descripcion " _

        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
    Rem  Listado.SQLQuery = "SELECT solic.Solicitud, solic.Fecha, solic.Planta, solic.Solicitante, Solic.Observaciones, solic.Entrega, solic.Cantidad, solic.Descripcion, solic.articulo " _
      rem         + "From " _
        rem            + DSQ + ".dbo.solic solic " _
        rem            + "Where " _
        rem            + "solic.Solicitud >= " + WSolicitud + " AND solic.Solicitud <= " + WSolicitud
                            
   Listado.SQLQuery = "SELECT solic.Solicitud, solic.Fecha, solic.Planta, solic.Solicitante, Solic.Observaciones, solic.Entrega, solic.Cantidad, solic.Descripcion, solic.articulo " _
                    + "From " _
                    + DSQ + ".dbo.solic solic, " _
                    + DSQ + ".dbo.articulo articulo " _
                    + "Where " _
                    + "solic.Solicitud >= " + WSolicitud + " AND solic.Solicitud <= " + WSolicitud _
                    + "AND  solic.articulo = articulo.codigo "
  
        
        
        
        Rem Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
        Listado.Connect = Connect()
        Listado.Action = 1
    
   Rem Next Ciclo
    
    Call Conecta_Empresa
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Form_Load()

    Datos.Clear

    Datos.ColWidth(0) = 50
    Datos.ColWidth(1) = 1400
    Datos.ColWidth(2) = 3620
    Datos.ColWidth(3) = 1100
    Datos.ColWidth(4) = 1100
    Datos.ColWidth(5) = 3800

    Datos.Row = 0
    
    Datos.Col = 1
    Datos.Text = "Producto"
    
    Datos.Col = 2
    Datos.Text = "Descripcion"
    
    Datos.Col = 3
    Datos.Text = "Cantidad"
    
    Datos.Col = 4
    Datos.Text = "F.Entrega"
    
    Datos.Col = 5
    Datos.Text = "Observaciones"


    Solicitud.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Observaciones.Text = ""
    Select Case Val(WEmpresa)
        Case 1
            Planta.Text = "Surfactan S.A. Planta I"
        Case 2
            Planta.Text = "Pellital S.A. Planta I"
        Case 3
            Planta.Text = "Surfactan S.A. Planta II"
        Case 4
            Planta.Text = "Pellital S.A: Planta II"
        Case 5
            Planta.Text = "Surfactan S.A. Planta III"
        Case 6
            Planta.Text = "Surfactan S.A: Planta IV"
        Case 7
            Planta.Text = "Surfactan S.A: Planta V"
        Case 8
            Planta.Text = "Pellital S.A: Planta V"
        Case 9
            Planta.Text = "Pellital S.A: Planta VI"
        Case 10
            Planta.Text = "Surfactan S.A: Planta VI"
        Case 11
            Planta.Text = "Surfactan S.A: Planta VII"
        Case Else
            Planta.Text = "Surfactan S.A: Planta I"
    End Select
    
    Solicitante.Text = ""
    Solicitud.Text = ""
 
    Rem spSolic = "ListaSolicitudNumero"
    Rem Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstSolic.RecordCount > 0 Then
    Rem     With rstSolic
    Rem         .MoveLast
    Rem         Solicitud.Text = rstSolic!Solicitud + 1
    Rem     End With
    Rem     rstSolic.Close
    Rem End If

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgSol.Caption = "Ingreso de Solicituded de Compras :  " + !Nombre
        End If
    End With
 
    Solicitud.Text = WXSol
    Call Proceso_Click
    ver = WEmpresa
End Sub

Private Sub Proceso_Click()

    
    Renglon = 0
    Erase Vector
    
    spSolic = "ListaSolicitud " + "'" + Solicitud.Text + "'"
    Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
    If rstSolic.RecordCount > 0 Then
            
        With rstSolic
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Fecha.Text = rstSolic!Fecha
                    Observaciones.Text = rstSolic!Observaciones
                    Planta.Text = rstSolic!Planta
                    Solicitante.Text = rstSolic!Solicitante
            
                    Renglon = Renglon + 1
                    Datos.Row = Renglon
                        
                    Datos.Col = 1
                    Datos.Text = rstSolic!Articulo
                        
                    If rstSolic!Marca = "X" Then
                        Datos.Col = 3
                        Datos.Text = ""
                            Else
                        Datos.Col = 3
                        Datos.Text = Pusing("###,###.##", rstSolic!Cantidad)
                    End If
                        
                    Datos.Col = 4
                    Datos.Text = rstSolic!Entrega
                        
                    Datos.Col = 5
                    Datos.Text = rstSolic!Obser
                    
                    Vector(Renglon, 1) = rstSolic!Articulo
            
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstSolic.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
        Auxi1 = Vector(Renglon, 1)
        Datos.Row = Renglon
    
        spArticulo = "ConsultaArticulo " + "'" + Auxi1 + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Datos.Col = 2
            Datos.Text = rstArticulo!Descripcion
            rstArticulo.Close
        End If
    Next Da

End Sub


Private Sub Impresion_Click()

        Rem Open "DAda.TXT" For Output As #1
        Open "lpt1" For Output As #1
        
        WObservaciones = Left$(Observaciones.Text + Space$(100), 100)
        
        With rstEmpresa
            .Index = "Empresa"
            Claveven$ = WEmpresa
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Impretit = !Nombre
                    Else
                Impretit = ""
            End If
        End With
    
        For Ci = 1 To 2

        '  Copia 1
        
        Print #1, Chr$(18)
        Print #1, ""
        Print #1, ""

        Print #1, Tab(1); "--------------------------------------------------------------------------------"
        
        Print #1, Tab(1); "|";
        Print #1, Impretit;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|                                                                              |"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Solicitud......: ";
        Print #1, Tab(25); Alinea("######", Solicitud.Text);
        Print #1, Tab(50); "Fecha : "; Fecha.Text;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Observaciones..:"; Tab(25); Left$(WObservaciones, 50);
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(25); Right$(WObservaciones, 50);
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Planta.........:"; Tab(25); Planta.Text;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Solicitante....:"; Tab(25); Solicitante.Text;
        Print #1, Tab(80); "|"
        
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, "|Producto  |  Descripcion      |Cantidad|Fecha Ent.|  Observaciones            |"
        Print #1, "--------------------------------------------------------------------------------"

        WCantidad = 0
        Valor = 0
        XRenglon = 0
        
        For a = 0 To 3
        
            For iRow = 0 To 9
            
                XRenglon = XRenglon + 1
                Datos.Row = XRenglon
                        
                Datos.Col = 1
                WArticulo = Datos.Text
                
                Datos.Col = 2
                WDescripcion = Datos.Text
                        
                Datos.Col = 3
                WCantidad = Val(Datos.Text)
                        
                Datos.Col = 4
                WEntrega = Datos.Text
                        
                Datos.Col = 5
                WObser = Datos.Text
                    
                If Left$(WArticulo, 2) <> "" And Left$(WArticulo, 2) <> Space$(2) And WCantidad <> 0 Then
                
                        Print #1, Tab(1); "|"; WArticulo;
                        Print #1, Tab(12); "|"; Left$(WDescripcion, 15);
                        Print #1, Tab(32); "|"; Alinea("###,###", Str$(WCantidad));
                        Print #1, Tab(41); "|"; WEntrega;
                        Print #1, Tab(52); "|"; Left$(WObser, 25);
                        Print #1, Tab(80); "|"

                End If
                                        
            Next iRow
        Next a

        For Ciclo = WCantidad To 15
            Print #1, "|          |                   |        |          |                           |"
        Next Ciclo

        Print #1, "--------------------------------------------------------------------------------"
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""

        Next Ci
        
        Close #1

 End Sub

