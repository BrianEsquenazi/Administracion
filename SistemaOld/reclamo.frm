VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Prgreclamo 
   AutoRedraw      =   -1  'True
   Caption         =   "seguimiento de reclamos "
   ClientHeight    =   6090
   ClientLeft      =   6555
   ClientTop       =   435
   ClientWidth     =   7350
   LinkTopic       =   "Form2"
   ScaleHeight     =   10.742
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   6.306
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1200
      Width           =   7095
   End
   Begin VB.CommandButton Confirma 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Frame Aviso 
      BackColor       =   &H00FF8080&
      Height          =   975
      Left            =   -120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "PedpenII.rpt"
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
Attribute VB_Name = "Prgreclamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XParam As String
Dim WVector(10000) As String
Dim ZPlanta(100) As String
Dim LeeAviso(100, 3) As String
Dim CargaEmpresa(12, 2) As String
Dim CargaEmpresaII(12, 2) As String
Dim Cliente As String

Dim spreclamo As String

Dim rstreclamo As Recordset
Dim LugarVector As Integer



Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Confirma_Click()

    XEmpresa = WEmpresa
    
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
    Rem Borra la solicitud original
        
    Sql1 = "DELETE reclamo"
    Sql2 = " Where codigo = " + "'" + Cliente + "'"
    spInsumo = Sql1 + Sql2
    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                      
  Sql1 = "INSERT INTO reclamo ("
            Sql2 = "codigo ,"
            Sql3 = "observacion )"
            Sql4 = "Values ("
            Sql5 = "'" + Cliente + "',"
            Sql6 = "'" + Text1 + "')"
           spreclamo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6
            Set rstreclamo = db.OpenRecordset(spreclamo, dbOpenSnapshot, dbSQLPassThrough)
                
    Prgreclamo.Hide
    Unload Me
    Rem PrgMiraInsumosII.Show
                
  
                
                
                
Rem                Call Conecta_Empresa
            Rem    Exit Sub
            
End Sub

Private Sub Form_Load()
        
        Cliente = cliente2
        Text2.Text = descliente2
        XEmpresa = WEmpresa
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
      Rem  If Val(Solicitud.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM reclamo"
            Sql3 = " Where reclamo.codigo = " + "'" + Cliente + "'"
            Sql4 = " Order by Codigo"
            spreclamo = Sql1 + Sql2 + Sql3 + Sql4
            Set rstreclamo = db.OpenRecordset(spreclamo, dbOpenSnapshot, dbSQLPassThrough)
            If rstreclamo.RecordCount > 0 Then
              comp = IIf(IsNull(rstreclamo!observacion), " ", rstreclamo!observacion)



                Text1 = comp
                                         
                rstreclamo.Close
                    Else
                
 End If
    
    
        Rem If Val(Solicitud.Text) = 0 Then
        Rem    Sql1 = "Select Max(Solicitud), Solicitud"
        Rem    Sql2 = " FROM Insumos"
        Rem    Sql3 = " Group By Solicitud"
        Rem    Sql4 = " Order By Solicitud"
        Rem    spInsumo = Sql1 + Sql2 + Sql3 + Sql4
        Rem    Set rstinsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
        Rem    If rstinsumo.RecordCount > 0 Then
        Rem        With rstinsumo
        Rem            .MoveLast
        Rem            Solicitud.Text = rstinsumo!Solicitud + 1
        Rem        End With
        Rem        rstinsumo.Close
        Rem    End If
        Rem End If
        
        Rem If Val(Solicitud.Text) = 0 Then
        Rem    Solicitud.Text = "1"
        Rem    End If
    
        Rem Borra la solicitud original
        
        Rem Sql1 = "DELETE Insumos"
        Rem Sql2 = " Where Solicitud = " + "'" + Solicitud.Text + "'"
        Rem   spInsumo = Sql1 + Sql2
        Rem Set rstinsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
        
       Rem Renglon = 0


End Sub
