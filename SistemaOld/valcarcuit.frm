VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgValcarcuit 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Valores en Cartera por Cuit"
   ClientHeight    =   3450
   ClientLeft      =   3150
   ClientTop       =   735
   ClientWidth     =   5400
   LinkTopic       =   "Form2"
   ScaleHeight     =   3450
   ScaleWidth      =   5400
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   3855
      Begin VB.TextBox Cuit 
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
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   9
         Text            =   " "
         Top             =   1200
         Width           =   1815
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
         Left            =   480
         TabIndex        =   7
         Top             =   2400
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
         Left            =   1920
         TabIndex        =   6
         Top             =   2400
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
         Left            =   720
         TabIndex        =   5
         Top             =   1680
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
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   1680
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hastafec 
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   720
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
      Begin MSMask.MaskEdBox Desdefec 
         Height          =   255
         Left            =   1920
         TabIndex        =   0
         Top             =   360
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
      Begin VB.Label Label2 
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
         TabIndex        =   12
         Top             =   360
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
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cuit"
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
         Top             =   1200
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5040
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "valcar.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Saldos de Cuenta Corriente de Proveedores"
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
      Left            =   4800
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4680
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgValcarcuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WAuxiliar As String
Private WLinea As Single
Private Cheques(10) As Double
Private Impre(10) As Double
Private WTotal(10) As Double
Private WRecibo As Double
Private WCheque As String
Private WBanco As String
Private Impre1 As String
Private Impre2 As String
Private Impre3 As String
Private Impre4 As String
Private Impre5 As String
Private Impre6 As String
Private da As Single
Private WCliente As String
Dim rstRecibo As Recordset
Dim spRecibo As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim XParam As String
Dim ZZTrabajo(1000, 3) As String

Private Sub Acepta_Click()

    If Trim(Cuit.Text) = "" Then Exit Sub

    Rem On Error GoTo Control_Error
    
    WDesdeFec = Right$(Desdefec.Text, 4) + Mid$(Desdefec.Text, 4, 2) + Left$(Desdefec.Text, 2)
    WHastaFec = Right$(Hastafec.Text, 4) + Mid$(Hastafec.Text, 4, 2) + Left$(Hastafec.Text, 2)
    
    Listado.WindowTitle = "Listado de Valores en Cartera por Cuit"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{Recibos.Cuit} in " + Chr$(34) + Cuit.Text + Chr$(34) + " to " + Chr$(34) + Cuit.Text + Chr$(34)
    Dos = " and {Recibos.FechaOrd2} in " + Chr$(34) + WDesdeFec + Chr$(34) + " to " + Chr$(34) + WHastaFec + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.ReportFileName = "ListaValoresCuit.rpt"
    
    
    
    Erase ZZTrabajo
    ZZLugar = 0


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.Cuit = " + "'" + Cuit.Text + "'"
    ZSql = ZSql + " and Recibos.FechaOrd2 >= " + "'" + WDesdeFec + "'"
    ZSql = ZSql + " and Recibos.FechaOrd2 <= " + "'" + WHastaFec + "'"
    ZSql = ZSql + " Order by Recibos.Clave"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
    
        With rstRecibos
            .MoveFirst
            Do
                If .EOF = False Then
                    
                    ZZLugar = ZZLugar + 1
                    
                    ZZTrabajo(ZZLugar, 1) = rstRecibos!Clave
                    ZZTrabajo(ZZLugar, 2) = rstRecibos!Destino
                    ZZTrabajo(ZZLugar, 3) = rstRecibos!Estado2
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstRecibos.Close
        
    End If
    
    For Ciclo = 1 To ZZLugar
    
        ZZClaveRecibo = ZZTrabajo(Ciclo, 1)
        ZZDestino = ZZTrabajo(Ciclo, 2)
        ZZEstado = ZZTrabajo(Ciclo, 3)
    
        If ZZEstado = "X" And Left$(ZZDestino, 8) <> "Deposito" Then
        
            ZZProveedor = ""
            ZZDesProveedor = ""
            ZZZOrden = ""
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pagos"
            ZSql = ZSql + " Where Pagos.ClaveRecibo = " + "'" + ZZClaveRecibo + "'"
            spPagos = ZSql
            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
            If rstPagos.RecordCount > 0 Then
                ZZProveedor = rstPagos!Proveedor
                ZZOrden = rstPagos!Orden
                rstPagos.Close
            End If
            
            If ZZProveedor <> "" Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Proveedor"
                ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + ZZProveedor + "'"
                spProveedor = ZSql
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    ZZDesProveedor = Trim(RstProveedor!Nombre)
                    RstProveedor.Close
                End If
            End If
            
            If ZZDesProveedor <> "" Then
                ZZDestino = Left$(ZZDesProveedor + "  O.P.:" + ZZOrden, 50)
            End If
            
        End If
                    
        ZSql = ""
        ZSql = ZSql + "UPDATE Recibos SET "
        ZSql = ZSql + " ImpreObserva = " + "'" + ZZDestino + "'"
        ZSql = ZSql + " Where Recibos.clave = " + "'" + ZZClaveRecibo + "'"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)

    Next Ciclo
    
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT Recibos.Recibo, Recibos.Renglon, Recibos.Cliente, Recibos.Fecha, Recibos.Numero2, Recibos.Fecha2, Recibos.banco2, Recibos.Importe2, Recibos.Estado2, Recibos.FechaOrd2, Recibos.Cuit, Recibos.ImpreObserva, " _
            + "Cliente.Razon " _
            + "From " _
            + DSQ + ".dbo.Recibos Recibos, " _
            + DSQ + ".dbo.Cliente Cliente " _
            + "Where " _
            + "Recibos.Cliente = Cliente.Cliente AND " _
            + "Recibos.Cuit >= '" + Cuit.Text + "' AND " _
            + "Recibos.Cuit <= '" + Cuit.Text + "' AND " _
            + "Recibos.FechaOrd2 >= '" + WDesdeFec + "' AND " _
            + "Recibos.FechaOrd2 <= '" + WHastaFec + "'"

    If Impresora.Value = True Then
       Listado.Destination = 1
           Else
       Listado.Destination = 0
    End If
    Listado.Action = 1
    Exit Sub
    
Control_Error:
     coderr = Err
     Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgValcarcuit.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desdefec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desdefec.Text, Auxi)
        If Auxi = "S" Then
            Hastafec.SetFocus
                Else
            Desdefec.SetFocus
        End If
    End If
End Sub

Private Sub Hastafec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hastafec.Text, Auxi)
        If Auxi = "S" Then
            Cuit.SetFocus
                Else
            Hastafec.SetFocus
        End If
    End If
End Sub

Private Sub Cuit_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdefec.SetFocus
    End If
End Sub

Sub Form_Load()

    Desdefec.Text = "  /  /    "
    Hastafec.Text = "  /  /    "
    Cuit.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

