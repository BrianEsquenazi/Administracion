VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgProveAdicionalParte3 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Datos Adicionales de Proveedore para Orden de Compra de Importacion"
   ClientHeight    =   3180
   ClientLeft      =   285
   ClientTop       =   435
   ClientWidth     =   11430
   LinkTopic       =   "Form2"
   ScaleHeight     =   3180
   ScaleWidth      =   11430
   Begin VB.TextBox Descri13 
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
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   8
      Top             =   1320
      Width           =   9015
   End
   Begin VB.TextBox Nombre 
      BackColor       =   &H00FFFFC0&
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
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   6
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox Proveedor 
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
      Left            =   1440
      MaxLength       =   11
      TabIndex        =   5
      Text            =   " "
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Descri12 
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
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   4
      Top             =   960
      Width           =   9015
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   8160
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Efluentes.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Efluentes de Lavado"
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
      Left            =   5760
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Descri11 
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
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   0
      Top             =   600
      Width           =   9015
   End
   Begin VB.Label lblLabels 
      Caption         =   "Incoterms"
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
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Condiciones de Pago"
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
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   2175
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   3840
      MouseIcon       =   "ProveAdicinalParte4.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "ProveAdicinalParte4.frx":030A
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   5400
      MouseIcon       =   "ProveAdicinalParte4.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "ProveAdicinalParte4.frx":0E56
      ToolTipText     =   "Salida"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Caption         =   "Instruccuiones de Envio"
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
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo "
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "PrgProveAdicionalParte3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstProveedor As Recordset
Dim spProveedor As String
Dim rstProveedorAdicional As Recordset
Dim spProveedorAdicional As String

Sub Imprime_Datos()
    Sql1 = "Select *"
    Sql2 = " FROM Proveedor"
    Sql3 = " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
    spProveedor = Sql1 + Sql2 + Sql3
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        Nombre.Text = Trim(rstProveedor!Nombre)
        rstProveedor.Close
    End If
    Sql1 = "Select *"
    Sql2 = " FROM ProveedorAdicional"
    Sql3 = " Where ProveedorAdicional.Proveedor = " + "'" + Proveedor.Text + "'"
    spProveedorAdicional = Sql1 + Sql2 + Sql3
    Set rstProveedorAdicional = db.OpenRecordset(spProveedorAdicional, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedorAdicional.RecordCount > 0 Then
        Descri11.Text = Trim(rstProveedorAdicional!Descri11)
        Descri12.Text = Trim(rstProveedorAdicional!Descri12)
        Descri13.Text = Trim(rstProveedorAdicional!Descri13)
        rstProveedorAdicional.Close
    End If
End Sub

Private Sub cmdAdd_Click()
    If Val(Proveedor.Text) <> 0 Then
    
        ZDescri11 = Descri11.Text
        ZDescri12 = Descri12.Text
        ZDescri13 = Descri13.Text
        ZDescri21 = ""
        ZDescri22 = ""
        ZDescri23 = ""
        ZDescri31 = ""
        ZDescri32 = ""
        ZDescri33 = ""
        ZDescri40 = ""
        ZDescri41 = ""
        ZDescri42 = ""
        ZDescri43 = ""
        ZDescri44 = ""
        ZDescri45 = ""
        ZDescri46 = ""
        ZDescri47 = ""
        ZDescri48 = ""
        ZDescri49 = ""
        ZDescri51 = ""
        ZDescri52 = ""
        ZDescri53 = ""
        ZDescri54 = ""
        ZDescri55 = ""
        ZDescri56 = ""
        ZDescri57 = ""
        
        Sql1 = "Select *"
        Sql2 = " FROM ProveedorAdicional"
        Sql3 = " Where ProveedorAdicional.Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedorAdicional = Sql1 + Sql2 + Sql3
        Set rstProveedorAdicional = db.OpenRecordset(spProveedorAdicional, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedorAdicional.RecordCount > 0 Then
            rstProveedorAdicional.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE ProveedorAdicional SET "
            ZSql = ZSql + " Descri11 = " + "'" + Descri11.Text + "',"
            ZSql = ZSql + " Descri12 = " + "'" + Descri12.Text + "',"
            ZSql = ZSql + " Descri13 = " + "'" + Descri13.Text + "'"
            ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
            spProveedorAdicional = ZSql
            Set rstProveedorAdicional = db.OpenRecordset(spProveedorAdicional, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ProveedorAdicional ("
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Descri11 ,"
            ZSql = ZSql + "Descri12 ,"
            ZSql = ZSql + "Descri13 ,"
            ZSql = ZSql + "Descri21 ,"
            ZSql = ZSql + "Descri22 ,"
            ZSql = ZSql + "Descri23 ,"
            ZSql = ZSql + "Descri31 ,"
            ZSql = ZSql + "Descri32 ,"
            ZSql = ZSql + "Descri33 ,"
            ZSql = ZSql + "Descri40 ,"
            ZSql = ZSql + "Descri41 ,"
            ZSql = ZSql + "Descri42 ,"
            ZSql = ZSql + "Descri43 ,"
            ZSql = ZSql + "Descri44 ,"
            ZSql = ZSql + "Descri45 ,"
            ZSql = ZSql + "Descri46 ,"
            ZSql = ZSql + "Descri47 ,"
            ZSql = ZSql + "Descri48 ,"
            ZSql = ZSql + "Descri49 ,"
            ZSql = ZSql + "Descri50 ,"
            ZSql = ZSql + "Descri51 ,"
            ZSql = ZSql + "Descri52 ,"
            ZSql = ZSql + "Descri53 ,"
            ZSql = ZSql + "Descri54 ,"
            ZSql = ZSql + "Descri55 ,"
            ZSql = ZSql + "Descri56 ,"
            ZSql = ZSql + "Descri57 )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Proveedor.Text + "',"
            ZSql = ZSql + "'" + ZDescri11 + "',"
            ZSql = ZSql + "'" + ZDescri12 + "',"
            ZSql = ZSql + "'" + ZDescri13 + "',"
            ZSql = ZSql + "'" + ZDescri21 + "',"
            ZSql = ZSql + "'" + ZDescri22 + "',"
            ZSql = ZSql + "'" + ZDescri23 + "',"
            ZSql = ZSql + "'" + ZDescri31 + "',"
            ZSql = ZSql + "'" + ZDescri32 + "',"
            ZSql = ZSql + "'" + ZDescri33 + "',"
            ZSql = ZSql + "'" + ZDescri40 + "',"
            ZSql = ZSql + "'" + ZDescri41 + "',"
            ZSql = ZSql + "'" + ZDescri42 + "',"
            ZSql = ZSql + "'" + ZDescri43 + "',"
            ZSql = ZSql + "'" + ZDescri44 + "',"
            ZSql = ZSql + "'" + ZDescri45 + "',"
            ZSql = ZSql + "'" + ZDescri46 + "',"
            ZSql = ZSql + "'" + ZDescri47 + "',"
            ZSql = ZSql + "'" + ZDescri48 + "',"
            ZSql = ZSql + "'" + ZDescri49 + "',"
            ZSql = ZSql + "'" + ZDescri50 + "',"
            ZSql = ZSql + "'" + ZDescri51 + "',"
            ZSql = ZSql + "'" + ZDescri52 + "',"
            ZSql = ZSql + "'" + ZDescri53 + "',"
            ZSql = ZSql + "'" + ZDescri54 + "',"
            ZSql = ZSql + "'" + ZDescri55 + "',"
            ZSql = ZSql + "'" + ZDescri56 + "',"
            ZSql = ZSql + "'" + ZDescri57 + "')"
            spProveedorAdicional = ZSql
            Set rstProveedorAdicional = db.OpenRecordset(spProveedorAdicional, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        Call cmdClose_Click
        
    End If
    
End Sub

Private Sub cmdClose_Click()

    PrgProveAdicionalParte1.Hide
    Unload Me
    PrgProveAdicional.Show
    
End Sub

Private Sub Descri11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri12.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri11.Text = ""
    End If
End Sub

Private Sub Descri12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri13.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri12.Text = ""
    End If
End Sub

Private Sub Descri13_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri11.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri13.Text = ""
    End If
End Sub

Sub Form_Load()

    Proveedor.Text = ""
    Descri11.Text = ""
    Descri12.Text = ""
    Descri13.Text = ""
    
    Proveedor.Text = WPasaProveedor
    Call Imprime_Datos
    
End Sub

