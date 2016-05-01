VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaPyme 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Deuda de Pyme banco Nacion"
   ClientHeight    =   3825
   ClientLeft      =   3240
   ClientTop       =   2025
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   3825
   ScaleWidth      =   5655
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1455
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Width           =   3735
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4680
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ListaPyme.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva Compras"
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
      Left            =   4320
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "ListaPyme.frx":0000
      Left            =   480
      List            =   "ListaPyme.frx":0007
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaPyme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim rstIvaComp As Recordset
Dim spIvaComp As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim XParam As String

Private Sub Acepta_Click()

    Listado.WindowTitle = "Listado de Deuda Pyme Nacion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.GroupSelectionFormula = "{CtaCtePrv.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Listado.SelectionFormula = "{CtaCtePrv.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT CtaCtePrv.Proveedor, CtaCtePrv.fecha, CtaCtePrv.Total, CtaCtePrv.Saldo, CtaCtePrv.OrdFecha, CtaCtePrv.Impre, CtaCtePrv.DesProveOriginal, CtaCtePrv.FacturaOriginal, CtaCtePrv.Cuota, CtaCtePrv.ImporteOriginal, CtaCtePrv.FechaOriginal, CtaCtePrv.Interes, CtaCtePrv.IvaInteres, CtaCtePrv.Referencia  " _
            + "From " _
            + DSQ + ".dbo.CtaCtePrv CtaCtePrv " _
            + "Where " _
            + "CtaCtePrv.Proveedor = '10077777777' AND " _
            + "CtaCtePrv.Saldo <> 0 AND " _
            + "CtaCtePrv.OrdFecha >= '" + WDesde + "' AND " _
            + "CtaCtePrv.OrdFecha <= '" + WHasta + "'"
                       
    Listado.Connect = Connect()
    
    Listado.Action = 1
End Sub

Private Sub Cancela_Click()
    PrgListaPyme.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
    Rem If KeyAscii >= 48 And KeyAscii <= 57 Then
    Rem     If Desde.SelStart = 1 Then
    Rem         Desde.Text = Mid$(Desde.Text, 1, Desde.SelStart) + Chr$(KeyAscii) + Mid$(Desde.Text, Desde.SelStart + 1, 10)
    Rem         If Mid$(Desde.Text, 3, 1) <> "/" Then
    Rem             Desde.Text = Desde.Text + "/"
    Rem         End If
    Rem         KeyAscii = 0
    Rem         Desde.SelStart = 3
    Rem     End If
    Rem     If Desde.SelStart = 4 Then
    Rem         Desde.Text = Mid$(Desde.Text, 1, Desde.SelStart) + Chr$(KeyAscii) + Mid$(Desde.Text, Desde.SelStart + 1, 10)
    Rem         If Mid$(Desde.Text, 6, 1) <> "/" Then
    Rem             Desde.Text = Desde.Text + "/"
    Rem         End If
    Rem         KeyAscii = 0
    Rem         Desde.SelStart = 6
    Rem     End If
    Rem End If
End Sub


Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

