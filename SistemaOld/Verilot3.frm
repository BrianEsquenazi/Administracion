VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgVerilot3 
   AutoRedraw      =   -1  'True
   Caption         =   "Control de Saldos de Lotes de Productos Terminados"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wlotemat.rpt"
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
      Left            =   6000
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "Verilot3.frx":0000
      Left            =   120
      List            =   "Verilot3.frx":0007
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   5880
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgVerilot3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WInicial As Double
Private WOrden As String
Private WClave As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim XParam As String
Dim Vector(10000, 4) As String
Private WDescripcion As String
Private WSaldo As Double
Private XSaldo As Double


Private Sub Acepta_Click()

    Rem Open "lpt1" For Output As #1
    Open "h.TXT" For Output As #1

    Erase Vector
    Renglon = 0
    
    spHoja = "ListaHojaTotal"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
            With rstHoja
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        If rstHoja!lote1 <> 0 Or rstHoja!Canti1 <> 0 Then
                        
                            WCantidad = 0
                            
                            If rstHoja!lote1 <> 0 Then
                                WCantidad = WCantidad + rstHoja!Canti1
                            End If
                            If rstHoja!lote2 <> 0 Then
                                WCantidad = WCantidad + rstHoja!Canti2
                            End If
                            If rstHoja!lote3 <> 0 Then
                                WCantidad = WCantidad + rstHoja!Canti3
                            End If
                            
                            If rstHoja!Cantidad <> WCantidad Then
                                Print #1, rstHoja!Hoja, rstHoja!Producto, rstHoja!Articulo, rstHoja!Terminado, rstHoja!Cantidad, WCantidad
                            End If
                        End If
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstHoja.Close
    End If
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Close
    PrgVerilot3.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub
