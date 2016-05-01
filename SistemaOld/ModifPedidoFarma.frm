VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgModifPedidoFarma 
   AutoRedraw      =   -1  'True
   Caption         =   "Modificacion de Pedidos de Colorantes / DY / DW  / DS"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11805
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11805
   Begin VB.CommandButton ImprePdfIII 
      Caption         =   "Cert. Ana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10440
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton ImprePdfII 
      Caption         =   "Hoja Seg."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9240
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton ImprePdf 
      Caption         =   "Hoja y Cert."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7920
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton PlantaIV 
      Caption         =   "Planta IV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5520
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton ImpreEti 
      Caption         =   "Etiquetas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4320
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin MSMask.MaskEdBox HastaFecha 
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   480
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
   Begin MSMask.MaskEdBox DesdeFecha 
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   120
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
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   6375
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11245
      _Version        =   327680
      Rows            =   4000
      Cols            =   11
      BackColor       =   16777215
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11160
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WImpreEtiDy.rpt"
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Lee datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6720
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
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
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "PrgModifPedidoFarma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim XParam As String
Dim TotalPedidos As Integer
Dim WGraba As String
Dim WSaldo As Double

Dim ZZRuta As String
Dim ZZEstado As String

Dim ZZLote1 As String
Dim ZZLote2 As String
Dim ZZLote3 As String
Dim ZZLote4 As String
Dim ZZLote5 As String
Dim ZLote(100) As String
Dim ZZLote(100) As String

Dim ZZProceso As Integer

Dim WCantiLote1 As Double
Dim WCantiLote2  As Double
Dim WCantiLote3  As Double
Dim WCantiLote4  As Double
Dim WCantiLote5  As Double
Dim WCantiLote  As Double

Dim XLote(100, 30) As String
Dim WLote(100, 5) As String
Dim WCanti(100, 5) As String
Dim WEti(100, 5) As String
Dim WTipo(100, 5) As String

Dim ZImpreConcepto(200) As String

Private Sub cmdClose_Click()
    With rstEmpresa
        .Close
    End With
    PrgModifColor.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Load()
    
    Call Limpia_Vector
    
    Muestra.Font.Bold = True
    
    Muestra.ColWidth(0) = 200
    Muestra.ColWidth(1) = 800
    Muestra.ColWidth(2) = 1200
    Muestra.ColWidth(3) = 800
    Muestra.ColWidth(4) = 2000
    Muestra.ColWidth(5) = 1200
    Muestra.ColWidth(6) = 800
    Muestra.ColWidth(7) = 1000
    Muestra.ColWidth(8) = 700
    Muestra.ColWidth(9) = 700
    Muestra.ColWidth(10) = 2000
    
    Muestra.ColAlignment(10) = flexAlignLeftCenter
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
    Muestra.Col = 7
    Muestra.Text = "Estado"
    
    Muestra.Col = 8
    Muestra.Text = "$ Pendiente"
    
    Call Proceso_Click
    
End Sub

Private Sub Proceso_Click()

    WSalida = "N"
    
    Call Limpia_Vector
    

    ZSql = ""
    ZSql = ZSql + "Select Pedido.TipoPedido, Pedido.FechaOrd, Pedido.Autorizo, Pedido.Clave, Pedido.Cantidad, Pedido.Facturado, Pedido.Terminado, Pedido.Clave, Pedido.Pedido, Pedido.Fecha, Pedido.Cliente, Pedido.FecEntrega, Pedido.TipoPed, Pedido.Impresion, Pedido.MarcaFactura, Pedido.Precio, Pedido.FechaInicial, Pedido.MarcaAutorizacion"
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where Pedido.MarcaFarma = " + "'" + "N" + "'"
    ZSql = ZSql + " and Pedido.Renglon = 1"
    ZSql = ZSql + " Order by Clave"
    spPedido = ZSql
        
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        
        With rstPedido
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                        
                    Renglon = Renglon + 1
                    Muestra.Row = Renglon
                        
    
                    Muestra.Col = 1
                    Muestra.Text = Pusing("######", Str$(rstPedido!Numero))
                    
                    Muestra.Col = 2
                    Muestra.Text = rstPedido!Fecha
            
                    Muestra.Col = 3
                    Muestra.Text = rstPedido!Cliente
                    
                    Muestra.Col = 5
                    Muestra.Text = rttpedido!FEntrega
                    
                    Select Case rstPedido!TipoPedido
                        Case 0
                            Muestra.Col = 6
                            Muestra.Text = "Normal"
                        Case 1
                            Muestra.Col = 6
                            Muestra.Text = "A Fecha"
                        Case 2
                            Muestra.Col = 6
                            Muestra.Text = "Fec.Lim."
                        Case 3
                            Muestra.Col = 6
                            Muestra.Text = "Urgente"
                        Case 4
                            Muestra.Col = 6
                            Muestra.Text = "Ret.Cli"
                        Case 5
                            Muestra.Col = 6
                            Muestra.Text = "Muestra"
                        Case Else
                            Muestra.Col = 6
                            Muestra.Text = ""
                    End Select
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            End If
            
        End With
        
    End If
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    Muestra.Row = 1
    Muestra.Col = 1
    Muestra.TopRow = 1
    
    Rem Muestra.SetFocus

End Sub

Private Sub Limpia_Vector()
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
End Sub

