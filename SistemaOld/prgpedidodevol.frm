VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPedidoDevol 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Devolucion de Mercaderia"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8625
   ScaleWidth      =   11775
   Visible         =   0   'False
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
      Height          =   1260
      Left            =   4800
      TabIndex        =   25
      Top             =   5880
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Observaciones 
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
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   27
      Text            =   " "
      Top             =   840
      Width           =   7095
   End
   Begin VB.CommandButton Inserta 
      Caption         =   "Inserta Renglon"
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
      Left            =   9720
      TabIndex        =   26
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   24
      Text            =   " "
      Top             =   5880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton BorraConsulta 
      Caption         =   "Borra Consulta"
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
      Left            =   10800
      TabIndex        =   23
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton ConsultaPro 
      Caption         =   "Consulta Producto"
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
      Left            =   9720
      TabIndex        =   22
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton ConsultaCli 
      Caption         =   "Consulta Cliente"
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
      Left            =   10800
      TabIndex        =   21
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Baja 
      Caption         =   "  Baja  Solicitud"
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
      Left            =   9720
      TabIndex        =   20
      Top             =   1440
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10200
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10800
      TabIndex        =   19
      Top             =   960
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   4560
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Cliente 
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
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   16
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4320
      TabIndex        =   14
      Top             =   120
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
   Begin VB.TextBox Pedido 
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
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9720
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10800
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9720
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   5160
      Width           =   8895
      Begin VB.TextBox WPartida 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   29
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Text            =   " "
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox WCantidad 
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
         Height          =   300
         Left            =   6120
         MaxLength       =   10
         TabIndex        =   5
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label WDescripcion 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10800
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3735
      Left            =   240
      OleObjectBlob   =   "prgpedidodevol.frx":0000
      TabIndex        =   2
      Top             =   1320
      Width           =   8895
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6840
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Observaciones"
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
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label DesCliente 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   3360
      TabIndex        =   17
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
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
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
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
      Left            =   3360
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Solicitud"
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
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgPedidodevol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 4 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Auxi As String
Private XLinea As Single
Private WDirentrega As String
Private WInicio As Integer
Private Auxiliar(100, 3) As String
Private WTermi As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstPedidoDevol As Recordset
Dim spPedidoDevol As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstImprePed As Recordset
Dim spImprePed As String
Dim WGraba As String
Dim XParam As String
Dim WPasa(100, 6) As String
Dim IngreVector(20000, 4) As String
Dim EntraVector As Integer
Dim Partida As String

Private Sub Baja_Click()

    Rem Renglon = Renglon + 1
    Rem Lugar1 = Int((Renglon - 1) / 10) * 10
    Rem Lugar2 = Renglon - Lugar1
                
    Rem DBGrid1.FirstRow = Lugar1
    Rem DBGrid1.Row = Lugar2 - 1
    
    Rem DBGrid1.Col = 0
    Rem DBGrid1.Text = ""
    
    Rem Erase Auxiliar
    Rem WRenglon = 0

    Rem spPedidoDevol = "ListaPedidoDevol " + "'" + Pedido.Text + "'"
    Rem Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPedidoDevol.RecordCount > 0 Then
    Rem
    Rem         With rstPedidoDevol
    Rem         .MoveFirst
    Rem         Do
    Rem             If .EOF = False Then
    Rem
    Rem                 WRenglon = WRenglon + 1
    Rem
    Rem                 Auxiliar(WRenglon, 1) = rstPedidoDevol!Terminado
    Rem                 Auxiliar(WRenglon, 2) = rstPedidoDevol!Cantidad
    Rem                 Auxiliar(WRenglon, 3) = IIf(IsNull(rstPedidoDevol!Tipopro), "", rstPedidoDevol!Tipopro)
    Rem
    Rem                 .MoveNext
    Rem                     Else
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem     End With
    Rem     rstPedidoDevol.Close
    Rem End If
    
    spPedidoDevol = "BorrarPedidoDevol " + "'" + Pedido.Text + "'"
    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenDynaset, dbSQLPassThrough)
    
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Pedido.SetFocus

End Sub

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = "  -     -   "
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    Rem DBGrid1.Col = 3
    Rem DBGrid1.Text = ""
    
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPartida.Text = ""
    Rem WPrecio.Caption = ""
    WLinea.Text = ""
    
    WArticulo.SetFocus
    Call DBGrid1_GotFocus
    
End Sub

Private Sub BorraConsulta_Click()
    Pantalla.Visible = False
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub cmdClose_Click()
    Call Limpia_Click
    PrgPedidodevol.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub ConsultaCli_Click()

    XIndice = 0

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear
    
    Ayuda.Height = 285
    Ayuda.Left = 2040
    Ayuda.Top = 0
    Ayuda.Width = 8055
    
    Ayuda.Visible = True
    Ayuda.Text = ""
    
    spClientes = "ListaClienteConsulta"
    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then

    With rstClientes
        .MoveFirst
        Do
            If .EOF = False Then
                IngresaItem = rstClientes!Cliente + " " + rstClientes!Razon
                Pantalla.AddItem IngresaItem
                IngresaItem = rstClientes!Cliente
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstClientes.Close

    End If
    
    Pantalla.Visible = True
    
    Pantalla.Height = 7740
    Pantalla.Left = 2040
    Pantalla.Top = 360
    Pantalla.Width = 8175
    
End Sub

Private Sub ConsultaPro_Click()

    XIndice = 1

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear
    
    spPrecios = "ListaPreciosCliente " + "'" + Cliente.Text + "'"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
    
        With rstPrecios
            .MoveFirst
            Do
                If .EOF = False Then
                    If Cliente.Text = rstPrecios!Cliente Then
                        If rstPrecios!Precio <> Null Then
                                IngresaItem = rstPrecios!Terminado + "   " + rstPrecios!Descripcion + Str$(rstPrecios!Precio)
                                    Else
                                Auxi$ = Str$(rstPrecios!Precio)
                                Call Mascara("###,###.##", Auxi$)
                                IngresaItem = rstPrecios!Terminado + "   " + Auxi$ + "  " + rstPrecios!Descripcion
                        End If
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstPrecios!Cliente + rstPrecios!Terminado
                        WIndice.AddItem IngresaItem
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPrecios.Close
    End If
    
    Erase IngreVector
    EntraVector = 0
    
    spPreciosMp = "ListaPreciosClienteMp " + "'" + Cliente.Text + "'"
    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
    If rstPreciosMp.RecordCount > 0 Then
    
        With rstPreciosMp
            .MoveFirst
            Do
                If .EOF = False Then
                    If Cliente.Text = rstPreciosMp!Cliente Then
                        ZArticulo = Left$(rstPreciosMp!Articulo, 3) + "00" + Right$(rstPreciosMp!Articulo, 7)
                        EntraVector = EntraVector + 1
                        IngreVector(EntraVector, 1) = ZArticulo
                        IngreVector(EntraVector, 2) = rstPreciosMp!Cliente
                        IngreVector(EntraVector, 3) = rstPreciosMp!Articulo
                        IngreVector(EntraVector, 4) = IIf(IsNull(rstPreciosMp!Precio), "0", rstPreciosMp!Precio)
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPreciosMp.Close
    End If
    
    For CicloVector = 1 To EntraVector
        
        ZTerminado = IngreVector(CicloVector, 1)
        WCliente = IngreVector(CicloVector, 2)
        WArti = IngreVector(CicloVector, 3)
        ZPrecio = IngreVector(CicloVector, 4)
        ZDescripcion = ""
        
        spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZDescripcion = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
        Auxi$ = ZPrecio
        Call Mascara("###,###.##", Auxi$)
        IngresaItem = ZTerminado + "  " + Auxi$ + "  " + ZDescripcion
        Pantalla.AddItem IngresaItem
        IngresaItem = WCliente + WArti
        WIndice.AddItem IngresaItem
        
    Next CicloVector
    
    Pantalla.Visible = True
    
    Pantalla.Height = 1740
    Pantalla.Left = 3480
    Pantalla.Top = 6720
    Pantalla.Width = 8175

End Sub

Private Sub Inserta_Click()

    WPrimer = DBGrid1.FirstRow
    WFila = DBGrid1.Row
    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
    
    Erase WPasa
    
    DBGrid1.Refresh

    Erase WPasa
    Salida = "N"
    
    XCounter = 0
    XLugar = 0

    For a = 0 To 5
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
    
            WRow = iRow
            DBGrid1.Col = 0
            DBGrid1.Row = iRow
    
            If DBGrid1.Text = "" Then
                Salida = "S"
                Exit For
            End If
    
            XLugar = XLugar + 1
            XCounter = XCounter + 1
                
            WRow = iRow
            DBGrid1.Row = iRow
            
            DBGrid1.Col = 0
            WPasa(XCounter, 1) = DBGrid1.Text
            DBGrid1.Text = ""
            
            DBGrid1.Col = 1
            WPasa(XCounter, 2) = DBGrid1.Text
            DBGrid1.Text = ""
            
            DBGrid1.Col = 2
            WPasa(XCounter, 3) = DBGrid1.Text
            DBGrid1.Text = ""
            
            Rem DBGrid1.Col = 3
            Rem WPasa(XCounter, 4) = DBGrid1.Text
            Rem DBGrid1.Text = ""
                
            If XLugar = WLugar - 1 Then
                XCounter = XCounter + 1
                WPasa(XCounter, 1) = "  -     -   "
                WPasa(XCounter, 2) = ""
                WPasa(XCounter, 3) = ""
                WPasa(XCounter, 4) = ""
            End If
    
        Next iRow
        If Salida = "S" Then
            Exit For
        End If
    Next a
     
    WLugar = 0
    Salida = "N"
            
    For a = 0 To 5
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            WLugar = WLugar + 1
            DBGrid1.Row = iRow
            DBGrid1.Col = 0
            DBGrid1.Text = WPasa(WLugar, 1)
            DBGrid1.Col = 1
            DBGrid1.Text = WPasa(WLugar, 2)
            DBGrid1.Col = 2
            DBGrid1.Text = WPasa(WLugar, 3)
            Rem DBGrid1.Col = 3
            Rem DBGrid1.Text = WPasa(WLugar, 4)
            If WLugar = XCounter Then
                Salida = "S"
                Exit For
            End If
        Next iRow
            
        If Salida = "S" Then
            Exit For
        End If
            
    Next a
    
    Renglon = Renglon + 1
    
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    DBGrid1.Col = 0
    
    DBGrid1.FirstRow = WPrimer
    DBGrid1.Row = WFila
    DBGrid1.Col = 0
    
    Call DBGrid1_GotFocus
    
End Sub

Private Sub DBGrid1_GotFocus()

    aa = Len(DBGrid1.Text)

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 12 Then
        WLinea.Text = DBGrid1.Row + 1
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -     -   "
        WLinea.Text = ""
    End If
    
    DBGrid1.Col = 1
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 2
    If Val(DBGrid1.Text) <> 0 Then
        WCantidad.Text = Pusing("###,###.##", DBGrid1.Text)
            Else
        WCantidad.Text = ""
    End If
    
    DBGrid1.Col = 3
    WPartida.Text = DBGrid1.Text
    
    Rem DBGrid1.Col = 3
    Rem WPrecio.Caption = Pusing("###,###.##", DBGrid1.Text)
    
    WTermi = WArticulo.Text
    
    WInicio = DBGrid1.FirstRow
    
    
    WArticulo.SetFocus

End Sub

Private Sub Graba_Click()

    Rem Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    
    Rem verifica que no halla dos articulos con
    Rem distinto tipo de condicion de pago
    
    XPasa = "S"
    WLLave = 0
    
    If Val(Wempresa) = 1 Then
    
    For a = 0 To 5
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Articulo = UCase(DBGrid1.Text)
                    
            DBGrid1.Col = 2
            Cantidad = DBGrid1.Text
            
            DBGrid1.Col = 3
            Partida = DBGrid1.Text
            
            If Val(Cantidad) <> 0 Then
            
            
                WCliente = UCase(Cliente.Text)
                WTerminado = UCase(Articulo)
                WClave = WCliente + WTerminado
    
                XCodigo = Val(Mid$(WTerminado, 4, 5))
                If Left$(WTerminado, 2) <> "PT" Then
                    Select Case Left$(WTerminado, 2)
                        Case "DY", "DS"
                            XTipoPro = "CO"
                        Case Else
                            XTipoPro = "PT"
                    End Select
                        Else
                    If XCodigo >= 0 And XCodigo <= 999 Then
                        XTipoPro = "CO"
                            Else
                        If XCodigo >= 11000 And XCodigo <= 12999 Then
                            XTipoPro = "CO"
                                Else
                            If XCodigo >= 25000 And XCodigo <= 25999 Then
                                XTipoPro = "FA"
                                    Else
                                If XCodigo >= 2300 And XCodigo <= 2399 Then
                                    XTipoPro = "BI"
                                        Else
                                    XTipoPro = "PT"
                                End If
                            End If
                        End If
                    End If
                End If
                
                ZLinea = 0
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZLinea = rstTerminado!Linea
                    rstTerminado.Close
                End If
                
                Select Case ZLinea
                    Case 8
                        XTipoPro = "PG"
                    Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                        XTipoPro = "FA"
                    Case Else
                End Select
                
                If WLLave = 0 Then
                    WLLave = 1
                    WTipopro = XTipoPro
                End If
                
                If WTipopro <> XTipoPro Then
                    XPasa = "2"
                End If
                        
            End If
                                        
        Next iRow
            
    Next a
    

    If XPasa = "2" Then
    
        m$ = "Se cargaron articulos PT, Biosidas, Farma, Pigmentos o Colorantes en forma conjunta un mismo Pedido"
        A1% = MsgBox(m$, 0, "INGRESO DE PEDIDOS")
    
        Renglon = Renglon + 1
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
        WAnterior = DBGrid1.Row
            
        DBGrid1.Col = 0
        DBGrid1.Text = ""
        Renglon = Renglon - 1

        Exit Sub
        
    End If
    
    End If

    spPedidoDevol = "BorrarPedidoDevol " + "'" + Pedido.Text + "'"
    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenDynaset, dbSQLPassThrough)

    Erase Auxiliar
    Renglon = 0
    WRenglon = 0
    DBGrid1.Refresh
        
    For a = 0 To 5
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Articulo = UCase(DBGrid1.Text)
                    
            DBGrid1.Col = 2
            Cantidad = DBGrid1.Text
            
            DBGrid1.Col = 3
            Partida = DBGrid1.Text
            
            XLote = Partida
                    
            If Left$(Articulo, 2) <> "PT" Then
                
                ZEntra = "N"
                                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Laudo"
                ZSql = ZSql + " Where Laudo.PartiOri = " + "'" + XLote + "'"
                ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                spLaudo = ZSql
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    With rstLaudo
                        .MoveFirst
                        Partida = IIf(IsNull(rstLaudo!Laudo), "", rstLaudo!Laudo)
                        ZEntra = "S"
                        rstLaudo.Close
                    End With
                End If
                        
                If ZEntra = "N" Then
            
                    ZZCodigo = Left$(Articulo, 3) + Mid$(Articulo, 6, 10)
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Guia"
                    ZSql = ZSql + " Where Guia.PartiOri = " + "'" + XLote + "'"
                    ZSql = ZSql + " and Guia.Articulo = " + "'" + ZZCodigo + "'"
                    ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                    spMovguia = ZSql
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        With rstMovguia
                            .MoveFirst
                            Partida = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                            rstMovguia.Close
                        End With
                    End If
                    
                End If
                   
            End If
            
            XPrecio = 0
            
            If Left$(Articulo, 2) <> "PT" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(Articulo, 3) + Right$(Articulo, 7)
                    WClaveMp = Cliente.Text + WArti
                    spPreciosMp = "ConsultaPreciosMp " + "'" + WClaveMp + "'"
                    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPreciosMp.RecordCount > 0 Then
                        XPrecio = Pusing("###,###.##", Str$(rstPreciosMp!Precio))
                        rstPreciosMp.Close
                    End If
                
                Case Else
                    WClave = Cliente.Text + Articulo
                    spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
                    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPrecios.RecordCount > 0 Then
                        XPrecio = Pusing("###,###.##", Str$(rstPrecios!Precio))
                        rstPrecios.Close
                    End If
            End Select
                    
            Rem DBGrid1.Col = 3
            Rem Precio = DBGrid1.Text
            
            WPrimer = DBGrid1.FirstRow
            WFila = DBGrid1.Row
            WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                                            
            WInicio = DBGrid1.FirstRow
                        
            If Val(Cantidad) <> 0 Then
            
                Renglon = Renglon + 1
                WRenglon = WRenglon + 1
                    
                Auxiliar(WRenglon, 1) = Articulo
                Auxiliar(WRenglon, 2) = Cantidad
                    
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Pedido.Text
                Call Ceros(Auxi1, 6)
                    
                WPedido = Pedido.Text
                WRenglon = Str$(Renglon)
                WCliente = Cliente.Text
                WFecha = Fecha.Text
                WObservaciones = Observaciones.Text
                WTerminado = Articulo
                WCantidad = Cantidad
                WPartida = Partida
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WPrecio = XPrecio
                WWLinea = Linea
                WFacturado = ""
                WImporte = ""
                WClave = Auxi1 + Auxi
                WAutorizo = "N"
                WImpresion = "N"
                If Left$(Articulo, 2) <> "PT" Then
                    WTipopro = "M"
                    WArti = Left$(Articulo, 3) + Right$(Articulo, 7)
                        Else
                    WTipopro = "T"
                    WArti = "  -   -   "
                End If
                
                XCodigo = Val(Mid$(WTerminado, 4, 5))
                If Left$(WTerminado, 2) <> "PT" Then
                    Select Case Left$(WTerminado, 2)
                        Case "DY", "DS"
                            XTipoPro = "CO"
                        Case Else
                            XTipoPro = "PT"
                    End Select
                        Else
                    If XCodigo >= 0 And XCodigo <= 999 Then
                        XTipoPro = "CO"
                            Else
                        If XCodigo >= 11000 And XCodigo <= 12999 Then
                            XTipoPro = "CO"
                                Else
                            If XCodigo >= 25000 And XCodigo <= 25999 Then
                                XTipoPro = "FA"
                                    Else
                                If XCodigo >= 2300 And XCodigo <= 2399 Then
                                    XTipoPro = "BI"
                                        Else
                                    XTipoPro = "PT"
                                End If
                            End If
                        End If
                    End If
                End If
                
                ZLinea = 0
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZLinea = rstTerminado!Linea
                    rstTerminado.Close
                End If
                
                Select Case ZLinea
                    Case 8
                        XTipoPro = "PG"
                    Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                        XTipoPro = "FA"
                    Case Else
                End Select
                
                WTipoPedido = XTipoPro
                
                XParam = "'" + WClave + "','" _
                         + WPedido + "','" _
                         + WRenglon + "','" _
                         + WCliente + "','" _
                         + WFecha + "','" _
                         + WObservaciones + "','" _
                         + WTerminado + "','" _
                         + WCantidad + "','" _
                         + WPrecio + "','" _
                         + WWLinea + "','" _
                         + WFacturado + "','" _
                         + WImporte + "','" _
                         + WAutorizo + "','" _
                         + WImpresion + "','" _
                         + WTipopro + "','" _
                         + WArti + "','" _
                         + WFechaord + "','" _
                         + WTipoPedido + "'"

                spPedidoDevol = "AltaPedidoDevolII " + XParam
                Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE PedidoDevol SET "
                ZSql = ZSql + " Partida =  " + "'" + WPartida + "',"
                ZSql = ZSql + " TipoProII =  " + "'" + XTipoPro + "'"
                ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                spPedidoDevol = ZSql
                Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                                        
        Next iRow
            
    Next a
    
    WImpreLaboI = "N"
    WImpreLaboII = "N"
    WImpreProdI = "N"
    WImpreProdII = "N"
    WImpreProdIII = "N"
    WImpreProdIV = "N"
    WImpresionII = "N"
    WBloqueo = ""
    WImpreTerminado = ""
                
    ZSql = ""
    ZSql = ZSql + "UPDATE PedidoDevol SET "
    ZSql = ZSql + " ImpreLaboI =  " + "'" + WImpreLaboI + "',"
    ZSql = ZSql + " ImpreLaboII =  " + "'" + WImpreLaboII + "',"
    ZSql = ZSql + " ImpreProdI =  " + "'" + WImpreProdI + "',"
    ZSql = ZSql + " ImpreProdII =  " + "'" + WImpreProdII + "',"
    ZSql = ZSql + " ImpreProdIII =  " + "'" + WImpreProdIII + "',"
    ZSql = ZSql + " ImpreProdIV =  " + "'" + WImpreProdIV + "',"
    ZSql = ZSql + " ImpresionII =  " + "'" + WImpresionII + "',"
    ZSql = ZSql + " Bloqueo =  " + "'" + WBloqueo + "',"
    ZSql = ZSql + " ImpreTerminado =  " + "'" + WImpreTerminado + "'"
    ZSql = ZSql + " Where Pedido = " + "'" + WPedido + "'"
    spPedidoDevol = ZSql
    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
    
    T$ = "Solicitud de Devolucion de Mercaderia"
    m$ = "Desea Imprimir la solicitud"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Call Impresion
    End If
        
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Pedido.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPartida.Text = ""
    Rem WPrecio.Caption = ""
    WArticulo.SetFocus
    
End Sub


Private Sub Limpia_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPartida.Text = ""
    Rem WPrecio.Caption = ""
    
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Observaciones.Text = ""
    
    Pantalla.Visible = False
    
    For a = 0 To 5
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 3
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    Pedido.Text = "1"
    spPedidoDevol = "ListaPedidoDevolNumero"
    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedidoDevol.RecordCount > 0 Then
        With rstPedidoDevol
            .MoveLast
            Pedido.Text = rstPedidoDevol!Pedido + 1
        End With
        rstPedidoDevol.Close
    End If
    
    DBGrid1.FirstRow = 0
    Renglon = 0

    Pedido.SetFocus

End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WArticulo.Text = UCase(WArticulo.Text)
        
        WCliente = Cliente.Text
        WTerminado = WArticulo.Text
        WArti = Left$(WTerminado, 3) + Right$(WTerminado, 7)
        WClave = Cliente.Text + WArticulo.Text
        WClaveMp = Cliente.Text + WArti
        
        If Left$(WArticulo.Text, 2) <> "PT" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                spPreciosMp = "ConsultaPreciosMp " + "'" + WClaveMp + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
                    WEntra = "S"
                    Rem WPrecio.Caption = Pusing("###,###.##", Str$(rstPreciosMp!Precio))
                    rstPreciosMp.Close
                    WCantidad.SetFocus
                        Else
                    WArticulo.SetFocus
                End If
                If WEntra = "S" Then
                    spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WDescripcion.Caption = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                End If
            
            Case Else
                spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    WEntra = "S"
                    WDescripcion.Caption = rstPrecios!Descripcion
                    Rem WPrecio.Caption = Pusing("###,###.##", Str$(rstPrecios!Precio))
                    rstPrecios.Close
                    WCantidad.SetFocus
                        Else
                    WArticulo.SetFocus
                End If
            
        End Select
        
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        WPartida.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WPartida_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Left$(WArticulo.Text, 2) <> "PT" Then
            ZTipopro = "M"
                Else
            ZTipopro = "T"
        End If
            
        Select Case ZTipopro
            Case "M"
                ZArti = Left$(WArticulo.Text, 3) + Right$(WArticulo.Text, 7)
                ZEntra = "N"
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Laudo"
                ZSql = ZSql + " Where Laudo.Articulo = " + "'" + ZArti + "'"
                ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WPartida.Text + "'"
                ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                spLaudo = ZSql
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    ZEntra = "S"
                    rstLaudo.Close
                End If
                        
                If ZEntra = "N" Then
                
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Guia"
                    ZSql = ZSql + " Where Guia.Articulo = " + "'" + ZArti + "'"
                    ZSql = ZSql + " and Guia.PartiOri = " + "'" + WPartida.Text + "'"
                    ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                    spMovguia = ZSql
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        ZEntra = "S"
                        rstMovguia.Close
                    End If
                    
                End If
                    
            Case Else
                ZEntra = "N"
                ZControla = 0
                spTerminado = "ConsultaTerminado " + "'" + WArticulo.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                    rstTerminado.Close
                End If
            
                If ZControla = 0 Then
                
                    WTerminado = WArticulo.Text
                    XCodigo = Val(Mid$(WTerminado, 4, 5))
                    If Left$(WTerminado, 2) <> "PT" Then
                        Select Case Left$(WTerminado, 2)
                            Case "DY", "DS"
                                WTipoPedido = "CO"
                            Case Else
                                WTipoPedido = "PT"
                        End Select
                            Else
                        If XCodigo >= 0 And XCodigo <= 999 Then
                            WTipoPedido = "CO"
                                Else
                            If XCodigo >= 11000 And XCodigo <= 12999 Then
                                WTipoPedido = "CO"
                                    Else
                                If XCodigo >= 25000 And XCodigo <= 25999 Then
                                    WTipoPedido = "FA"
                                        Else
                                    If XCodigo >= 2300 And XCodigo <= 2399 Then
                                        WTipoPedido = "BI"
                                            Else
                                        WTipoPedido = "PT"
                                    End If
                                End If
                            End If
                        End If
                    End If
                
                    ZLinea = 0
                    spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        ZLinea = rstTerminado!Linea
                        rstTerminado.Close
                    End If
                
                    Select Case ZLinea
                        Case 8
                            WTipoPedido = "PG"
                        Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                            WTipoPedido = "FA"
                        Case Else
                    End Select
                    
                    If Left$(WTerminado, 4) = "PT-4" Then
                        WTipoPedido = "TA"
                    End If
                
                    XEmpresa = Wempresa
                    If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                        Select Case WTipoPedido
                            Case "PG", "CO"
                                Wempresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case "TA"
                                Wempresa = "0003"
                                txtOdbc = "Empresa03"
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
                    End If
                
                    XParam = "'" + WPartida.Text + "','" _
                            + WArticulo.Text + "'"
                    spHoja = "ListaHojaProducto " + XParam
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHoja.RecordCount > 0 Then
                        ZEntra = "S"
                        rstHoja.Close
                    End If
                
                    If ZEntra = "N" Then
                        XParam = "'" + WArticulo.Text + "','" _
                                + WPartida.Text + "'"
                        spMovguia = "ListaMovguiaLote1 " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            ZEntra = "S"
                            rstMovguia.Close
                        End If
                    End If
                    
                    Call Conecta_Empresa
                
                        Else
                    
                    ZEntra = "S"
                
                End If
        End Select
        
        If Trim(WPartida.Text) = "" Then
            ZEntra = "N"
        End If
                
        If ZEntra = "N" Then
            m$ = WArticulo.Text + " Producto inexistente o Lote nro. " + WPartida.Text + " inexistente"
            G% = MsgBox(m$, 0, "Ingreso de Solicitud de Devolucion de Mercaderia")
                Else
            Call Alta_Vector
            Call Ingresa_Click
            WArticulo.SetFocus
        End If
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub pantalla_Click()
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spCliente = "ConsultaCliente " + "'" + Claveven$ + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = Claveven$
                DesCliente.Caption = rstCliente!Razon
                WDirentrega = rstCliente!DirEntrega
                Rem Observaciones.Text = RTrim(rstCliente!Observaciones)
                rstCliente.Close
            End If
            Pantalla.Visible = False
            Observaciones.SetFocus
            
        Case 1
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            
            If Left$(Claveven$, 2) <> "PT" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    Claveven$ = Left$(Claveven$, 9) + Right$(Claveven$, 7)
                    spPreciosMp = "ConsultaPreciosMp " + "'" + Claveven$ + "'"
                    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPreciosMp.RecordCount > 0 Then
                        WArti = rstPreciosMp!Articulo
                        WArticulo.Text = Left$(rstPreciosMp!Articulo, 3) + "00" + Right$(rstPreciosMp!Articulo, 7)
                        Rem WPrecio.Caption = rstPreciosMp!Precio
                    
                        DBGrid1.Col = 0
                        DBGrid1.Text = WArticulo.Text
                        Rem DBGrid1.Col = 3
                        Rem DBGrid1.Text = Pusing("###,###.##", rstPreciosMp!Precio)
                
                        rstPreciosMp.Close
                        
                        spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            WDescripcion.Caption = rstArticulo!Descripcion
                            rstArticulo.Close
                        End If
                        
                        DBGrid1.Col = 1
                        DBGrid1.Text = WDescripcion.Caption
                
                        Call Alta_Vector
                        WLinea.Text = WAnterior + 1
                        If Val(WLinea.Text) > 0 Then
                            DBGrid1.Row = Val(WLinea.Text) - 1
                        End If
                    
                        Call DBGrid1.SetFocus
                        WCantidad.SetFocus
                    
                    End If
            
                Case "T"
                    spPrecios = "ConsultaPrecios " + "'" + Claveven$ + "'"
                    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPrecios.RecordCount > 0 Then
                        WArticulo.Text = rstPrecios!Terminado
                        WDescripcion.Caption = rstPrecios!Descripcion
                        Rem WPrecio.Caption = rstPrecios!Precio
                    
                        DBGrid1.Col = 0
                        DBGrid1.Text = rstPrecios!Terminado
                        DBGrid1.Col = 1
                        DBGrid1.Text = rstPrecios!Descripcion
                        Rem DBGrid1.Col = 3
                        Rem DBGrid1.Text = Pusing("###,###.##", rstPrecios!Precio)
                
                        rstPrecios.Close
                        
                        Call Alta_Vector
                        WLinea.Text = WAnterior + 1
                        If Val(WLinea.Text) > 0 Then
                            DBGrid1.Row = Val(WLinea.Text) - 1
                        End If
                    
                        Call DBGrid1.SetFocus
                        WCantidad.SetFocus
                    
                    End If
            
                Case Else
            End Select
            
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 40 Then
                            DBGrid1.Row = DBGrid1.Row + 1
                            DBGrid1.Col = 0
                            KeyCode = 0
                        End If
                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
    End Select

    
End Sub


' Cuando el usuario hace clic en el icono Agregar, esta subrutina agrega una
' nueva fila a la variable RowBuf y un marcador a la variable NewRowBookmark
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows = mTotalRows + 1
ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
NewRowBookmark = mTotalRows - 1 'Establece el marcador a la última fila.

' El bucle siguiente agrega un nuevo registro a la base de datos.
For iCol = 0 To UBound(UserData, 1)
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
    Else
        ' Si no se establece ningún valor para la columna, usa DefaultValue
        UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
    End If
Next iCol

End Sub

' Esta subrutina elimina una fila basándose en su marcador.
Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

' Mueve todas las filas encima de la fila eliminada de
' la matriz.

For iRow = Bookmark + 1 To mTotalRows - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows = mTotalRows - 1

End Sub

' Se llama a esta subrutina cada vez que DBGrid quiere mostrar
' datos nuevos.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid está solicitando filas, así que se las damos

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' Si StartLocation es Null, empieza a leer por el final
' o por el principio del conjunto de datos.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Busca la posición para empezar a leer, basándose en el marcador
    ' StartLocation y en la variable iIncr
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfiere datos de nuestra matriz de conjunto de datos al objeto RowBuf
' que DBGrid utiliza para presentar los datos
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
    For iCol = 0 To UBound(UserData, 1)
        RowBuf.Value(iRow, iCol) = UserData(iCol, CurRow&)
    Next iCol
    ' Establece el marcador mediante CurRow&, que es también
    ' nuestro índice de matriz
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

' Esta subrutina actualiza los datos de la matriz después de
' haberse modificado.

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Se están actualizando los datos

' Actualiza cada columna de la matriz de conjuntos de datos
For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub


Private Sub Form_Load()

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 3, 0 To 80)

mTotalRows& = 80

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DBGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DBGrid1.Columns.Count - 1 To 0 Step -1
      DBGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 3
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 4000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Partida"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 

    Pantalla.Visible = False
    Pedido.Text = "1"
    spPedidoDevol = "ListaPedidoDevolNumero"
    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedidoDevol.RecordCount > 0 Then
        With rstPedidoDevol
            .MoveLast
            Pedido.Text = rstPedidoDevol!Pedido + 1
        End With
        rstPedidoDevol.Close
    End If
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Pedido.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For a = 0 To 5
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 3
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    Erase Auxiliar
    WRenglon = 0

    spPedidoDevol = "ListaPedidoDevol " + "'" + Pedido.Text + "'"
    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedidoDevol.RecordCount > 0 Then
            WGraba = "N"
            With rstPedidoDevol
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        Renglon = Renglon + 1
                
                        Lugar1 = Int((Renglon - 1) / 10) * 10
                        Lugar2 = Renglon - Lugar1
                
                        DBGrid1.FirstRow = Lugar1
                        DBGrid1.Row = Lugar2 - 1
                
                        DBGrid1.Col = 0
                        DBGrid1.Text = rstPedidoDevol!Terminado
                        Auxi1 = rstPedidoDevol!Terminado
                
                        DBGrid1.Col = 2
                        DBGrid1.Text = Pusing("###,###.##", rstPedidoDevol!Cantidad - rstPedidoDevol!Facturado)
                        
                        DBGrid1.Col = 3
                        DBGrid1.Text = rstPedidoDevol!Partida
                        
                        WRenglon = WRenglon + 1
                    
                        Auxiliar(WRenglon, 1) = rstPedidoDevol!Cliente
                        Auxiliar(WRenglon, 2) = rstPedidoDevol!Terminado
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedidoDevol.Close
    End If
    
    Renglon = 0
    
    For DA = 1 To WRenglon
    
        Cliente = Auxiliar(DA, 1)
        Terminado = Auxiliar(DA, 2)
        If Left$(Terminado, 2) <> "PT" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Terminado, 3) + Right$(Terminado, 7)
                spPreciosMp = "ConsultaPreciosMp " + "'" + Cliente + WArti + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
        
                    Renglon = Renglon + 1
                
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    Rem DBGrid1.Col = 3
                    Rem DBGrid1.Text = Pusing("###,###.##", rstPreciosMp!Precio)
                    
                    rstPreciosMp.Close
                    
                    spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        DBGrid1.Col = 1
                        DBGrid1.Text = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                    
                    WArticulo.SetFocus
                    
                End If
            
            Case Else
                spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
        
                    Renglon = Renglon + 1
                
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstPrecios!Descripcion
            
                    Rem DBGrid1.Col = 3
                    Rem DBGrid1.Text = Pusing("###,###.##", rstPrecios!Precio)
                    
                    rstPrecios.Close
                
                    WArticulo.SetFocus
                    
                End If
                
        End Select
        
    Next DA

    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    Call DBGrid1_GotFocus
   
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    WArticulo.SetFocus
    Call DBGrid1_GotFocus

End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) = 0 Then

            Renglon = Renglon + 1
            
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
            
            DBGrid1.Col = 3
            DBGrid1.Text = WPartida.Text
            
            Rem DBGrid1.Col = 3
            Rem DBGrid1.Text = Pusing("###,###.##", WPrecio.Caption)
            
            Rem DbGrid1.Row = Renglon
            DBGrid1.Row = Lugar2 - 1
            DBGrid1.Col = 0
    
            
            DBGrid1.Col = 0
            DBGrid1.Text = WArticulo.Text
            
            If Renglon = 1 Then
                 DBGrid1.Row = DBGrid1.Row + 1
                 DBGrid1.Col = 0
                 DBGrid1.Text = ""
            End If
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
            
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
            
            DBGrid1.Col = 3
            DBGrid1.Text = WPartida.Text
            
            Rem DBGrid1.Col = 3
            Rem DBGrid1.Text = Pusing("###,###.##", WPrecio.Caption)
                
            Rem DbGrid1.Row = Renglon
            DBGrid1.Row = Lugar2 - 1
            DBGrid1.Col = 0
            WInicio = DBGrid1.FirstRow
            
            DBGrid1.Row = DBGrid1.Row + 1
            
    End If

End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)

    On Error GoTo WError

    If KeyAscii = 13 Then
    
        spPedidoDevol = "ListaPedidoDevol " + "'" + Pedido.Text + "'"
        Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedidoDevol.RecordCount > 0 Then
            Fecha.Text = rstPedidoDevol!Fecha
            Cliente.Text = rstPedidoDevol!Cliente
            Observaciones.Text = RTrim(rstPedidoDevol!Observaciones)
            rstPedidoDevol.Close
            
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WDirentrega = rstCliente!DirEntrega
                Rem Observaciones.Text = rstCliente!Observaciones
                rstCliente.Close
            End If
            Call Proceso_Click
                Else
            WPedido = Pedido.Text
            Call Limpia_Click
            Pedido.Text = WPedido
            Fecha.SetFocus
        End If
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Cliente.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        If Cliente.Text <> "" Then
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WDirentrega = rstCliente!DirEntrega
                Rem Observaciones.Text = RTrim(rstCliente!Observaciones)
                rstCliente.Close
                Observaciones.SetFocus
                    Else
                Cliente.Text = Claveven$
                Cliente.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WArticulo.SetFocus
    End If
End Sub

Private Sub Impresion()

    On Error GoTo WError
        
    spImprePed = "Delete ImprePed"
    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    WObservaciones = Left$(RTrim(Observaciones.Text) + Space$(100), 100)
        
    XLinea = 0
    WCounter = 0
    WRenglon = 0
                    
    For a = 0 To 5
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WCounter = WCounter + 1
        
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
                    
            If DBGrid1.Text <> "" Then
                    
                WArticulo = DBGrid1.Text
                    
                DBGrid1.Col = 1
                WDescripcion = DBGrid1.Text
                    
                DBGrid1.Col = 2
                WCantidad = Val(DBGrid1.Text)
                
                DBGrid1.Col = 3
                WPartida = DBGrid1.Text
                    
                Rem DBGrid1.Col = 3
                Rem WPrecio = Val(DBGrid1.Text)
                
                If WCantidad <> 0 Then
                
                    WRenglon = WRenglon + 1
                    
                    Auxi = Pedido.Text
                    Call Ceros(Auxi, 6)
                    Auxi1 = WRenglon
                    Call Ceros(Auxi1, 2)
                    ZClave = "1" + Auxi + Auxi1
                    ZTipo = "1"
                    ZPedido = Pedido.Text
                    ZRenglon = Str$(WRenglon)
                    ZEmpresa = WNombreEmpresa
                    ZVersion = ""
                    ZCliente = Cliente.Text
                    ZNombre = DesCliente.Caption
                    ZFecha = Fecha.Text
                    ZFechaent = "  /  /    "
                    ZTipoPedido = WPartida
                    ZCondicion = ""
                    ZEntrega = WDirentrega
                    ZObservaciones1 = Left$(WObservaciones, 50)
                    ZObservaciones2 = Right$(WObservaciones, 50)
                    ZOrden = ""
                    ZArticulo = WArticulo
                    ZDescripcion = WDescripcion
                    ZPrecio = Str$(WPrecio)
                    ZCantidad = Str$(WCantidad)
                    ZEnvase = ""
                    
                    spImprePed = "INSERT INTO ImprePed (" + _
                                "Clave ," + _
                                "Tipo , Pedido ," + _
                                "Renglon , Empresa ," + _
                                "Version , Cliente ," + _
                                "Nombre , Fecha ," + _
                                "Fechaent , TipoPedido ," + _
                                "Condicion , Entrega ," + _
                                "Observaciones1 , Observaciones2 ," + _
                                "Orden , Articulo ," + _
                                "Descripcion , Precio ," + _
                                "Cantidad , Envase )" + _
                                "Values (" + _
                                "'" + ZClave + "'," + _
                                "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                                "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                                "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                                "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                                "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                                "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                                "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                                "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                                "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                                "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                                
                    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                    
            End If
                                        
        Next iRow
            
    Next a
    
    For Ciclo = WRenglon + 1 To 12
    
        WRenglon = WRenglon + 1
                    
        Auxi = Pedido.Text
        Call Ceros(Auxi, 6)
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        ZClave = "1" + Auxi + Auxi1
        ZTipo = "1"
        ZPedido = Pedido.Text
        ZRenglon = Str$(WRenglon)
        ZEmpresa = WNombreEmpresa
        ZVersion = ""
        ZCliente = Cliente.Text
        ZNombre = DesCliente.Caption
        ZFecha = Fecha.Text
        ZFechaent = "  /  /    "
        ZTipoPedido = ""
        ZCondicion = ""
        ZEntrega = WDirentrega
        ZObservaciones1 = Left$(WObservaciones, 50)
        ZObservaciones2 = Right$(WObservaciones, 50)
        ZOrden = ""
        ZArticulo = ""
        ZDescripcion = ""
        ZPrecio = ""
        ZCantidad = ""
        ZEnvase = ""
                    
        spImprePed = "INSERT INTO ImprePed (" + _
                    "Clave ," + _
                    "Tipo , Pedido ," + _
                    "Renglon , Empresa ," + _
                    "Version , Cliente ," + _
                    "Nombre , Fecha ," + _
                    "Fechaent , TipoPedido ," + _
                    "Condicion , Entrega ," + _
                    "Observaciones1 , Observaciones2 ," + _
                    "Orden , Articulo ," + _
                    "Descripcion , Precio ," + _
                    "Cantidad , Envase )" + _
                    "Values (" + _
                    "'" + ZClave + "'," + _
                    "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                    "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                    "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                    "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                    "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                    "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                    "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                    "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                    "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                    "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                                
        Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT ImprePed.Pedido, ImprePed.Empresa, ImprePed.Cliente, ImprePed.Nombre, ImprePed.Fecha, ImprePed.TipoPedido, ImprePed.Observaciones1, ImprePed.Observaciones2, ImprePed.Articulo, ImprePed.Descripcion, ImprePed.Cantidad " _
                    + "From " _
                    + DSQ + ".dbo.ImprePed ImprePed " _
                    + "Where " _
                    + "ImprePed.Pedido >= 0 AND ImprePed.Pedido <= 999999"
                        
    Listado.Connect = Connect()
    Listado.ReportFileName = "ImprepedidevolSQL.rpt"
    Listado.Destination = 1
    Listado.CopiesToPrinter = 1
    Listado.Action = 1
    
    Exit Sub
        
WError:
    Resume Next

End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
            
                    DA = Len(rstCliente!Razon) - WEspacios
                
                    For aa = 1 To DA
                        If Left$(Ayuda.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                            Auxi = rstCliente!Cliente
                            IngresaItem = Auxi + "    " + rstCliente!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
                            WIndice.AddItem IngresaItem
                            Exit For
                        End If
                    Next aa
                    .MoveNext
                    
                        Else
                        
                    Exit Do
                
                End If
            Loop
        End With
        rstCliente.Close
    End If
    End If

End Sub

Private Sub WPartida_DblClick()
    WPasaCliente = Cliente.Text
    WPasaTerminado = WArticulo.Text
    WPartida.SetFocus
    PrgConsultaPartida.Show
End Sub




