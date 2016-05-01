VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsig 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Remitos de Mercaderia en Consignacion"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   240
      TabIndex        =   55
      Top             =   6600
      Visible         =   0   'False
      Width           =   6255
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
      Left            =   240
      TabIndex        =   54
      Text            =   " "
      Top             =   6240
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.CommandButton BorraConsulta 
      Caption         =   "Borra Consulta"
      Height          =   495
      Left            =   10440
      TabIndex        =   53
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton ConsultaPro 
      Caption         =   "Consulta Producto"
      Height          =   495
      Left            =   10440
      TabIndex        =   52
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton ConsultaCli 
      Caption         =   "Consulta Cliente"
      Height          =   495
      Left            =   10440
      TabIndex        =   51
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Baja 
      Caption         =   "  Baja  Consig."
      Height          =   495
      Left            =   8280
      TabIndex        =   50
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame IngreEnvases 
      Caption         =   "Ingreso de Envases"
      Height          =   2055
      Left            =   6600
      TabIndex        =   23
      Top             =   6480
      Width           =   4815
      Begin VB.TextBox Canti3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   31
         Text            =   " "
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Canti2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   30
         Text            =   " "
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Canti1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   29
         Text            =   " "
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Envase3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   28
         Text            =   " "
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Envase2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   27
         Text            =   " "
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Envase1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   26
         Text            =   " "
         Top             =   600
         Width           =   855
      End
      Begin VB.Label WAbre6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3000
         TabIndex        =   49
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label WAbre5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3000
         TabIndex        =   48
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label WAbre4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3000
         TabIndex        =   47
         Top             =   960
         Width           =   615
      End
      Begin VB.Label WAbre3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3000
         TabIndex        =   46
         Top             =   720
         Width           =   615
      End
      Begin VB.Label WAbre2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3000
         TabIndex        =   45
         Top             =   480
         Width           =   615
      End
      Begin VB.Label WAbre1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3000
         TabIndex        =   44
         Top             =   240
         Width           =   615
      End
      Begin VB.Label WCapa6 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3720
         TabIndex        =   43
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label WCapa5 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3720
         TabIndex        =   42
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label WCapa4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3720
         TabIndex        =   41
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Wcapa3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   720
         Width           =   855
      End
      Begin VB.Label WCapa2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3720
         TabIndex        =   39
         Top             =   480
         Width           =   855
      End
      Begin VB.Label WCapa1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3720
         TabIndex        =   38
         Top             =   240
         Width           =   855
      End
      Begin VB.Label WEnvase6 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   2280
         TabIndex        =   37
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label WEnvase5 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   2280
         TabIndex        =   36
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label WEnvase4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   2280
         TabIndex        =   35
         Top             =   960
         Width           =   615
      End
      Begin VB.Label WEnvase3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   2280
         TabIndex        =   34
         Top             =   720
         Width           =   615
      End
      Begin VB.Label WEnvase2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   480
         Width           =   615
      End
      Begin VB.Label WEnvase1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Envase 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Envase"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   500
      Left            =   9360
      TabIndex        =   22
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
      Left            =   3360
      TabIndex        =   21
      Top             =   6840
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   19
      Text            =   " "
      Top             =   840
      Width           =   5895
   End
   Begin VB.TextBox Cliente 
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   16
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   14
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Numero 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   500
      Left            =   8280
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      Height          =   500
      Left            =   9360
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   500
      Left            =   8280
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   5640
      Width           =   10095
      Begin VB.TextBox WLote 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8520
         MaxLength       =   6
         TabIndex        =   56
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Text            =   " "
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox WCantidad 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   6120
         MaxLength       =   10
         TabIndex        =   5
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label WPrecio 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   7320
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label WDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   500
      Left            =   9360
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3975
      Left            =   240
      OleObjectBlob   =   "Consig.frx":0000
      TabIndex        =   2
      Top             =   1560
      Width           =   10215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6960
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Observaciones"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3000
      TabIndex        =   17
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Remito"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgConsig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 5 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Auxi As String
Private WImpre(10) As String
Private WEnvase(10) As String
Private WVector(6, 3) As String
Private XEnvase(40, 6) As String
Private XLinea As Single
Private WInicio As Integer
Private Auxiliar(100, 3) As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstConsig As Recordset
Dim spConsig As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstMovenv As Recordset
Dim spMovenv As String
Dim XParam As String
Private WCodIva As String
Private WProvincia As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WProv As String
Private WPostal As String
Private WCuit As String
Private WDirentrega As String
Private Iva(0 To 30) As String
Private Provincia(0 To 30) As String
Private Stk(19, 4) As String
Private XSaldo1 As String

Private Sub Baja_Click()
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Erase Auxiliar
    WRenglon = 0

    spConsig = "ListaConsig " + "'" + Numero.Text + "'"
    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)

    If rstConsig.RecordCount > 0 Then
            With rstConsig
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    
                    Auxiliar(WRenglon, 1) = rstConsig!Terminado
                    Auxiliar(WRenglon, 2) = rstConsig!Cantidad
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstConsig.Close
    End If
    
    For DA = 1 To WRenglon
    
        Terminado = Auxiliar(DA, 1)
        Cantidad = Auxiliar(DA, 2)
        
        spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WCodigo = Articulo
            WPedido = Str$(rstTerminado!Pedido)
            WSalidas = Str$(rstTerminado!Salidas - Cantidad)
            WDate = Date$
            WLinea = rstTerminado!Linea
            rstTerminado.Close
                
            XParam = "'" + WCodigo + "','" _
                            + WPedido + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
            spTerminado = "ModificaTerminadoFacturas " + XParam
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Next DA
    
    spConsig = "BorrarConsig " + "'" + Numero.Text + "'"
    Set rstConsig = db.OpenRecordset(spConsig, dbOpenDynaset, dbSQLPassThrough)
    
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Numero.SetFocus

End Sub

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    DBGrid1.Col = 3
    DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WLote.Text = ""
    WLinea.Text = ""
    
    WEnvase1.Caption = ""
    WEnvase2.Caption = ""
    WEnvase3.Caption = ""
    WEnvase4.Caption = ""
    WEnvase5.Caption = ""
    WEnvase6.Caption = ""
    
    WAbre1.Caption = ""
    WAbre2.Caption = ""
    WAbre3.Caption = ""
    WAbre4.Caption = ""
    WAbre5.Caption = ""
    WAbre6.Caption = ""
    
    WCapa1.Caption = ""
    WCapa2.Caption = ""
    Wcapa3.Caption = ""
    WCapa4.Caption = ""
    WCapa5.Caption = ""
    WCapa6.Caption = ""
    
    Envase1.Text = ""
    Envase2.Text = ""
    Envase3.Text = ""
    
    Canti1.Text = ""
    Canti2.Text = ""
    Canti3.Text = ""

    WArticulo.SetFocus
    
End Sub

Private Sub BorraConsulta_Click()
    Pantalla.Visible = False
    Opcion.Visible = False

End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstEmpresa
        .Close
    End With
    PrgConsig.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Clientes"
     Opcion.AddItem "Productos"

     Opcion.Visible = True
     
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
        rstClientes.Close
    End With
    
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
    
    spTerminado = "ListaTerminado "
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
        With rstTerminado
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = rstTerminado!Codigo + "   " + rstTerminado!Descripcion
                    Pantalla.AddItem IngresaItem
                    IngresaItem = rstTerminado!Codigo
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTerminado.Close
    End If
    
    Pantalla.Visible = True
    
    Pantalla.Height = 1740
    Pantalla.Left = 3600
    Pantalla.Top = 6360
    Pantalla.Width = 8055

End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

 Private Sub Opcion_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Rem XIndice = 0
    
    Select Case XIndice
        Case 0
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
            
        Case 1
            spTerminado = "ListaTerminado"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstTerminado!Codigo + "   " + rstTerminado!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstTerminado!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTerminado.Close
            End If
        Case Else
    End Select
    Pantalla.Visible = True

End Sub

Private Sub DBGrid1_GotFocus()

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
    WPrecio.Caption = Pusing("###,###.##", DBGrid1.Text)
    
    DBGrid1.Col = 4
    WLote.Text = DBGrid1.Text
    
    Erase WVector
    
    spTerminado = "ConsultaTerminado " + "'" + WArticulo.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WVector(1, 1) = rstTerminado!Envase1
        WVector(2, 1) = rstTerminado!Envase2
        WVector(3, 1) = rstTerminado!Envase3
        WVector(4, 1) = rstTerminado!Envase4
        WVector(5, 1) = rstTerminado!Envase5
        WVector(6, 1) = rstTerminado!Envase6
        rstTerminado.Close
        Call Carga_Envases
    End If
    
    For WDa = 1 To 6
        If Val(WVector(WDa, 1)) <> 0 Then
            spEnvase = "ConsultaEnvases " + "'" + WVector(WDa, 1) + "'"
            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvase.RecordCount > 0 Then
                WVector(WDa, 3) = rstEnvase!Abreviatura
            End If
            rstEnvase.Close
        End If
    Next WDa
    
    WInicio = DBGrid1.FirstRow
    
    If Val(WLinea.Text) <> 0 Then
        Envase1.Text = XEnvase(Val(WLinea.Text) + WInicio, 1)
        Canti1.Text = XEnvase(Val(WLinea.Text) + WInicio, 2)
        Envase2.Text = XEnvase(Val(WLinea.Text) + WInicio, 3)
        Canti2.Text = XEnvase(Val(WLinea.Text) + WInicio, 4)
        Envase3.Text = XEnvase(Val(WLinea.Text) + WInicio, 5)
        Canti3.Text = XEnvase(Val(WLinea.Text) + WInicio, 6)
    End If
    
    WArticulo.SetFocus

End Sub

Private Sub Graba_Click()

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""

    Erase Auxiliar
    WRenglon = 0

    spConsig = "ListaConsig " + "'" + Numero.Text + "'"
    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)

    If rstConsig.RecordCount > 0 Then
    
            With rstConsig
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    
                    Auxiliar(WRenglon, 1) = rstConsig!Terminado
                    Auxiliar(WRenglon, 2) = rstConsig!Cantidad
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstConsig.Close
    End If
    
    For DA = 1 To WRenglon
    
        Terminado = Auxiliar(DA, 1)
        Cantidad = Auxiliar(DA, 2)
        
        spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WCodigo = Articulo
            WPedido = Str$(rstTerminado!Pedido)
            WSalidas = Str$(rstTerminado!Salidas - Cantidad)
            WDate = Date$
            WLinea = rstTerminado!Linea
            rstTerminado.Close
                
            XParam = "'" + WCodigo + "','" _
                            + WPedido + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
            spTerminado = "ModificaTerminadoFacturas " + XParam
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Next DA
    
    spConsig = "BorrarConsig " + "'" + Numero.Text + "'"
    Set rstConsig = db.OpenRecordset(spConsig, dbOpenDynaset, dbSQLPassThrough)
    
    Erase Auxiliar
    Renglon = 0
    WRenglon = 0
    DBGrid1.Refresh
        
    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Terminado = UCase(DBGrid1.Text)
                    
            DBGrid1.Col = 2
            Cantidad = DBGrid1.Text
                    
            DBGrid1.Col = 3
            Precio = DBGrid1.Text
            
            DBGrid1.Col = 4
            Lote = DBGrid1.Text
                                            
            WInicio = DBGrid1.FirstRow
            Renglon = Renglon + 1
                        
            If Val(Cantidad) <> 0 Then
            
                WRenglon = WRenglon + 1
                    
                Auxiliar(WRenglon, 1) = Terminado
                Auxiliar(WRenglon, 2) = Cantidad
                Auxiliar(WRenglon, 3) = Lote
                    
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Numero)
                Call Ceros(Auxi1, 6)
                    
                WNumero = Numero.Text
                WRenglon = Str$(Renglon)
                WCliente = Cliente.Text
                WFecha = Fecha.Text
                WObservaciones = Observaciones.Text
                WTerminado = Terminado
                WCantidad = Cantidad
                WEnvase1 = XEnvase(Renglon, 1)
                WCanti1 = XEnvase(Renglon, 2)
                WEnvase2 = XEnvase(Renglon, 3)
                WCanti2 = XEnvase(Renglon, 4)
                WEnvase3 = XEnvase(Renglon, 5)
                WCanti3 = XEnvase(Renglon, 6)
                WEnvase4 = 0
                WCanti4 = ""
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WPrecio = Precio
                WLinea = Linea
                WFacturado = ""
                WImporte = ""
                WMarca = ""
                WLote = Lote
                WClave = Auxi1 + Auxi
                
                XParam = "'" + WClave + "','" _
                         + WNumero + "','" + WRenglon + "','" _
                         + WCliente + "','" + WFecha + "','" _
                         + WObservaciones + "','" _
                         + WTerminado + "','" _
                         + WCantidad + "','" _
                         + WEnvase1 + "','" _
                         + WCanti1 + "','" _
                         + WEnvase2 + "','" _
                         + WCanti2 + "','" _
                         + WEnvase3 + "','" _
                         + WCanti3 + "','" _
                         + WEnvase4 + "','" _
                         + WCanti4 + "','" _
                         + WFechaord + "','" _
                         + WPrecio + "','" _
                         + WLinea + "','" _
                         + WFacturado + "','" _
                         + WImporte + "','" _
                         + WMarca + "','" _
                         + WLote + "'"
                         
                spConsig = "AltaConsig " + XParam
                Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
                        
            End If
                                        
        Next iRow
            
    Next a
    
    For DA = 1 To WRenglon
    
        Terminado = Auxiliar(DA, 1)
        Cantidad = Val(Auxiliar(DA, 2))
        Lote = Val(Auxiliar(DA, 3))
        
        spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WCodigo = Terminado
            WPedido = Str$(rstTerminado!Pedido)
            WSalidas = Str$(rstTerminado!Salidas + Cantidad)
            WDate = Date$
            WLinea = rstTerminado!Linea
            rstTerminado.Close
                
            XParam = "'" + WCodigo + "','" _
                            + WPedido + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
            spTerminado = "ModificaTerminadoFacturas " + XParam
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            If WControla = 0 Then
                    XParam = "'" + Lote + "','" _
                                + Terminado + "'"
                    spHoja = "ListaHojaProducto " + XParam
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHoja.RecordCount > 0 Then
                        WClave = rstHoja!Clave
                        WSaldo = Str$(rstHoja!Saldo - Cantidad)
                        WDate = Date$
                        rstHoja.Close
                            
                        XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                        spHoja = "ModificaHojaSaldo " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            
                                    Else
                                
                        XParam = "'" + Terminado + "','" _
                                        + Lote + "'"
                        spMovguia = "ListaMovguiaLote1 " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            WClave = rstMovguia!Clave
                            WSaldo = Str$(rstMovguia!Saldo - Cantidad)
                            WDate = Date$
                            rstMovguia.Close
                            
                            XParam = "'" + WClave + "','" _
                                        + WDate + "','" _
                                        + WSaldo + "'"
                            spMovguia = "ModificaMovguiaSaldo " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                            
                    End If
            End If
            
        End If
        
    Next DA
    
    Renglon = 0
        
    For DA = 1 To 40
    
        XEnv1 = XEnvase(DA, 1)
        XCanti1 = XEnvase(DA, 2)
        XEnv2 = XEnvase(DA, 3)
        XCanti2 = XEnvase(DA, 4)
        XEnv3 = XEnvase(DA, 5)
        XCanti3 = XEnvase(DA, 6)
        
        If Val(XEnv1) <> 0 Then
            
                Renglon = Renglon + 1
                    
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Val(Numero.Text))
                Call Ceros(Auxi1, 6)
                    
                WTipo = "9"
                WCodigo = Numero.Text
                WRenglon = Str$(Renglon)
                WCliente = Cliente.Text
                WFecha = Fecha.Text
                WEnv = XEnv1
                WCantidad = XCanti1
                WMovimiento = "S"
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WClave = Auxi1 + Auxi
                
                XParam = "'" + WClave + "','" _
                        + WTipo + "','" _
                        + WCodigo + "','" _
                        + WRenglon + "','" _
                        + WFecha + "','" _
                        + WFechaord + "','" _
                        + WCliente + "','" _
                        + WEnv + "','" _
                        + WMovimiento + "','" _
                        + WCantidad + "'"
                    
                spMovenv = "AltaMovenv " + XParam
                Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
                
        End If
        
        If Val(XEnv2) <> 0 Then
            
                Renglon = Renglon + 1
                    
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Val(Numero.Text))
                Call Ceros(Auxi1, 6)
                    
                WTipo = "9"
                WCodigo = Numero.Text
                WRenglon = Str$(Renglon)
                WCliente = Cliente.Text
                WFecha = Fecha.Text
                WEnv = XEnv2
                WCantidad = XCanti2
                WMovimiento = "S"
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WClave = Auxi1 + Auxi
                
                XParam = "'" + WClave + "','" _
                        + WTipo + "','" _
                        + WCodigo + "','" _
                        + WRenglon + "','" _
                        + WFecha + "','" _
                        + WFechaord + "','" _
                        + WCliente + "','" _
                        + WEnv + "','" _
                        + WMovimiento + "','" _
                        + WCantidad + "'"
                    
                spMovenv = "AltaMovenv " + XParam
                Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
                
        End If
        
        If Val(XEnv3) <> 0 Then
            
                Renglon = Renglon + 1
                    
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Val(Numero.Text))
                Call Ceros(Auxi1, 6)
                    
                WTipo = "9"
                WCodigo = Numero.Text
                WRenglon = Str$(Renglon)
                WCliente = Cliente.Text
                WFecha = Fecha.Text
                WEnv = XEnv3
                WCantidad = XCanti3
                WMovimiento = "S"
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WClave = Auxi1 + Auxi
                
                XParam = "'" + WClave + "','" _
                        + WTipo + "','" _
                        + WCodigo + "','" _
                        + WRenglon + "','" _
                        + WFecha + "','" _
                        + WFechaord + "','" _
                        + WCliente + "','" _
                        + WEnv + "','" _
                        + WMovimiento + "','" _
                        + WCantidad + "'"
                    
                spMovenv = "AltaMovenv " + XParam
                Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
                
        End If

    Next DA
        
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    T$ = "Remitos en Consignacion"
    m$ = "Desea Imprimir el Remito"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Call Impresion
    End If
        
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Numero.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WLote.Text = ""
    
    WEnvase1.Caption = ""
    WEnvase2.Caption = ""
    WEnvase3.Caption = ""
    WEnvase4.Caption = ""
    WEnvase5.Caption = ""
    WEnvase6.Caption = ""
    
    WAbre1.Caption = ""
    WAbre2.Caption = ""
    WAbre3.Caption = ""
    WAbre4.Caption = ""
    WAbre5.Caption = ""
    WAbre6.Caption = ""
    
    WCapa1.Caption = ""
    WCapa2.Caption = ""
    Wcapa3.Caption = ""
    WCapa4.Caption = ""
    WCapa5.Caption = ""
    WCapa6.Caption = ""
    Envase1.Text = ""
    Envase2.Text = ""
    Envase3.Text = ""
    
    Canti1.Text = ""
    Canti2.Text = ""
    Canti3.Text = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub Limpia_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WLote.Text = ""
    
    WEnvase1.Caption = ""
    WEnvase2.Caption = ""
    WEnvase3.Caption = ""
    WEnvase4.Caption = ""
    WEnvase5.Caption = ""
    WEnvase6.Caption = ""
    
    WAbre1.Caption = ""
    WAbre2.Caption = ""
    WAbre3.Caption = ""
    WAbre4.Caption = ""
    WAbre5.Caption = ""
    WAbre6.Caption = ""
    
    WCapa1.Caption = ""
    WCapa2.Caption = ""
    Wcapa3.Caption = ""
    WCapa4.Caption = ""
    WCapa5.Caption = ""
    WCapa6.Caption = ""
    
    Envase1.Text = ""
    Envase2.Text = ""
    Envase3.Text = ""
    
    Canti1.Text = ""
    Canti2.Text = ""
    Canti3.Text = ""
    
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Observaciones.Text = ""
    
    Pantalla.Visible = False
    
    Erase XEnvase
    
    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 4
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    Numero.Text = ""
    spConsig = "ListaConsigNumero"
    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
    If rstConsig.RecordCount > 0 Then
        With rstConsig
            .MoveLast
            Numero.Text = rstConsig!Numero + 1
        End With
        rstConsig.Close
    End If
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Graba.Enabled = True


    Numero.SetFocus

End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WArticulo.Text = UCase(WArticulo.Text)
        
        Erase WVector
        spTerminado = "ConsultaTerminado " + "'" + WArticulo.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WDescripcion.Caption = rstTerminado!Descripcion
            WVector(1, 1) = rstTerminado!Envase1
            WVector(2, 1) = rstTerminado!Envase2
            WVector(3, 1) = rstTerminado!Envase3
            WVector(4, 1) = rstTerminado!Envase4
            WVector(5, 1) = rstTerminado!Envase5
            WVector(6, 1) = rstTerminado!Envase6
            rstTerminado.Close
            Call Carga_Envases
        End If
        
        WCliente = Cliente.Text
        WTerminado = WArticulo.Text
        WClave = Cliente.Text + WArticulo.Text
    
        spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
                WDescripcion.Caption = rstPrecios!Descripcion
                WPrecio.Caption = Pusing("###,###.##", Str$(rstPrecios!Precio))
                rstPrecios.Close
        End If
        
        WCantidad.SetFocus
        
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        WLote.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WLote_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WEntra = "N"
            
        WControla = 0
        spTerminado = "ConsultaTerminado " + "'" + WArticulo.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
            rstTerminado.Close
        End If
            
        If WControla = 0 Then
            XParam = "'" + WLote.Text + "','" _
                     + WArticulo.Text + "'"
            spHoja = "ListaHojaProducto " + XParam
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                WEntra = "S"
                rstHoja.Close
            End If
                
            If WEntra = "N" Then
                XParam = "'" + WArticulo.Text + "','" _
                            + WLote.Text + "'"
                spMovguia = "ListaMovguiaLote1 " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    WEntra = "S"
                    rstMovguia.Close
                End If
            End If
                
                Else
                    
            WEntra = "S"
                
        End If
        
        If WEntra = "N" Then
            m$ = WArticulo.Text + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
            G% = MsgBox(m$, 0, "Mercaderia en Consignacion")
                Else
            If WSaldo1 >= Val(WCantidad.Text) Then
                Envase1.SetFocus
                     Else
                XSaldo1 = WSaldo1
                XSaldo1 = Pusing("###,###.##", XSaldo1)
                m$ = WArticulo.Text + " Cantidad Insuficiente Stock : " + XSaldo1
                G% = MsgBox(m$, 0, "Mercaderia en Consignacion")
                WLote.SetFocus
            End If
            Rem Envase1.SetFocus
        End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase1_Keypress(KeyAscii As Integer)
    Rem caca
    If KeyAscii = 13 Then
        If Val(Envase1.Text) <> 0 Then
            Ingre = "N"
            For DA = 1 To 6
                If Val(WVector(DA, 1)) = Val(Envase1.Text) Then
                    Ingre = "S"
                    Exit For
                End If
            Next DA
            If Ingre = "S" Then
                Canti1.SetFocus
                    Else
                Envase1.SetFocus
            End If
                Else
            Call Alta_Vector
            Call Ingresa_Click
            WArticulo.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Envase2.Text) <> 0 Then
            Ingre = "N"
            For DA = 1 To 6
                If Val(WVector(DA, 1)) = Val(Envase2.Text) Then
                    Ingre = "S"
                    Exit For
                End If
            Next DA
            If Ingre = "S" Then
                Canti2.SetFocus
                    Else
                Envase2.SetFocus
            End If

                Else
            Call Alta_Vector
            Call Ingresa_Click
            WArticulo.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
Private Sub Envase3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Envase3.Text) <> 0 Then
            Ingre = "N"
            For DA = 1 To 6
                If Val(WVector(DA, 1)) = Val(Envase3.Text) Then
                    Ingre = "S"
                    Exit For
                End If
            Next DA
            If Ingre = "S" Then
                Canti3.SetFocus
                    Else
                Envase3.SetFocus
            End If
            
                Else
            Call Alta_Vector
            Call Ingresa_Click
            WArticulo.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Alta_Vector
        Call Ingresa_Click
        WArticulo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub pantalla_Click()
    Ayuda.Visible = False
    Opcion.Visible = False
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
                WProv = rstCliente!Provincia
                WCodIva = rstCliente!Iva
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                rstCliente.Close
                Observaciones.SetFocus
            End If
            Pantalla.Visible = False
            
        Case 1
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WArticulo.Text = rstTerminado!Codigo
                WDescripcion.Caption = rstTerminado!Descripcion
                WPrecio.Caption = ""
                    
                DBGrid1.Col = 0
                DBGrid1.Text = rstTerminado!Codigo
                DBGrid1.Col = 1
                DBGrid1.Text = rstTerminado!Descripcion
                DBGrid1.Col = 3
                DBGrid1.Text = ""
                    
                Erase WVector

                Call Alta_Vector
                WLinea.Text = WAnterior + 1
                If Val(WLinea.Text) > 0 Then
                    DBGrid1.Row = Val(WLinea.Text) - 1
                End If
                    
                Call DBGrid1.SetFocus
                WCantidad.SetFocus
                    
            End If
            
            spTerminado = "ConsultaTerminado " + "'" + WArticulo.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WVector(1, 1) = rstTerminado!Envase1
                WVector(2, 1) = rstTerminado!Envase2
                WVector(3, 1) = rstTerminado!Envase3
                WVector(4, 1) = rstTerminado!Envase4
                WVector(5, 1) = rstTerminado!Envase5
                WVector(6, 1) = rstTerminado!Envase6
                rstTerminado.Close
                Call Carga_Envases
            End If
            
            WCliente = Cliente.Text
            WTerminado = WArticulo.Text
            WClave = Cliente.Text + WArticulo.Text
            
            spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                DBGrid1.Col = 0
                DBGrid1.Text = rstPrecios!Descripcion
                DBGrid1.Col = 3
                DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio))
                rstPrecios.Close
            End If
            
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

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 4, 0 To 50)

mTotalRows& = 50

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
For i = 0 To 4
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
             DBGrid1.Columns(newcnt).Caption = "Precio"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
             
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Lote"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
 WEnvase(1) = 20
 WEnvase(2) = 21
 WEnvase(3) = 22
 WEnvase(4) = 23
 WEnvase(5) = 24
 WEnvase(6) = 25
 WEnvase(7) = 26
 WEnvase(8) = 30
 WEnvase(9) = 28

 For Cicla = 1 To 9
    spEnvase = "ConsultaEnvases " + "'" + WEnvase(Cicla) + "'"
    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvase.RecordCount > 0 Then
        WImpre(Cicla) = Left$(rstEnvase!Abreviatura, 7)
        rstEnvase.Close
            Else
        WImpre(Cicla) = ""
    End If
Next Cicla

    Erase XEnvase

    Pantalla.Visible = False
    Numero.Text = ""
    spConsig = "ListaConsigNumero"
    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
    If rstConsig.RecordCount > 0 Then
        With rstConsig
            .MoveLast
            Numero.Text = rstConsig!Numero + 1
        End With
        rstConsig.Close
    End If

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Graba.Enabled = True
    
    Numero.SetFocus
    
End Sub

Private Sub Proceso_Click()

    Erase XEnvase

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 4
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    Erase Auxiliar
    WRenglon = 0

    spConsig = "ListaConsig " + "'" + Numero.Text + "'"
    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)

    If rstConsig.RecordCount > 0 Then
            With rstConsig
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        Renglon = Renglon + 1
                
                        Lugar1 = Int((Renglon - 1) / 10) * 10
                        Lugar2 = Renglon - Lugar1
                
                        DBGrid1.FirstRow = Lugar1
                        DBGrid1.Row = Lugar2 - 1
                
                        DBGrid1.Col = 0
                        DBGrid1.Text = rstConsig!Terminado
                        Auxi1 = rstConsig!Terminado
                        
                        DBGrid1.Col = 2
                        DBGrid1.Text = Pusing("###,###.##", rstConsig!Cantidad - rstConsig!Facturado)
                        
                        DBGrid1.Col = 4
                        DBGrid1.Text = rstConsig!Lote
                
                        XEnvase(Renglon, 1) = rstConsig!Envase1
                        XEnvase(Renglon, 2) = rstConsig!Canti1
                        XEnvase(Renglon, 3) = rstConsig!Envase2
                        XEnvase(Renglon, 4) = rstConsig!Canti2
                        XEnvase(Renglon, 5) = rstConsig!Envase3
                        XEnvase(Renglon, 6) = rstConsig!Canti3
                        
                        WRenglon = WRenglon + 1
                    
                        Auxiliar(WRenglon, 1) = rstConsig!Cliente
                        Auxiliar(WRenglon, 2) = rstConsig!Terminado
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstConsig.Close
    End If
    
    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    
    Renglon = 0
    
    For DA = 1 To WRenglon
        Cliente = Auxiliar(DA, 1)
        Terminado = Auxiliar(DA, 2)
        
        spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
        
            Renglon = Renglon + 1
                
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
                
            DBGrid1.Col = 0
            DBGrid1.Text = rstTerminado!Codigo
            
            DBGrid1.Col = 1
            DBGrid1.Text = rstTerminado!Descripcion
            
            DBGrid1.Col = 3
            DBGrid1.Text = ""
            
            rstTerminado.Close
            
        End If
        
        WClave = Cliente + Terminado
        
        spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
                DBGrid1.Col = 1
                DBGrid1.Text = rstPrecios!Descripcion
                DBGrid1.Col = 3
                DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio))
                rstPrecios.Close
        End If
        
    Next DA

    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    Graba.Enabled = False
    
    WArticulo.SetFocus

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
            DBGrid1.Text = Pusing("###,###.##", WPrecio.Caption)
            
            DBGrid1.Col = 4
            DBGrid1.Text = WLote.Text
            
            Rem DbGrid1.Row = Renglon
            DBGrid1.Row = Lugar2 - 1
            DBGrid1.Col = 0
            
            XEnvase(Renglon, 1) = Envase1.Text
            XEnvase(Renglon, 2) = Canti1.Text
            XEnvase(Renglon, 3) = Envase2.Text
            XEnvase(Renglon, 4) = Canti2.Text
            XEnvase(Renglon, 5) = Envase3.Text
            XEnvase(Renglon, 6) = Canti3.Text
            
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
            DBGrid1.Text = Pusing("###,###.##", WPrecio.Caption)
            
            DBGrid1.Col = 4
            DBGrid1.Text = WLote.Text
                
            Rem DbGrid1.Row = Renglon
            DBGrid1.Row = Lugar2 - 1
            DBGrid1.Col = 0
            WInicio = DBGrid1.FirstRow
            
            XEnvase(Val(WLinea.Text) + WInicio, 1) = Envase1.Text
            XEnvase(Val(WLinea.Text) + WInicio, 2) = Canti1.Text
            XEnvase(Val(WLinea.Text) + WInicio, 3) = Envase2.Text
            XEnvase(Val(WLinea.Text) + WInicio, 4) = Canti2.Text
            XEnvase(Val(WLinea.Text) + WInicio, 5) = Envase3.Text
            XEnvase(Val(WLinea.Text) + WInicio, 6) = Canti3.Text
            
            DBGrid1.Row = DBGrid1.Row + 1
            
    End If

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)

    On Error GoTo WError

    If KeyAscii = 13 Then
    
        spConsig = "ListaConsig " + "'" + Numero.Text + "'"
        Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
        If rstConsig.RecordCount > 0 Then
            Fecha.Text = rstConsig!Fecha
            Cliente.Text = rstConsig!Cliente
            Observaciones.Text = rstConsig!Observaciones
            rstConsig.Close
            
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WDirentrega = rstCliente!DirEntrega
                WProv = rstCliente!Provincia
                WCodIva = rstCliente!Iva
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                rstCliente.Close
            End If
            Call Proceso_Click
                Else
            WNumero = Numero.Text
            Call Limpia_Click
            Numero.Text = WNumero
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
                WProv = rstCliente!Provincia
                WCodIva = rstCliente!Iva
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
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
        Call Ingresa_Click
        WArticulo.SetFocus
    End If
End Sub

Sub Carga_Envases()

 For Cicla = 1 To 6
    spEnvase = "ConsultaEnvases " + "'" + WVector(Cicla, 1) + "'"
    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvase.RecordCount > 0 Then
        WVector(Cicla, 2) = rstEnvase!Kilos
        WVector(Cicla, 3) = rstEnvase!Abreviatura
                Else
        WVector(Cicla, 2) = ""
        WVector(Cicla, 3) = ""
    End If
Next Cicla

WEnvase1.Caption = WVector(1, 1)
WEnvase2.Caption = WVector(2, 1)
WEnvase3.Caption = WVector(3, 1)
WEnvase4.Caption = WVector(4, 1)
WEnvase5.Caption = WVector(5, 1)
WEnvase6.Caption = WVector(6, 1)

WCapa1.Caption = WVector(1, 2)
WCapa2.Caption = WVector(2, 2)
Wcapa3.Caption = WVector(3, 2)
WCapa4.Caption = WVector(4, 2)
WCapa5.Caption = WVector(5, 2)
WCapa6.Caption = WVector(6, 2)

WAbre1.Caption = WVector(1, 3)
WAbre2.Caption = WVector(2, 3)
WAbre3.Caption = WVector(3, 3)
WAbre4.Caption = WVector(4, 3)
WAbre5.Caption = WVector(5, 3)
WAbre6.Caption = WVector(6, 3)


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

Sub Impresion()

        If Val(WEmpresa) = 1 Then
            Rem Open "DADA.TXT" For Output As #1
            Open "lpt1" For Output As #1
                Else
            Rem Open "DADA.TXT" For Output As #1
            Open "lpt1" For Output As #1
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
        End If
  
        Rem  #1, 255

        For FF = 1 To 2

        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "2" + Chr$(72)
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        Print #1, ""
        Print #1, Tab(48); "REMITO EN CONSIGNACION"
        Print #1, ""
        Print #1, ""
        Print #1, Tab(53); Fecha.Text
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(7); WRazon
        Print #1, Tab(7); Left$(WDireccion, 33)
        Print #1, Tab(7); Left$(WLocalidad, 33);
        Print #1, Tab(57); Cliente.Text;
        Print #1, Tab(7); Provincia(Val(WProv)); "("; WPostal; ")"
        Print #1, ""
        Print #1, Tab(7); Iva(Val(WCodIva));
        Print #1, Tab(48); WCuit
        Print #1, ""
        Print #1, Tab(30); WDirentrega;
        Print #1, ""
        If FF = 1 Then
            Print #1, Tab(60); "ORIGINAL"
                Else
            Print #1, Tab(60); "DUPLICADO"
        End If
        Print #1, ""
        
        Impre = 0

        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Producto = DBGrid1.Text
                
                DBGrid1.Col = 1
                Descri = DBGrid1.Text
                
                DBGrid1.Col = 3
                Precio = Val(DBGrid1.Text)
            
                DBGrid1.Col = 2
                Cantidad = Val(DBGrid1.Text)
                
                If Cantidad <> 0 Then
                        
                        Print #1, Tab(14); Left$(Descri, 40);
                        Print #1, Tab(58); Alinea("#####.##", Str$(Cantidad));
                        Print #1, " Kg";
                        Print #1, Tab(71); "Netos"
                        Impre = Impre + 1
                End If
                    
            Next iRow
            
        Next a
        
        For aa = Impre To 20
                Print #1, ""
        Next aa
        
        Print #1, ""
        Print #1, Tab(10); "Lugar de Pago : Ayacucho 1231 5to Piso Dto. 'A' Capital Federal"
        Print #1, ""
        
        Call Calcula_Saldo
        
        For XDa = 1 To 40
            XEnv1 = XEnvase(XDa, 1)
            XCanti1 = XEnvase(XDa, 2)
            XEnv2 = XEnvase(XDa, 3)
            XCanti2 = XEnvase(XDa, 4)
            XEnv3 = XEnvase(XDa, 5)
            XCanti3 = XEnvase(XDa, 6)
        
            For DA = 1 To 9
                If Val(XEnv1) = Val(Stk(DA, 1)) Then
                    Stk(DA, 3) = Str$(Val(Stk(DA, 3)) + Val(XCanti1))
                End If
                If Val(XEnv2) = Val(Stk(DA, 1)) Then
                    Stk(DA, 3) = Str$(Val(Stk(DA, 3)) + Val(XCanti2))
                End If
                If Val(XEnv3) = Val(Stk(DA, 1)) Then
                    Stk(DA, 3) = Str$(Val(Stk(DA, 3)) + Val(XCanti3))
                End If
            Next DA
        Next XDa
        
        For DA = 1 To 9
            Stk(DA, 4) = Str$(Val(Stk(DA, 2)) + Val(Stk(DA, 3)))
        Next DA
        
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)

        For XDa = 1 To 1
                For DA = 1 To 9
                        If Val(Stk(DA, 4)) <> 0 Then
                                        
                                Select Case DA
                                        Case 1
                                                Lugar = 22
                                        Case 2
                                                Lugar = 33
                                        Case 3
                                                Lugar = 44
                                        Case 4
                                                Lugar = 55
                                        Case 5
                                                Lugar = 66
                                        Case 6
                                                Lugar = 77
                                        Case 7
                                                Lugar = 89
                                        Case 8
                                                Lugar = 101
                                        Case 9
                                                Lugar = 113
                                        Case Else
                                End Select
                                                         
                                If DA = 9 Then
                                    Digi = 7
                                            Else
                                    Digi = 7
                                End If
                                
                                spEnvases = "ConsultaEnvases " + "'" + Str$(Val(Stk(DA, XDa))) + "'"
                                Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                                If rstEnvases.RecordCount > 0 Then
                                    Print #1, Tab(Lugar); Left$(rstEnvases!Abreviatura, Digi);
                                    rstEnvases.Close
                                            Else
                                    Print #1, Tab(Lugar); Stk(DA, XDa);
                                End If
                            End If
        
                Next DA
                Print #1, ""
        
        Next XDa
        
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        
        For XDa = 2 To 4
                For DA = 1 To 9
        
                        If Val(Stk(DA, 4)) <> 0 Then
        
                                Select Case DA
                                        Case 1
                                                Lugar = 14
                                        Case 2
                                                Lugar = 21
                                        Case 3
                                                Lugar = 29
                                        Case 4
                                                Lugar = 36
                                        Case 5
                                                Lugar = 43
                                        Case 6
                                                Lugar = 50
                                        Case 7
                                                Lugar = 57
                                        Case 8
                                                Lugar = 64
                                        Case 9
                                                Lugar = 71
                                        Case Else
                                End Select
        
                                If Val(Stk(DA, XDa)) <> 0 Then
                                        Print #1, Tab(Lugar); Alinea("####", Str$(Val(Stk(DA, XDa))));
                                End If
        
                         End If
                Next DA
        
                Print #1, ""
                Print #1, ""
        
        Next XDa
        
        Print #1, ""
        Select Case XX
                Case 1
                        Print #1, Tab(10); "ORIGINAL";
                Case 2
                        Print #1, Tab(10); "DUPLICADO";
                Case 3
                        Print #1, Tab(10); "TRIPLICADO";
                Case Else
        End Select
        Print #1, Tab(10); "Nro. Control : "; Numero.Text
        Print #1, Chr$(12)

        Next FF

        Close #1

End Sub

Private Sub Calcula_Saldo()

    Rem On Error GoTo Error_saldo


    Erase Stk

    If Val(WEmpresa) = 8 Then
        Stk(1, 1) = "005"
        Stk(2, 1) = "011"
        Stk(3, 1) = "021"
        Stk(4, 1) = "027"
        Stk(5, 1) = "004"
        Stk(6, 1) = "012"
        Stk(7, 1) = "000"
        Stk(8, 1) = "000"
        Stk(9, 1) = "000"
            Else
        Stk(1, 1) = "020"
        Stk(2, 1) = "021"
        Stk(3, 1) = "022"
        Stk(4, 1) = "023"
        Stk(5, 1) = "024"
        Stk(6, 1) = "025"
        Stk(7, 1) = "026"
        Stk(8, 1) = "030"
        Stk(9, 1) = "028"
    End If

    XParam = "'" + Cliente.Text + "','" _
                + Cliente.Text + "'"

    spMovenv = "ListaMovenvDesdeHastaCliente " + XParam
    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovenv.RecordCount > 0 Then
    
        With rstMovenv
            .MoveFirst
            Do
                If .EOF = False Then

                    For DA = 1 To 9
                        If Val(Stk(DA, 1)) = !Envase Then
                            If !Movimiento = "S" Then
                                Stk(DA, 2) = Str$(Val(Stk(DA, 2)) + !Cantidad)
                                    Else
                                Stk(DA, 2) = Str$(Val(Stk(DA, 2)) - !Cantidad)
                            End If
                        End If
                    
                    Next DA
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovenv.Close
    End If

End Sub





