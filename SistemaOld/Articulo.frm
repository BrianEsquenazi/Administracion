VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgArticulo 
   Caption         =   "Ingreso de Materia Prima"
   ClientHeight    =   6465
   ClientLeft      =   1755
   ClientTop       =   285
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6465
   ScaleWidth      =   6810
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin VB.TextBox Costo1 
      Height          =   285
      Left            =   2280
      TabIndex        =   48
      Text            =   " "
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Rs 
      Height          =   285
      Left            =   5520
      MaxLength       =   1
      TabIndex        =   47
      Text            =   " "
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Envase 
      Height          =   285
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   46
      Text            =   " "
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Pedido 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   45
      Text            =   " "
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Laboratorio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   44
      Text            =   " "
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Deposito 
      Height          =   285
      Left            =   2280
      TabIndex        =   39
      Text            =   " "
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Unidad 
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   38
      Text            =   " "
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Minimo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   37
      Text            =   " "
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Salidas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   36
      Text            =   " "
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Entradas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   35
      Text            =   " "
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Inicial 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   34
      Text            =   " "
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Costo2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   33
      Text            =   " "
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1455
      Left            =   240
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
      Width           =   3735
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1440
         TabIndex        =   50
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1440
         TabIndex        =   49
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   23
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5520
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "materia.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "Articulo.frx":0000
      Left            =   840
      List            =   "Articulo.frx":0007
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton lista 
      Caption         =   "Listado"
      Height          =   300
      Left            =   1800
      TabIndex        =   14
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   600
      TabIndex        =   13
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles"
      Height          =   1335
      Left            =   4200
      TabIndex        =   8
      Top             =   3720
      Width           =   1575
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   3000
      TabIndex        =   0
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   3000
      TabIndex        =   7
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   1800
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   600
      TabIndex        =   5
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox Descripcion 
      Height          =   285
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   4
      Top             =   360
      Width           =   3375
   End
   Begin VB.ListBox Opcion 
      Height          =   1035
      ItemData        =   "Articulo.frx":0015
      Left            =   240
      List            =   "Articulo.frx":0017
      TabIndex        =   53
      Top             =   5040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label DescriEnvase 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3840
      TabIndex        =   52
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Stock 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2280
      TabIndex        =   51
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Rs"
      Height          =   255
      Left            =   3840
      TabIndex        =   43
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Envase"
      Height          =   255
      Left            =   3840
      TabIndex        =   42
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "Pedido Pendiente"
      Height          =   255
      Left            =   3840
      TabIndex        =   41
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "Stock Laboratorio"
      Height          =   255
      Left            =   3840
      TabIndex        =   40
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Cantidad Final"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Deposito"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Unidad de Medida"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Stock Minimo"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Cantidad Salida"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Cantidad Entrada"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Cantidad Inicial"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Costo Standard"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Costo Ultima compra"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "PrgArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Verifica_datos()
    If Val(Costo1.Text) = 0 Then
        Costo1.Text = "0"
    End If
    If Val(Costo2.Text) = 0 Then
        Costo2.Text = "0"
    End If
    If Val(Entradas.Text) = 0 Then
        Entradas.Text = "0"
    End If
    If Val(Salidas.Text) = 0 Then
        Salidas.Text = "0"
    End If
    If Val(Minimo.Text) = 0 Then
        Minimo.Text = "0"
    End If
    If Val(Laboratorio.Text) = 0 Then
        Laboratorio.Text = "0"
    End If
    If Val(Pedido.Text) = 0 Then
        Pedido.Text = "0"
    End If
    
End Sub





Sub Format_datos()
    Costo1.Text = Pusing("###,###.##", Costo1.Text)
    Costo2.Text = Pusing("###,###.##", Costo2.Text)
    Inicial.Text = Pusing("###,###.##", Inicial.Text)
    Entradas.Text = Pusing("###,###.##", Entradas.Text)
    Salidas.Text = Pusing("###,###.##", Salidas.Text)
    Minimo.Text = Pusing("###,###.##", Minimo.Text)
    Laboratorio.Text = Pusing("###,###.##", Laboratorio.Text)
    Pedido.Text = Pusing("###,###.##", Pedido.Text)
    Stock.Caption = Pusing("###,###.##", CDbl(Inicial.Text) + CDbl(Entradas.Text) - CDbl(Salidas.Text))
End Sub
Private Sub Acepta_Click()
    Listado.GroupSelectionFormula = "{Producto.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Producto.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Producto.Text <> "" Then
        With rstProductos
            .Index = "Producto"
            .Seek "=", Producto.Text
            If .NoMatch Then
                .AddNew
                Call Verifica_datos
                !Producto = Producto.Text
                !Descripcion = Descripcion.Text
                !Costo1 = CDbl(Costo1.Text)
                !Costo2 = CDbl(Costo2.Text)
                !Inicial = CDbl(Inicial.Text)
                !Entradas = CDbl(Entradas.Text)
                !Salidas = CDbl(Salidas.Text)
                !Minimo = CDbl(Minimo.Text)
                !Unidad = Unidad.Text
                !Deposito = Deposito.Text
                !Laboratorio = CDbl(Laboratorio.Text)
                !Pedido = CDbl(Pedido.Text)
                !Envase = Envase.Text
                !Rs = Rs.Text
                .Update
                .Bookmark = .LastModified
                    Else
                .Edit
                Call Verifica_datos
                !Producto = Producto.Text
                !Descripcion = Descripcion.Text
                !Costo1 = CDbl(Costo1.Text)
                !Costo2 = CDbl(Costo2.Text)
                !Inicial = CDbl(Inicial.Text)
                !Entradas = CDbl(Entradas.Text)
                !Salidas = CDbl(Salidas.Text)
                !Minimo = CDbl(Minimo.Text)
                !Unidad = Unidad.Text
                !Deposito = Deposito.Text
                !Laboratorio = CDbl(Laboratorio.Text)
                !Pedido = CDbl(Pedido.Text)
                !Envase = Envase.Text
                !Rs = Rs.Text
                .Update
                .Bookmark = .LastModified
            End If
        End With
        Call CmdLimpiar_Click
        Producto.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Producto.Text <> "" Then
        With rstProductos
            .Index = "Producto"
            .Seek "=", Producto.Text
            If .NoMatch = False Then
                T$ = "Borrar Registro"
                M$ = "Desea Borrar el Registro "
                Respuesta% = MsgBox(M$, 32 + 4, T$)
                If Respuesta% = 6 Then
                    .Delete
                    Call CmdLimpiar_Click
                End If
            End If
        End With
    End If
    Producto.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Producto.Text = "  -   -   "
    Descripcion.Text = ""
    Costo1.Text = ""
    Costo2.Text = ""
    Inicial.Text = "0"
    Entradas.Text = "0"
    Salidas.Text = "0"
    Minimo.Text = ""
    Unidad.Text = ""
    Deposito.Text = ""
    Rem Fecha.Text = ""
    Rem Orden.Text = ""
    Laboratorio.Text = ""
    Pedido.Text = ""
    Envase.Text = ""
    Rs.Text = ""
    Stock.Caption = ""
    DescriEnvase.Caption = ""
    Producto.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgArticulo.Hide
    Menu.SetFocus
End Sub

Private Sub Anterior_Click()
    With rstProductos
        .Index = "Producto"
        .Seek "=", Producto.Text
        If .NoMatch = False Then
            .MovePrevious
            If .BOF = True Then
                .MoveFirst
                M$ = "No exsite registro Anterior"
                A% = MsgBox(M$, 0, "Archivo de Articulos")
                .MoveFirst
            End If
            Producto.Text = !Producto
            Descripcion.Text = !Descripcion
            Costo1.Text = Str$(!Costo1)
            Costo2.Text = Str$(!Costo2)
            Inicial.Text = Str$(!Inicial)
            Entradas.Text = Str$(!Entradas)
            Salidas.Text = Str$(!Salidas)
            Minimo.Text = Str$(!Minimo)
            Unidad.Text = !Unidad
            Deposito.Text = !Deposito
            Rem Fecha.Text = !Fecha
            Rem Orden.Text = !Orden
            Laboratorio.Text = Str$(!Laboratorio)
            Pedido.Text = Str$(!Pedido)
            Envase.Text = !Envase
            Rs.Text = !Rs
            With rstEnvases
                .Index = "Envases"
                .Seek "=", Envase.Text
                If .NoMatch = False Then
                    DescriEnvase.Caption = !Descripcion
                End If
            End With
            Call Format_datos
            Producto.SetFocus
        End If
    End With
End Sub
Private Sub Lista_Click()
    Desde.Text = "AA-000-000"
    Hasta.Text = "ZZ-999-999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Costo1.SetFocus
    End If
End Sub

Private Sub Costo1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Costo1.Text = Pusing("###,###.##", Costo1.Text)
        Costo2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
Private Sub Costo2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Costo2.Text = Pusing("###,###.##", Costo2.Text)
        Inicial.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
Private Sub Inicial_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Stock.Caption = Pusing("###,###.##", CDbl(Inicial.Text) + CDbl(Entradas.Text) - CDbl(Salidas.Text))
        Inicial.Text = Pusing("###,###.##", Inicial.Text)
        Entradas.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
Private Sub Entradas_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Stock.Caption = Pusing("###,###.##", CDbl(Inicial.Text) + CDbl(Entradas.Text) - CDbl(Salidas.Text))
        Entradas.Text = Pusing("###,###.##", Entradas.Text)
        Salidas.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Salidas_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Stock.Caption = Pusing("###,###.##", CDbl(Inicial.Text) + CDbl(Entradas.Text) - CDbl(Salidas.Text))
        Salidas.Text = Pusing("###,###.##", Salidas.Text)
        Minimo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)

End Sub
Private Sub Minimo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Minimo.Text = Pusing("###,###.##", Minimo.Text)
        Unidad.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
Private Sub Unidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Deposito.SetFocus
    End If
End Sub
Private Sub Deposito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Laboratorio.SetFocus
    End If
End Sub
Private Sub Laboratorio_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Laboratorio.Text = Pusing("###,###.##", Laboratorio.Text)
        Pedido.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
Private Sub Pedido_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Pedido.Text = Pusing("###,###.##", Pedido.Text)
        Envase.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
Private Sub Envase_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstEnvases
            .Index = "Envases"
            .Seek "=", Envase.Text
            If .NoMatch = False Then
                DescriEnvase.Caption = !Descripcion
                Rs.SetFocus
            End If
        End With
    End If
End Sub
Private Sub Rs_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Rs.Text = "1" Or Rs.Text = "2" Or Rs.Text = "3" Then
            Descripcion.SetFocus
        End If
    End If
End Sub
Sub Producto_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "" Then
            With rstProductos
                .Index = "Producto"
                ClaveProd$ = Producto.Text
                .Seek "=", Producto.Text
                If .NoMatch Then
                    CmdLimpiar_Click
                    Producto.Text = ClaveProd$
                        Else
                    Producto.Text = !Producto
                    Descripcion.Text = !Descripcion
                    Costo1.Text = Str$(!Costo1)
                    Costo2.Text = Str$(!Costo2)
                    Inicial.Text = Str$(!Inicial)
                    Entradas.Text = Str$(!Entradas)
                    Salidas.Text = Str$(!Salidas)
                    Minimo.Text = Str$(!Minimo)
                    Unidad.Text = !Unidad
                    Deposito.Text = !Deposito
                    Rem Fecha.Text = !Fecha
                    Rem Orden.Text = !Orden
                    Laboratorio.Text = Str$(!Laboratorio)
                    Pedido.Text = Str$(!Pedido)
                    Envase.Text = !Envase
                    Rs.Text = !Rs
                    With rstEnvases
                        .Index = "Envases"
                        .Seek "=", Envase.Text
                        If .NoMatch = False Then
                            DescriEnvase.Caption = !Descripcion
                        End If
                    End With
                    Call Format_datos
                End If
            End With
        End If
        Descripcion.SetFocus
    End If
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Articulos"
     Opcion.AddItem "Envases"

     Opcion.Visible = True
     
 End Sub

 Private Sub Opcion_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    Rem XIndice = 0
    
    Select Case XIndice
        Case 0
            With rstProductos
                .Index = "Producto"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Producto + " " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Producto
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
        Case 1
            With rstEnvases
                .Index = "Envases"
                .MoveFirst
                Do
                     If .EOF = False Then
                         IngresaItem = Str$(!Envases) + " " + !Descripcion
                         Pantalla.AddItem IngresaItem
                         IngresaItem = !Envases
                         WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
             End With
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            With rstProductos

                Indice = Pantalla.ListIndex
                ClaveProd$ = WIndice.List(Indice)
                Producto.Text = ClaveProd$
                .Index = "Producto"
                ClaveProd$ = Producto.Text
                .Seek "=", ClaveProd$
                If .NoMatch = False Then
                    Producto.Text = !Producto
                    Descripcion.Text = !Descripcion
                    Costo1.Text = Val(!Costo1)
                    Costo2.Text = Val(!Costo2)
                    Inicial.Text = Val(!Inicial)
                    Entradas.Text = Val(!Entradas)
                    Salidas.Text = Val(!Salidas)
                    Minimo.Text = Val(!Minimo)
                    Unidad.Text = !Unidad
                    Deposito.Text = !Deposito
                    Rem Fecha.Text = !Fecha
                    Rem Orden.Text = !Orden
                    Laboratorio.Text = Val(!Laboratorio)
                    Pedido.Text = Val(!Pedido)
                    Envase.Text = !Envase
                    Rs.Text = !Rs
                    With rstEnvases
                        .Index = "Envases"
                        .Seek "=", Envase.Text
                        If .NoMatch = False Then
                            DescriEnvase.Caption = !Descripcion
                        End If
                    End With
                    Call Format_datos
                        Else
                    CmdLimpiar_Click
                    Producto.Text = ClaveProd$
                End If
            End With
            Producto.SetFocus
        Case 1
            With rstEnvases
                Indice = Pantalla.ListIndex
                ClaveEnv$ = WIndice.List(Indice)
                Envase.Text = ClaveEnv$
                .Index = "Envases"
                .Seek "=", Envase.Text
                If .NoMatch = False Then
                    DescriEnvase.Caption = !Descripcion
                End If
            End With
            Rs.SetFocus
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()
    Rem On Error GoTo Error_primer
    With rstProductos
        .Index = "Producto"
        .MoveFirst
        Producto.Text = !Producto
        Descripcion.Text = !Descripcion
        Costo1.Text = Val(!Costo1)
        Costo2.Text = Val(!Costo2)
        Inicial.Text = Val(!Inicial)
        Entradas.Text = Val(!Entradas)
        Salidas.Text = Val(!Salidas)
        Minimo.Text = Val(!Minimo)
        Unidad.Text = !Unidad
        Deposito.Text = !Deposito
        Rem Fecha.Text = !Fecha
        Rem Orden.Text = !Orden
        Laboratorio.Text = Val(!Laboratorio)
        Pedido.Text = Val(!Pedido)
        Envase.Text = !Envase
        Rs.Text = !Rs
        With rstEnvases
                .Index = "Envases"
                .Seek "=", Envase.Text
                If .NoMatch = False Then
                    DescriEnvase.Caption = !Descripcion
                End If
        End With
        Call Format_datos
        Producto.SetFocus
    End With
    Exit Sub

Error_primer:
     coderr = Err
     Call Errores(coderr, "Productos", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Prod1.SetFocus
 End Sub

Private Sub Ultimo_Click()
    Rem On Error GoTo Error_ultimo
    With rstProductos
        .Index = "Producto"
        .MoveLast
        Producto.Text = !Producto
        Descripcion.Text = !Descripcion
        Costo1.Text = Val(!Costo1)
        Costo2.Text = Val(!Costo2)
        Inicial.Text = Val(!Inicial)
        Entradas.Text = Val(!Entradas)
        Salidas.Text = Val(!Salidas)
        Minimo.Text = Val(!Minimo)
        Unidad.Text = !Unidad
        Deposito.Text = !Deposito
        Rem  Fecha.Text = !Fecha
        Rem Orden.Text = !Orden
        Laboratorio.Text = Val(!Laboratorio)
        Pedido.Text = Val(!Pedido)
        Envase.Text = !Envase
        Rs.Text = !Rs
        With rstEnvases
                .Index = "Envases"
                .Seek "=", Envase.Text
                If .NoMatch = False Then
                    DescriEnvase.Caption = !Descripcion
                End If
        End With
        Call Format_datos
        Producto.SetFocus
    End With
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Prodcuto", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Prod1.SetFocus
 End Sub

Private Sub Siguiente_Click()
    With rstProductos
        .Index = "Producto"
        ClaveProd$ = Producto.Text
        .Seek "=", ClaveProd$
        If .NoMatch = False Then
            .MoveNext
            If .EOF = True Then
                M$ = "No exsite registro Posterior"
                A% = MsgBox(M$, 0, "Archivo de Articulos")
                Call Ultimo_Click
            End If
            Producto.Text = !Producto
            Descripcion.Text = !Descripcion
            Costo1.Text = Val(!Costo1)
            Costo2.Text = Val(!Costo2)
            Inicial.Text = Val(!Inicial)
            Entradas.Text = Val(!Entradas)
            Salidas.Text = Val(!Salidas)
            Minimo.Text = Val(!Minimo)
            Unidad.Text = !Unidad
            Deposito.Text = !Deposito
            Rem  Fecha.Text = !Fecha
            Rem Orden.Text = !Orden
            Laboratorio.Text = Val(!Laboratorio)
            Pedido.Text = Val(!Pedido)
            Envase.Text = !Envase
            Rs.Text = !Rs
            With rstEnvases
                .Index = "Envases"
                .Seek "=", Envase.Text
                If .NoMatch = False Then
                    DescriEnvase.Caption = !Descripcion
                End If
            End With
            Call Format_datos
            Producto.SetFocus
        End If
    End With
End Sub

