VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresos de Pedidos"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de pedido"
      Height          =   855
      Left            =   6840
      TabIndex        =   62
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton Tipo2 
         Caption         =   "Compra para Stock"
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Tipo1 
         Caption         =   "Venta a Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton Acepta 
      Caption         =   "Impresion"
      Height          =   220
      Left            =   1920
      TabIndex        =   61
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Precio5 
      Height          =   285
      Left            =   6960
      MaxLength       =   15
      TabIndex        =   60
      Text            =   " "
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox Precio4 
      Height          =   285
      Left            =   6960
      MaxLength       =   15
      TabIndex        =   59
      Text            =   " "
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox Precio3 
      Height          =   285
      Left            =   6960
      MaxLength       =   15
      TabIndex        =   58
      Text            =   " "
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox Precio2 
      Height          =   285
      Left            =   6960
      MaxLength       =   15
      TabIndex        =   57
      Text            =   " "
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Precio1 
      Height          =   285
      Left            =   6960
      MaxLength       =   15
      TabIndex        =   56
      Text            =   " "
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Repuesto5 
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   51
      Text            =   " "
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Repuesto4 
      Height          =   285
      Left            =   2280
      TabIndex        =   49
      Text            =   " "
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Repuesto3 
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   48
      Text            =   " "
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Repuesto2 
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   47
      Text            =   " "
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Repuesto1 
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   46
      Text            =   " "
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Obser5 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   44
      Text            =   " "
      Top             =   3480
      Width           =   7095
   End
   Begin VB.TextBox Obser4 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   43
      Text            =   " "
      Top             =   3240
      Width           =   7095
   End
   Begin VB.TextBox Obser3 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   42
      Text            =   " "
      Top             =   3000
      Width           =   7095
   End
   Begin VB.TextBox Obser2 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   41
      Text            =   " "
      Top             =   2760
      Width           =   7095
   End
   Begin VB.TextBox Obser1 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   40
      Text            =   " "
      Top             =   2520
      Width           =   7095
   End
   Begin VB.TextBox Cuotas 
      Height          =   285
      Left            =   5640
      MaxLength       =   4
      TabIndex        =   38
      Text            =   " "
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Interes 
      Height          =   285
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   36
      Text            =   " "
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox OCompra 
      Height          =   285
      Left            =   5400
      MaxLength       =   10
      TabIndex        =   34
      Text            =   " "
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Descuento3 
      Height          =   285
      Left            =   6840
      MaxLength       =   6
      TabIndex        =   33
      Text            =   " "
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Descuento2 
      Height          =   285
      Left            =   5880
      MaxLength       =   6
      TabIndex        =   32
      Text            =   " "
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Descuento1 
      Height          =   285
      Left            =   4920
      MaxLength       =   6
      TabIndex        =   31
      Text            =   " "
      Top             =   1800
      Width           =   855
   End
   Begin MSMask.MaskEdBox FEntrega 
      Height          =   285
      Left            =   2280
      TabIndex        =   28
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   5280
      TabIndex        =   27
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Pedido 
      Height          =   285
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   24
      Text            =   " "
      Top             =   0
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   6240
      Width           =   3855
      Begin VB.CommandButton Anterior 
         Caption         =   "Anterior"
         Height          =   220
         Left            =   2880
         TabIndex        =   10
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Siguiente"
         Height          =   220
         Left            =   960
         TabIndex        =   9
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
         Height          =   220
         Left            =   1920
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
         Height          =   220
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.TextBox Precio 
      Height          =   315
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   21
      Text            =   " "
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Pago 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   20
      Text            =   " "
      Top             =   1080
      Width           =   4695
   End
   Begin VB.TextBox Producto 
      Height          =   285
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   19
      Text            =   " "
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Cliente 
      Height          =   285
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   15
      Text            =   " "
      Top             =   360
      Width           =   735
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   8040
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "pedido.rpt"
      Destination     =   1
      WindowTitle     =   "Pedido"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   220
      Left            =   0
      TabIndex        =   11
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   220
      Left            =   1920
      TabIndex        =   0
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   220
      Left            =   960
      TabIndex        =   5
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   220
      Left            =   960
      TabIndex        =   4
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   220
      Left            =   0
      TabIndex        =   3
      Top             =   5640
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1035
      ItemData        =   "PEDIDO.frx":0000
      Left            =   3960
      List            =   "PEDIDO.frx":0007
      TabIndex        =   12
      Top             =   5520
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.ListBox Opcion 
      Height          =   1035
      Left            =   5040
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7680
      TabIndex        =   13
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label DesRepuesto5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3480
      TabIndex        =   55
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Label DesRepuesto4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3480
      TabIndex        =   54
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label DesRepuesto3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3480
      TabIndex        =   53
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Label DesRepuesto2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3480
      TabIndex        =   52
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label DesRepuesto1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3480
      TabIndex        =   50
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label Label9 
      Caption         =   "Adicionales"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Cantidad de Cuotas"
      Height          =   255
      Left            =   3840
      TabIndex        =   37
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "% de Interes"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Descuento"
      Height          =   255
      Left            =   3840
      TabIndex        =   30
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Orden de Compra"
      Height          =   255
      Left            =   3840
      TabIndex        =   29
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label DesProducto 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   4440
      TabIndex        =   26
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   3120
      TabIndex        =   25
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label10 
      Caption         =   "Fecha de entrega"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Precio"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Condicion de pago"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Poblaci 
      Caption         =   "Codigo de Producto"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Codigo de Cliente"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fecha del Pedido"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Numero de Pedido"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "PrgPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WModelo As Integer
Private WVendedor As Integer
Private WCiudad As Integer


Sub Imprime_Descripcion()
    With rstCliente
        .Index = "Cliente"
        .Seek "=", Cliente.Text
        If .NoMatch = False Then
            DesCliente.Caption = !Razon
            WVendedor = !Vendedor
            WCiudad = !Ciudad
                Else
            DesCliente.Caption = ""
        End If
    End With
    With rstProducto
        .Index = "Producto"
        .Seek "=", Producto.Text
        If .NoMatch = False Then
            DesProducto.Caption = !Descripcion1
            WModelo = !Modelo
                Else
            DesProducto.Caption = ""
        End If
    End With
    With rstRepuesto
        .Index = "Repuesto"
        
        .Seek "=", Val(Repuesto1.Text)
        If .NoMatch = False Then
            DesRepuesto1.Caption = !Descripcion
                Else
            DesRepuesto1.Caption = ""
        End If
        
        .Seek "=", Val(Repuesto2.Text)
        If .NoMatch = False Then
            DesRepuesto2.Caption = !Descripcion
                Else
            DesRepuesto2.Caption = ""
        End If
        
        .Seek "=", Val(Repuesto3.Text)
        If .NoMatch = False Then
            DesRepuesto3.Caption = !Descripcion
                Else
            DesRepuesto3.Caption = ""
        End If
        
        .Seek "=", Val(Repuesto4.Text)
        If .NoMatch = False Then
            DesRepuesto4.Caption = !Descripcion
                Else
            DesRepuesto4.Caption = ""
        End If
        
        .Seek "=", Val(Repuesto5.Text)
        If .NoMatch = False Then
            DesRepuesto5.Caption = !Descripcion
                Else
            DesRepuesto5.Caption = ""
        End If
        
    End With

End Sub

Sub Verifica_datos()
    If Val(Cliente.Text) = 0 Then
        Cliente.Text = "0"
    End If
    If Val(Precio.Text) = 0 Then
        Precio.Text = "0"
    End If
    If Val(Descuento1.Text) = 0 Then
        Descuento1.Text = "0"
    End If
    If Val(Descuento2.Text) = 0 Then
        Descuento2.Text = "0"
    End If
    If Val(Descuento3.Text) = 0 Then
        Descuento3.Text = "0"
    End If
    If Val(Interes.Text) = 0 Then
        Interes.Text = "0"
    End If
    If Val(Cuotas.Text) = 0 Then
        Cuotas.Text = "0"
    End If
    If Val(Repuesto1.Text) = 0 Then
        Repuesto1.Text = "0"
    End If
    If Val(Repuesto2.Text) = 0 Then
        Repuesto2.Text = "0"
    End If
    If Val(Repuesto3.Text) = 0 Then
        Repuesto3.Text = "0"
    End If
    If Val(Repuesto4.Text) = 0 Then
        Repuesto4.Text = "0"
    End If
    If Val(Repuesto5.Text) = 0 Then
        Repuesto5.Text = "0"
    End If
    If Val(Precio1.Text) = 0 Then
        Precio1.Text = "0"
    End If
    If Val(Precio2.Text) = 0 Then
        Precio2.Text = "0"
    End If
    If Val(Precio3.Text) = 0 Then
        Precio3.Text = "0"
    End If
    If Val(Precio4.Text) = 0 Then
        Precio4.Text = "0"
    End If
    If Val(Precio5.Text) = 0 Then
        Precio5.Text = "0"
    End If
End Sub

Sub Format_datos()
    Rem Precio.text = PUsing("###,###.##", Precio.text)
    Rem Precio1.text = PUsing("###,###.##", Precio1.text)
    Rem Precio2.text = PUsing("###,###.##", Precio2.text)
    Rem Precio3.text = PUsing("###,###.##", Precio3.text)
    Rem Precio4.text = PUsing("###,###.##", Precio4.text)
    Rem Precio5.text = PUsing("###,###.##", Precio5.text)
    Rem Interes.text = PUsing("###.##", Interes.text)
    Rem Descuento1.text = PUsing("###.##", Descuento1.text)
    Rem Descuento2.text = PUsing("###.##", Descuento2.text)
    Rem Descuento3.text = PUsing("###.##", Descuento3.text)
End Sub

Sub Imprime_Datos()
    With rstPedido
        .Index = "Pedido"
        .Seek "=", Pedido.Text
        If .NoMatch = False Then
            Pedido.Text = !Pedido
            Fecha.Text = !Fecha
            Cliente.Text = !Cliente
            Producto.Text = !Producto
            Pago.Text = !Pago
            Precio.Text = !Precio
            FEntrega.Text = !FEntrega
            Interes.Text = !Interes
            Cuotas.Text = !Cuotas
            Descuento1.Text = !Descuento1
            Descuento2.Text = !Descuento2
            Descuento3.Text = !Descuento3
            OCompra.Text = !OCompra
            Obser1.Text = !Obser1
            Obser2.Text = !Obser2
            Obser3.Text = !Obser3
            Obser4.Text = !Obser4
            Obser5.Text = !Obser5
            Repuesto1.Text = !Repuesto1
            Repuesto2.Text = !Repuesto2
            Repuesto3.Text = !Repuesto3
            Repuesto4.Text = !Repuesto4
            Repuesto5.Text = !Repuesto5
            Precio1.Text = !Precio1
            Precio2.Text = !Precio2
            Precio3.Text = !Precio3
            Precio4.Text = !Precio4
            Precio5.Text = !Precio5
            Tipo1.Value = False
            Tipo2.Value = False
            Select Case Val(!Tipo)
                Case 1
                    Tipo1.Value = True
                Case 2
                    Tipo2.Value = True
                Case Else
            End Select

            Call Format_datos
            Call Imprime_Descripcion
        End If
    End With
End Sub

Private Sub Acepta_Click()
    Listado.GroupSelectionFormula = "{Pedido.Pedido} in " + Chr$(34) + Pedido.Text + Chr$(34) + " to " + Chr$(34) + Pedido.Text + Chr$(34)
    Rem If Impresora.Value = True Then
    Rem    Listado.Destination = 1
    Rem        Else
    Rem    Listado.Destination = 0
    Rem End If
    Listado.Destination = 0
    Listado.Action = 1
    Rem Pedido.SetFocus
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()

    If Pedido.Text <> "" Then
    
        Call Verifica_datos
        
        WPasa = "S"
        
        With rstCliente
            .Index = "Cliente"
            .Seek "=", Cliente.Text
            If .NoMatch = True Then
                WPasa = "N"
                M$ = "Codigo de Cliente Incorrecto"
                A% = MsgBox(M$, 0, "Archivo de Pedidos")
            End If
        End With
        
        With rstProducto
            .Index = "Producto"
            .Seek "=", Producto.Text
            If .NoMatch = True Then
                WPasa = "N"
                M$ = "Codigo de Producto Incorrecto"
                A% = MsgBox(M$, 0, "Archivo de Pedidos")
            End If
        End With
        
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi <> "S" Then
            WPasa = "N"
            M$ = "Formato de Fecha Incorrecto, formato valido : dd/mm/aaaa"
            A% = MsgBox(M$, 0, "Archivo de Pedidos")
        End If
        
        With rstRepuesto
        
            .Index = "Repuesto"
            
            If Val(Repuesto1.Text) <> 0 Then
                .Seek "=", Val(Repuesto1.Text)
                If .NoMatch = True Then
                    WPasa = "N"
                    M$ = "Codigo de Adicionales Incorrecto"
                    A% = MsgBox(M$, 0, "Archivo de Pedidos")
                End If
            End If
            
            If Val(Repuesto2.Text) <> 0 Then
                .Seek "=", Val(Repuesto2.Text)
                If .NoMatch = True Then
                    WPasa = "N"
                    M$ = "Codigo de Adicionales Incorrecto"
                    A% = MsgBox(M$, 0, "Archivo de Pedidos")
                End If
            End If
            
            If Val(Repuesto3.Text) <> 0 Then
                .Seek "=", Val(Repuesto3.Text)
                If .NoMatch = True Then
                    WPasa = "N"
                    M$ = "Codigo de Adicionales Incorrecto"
                    A% = MsgBox(M$, 0, "Archivo de Pedidos")
                End If
            End If
            
            If Val(Repuesto4.Text) <> 0 Then
                .Seek "=", Val(Repuesto4.Text)
                If .NoMatch = True Then
                    WPasa = "N"
                    M$ = "Codigo de Adicionales Incorrecto"
                    A% = MsgBox(M$, 0, "Archivo de Pedidos")
                End If
            End If
            
            If Val(Repuesto5.Text) <> 0 Then
                .Seek "=", Val(Repuesto5.Text)
                If .NoMatch = True Then
                    WPasa = "N"
                    M$ = "Codigo de Adicionales Incorrecto"
                    A% = MsgBox(M$, 0, "Archivo de Pedidos")
                End If
            End If
            
        End With
        
        If WPasa = "S" Then
        
        With rstEmpresa
        
            .Index = "Empresa"
            .Seek "=", 1
            If .NoMatch = False Then
                WNroEmpresa = !NroEmpresa
                    Else
                WNroEmpresa = 0
            End If
        End With
    
        With rstPedido
            .Index = "Pedido"
            .Seek "=", Val(Pedido.Text)
            If .NoMatch Then
                .AddNew
                Call Verifica_datos
                !Pedido = Pedido
                !Fecha = Fecha.Text
                !Cliente = Cliente.Text
                !Producto = Producto.Text
                !Pago = Pago.Text
                !Precio = CDbl(Precio.Text)
                !FEntrega = FEntrega.Text
                !Vendedor = WVendedor
                !Modelo = WModelo
                !Empresa = 1
                !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                !OrdFEntrega = Right$(FEntrega.Text, 4) + Mid$(FEntrega.Text, 4, 2) + Left$(FEntrega.Text, 2)
                !Interes = CDbl(Interes.Text)
                !Cuotas = CDbl(Cuotas.Text)
                !Descuento1 = CDbl(Descuento1.Text)
                !Descuento2 = CDbl(Descuento2.Text)
                !Descuento3 = CDbl(Descuento3.Text)
                !OCompra = OCompra.Text
                !Obser1 = Obser1.Text
                !Obser2 = Obser2.Text
                !Obser3 = Obser3.Text
                !Obser4 = Obser4.Text
                !Obser5 = Obser5.Text
                !Repuesto1 = Repuesto1.Text
                !Repuesto2 = Repuesto2.Text
                !Repuesto3 = Repuesto3.Text
                !Repuesto4 = Repuesto4.Text
                !Repuesto5 = Repuesto5.Text
                !Precio1 = Precio1.Text
                !Precio2 = Precio2.Text
                !Precio3 = Precio3.Text
                !Precio4 = Precio4.Text
                !Precio5 = Precio5.Text
                !Descri1 = DesRepuesto1.Caption
                !Descri2 = DesRepuesto2.Caption
                !Descri3 = DesRepuesto3.Caption
                !Descri4 = DesRepuesto4.Caption
                !Descri5 = DesRepuesto5.Caption
                !Ciudad = WCiudad
                !NroEmpresa = WNroEmpresa
                If Tipo1.Value = True Then
                    !Tipo = "1"
                End If
                If Tipo2.Value = True Then
                    !Tipo = "2"
                End If
                
                Rem !Km = CDbl(Km.text)
                .Update
                .Bookmark = .LastModified
                    Else
                .Edit
                Call Verifica_datos
                !Pedido = Pedido
                !Fecha = Fecha.Text
                !Cliente = Cliente.Text
                !Producto = Producto.Text
                !Pago = Pago.Text
                !Precio = CDbl(Precio.Text)
                !FEntrega = FEntrega.Text
                !Vendedor = WVendedor
                !Modelo = WModelo
                !Empresa = 1
                !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                !OrdFEntrega = Right$(FEntrega.Text, 4) + Mid$(FEntrega.Text, 4, 2) + Left$(FEntrega.Text, 2)
                !Interes = CDbl(Interes.Text)
                !Cuotas = CDbl(Cuotas.Text)
                !Descuento1 = CDbl(Descuento1.Text)
                !Descuento2 = CDbl(Descuento2.Text)
                !Descuento3 = CDbl(Descuento3.Text)
                !OCompra = OCompra.Text
                !Obser1 = Obser1.Text
                !Obser2 = Obser2.Text
                !Obser3 = Obser3.Text
                !Obser4 = Obser4.Text
                !Obser5 = Obser5.Text
                !Repuesto1 = Repuesto1.Text
                !Repuesto2 = Repuesto2.Text
                !Repuesto3 = Repuesto3.Text
                !Repuesto4 = Repuesto4.Text
                !Repuesto5 = Repuesto5.Text
                !Precio1 = Precio1.Text
                !Precio2 = Precio2.Text
                !Precio3 = Precio3.Text
                !Precio4 = Precio4.Text
                !Precio5 = Precio5.Text
                !Descri1 = DesRepuesto1.Caption
                !Descri2 = DesRepuesto2.Caption
                !Descri3 = DesRepuesto3.Caption
                !Descri4 = DesRepuesto4.Caption
                !Descri5 = DesRepuesto5.Caption
                !Ciudad = WCiudad
                !NroEmpresa = WNroEmpresa
                If Tipo1.Value = True Then
                    !Tipo = "1"
                End If
                If Tipo2.Value = True Then
                    !Tipo = "2"
                End If
                Rem !Km = CDbl(Km.text)
                .Update
                .Bookmark = .LastModified
            End If
        End With
        
        Call CmdLimpiar_Click
        
        End If
        
        Pedido.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()
    If Pedido.Text <> "" Then
        With rstPedido
            .Index = "Pedido"
            .Seek "=", Pedido.Text
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
    Pedido.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Pedido.Text = ""
    Fecha.Text = "  /  /    "
    Cliente.Text = ""
    Producto.Text = ""
    Pago.Text = ""
    Precio.Text = ""
    FEntrega.Text = "  /  /    "
    DesCliente.Caption = ""
    DesProducto.Caption = ""
    Cuotas.Text = ""
    Interes.Text = ""
    Descuento1.Text = ""
    Descuento2.Text = ""
    Descuento3.Text = ""
    OCompra.Text = ""
    Obser1.Text = ""
    Obser2.Text = ""
    Obser3.Text = ""
    Obser4.Text = ""
    Obser5.Text = ""
    Precio1.Text = ""
    Precio2.Text = ""
    Precio3.Text = ""
    Precio4.Text = ""
    Precio5.Text = ""
    Repuesto1.Text = ""
    Repuesto2.Text = ""
    Repuesto3.Text = ""
    Repuesto4.Text = ""
    Repuesto5.Text = ""
    DesRepuesto1.Caption = ""
    DesRepuesto2.Caption = ""
    DesRepuesto3.Caption = ""
    DesRepuesto4.Caption = ""
    DesRepuesto5.Caption = ""
    Tipo1.Value = True
    Tipo2.Value = False
    Pedido.SetFocus
End Sub

Private Sub cmdClose_Click()
    CmdLimpiar_Click
    Pedido.SetFocus
    PrgPedido.Hide
    Menu.SetFocus
End Sub

Private Sub Anterior_Click()
    With rstPedido
        .Index = "Pedido"
        .Seek "=", Val(Pedido.Text)
        If .NoMatch = False Then
            .MovePrevious
            If .BOF = True Then
                .MoveFirst
                M$ = "No exsite registro Anterior"
                A% = MsgBox(M$, 0, "Archivo de Pedido")
                .MoveFirst
            End If
            Pedido.Text = !Pedido
            Call Imprime_Datos
            Pedido.SetFocus
        End If
    End With
End Sub


Private Sub Lista_Click()
    Desde.Text = "0"
    Hasta.Text = "999999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
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
        With rstCliente
            .Index = "Cliente"
            .Seek "=", Val(Cliente.Text)
            If .NoMatch = False Then
                DesCliente.Caption = !Razon
                WCiudad = !Ciudad
                WVendedor = !Vendedor
                Call Imprime_Descripcion
                Producto.SetFocus
                    Else
                Cliente.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstProducto
            .Index = "Producto"
            .Seek "=", Producto.Text
            If .NoMatch = False Then
                DesProducto.Caption = !Descripcion1
                WModelo = !Modelo
                Pago.SetFocus
                    Else
                Producto.SetFocus
            End If
        End With
    End If
End Sub

Private Sub Pago_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Interes.SetFocus
    End If
End Sub

Private Sub Interes_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Interes.text = PUsing("###.##", Interes.text)
        Cuotas.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
Private Sub Cuotas_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Cuotas.text = PUsing("###", Cuotas.text)
        Precio.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Precio_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Precio.text = PUsing("###,###.##", Precio.text)
        Descuento1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descuento1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Descuento1.text = PUsing("###.##", Descuento1.text)
        Descuento2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descuento2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Descuento2.text = PUsing("###.##", Descuento2.text)
        Descuento3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descuento3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Descuento3.text = PUsing("###.##", Descuento3.text)
        FEntrega.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FEntrega_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FEntrega.Text, Auxi)
        If Auxi = "S" Then
            OCompra.SetFocus
                Else
            FEntrega.SetFocus
        End If
    End If
End Sub

Private Sub OCompra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Obser1.SetFocus
    End If
End Sub

Private Sub Obser1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Obser2.SetFocus
    End If
End Sub

Private Sub Obser2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Obser3.SetFocus
    End If
End Sub

Private Sub Obser3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Obser4.SetFocus
    End If
End Sub

Private Sub Obser4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Obser5.SetFocus
    End If
End Sub

Private Sub Obser5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Repuesto1.SetFocus
    End If
End Sub

Private Sub Repuesto1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Repuesto1.Text) <> 0 Then
            With rstRepuesto
                .Index = "Repuesto"
                .Seek "=", Val(Repuesto1.Text)
                If .NoMatch = False Then
                    DesRepuesto1.Caption = !Descripcion
                    If Val(Precio1.Text) = 0 Then
                        Precio1.Text = !Precio
                        Call Format_datos
                    End If
                    Precio1.SetFocus
                        Else
                    Repuesto1.SetFocus
                End If
            End With
                Else
            Precio1.SetFocus
        End If
    End If
End Sub

Private Sub precio1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Precio1.text = PUsing("###,###.##", Precio1.text)
        Repuesto2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub Repuesto2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Repuesto2.Text) <> 0 Then
            With rstRepuesto
                .Index = "Repuesto"
                .Seek "=", Val(Repuesto2.Text)
                If .NoMatch = False Then
                    DesRepuesto2.Caption = !Descripcion
                    If Val(Precio2.Text) = 0 Then
                        Precio2.Text = !Precio
                        Call Format_datos
                    End If
                    Precio2.SetFocus
                        Else
                    Repuesto2.SetFocus
                End If
            End With
                Else
            Precio2.SetFocus
        End If
    End If
End Sub

Private Sub precio2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Precio2.text = PUsing("###,###.##", Precio2.text)
        Repuesto3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Repuesto3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Repuesto3.Text) <> 0 Then
            With rstRepuesto
                .Index = "Repuesto"
                .Seek "=", Val(Repuesto3.Text)
                If .NoMatch = False Then
                    DesRepuesto3.Caption = !Descripcion
                    If Val(Precio3.Text) = 0 Then
                        Precio3.Text = !Precio
                        Call Format_datos
                    End If
                    Precio3.SetFocus
                        Else
                    Repuesto3.SetFocus
                End If
            End With
                Else
            Precio3.SetFocus
        End If
    End If
End Sub

Private Sub precio3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Precio3.text = PUsing("###,###.##", Precio3.text)
        Repuesto4.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Repuesto4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Repuesto4.Text) <> 0 Then
            With rstRepuesto
                .Index = "Repuesto"
                .Seek "=", Val(Repuesto4.Text)
                If .NoMatch = False Then
                    DesRepuesto4.Caption = !Descripcion
                    If Val(Precio4.Text) = 0 Then
                        Precio4.Text = !Precio
                        Call Format_datos
                    End If
                    Precio4.SetFocus
                        Else
                    Repuesto4.SetFocus
                End If
            End With
                Else
            Precio4.SetFocus
        End If
    End If
End Sub

Private Sub precio4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Precio4.text = PUsing("###,###.##", Precio4.text)
        Repuesto5.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Repuesto5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Repuesto5.Text) <> 0 Then
            With rstRepuesto
                .Index = "Repuesto"
                .Seek "=", Val(Repuesto5.Text)
                If .NoMatch = False Then
                    DesRepuesto5.Caption = !Descripcion
                    If Val(Precio5.Text) = 0 Then
                        Precio5.Text = !Precio
                        Call Format_datos
                    End If
                    Precio5.SetFocus
                        Else
                    Repuesto5.SetFocus
                End If
            End With
                Else
            Precio5.SetFocus
        End If
    End If
End Sub

Private Sub precio5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Precio5.text = PUsing("###,###.##", Precio5.text)
        Fecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Pedido.Text) <> 0 Then
            Existe = "N"
            WProducto = ""
            With rstPedido
                .Index = "Pedido"
                Claveven$ = Pedido.Text
                .Seek "=", Pedido.Text
                If .NoMatch Then
                    CmdLimpiar_Click
                    Pedido.Text = Claveven$
                        Else
                    Pedido.Text = !Pedido
                    Call Imprime_Datos
                End If
            End With
        End If
        Fecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False

     Opcion.Clear

     Opcion.AddItem "Pedidos"
     Opcion.AddItem "Clientes"
     Opcion.AddItem "Producto"
     Opcion.AddItem "Adicionales"

     Opcion.Visible = True
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            With rstPedido
                .Index = "Pedido"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Pedido + " " + !Producto
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Pedido
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
        Case 1
            With rstCliente
                .Index = "Razon"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(!Cliente) + " " + !Razon
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Cliente
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case 2
            With rstProducto
                .Index = "Producto"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Producto + " " + !Descripcion1
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Producto
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case 3
            With rstRepuesto
                .Index = "Repuesto"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(!Repuesto) + " " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Repuesto
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
    Select Case XIndice
        Case 0
            With rstPedido

                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                Pedido.Text = Claveven$
                .Index = "Pedido"
                Claveven$ = Pedido.Text
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    Pedido.Text = !Pedido
                    Call Imprime_Datos
                        Else
                    CmdLimpiar_Click
                    Pedido.Text = Claveven$
                End If
            End With
            Pedido.SetFocus
            
        Case 1
            With rstCliente

                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                Cliente.Text = Claveven$
                .Index = "Cliente"
                Claveven$ = Cliente.Text
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    Cliente.Text = !Cliente
                    Call Imprime_Descripcion
                        Else
                    CmdLimpiar_Click
                    Cliente.Text = Claveven$
                End If
            End With
            Cliente.SetFocus
            
        Case 2
            With rstProducto

                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                Producto.Text = Claveven$
                .Index = "Producto"
                Claveven$ = Producto.Text
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    Producto.Text = !Producto
                    Call Imprime_Descripcion
                        Else
                    CmdLimpiar_Click
                    Producto.Text = Claveven$
                End If
            End With
            Producto.SetFocus
            
        Case 3
            With rstRepuesto

                lugar = 1
                
                If Val(Repuesto1.Text) = 0 Then
                    lugar = 1
                        Else
                    If Val(Repuesto2.Text) = 0 Then
                        lugar = 2
                            Else
                        If Val(Repuesto3.Text) = 0 Then
                            lugar = 3
                                Else
                            If Val(Repuesto4.Text) = 0 Then
                                lugar = 4
                                    Else
                                If Val(Repuesto5.Text) = 0 Then
                                    lugar = 5
                                End If
                            End If
                        End If
                    End If
                End If

                Select Case lugar
                    Case 1
                        Indice = Pantalla.ListIndex
                        Claveven$ = WIndice.List(Indice)
                        Repuesto1.Text = Claveven$
                        .Index = "Repuesto"
                        Claveven$ = Repuesto1.Text
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            Repuesto1.Text = !Repuesto
                            If Val(Precio1.Text) = 0 Then
                                Precio1.Text = !Precio
                                Call Format_datos
                            End If
                            Call Imprime_Descripcion
                                Else
                            CmdLimpiar_Click
                            Repuesto1.Text = Claveven$
                        End If
                        Repuesto1.SetFocus
                    
                    Case 2
                        Indice = Pantalla.ListIndex
                        Claveven$ = WIndice.List(Indice)
                        Repuesto2.Text = Claveven$
                        .Index = "Repuesto"
                        Claveven$ = Repuesto2.Text
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            Repuesto2.Text = !Repuesto
                            If Val(Precio2.Text) = 0 Then
                                Precio2.Text = !Precio
                                Call Format_datos
                            End If
                            Call Imprime_Descripcion
                                Else
                            CmdLimpiar_Click
                            Repuesto2.Text = Claveven$
                        End If
                        Repuesto2.SetFocus
                    
                    Case 3
                        Indice = Pantalla.ListIndex
                        Claveven$ = WIndice.List(Indice)
                        Repuesto3.Text = Claveven$
                        .Index = "Repuesto"
                        Claveven$ = Repuesto3.Text
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            Repuesto3.Text = !Repuesto
                            If Val(Precio3.Text) = 0 Then
                                Precio3.Text = !Precio
                                Call Format_datos
                            End If
                            Call Imprime_Descripcion
                                Else
                            CmdLimpiar_Click
                            Repuesto3.Text = Claveven$
                        End If
                        Repuesto3.SetFocus
                    
                    Case 4
                        Indice = Pantalla.ListIndex
                        Claveven$ = WIndice.List(Indice)
                        Repuesto4.Text = Claveven$
                        .Index = "Repuesto"
                        Claveven$ = Repuesto4.Text
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            Repuesto4.Text = !Repuesto
                            If Val(Precio4.Text) = 0 Then
                                Precio4.Text = !Precio
                                Call Format_datos
                            End If
                            Call Imprime_Descripcion
                                Else
                            CmdLimpiar_Click
                            Repuesto4.Text = Claveven$
                        End If
                        Repuesto4.SetFocus
                                        
                    Case 5
                        Indice = Pantalla.ListIndex
                        Claveven$ = WIndice.List(Indice)
                        Repuesto5.Text = Claveven$
                        .Index = "Repuesto"
                        Claveven$ = Repuesto5.Text
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            Repuesto5.Text = !Repuesto
                            If Val(Precio5.Text) = 0 Then
                                Precio5.Text = !Precio
                                Call Format_datos
                            End If
                            Call Imprime_Descripcion
                                Else
                            CmdLimpiar_Click
                            Repuesto5.Text = Claveven$
                        End If
                        Repuesto5.SetFocus
                        
                    Case Else
                End Select
            End With
            
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()
    Rem On Error GoTo Error_primer
    With rstPedido
        .Index = "Pedido"
        .MoveFirst
        Pedido.Text = !Pedido
        Call Imprime_Datos
        Pedido.SetFocus
    End With
    Exit Sub

Error_primer:
     coderr = Err
     Call Errores(coderr, "Pedido", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Pedido.SetFocus
 End Sub

Private Sub Ultimo_Click()
    Rem On Error GoTo Error_ultimo
    With rstPedido
        .Index = "Pedido"
        .MoveLast
        Pedido.Text = !Pedido
        Call Imprime_Datos
        Pedido.SetFocus
    End With
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Pedido", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Pedido.SetFocus
 End Sub

Private Sub Siguiente_Click()
    With rstPedido
        .Index = "Pedido"
        Claveven$ = Pedido.Text
        .Seek "=", Val(Claveven$)
        If .NoMatch = False Then
            .MoveNext
            If .EOF = True Then
                M$ = "No exsite registro Posterior"
                A% = MsgBox(M$, 0, "Archivo de Pedido")
                Call Ultimo_Click
            End If
            Pedido.Text = !Pedido
            Call Imprime_Datos
            Pedido.SetFocus
        End If
    End With
End Sub

Sub Form_load()
    Pedido.Text = ""
    Fecha.Text = "  /  /    "
    Cliente.Text = ""
    Producto.Text = ""
    Pago.Text = ""
    Precio.Text = ""
    FEntrega.Text = "  /  /    "
    DesCliente.Caption = ""
    DesProducto.Caption = ""
    Cuotas.Text = ""
    Interes.Text = ""
    Descuento1.Text = ""
    Descuento2.Text = ""
    Descuento3.Text = ""
    OCompra.Text = ""
    Obser1.Text = ""
    Obser2.Text = ""
    Obser3.Text = ""
    Obser4.Text = ""
    Obser5.Text = ""
    Precio1.Text = ""
    Precio2.Text = ""
    Precio3.Text = ""
    Precio4.Text = ""
    Precio5.Text = ""
    Repuesto1.Text = ""
    Repuesto2.Text = ""
    Repuesto3.Text = ""
    Repuesto4.Text = ""
    Repuesto5.Text = ""
    DesRepuesto1.Caption = ""
    DesRepuesto2.Caption = ""
    DesRepuesto3.Caption = ""
    DesRepuesto4.Caption = ""
    DesRepuesto5.Caption = ""
    Tipo1.Value = True
    Tipo2.Value = False
    
    
End Sub
