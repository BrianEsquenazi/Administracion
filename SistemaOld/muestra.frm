VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMuestra 
   Caption         =   "Solicitud de Muestras para Clientes"
   ClientHeight    =   8415
   ClientLeft      =   1665
   ClientTop       =   405
   ClientWidth     =   8580
   LinkTopic       =   "Form2"
   ScaleHeight     =   8415
   ScaleWidth      =   8580
   Begin VB.TextBox Razon 
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
      Left            =   3720
      MaxLength       =   50
      TabIndex        =   30
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox Ayuda 
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
      TabIndex        =   29
      Top             =   4440
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.TextBox Ensayo 
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
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   24
      Text            =   " "
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Vendedor 
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
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   22
      Text            =   " "
      Top             =   2640
      Width           =   1335
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
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   20
      Text            =   " "
      Top             =   1920
      Width           =   1335
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   18
      Text            =   " "
      Top             =   3000
      Width           =   6015
   End
   Begin VB.TextBox DescriCliente 
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   16
      Text            =   " "
      Top             =   2280
      Width           =   6015
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.TextBox Cantidad 
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
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   13
      Text            =   " "
      Top             =   1560
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
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
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   1080
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7320
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      ItemData        =   "muestra.frx":0000
      Left            =   120
      List            =   "muestra.frx":0007
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "   Consulta        Datos           (F3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4680
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "    Limpia         Pantalla          (F2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3360
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "    Fin de         Ingreso          (F10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   6000
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdGraba 
      Caption         =   "    Graba            (F1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2040
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin MSMask.MaskEdBox Articulo 
      Height          =   285
      Left            =   2280
      TabIndex        =   26
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
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
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin VB.Label DesArticulo 
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
      Left            =   3720
      TabIndex        =   28
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label8 
      Caption         =   "Codigo de M.Prima"
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
      TabIndex        =   27
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Codigo de Ensayo"
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
      TabIndex        =   25
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label DesVendedor 
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
      Left            =   3720
      TabIndex        =   23
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Label Label4 
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
      TabIndex        =   21
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Vendedor"
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
      TabIndex        =   17
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label15 
      Caption         =   "Fecha de Solicitud"
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
      TabIndex        =   14
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Nombre para el Cliente"
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
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Cantidad Solicitada"
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
      TabIndex        =   11
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo de Producto"
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
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label DesProducto 
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
      Left            =   3720
      TabIndex        =   8
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   7
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "PrgMuestra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstVendedor As Recordset
Dim spVendedor As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstMuestra As Recordset
Dim spMuestra As String
Dim XParam As String
Dim EmpresaActual As String
Dim XIndice As Integer

Private Sub cmdGraba_Click()

    WNombre = ""
    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WNombre = rstTerminado!Descripcion
        rstTerminado.Close
    End If
    
    spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        WNombre = rstArticulo!Descripcion
        rstArticulo.Close
    End If
    
    WAutoriza = "X"
    WImpresion = "X"

    If Val(WMuestra) <> 0 Then
    
        WFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        
        Sql1 = "UPDATE Muestra SET "
        Sql2 = "Codigo =  " + "'" + WMuestra + "',"
        Sql3 = "Producto =  " + "'" + Producto.Text + "',"
        Sql4 = "Articulo =  " + "'" + Articulo.Text + "',"
        Sql5 = "Ensayo =  " + "'" + Ensayo.Text + "',"
        Sql6 = "Nombre =  " + "'" + WNombre + "',"
        Sql7 = "Fecha =  " + "'" + Fecha.Text + "',"
        Sql8 = "OrdFecha =  " + "'" + WFechaOrd + "',"
        Sql9 = "Cantidad =  " + "'" + Cantidad.Text + "',"
        Sql10 = "Cliente =  " + "'" + Cliente.Text + "',"
        Sql11 = "Razon =  " + "'" + Razon.Text + "',"
        Sql12 = "DescriCliente =  " + "'" + DescriCliente.Text + "',"
        Sql13 = "Vendedor =  " + "'" + Vendedor.Text + "',"
        Sql14 = "DesVendedor =  " + "'" + DesVendedor.Caption + "',"
        Sql15 = "Observaciones =  " + "'" + Observaciones.Text + "',"
        Sql16 = "Autoriza =  " + "'" + WAutoriza + "',"
        Sql17 = "Impresion =  " + "'" + WImpresion + "'"
        Sql18 = " Where Codigo = " + "'" + WMuestra + "'"
        spMuestra = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                    Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18
        Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
        
        Call cmdClose_Click

            Else

        WCodigo = 1
        spMuestra = "ListaMuestraNumero"
        Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
        If rstMuestra.RecordCount > 0 Then
            With rstMuestra
                .MoveLast
                WCodigo = rstMuestra!Codigo + 1
            End With
            rstMuestra.Close
        End If
    
        XCodigo = Str$(WCodigo)
        WFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        
        Sql1 = "INSERT INTO Muestra ("
        Sql2 = "Codigo ,"
        Sql3 = "Producto ,"
        Sql4 = "Articulo ,"
        Sql5 = "Ensayo ,"
        Sql6 = "Nombre ,"
        Sql7 = "Fecha ,"
        Sql8 = "OrdFecha ,"
        Sql9 = "Cantidad ,"
        Sql10 = "Cliente ,"
        Sql11 = "Razon ,"
        Sql12 = "DescriCliente ,"
        Sql13 = "Vendedor ,"
        Sql14 = "DesVendedor ,"
        Sql15 = "Observaciones ,"
        Sql16 = "Autoriza ,"
        Sql17 = "Impresion) "
        Sql18 = "Values ("
        Sql19 = "'" + XCodigo + "',"
        Sql20 = "'" + Producto.Text + "',"
        Sql21 = "'" + Articulo.Text + "',"
        Sql22 = "'" + Ensayo.Text + "',"
        Sql23 = "'" + WNombre + "',"
        Sql24 = "'" + Fecha.Text + "',"
        Sql25 = "'" + WFechaOrd + "',"
        Sql26 = "'" + Cantidad.Text + "',"
        Sql27 = "'" + Cliente.Text + "',"
        Sql28 = "'" + Razon.Text + "',"
        Sql29 = "'" + DescriCliente.Text + "',"
        Sql30 = "'" + Vendedor.Text + "',"
        Sql31 = "'" + DesVendedor.Caption + "',"
        Sql32 = "'" + Observaciones.Text + "',"
        Sql33 = "'" + WAutoriza + "',"
        Sql34 = "'" + WImpresion + "')"
      
        spMuestra = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                    Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                    Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                    Sql31 + Sql32 + Sql33 + Sql34
        Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
    
        Call CmdLimpiar_Click
        Producto.SetFocus
    End If
        
End Sub

Private Sub CmdLimpiar_Click()

    Producto.Text = "  -     -   "
    DesProducto.Caption = ""
    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
    Ensayo.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cantidad.Text = ""
    Rem Cliente.Text = ""
    Rem Razon.Text = ""
    DescriCliente.Text = ""
    Rem Vendedor.Text = ""
    Rem DesVendedor.Caption = ""
    Rem Observaciones.Text = ""
    Producto.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgMuestra.Hide
    Unload Me
    PrgAju.Show
End Sub

Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "  -     -   " Then
        
            Producto.Text = UCase(Producto.Text)
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                DesProducto.Caption = rstTerminado!Descripcion
                If DescriCliente.Text = "" Then
                    DescriCliente.Text = DesProducto.Caption
                End If
                rstTerminado.Close
                Fecha.SetFocus
                    Else
                Producto.SetFocus
            End If
            
                Else
                
            Articulo.SetFocus
            
        End If
    End If
    If KeyAscii = 27 Then
        Producto.Text = "  -     -   "
        DesProducto.Caption = ""
    End If
End Sub

Sub Articulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Articulo.Text <> "  -   -   " Then
        
            Articulo.Text = UCase(Articulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                DesArticulo.Caption = rstArticulo!Descripcion
                If DescriCliente.Text = "" Then
                    DescriCliente.Text = DesArticulo.Caption
                End If
                rstArticulo.Close
                Fecha.SetFocus
                    Else
                Articulo.SetFocus
            End If
            
                Else
                
            Ensayo.SetFocus
            
        End If
    End If
    If KeyAscii = 27 Then
        Articulo.Text = "  -   -   "
        DesArticulo.Caption = ""
    End If
End Sub

Private Sub Ensayo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Ensayo.Text = ""
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Cantidad.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.SetFocus
    End If
    If KeyAscii = 27 Then
        Cantidad.Text = ""
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cliente.Text <> "" And Cliente.Text <> Space$(6) Then
            Cliente.Text = UCase(Cliente.Text)
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Razon.Text = rstCliente!Razon
                rstCliente.Close
                DescriCliente.SetFocus
                    Else
                Cliente.SetFocus
            End If
                Else
            Razon.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
    End If
End Sub

Private Sub Razon_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescriCliente.SetFocus
    End If
    If KeyAscii = 27 Then
        Razon.Text = ""
    End If
End Sub

Private Sub DescriCliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Vendedor.SetFocus
    End If
    If KeyAscii = 27 Then
        DescriCliente.Text = ""
    End If
End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spVendedor = "ConsultaVendedor " + "'" + Vendedor.Text + "'"
        Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstVendedor.RecordCount > 0 Then
            DesVendedor.Caption = rstVendedor!Nombre
            rstVendedor.Close
            Observaciones.SetFocus
                Else
            Vendedor.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vendedor.Text = ""
        DesVendedor.Caption = ""
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Producto.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub Consulta_Click()
    Opcion.Clear
    
    Opcion.AddItem "Productos"
    Opcion.AddItem "Materias Primas"
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Vendedores"
    
    Opcion.Visible = True
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
    Dim IngresaItem As String

    pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            spTerminado = "ListaTerminadoConsulta"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
            With rstTerminado
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Left$(rstTerminado!Codigo, 2) = "PT" Then
                            IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                            pantalla.AddItem IngresaItem
                            IngresaItem = rstTerminado!Codigo
                            WIndice.AddItem IngresaItem
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstTerminado.Close
            End If
            
        Case 1
            spArticulo = "ListaArticuloConsulta"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                        pantalla.AddItem IngresaItem
                        IngresaItem = rstArticulo!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstArticulo.Close
            End If
            
            
        Case 2
            spClientes = "ListaClienteConsulta"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                With rstClientes
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstClientes!Cliente + " " + rstClientes!Razon
                            pantalla.AddItem IngresaItem
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
            
        Case 3
            spVendedor = "ListaVendedor"
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                With rstVendedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstVendedor!Vendedor) + " " + rstVendedor!Nombre
                            pantalla.AddItem IngresaItem
                            IngresaItem = rstVendedor!Vendedor
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstVendedor.Close
            End If
        
        Case Else
    End Select
            
    pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus

End Sub

Private Sub Pantalla_Click()
    pantalla.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = pantalla.ListIndex
            ClavePro$ = WIndice.List(Indice)
            If Left$(ClavePro$, 2) = "PT" Then
                spTerminado = "ConsultaTerminado " + "'" + ClavePro$ + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    Producto.Text = rstTerminado!Codigo
                    DesProducto.Caption = rstTerminado!Descripcion
                    If DescriCliente.Text = "" Then
                        DescriCliente.Text = DesProducto.Caption
                    End If
                    rstTerminado.Close
                        Else
                    Producto.Text = "  -     -   "
                    DesProducto.Caption = ""
                End If
                Producto.SetFocus
                    Else
                spArticulo = "ConsultaArticulo " + "'" + Left$(ClavePro$, 3) + Right$(ClavePro$, 7) + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Producto.Text = Left$(rstArticulo!Codigo, 3) + "00" + Right$(rstArticulo!Codigo, 7)
                    DesProducto.Caption = rstArticulo!Descripcion
                    If DescriCliente.Text = "" Then
                        DescriCliente.Text = DesProducto.Caption
                    End If
                    rstArticulo.Close
                        Else
                    Producto.Text = "  -   -   "
                    DesProducto.Caption = ""
                End If
                Producto.SetFocus
            End If
            
        Case 1
            Indice = pantalla.ListIndex
            ClavePro$ = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + Left$(ClavePro$, 3) + Right$(ClavePro$, 7) + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Articulo.Text = rstArticulo!Codigo
                DesArticulo.Caption = rstArticulo!Descripcion
                If DescriCliente.Text = "" Then
                    DescriCliente.Text = DesArticulo.Caption
                End If
                rstArticulo.Close
                    Else
                Articulo.Text = "  -   -   "
                DesArticulo.Caption = ""
            End If
            Articulo.SetFocus
            
        Case 2
            Indice = pantalla.ListIndex
            Cliente.Text = WIndice.List(Indice)
            Call Cliente_KeyPress(13)
            
        Case 3
            Indice = pantalla.ListIndex
            Vendedor.Text = WIndice.List(Indice)
            Call Vendedor_KeyPress(13)
        
        Case Else
    End Select
    
End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
    Opcion.Visible = False
    Dim IngresaItem As String

    pantalla.Clear
    WIndice.Clear

    WEspacios = Len(Ayuda.Text)
    
    Select Case XIndice
        Case 0
            spTerminado = "ListaTerminadoConsulta"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
            With rstTerminado
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Left$(rstTerminado!Codigo, 2) = "PT" Then
                            da = Len(rstTerminado!Descripcion) - WEspacios
                            For Aaa = 1 To da
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstTerminado!Descripcion, Aaa, WEspacios) Then
                                    IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                                    pantalla.AddItem IngresaItem
                                    IngresaItem = rstTerminado!Codigo
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next Aaa
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstTerminado.Close
            End If
    
        Case 1
            spArticulo = "ListaArticuloConsulta"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
    
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            da = Len(rstArticulo!Descripcion) - WEspacios
                            For Aaa = 1 To da
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstArticulo!Descripcion, Aaa, WEspacios) Then
                                    IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                                    pantalla.AddItem IngresaItem
                                    IngresaItem = rstArticulo!Codigo
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next Aaa
                            .MoveNext
                    
                                    Else
                        
                            Exit Do
                
                        End If
                    Loop
                End With
    
                rstArticulo.Close
            End If
            
        Case 2
            spClientes = "ListaClienteConsulta"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                With rstClientes
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstClientes!Razon) - WEspacios
                            For Aaa = 1 To da
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstClientes!Razon, Aaa, WEspacios) Then
                                    IngresaItem = rstClientes!Cliente + " " + rstClientes!Razon
                                    pantalla.AddItem IngresaItem
                                    IngresaItem = rstClientes!Cliente
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next Aaa
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstClientes.Close
            End If
            
        Case 3
            spVendedor = "ListaVendedor"
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                With rstVendedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstVendedor!Nombre) - WEspacios
                            For Aaa = 1 To da
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstVendedor!Nombre, Aaa, WEspacios) Then
                                    IngresaItem = Str$(rstVendedor!Vendedor) + " " + rstVendedor!Nombre
                                    pantalla.AddItem IngresaItem
                                    IngresaItem = rstVendedor!Vendedor
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next Aaa
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstVendedor.Close
            End If
    
        Case Else
    End Select
    
    End If

End Sub

Private Sub Form_Load()
    Producto.Text = "  -     -   "
    DesProducto.Caption = ""
    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
    Ensayo.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cantidad.Text = ""
    Cliente.Text = ""
    Razon.Text = ""
    DescriCliente.Text = ""
    Vendedor.Text = ""
    DesVendedor.Caption = ""
    Observaciones.Text = ""
    
    
    If Val(WMuestra) <> 0 Then
        spMuestra = "ConsultaMuestra " + "'" + WMuestra + "'"
        Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
        If rstMuestra.RecordCount > 0 Then
            Producto.Text = rstMuestra!Producto
            Articulo.Text = rstMuestra!Articulo
            Ensayo.Text = rstMuestra!Ensayo
            Cantidad.Text = rstMuestra!Cantidad
            Cliente.Text = rstMuestra!Cliente
            Razon.Text = rstMuestra!Razon
            DescriCliente.Text = rstMuestra!DescriCliente
            Vendedor.Text = rstMuestra!Vendedor
            Observaciones.Text = rstMuestra!Observaciones
            rstMuestra.Close
        End If
        
        spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesProducto.Caption = rstTerminado!Descripcion
            rstTerminado.Close
        End If
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DesArticulo.Caption = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Razon.Text = rstCliente!Razon
            rstCliente.Close
        End If
        
        spVendedor = "ConsultaVendedor " + "'" + Vendedor.Text + "'"
        Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstVendedor.RecordCount > 0 Then
            DesVendedor.Caption = rstVendedor!Nombre
            rstVendedor.Close
        End If
        
    End If
        
End Sub

Private Sub Producto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Articulo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ensayo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Cantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Razon_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub DescriCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Observaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call cmdGraba_Click
        Case 113
            Call CmdLimpiar_Click
        Case 114
            Call Consulta_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub






