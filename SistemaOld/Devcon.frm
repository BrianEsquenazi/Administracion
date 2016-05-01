VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDevcon 
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
   Begin VB.TextBox Remito 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   30
      Top             =   1200
      Width           =   1095
   End
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
      TabIndex        =   28
      Top             =   6360
      Visible         =   0   'False
      Width           =   6975
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
      Left            =   4440
      TabIndex        =   27
      Text            =   " "
      Top             =   7080
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.CommandButton BorraConsulta 
      Caption         =   "Borra Consulta"
      Height          =   495
      Left            =   8280
      TabIndex        =   26
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton ConsultaPro 
      Caption         =   "Consulta Producto"
      Height          =   495
      Left            =   8280
      TabIndex        =   25
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton ConsultaCli 
      Caption         =   "Consulta Cliente"
      Height          =   495
      Left            =   8280
      TabIndex        =   24
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Baja 
      Caption         =   "  Baja  Devol."
      Height          =   495
      Left            =   9240
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   975
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
      Left            =   10320
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
      Left            =   3840
      TabIndex        =   21
      Top             =   6600
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
      Left            =   9240
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      Height          =   500
      Left            =   10320
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   500
      Left            =   9240
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
      Width           =   9975
      Begin VB.TextBox WLote 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8640
         MaxLength       =   6
         TabIndex        =   31
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
      Left            =   10320
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3975
      Left            =   240
      OleObjectBlob   =   "Devcon.frx":0000
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
   Begin VB.Label Label4 
      Caption         =   "Remito"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1200
      Width           =   1575
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
      Caption         =   "Numero de Devolucion"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgDevcon"
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
Dim rstDevcon As Recordset
Dim spDevcon As String
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
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String

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

    spDevcon = "ListaDevcon " + "'" + Numero.Text + "'"
    Set rstDevcon = db.OpenRecordset(spDevcon, dbOpenSnapshot, dbSQLPassThrough)

    If rstDevcon.RecordCount > 0 Then
            With rstDevcon
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    
                    Auxiliar(WRenglon, 1) = rstDevcon!Terminado
                    Auxiliar(WRenglon, 2) = rstDevcon!Cantidad
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstDevcon.Close
    End If
    
    For da = 1 To WRenglon
    
        Terminado = Auxiliar(da, 1)
        Cantidad = Auxiliar(da, 2)
        
        XParam = "'" + Remito.Text + "','" _
                + Terminado + "'"
                
        spConsig = "ListaConsigFactura " + XParam
        Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
        If rstConsig.RecordCount > 0 Then
            WClave = rstConsig!Clave
            WFacturado = Str$(rstConsig!Facturado - Cantidad)
            rstConsig.Close
                
            XParam = "'" + WClave + "','" _
                    + WFacturado + "'"
                                           
            spConsig = "ModificaConsigFacturado " + XParam
            Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Next da
    
    spDevcon = "BorrarDevcon " + "'" + Numero.Text + "'"
    Set rstDevcon = db.OpenRecordset(spDevcon, dbOpenDynaset, dbSQLPassThrough)
    
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
    WLinea.Text = ""
    WLote.Text = ""

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
    PrgDevcon.Hide
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
    
    WInicio = DBGrid1.FirstRow
    WArticulo.SetFocus

End Sub

Private Sub Graba_Click()

    spConsig = "ListaConsig " + "'" + Remito.Text + "'"
    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
    If rstConsig.RecordCount = 0 Then
        m$ = "No Existe el Remito de mercaderia en Consignacion Especificado"
        A% = MsgBox(m$, 0, "MODULO DE DEVOLUCION DE MERCADERIA")
        Exit Sub
            Else
        If Cliente.Text <> rstConsig!Cliente Then
            m$ = "No coincide el cliente informado con el especificado en el remito"
            A% = MsgBox(m$, 0, "MODULO DE DEVOLUCION DE MERCADERIA")
            Exit Sub
        End If
        rstConsig.Close
    End If

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""

    DBGrid1.Refresh
        
    For A = 0 To 3
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Terminado = UCase(DBGrid1.Text)
                    
            DBGrid1.Col = 2
            Cantidad = Val(DBGrid1.Text)
            
            If Cantidad <> 0 Then
            
                XParam = "'" + Remito.Text + "','" _
                            + Terminado + "'"
                spConsig = "ListaConsigFactura " + XParam
                Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
                If rstConsig.RecordCount > 0 Then
                    WSaldo = rstConsig!Cantidad - rstConsig!Facturado
                    If Cantidad > WSaldo Then
                        m$ = "Cantidad insuficiente en consignacion Articulo " + Terminado + " Saldo : " + Str$(WSaldo)
                        A% = MsgBox(m$, 0, "MODULO DE DEVOLUCION DE MERCADERIA")
                        Exit Sub
                    End If
                    rstConsig.Close
                        Else
                    m$ = "No existe este producto en consignacion Articulo " + Terminado
                    A% = MsgBox(m$, 0, "MODULO DE DEVOLUCION DE MERCEDERIA")
                    Exit Sub
                
                End If
                
            End If
                                        
        Next iRow
            
    Next A


    Erase Auxiliar
    WRenglon = 0

    spDevcon = "ListaDevcon " + "'" + Numero.Text + "'"
    Set rstDevcon = db.OpenRecordset(spDevcon, dbOpenSnapshot, dbSQLPassThrough)

    If rstDevcon.RecordCount > 0 Then
    
            With rstDevcon
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    
                    Auxiliar(WRenglon, 1) = rstDevcon!Terminado
                    Auxiliar(WRenglon, 2) = rstDevcon!Cantidad
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstDevcon.Close
    End If
    
    For da = 1 To WRenglon
    
        Terminado = Auxiliar(da, 1)
        Cantidad = Auxiliar(da, 2)
        
        XParam = "'" + Remito.Text + "','" _
                + Terminado + "'"
                
        spConsig = "ListaConsigFactura " + XParam
        Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
        If rstConsig.RecordCount > 0 Then
            WClave = rstConsig!Clave
            WFacturado = Str$(rstConsig!Facturado - Cantidad)
            rstConsig.Close
                
            XParam = "'" + WClave + "','" _
                    + WFacturado + "'"
                                           
            spConsig = "ModificaConsigFacturado " + XParam
            Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Next da
    
    spDevcon = "BorrarDevcon " + "'" + Numero.Text + "'"
    Set rstDevcon = db.OpenRecordset(spDevcon, dbOpenDynaset, dbSQLPassThrough)
    
    Erase Auxiliar
    Renglon = 0
    WRenglon = 0
    DBGrid1.Refresh
        
    For A = 0 To 3
        
        Suma = A * 10
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
                WFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WPrecio = Precio
                WLinea = Linea
                WImporte = ""
                WRemito = Remito.Text
                WClave = Auxi1 + Auxi
                WLote = Lote
                
                XParam = "'" + WClave + "','" _
                         + WNumero + "','" + WRenglon + "','" _
                         + WCliente + "','" + WFecha + "','" _
                         + WObservaciones + "','" _
                         + WTerminado + "','" _
                         + WCantidad + "','" _
                         + WFechaOrd + "','" _
                         + WPrecio + "','" _
                         + WLinea + "','" _
                         + WImporte + "','" _
                         + WRemito + "','" _
                         + WLote + "'"
                         
                spDevcon = "AltaDevcon " + XParam
                Set rstDevcon = db.OpenRecordset(spDevcon, dbOpenSnapshot, dbSQLPassThrough)
                        
            End If
                                        
        Next iRow
            
    Next A
    
    For da = 1 To WRenglon
    
        Articulo = Auxiliar(da, 1)
        Terminado = Auxiliar(da, 1)
        Cantidad = Val(Auxiliar(da, 2))
        Lote = Val(Auxiliar(da, 3))
        
        XParam = "'" + Remito.Text + "','" _
                + Terminado + "'"
                
        spConsig = "ListaConsigFactura " + XParam
        Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
        If rstConsig.RecordCount > 0 Then
            WClave = rstConsig!Clave
            WFacturado = Str$(rstConsig!Facturado + Cantidad)
            rstConsig.Close
                
            XParam = "'" + WClave + "','" _
                    + WFacturado + "'"
                                           
            spConsig = "ModificaConsigFacturado " + XParam
            Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        WControla = 0
        spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                WCodigo = Articulo
                rstTerminado.Close
                
                If WControla = 0 Then
                    XParam = "'" + Lote + "','" _
                                + Articulo + "'"
                    spHoja = "ListaHojaProducto " + XParam
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHoja.RecordCount > 0 Then
                        WClave = rstHoja!Clave
                        WSaldo = Str$(rstHoja!Saldo + Cantidad)
                        WDate = Date$
                        WMarca = ""
                        rstHoja.Close
                            
                        XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "','" _
                                    + WMarca + "'"
                        spHoja = "ModificaHojaSaldo2 " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            
                                    Else
                                
                        XParam = "'" + Articulo + "','" _
                                        + Lote + "'"
                        spMovguia = "ListaMovguiaLote1 " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            WClave = rstMovguia!Clave
                            WSaldo = Str$(rstMovguia!Saldo + Cantidad)
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
        
    Next da
    
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
    
    WArticulo.SetFocus
    
End Sub

Private Sub Limpia_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WLote.Text = ""
    
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Observaciones.Text = ""
    Remito.Text = ""
    
    Pantalla.Visible = False
    
    For A = 0 To 3
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 4
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next A
    
    Numero.Text = ""
    spDevcon = "ListaDevconNumero"
    Set rstDevcon = db.OpenRecordset(spDevcon, dbOpenSnapshot, dbSQLPassThrough)
    If rstDevcon.RecordCount > 0 Then
        With rstDevcon
            .MoveLast
            Numero.Text = rstDevcon!Numero + 1
        End With
        rstDevcon.Close
    End If
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Graba.Enabled = True


    Numero.SetFocus

End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WArticulo.Text = UCase(WArticulo.Text)
        
        spTerminado = "ConsultaTerminado " + "'" + WArticulo.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WDescripcion.Caption = rstTerminado!Descripcion
            rstTerminado.Close
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
                WEntra = "S"
                rstHoja.Close
            End If
                
            If WEntra = "N" Then
                XParam = "'" + WArticulo.Text + "','" _
                            + WLote.Text + "'"
                spMovguia = "ListaMovguiaLote1 " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    WEntra = "S"
                    rstMovguia.Close
                End If
            End If
                
                Else
                    
            WEntra = "S"
                
        End If
        
        If WEntra = "N" Then
            m$ = WArticulo.Text + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
            G% = MsgBox(m$, 0, "Nota de Credito por Devolucion")
                Else
            Call Alta_Vector
            Call Ingresa_Click
            WArticulo.SetFocus
        End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Pantalla_Click()
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
                    
                Call Alta_Vector
                WLinea.Text = WAnterior + 1
                If Val(WLinea.Text) > 0 Then
                    DBGrid1.Row = Val(WLinea.Text) - 1
                End If
                    
                Call DBGrid1.SetFocus
                WCantidad.SetFocus
                    
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
 
    Pantalla.Visible = False
    Numero.Text = ""
    spDevcon = "ListaDevconNumero"
    Set rstDevcon = db.OpenRecordset(spDevcon, dbOpenSnapshot, dbSQLPassThrough)
    If rstDevcon.RecordCount > 0 Then
        With rstDevcon
            .MoveLast
            Numero.Text = rstDevcon!Numero + 1
        End With
        rstDevcon.Close
    End If

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Graba.Enabled = True
    
    Remito.Text = ""
    
    Numero.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For A = 0 To 3
    Suma = A * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 4
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
    Renglon = 0
    Erase Auxiliar
    WRenglon = 0

    spDevcon = "ListaDevcon " + "'" + Numero.Text + "'"
    Set rstDevcon = db.OpenRecordset(spDevcon, dbOpenSnapshot, dbSQLPassThrough)

    If rstDevcon.RecordCount > 0 Then
            With rstDevcon
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        Renglon = Renglon + 1
                
                        Lugar1 = Int((Renglon - 1) / 10) * 10
                        Lugar2 = Renglon - Lugar1
                
                        DBGrid1.FirstRow = Lugar1
                        DBGrid1.Row = Lugar2 - 1
                
                        DBGrid1.Col = 0
                        DBGrid1.Text = rstDevcon!Terminado
                        Auxi1 = rstDevcon!Terminado
                
                        DBGrid1.Col = 2
                        DBGrid1.Text = Pusing("###,###.##", rstDevcon!Cantidad)
                        
                        DBGrid1.Col = 4
                        DBGrid1.Text = rstDevcon!Lote
                
                        WRenglon = WRenglon + 1
                    
                        Auxiliar(WRenglon, 1) = rstDevcon!Cliente
                        Auxiliar(WRenglon, 2) = rstDevcon!Terminado
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstDevcon.Close
    End If
    
    Renglon = 0
    
    For da = 1 To WRenglon
        Cliente = Auxiliar(da, 1)
        Terminado = Auxiliar(da, 2)
        
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
        
        WClave = Cliente.Text + Terminado
        
        spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
                DBGrid1.Col = 1
                DBGrid1.Text = rstPrecios!Descripcion
                DBGrid1.Col = 3
                DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio))
                rstPrecios.Close
        End If
        
    Next da

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
            
            DBGrid1.Row = DBGrid1.Row + 1
            
    End If

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)

    On Error GoTo WError

    If KeyAscii = 13 Then
    
        spDevcon = "ListaDevcon " + "'" + Numero.Text + "'"
        Set rstDevcon = db.OpenRecordset(spDevcon, dbOpenSnapshot, dbSQLPassThrough)
        If rstDevcon.RecordCount > 0 Then
            Fecha.Text = rstDevcon!Fecha
            Cliente.Text = rstDevcon!Cliente
            Observaciones.Text = rstDevcon!Observaciones
            Remito.Text = rstDevcon!Remito
            rstDevcon.Close
            
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
        Remito.SetFocus
    End If
End Sub

Private Sub Remito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Ingresa_Click
        WArticulo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
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
            
                    da = Len(rstCliente!Razon) - WEspacios
                
                    For aa = 1 To da
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

