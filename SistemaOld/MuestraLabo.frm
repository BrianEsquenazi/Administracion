VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMuestraLabo 
   Caption         =   "Actualizacion de Datos en Laboratorio"
   ClientHeight    =   7260
   ClientLeft      =   915
   ClientTop       =   480
   ClientWidth     =   9555
   LinkTopic       =   "Form2"
   ScaleHeight     =   7260
   ScaleWidth      =   9555
   Begin VB.ComboBox Planta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      TabIndex        =   27
      Top             =   1560
      Width           =   3015
   End
   Begin VB.ComboBox Actualiza 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   26
      Top             =   1200
      Width           =   2415
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
      TabIndex        =   24
      Top             =   2400
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
      TabIndex        =   23
      Top             =   2400
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
      TabIndex        =   22
      Top             =   2400
      Width           =   1215
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
      TabIndex        =   21
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Lote 
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
      Left            =   7920
      MaxLength       =   10
      TabIndex        =   19
      Text            =   " "
      Top             =   1560
      Width           =   1455
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
      TabIndex        =   18
      Top             =   3240
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
      TabIndex        =   13
      Text            =   " "
      Top             =   840
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
      TabIndex        =   11
      Text            =   " "
      Top             =   1920
      Width           =   6015
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   2280
      TabIndex        =   10
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
      MaxLength       =   15
      TabIndex        =   8
      Text            =   " "
      Top             =   1560
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
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
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7320
      TabIndex        =   2
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
      ItemData        =   "MuestraLabo.frx":0000
      Left            =   120
      List            =   "MuestraLabo.frx":0007
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   8175
   End
   Begin MSMask.MaskEdBox Articulo 
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Top             =   480
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
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      Caption         =   "Actualiza Stock"
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
      Left            =   3720
      TabIndex        =   25
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "Lote"
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
      Left            =   3720
      TabIndex        =   20
      Top             =   1560
      Width           =   1935
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
      Left            =   3960
      TabIndex        =   17
      Top             =   480
      Width           =   4335
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
      TabIndex        =   16
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
      TabIndex        =   14
      Top             =   840
      Width           =   1935
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
      TabIndex        =   12
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label15 
      Caption         =   "Fecha de Realizacion"
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
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Cantidad Entregada"
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   3
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "PrgMuestraLabo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstMuestra As Recordset
Dim spMuestra As String
Dim XParam As String
Dim EmpresaActual As String
Dim WActualiza As String
Dim WGraba As String
Dim WTipoMov As String
Dim XIndice As Integer


Private Sub cmdGraba_Click()

    If WGraba = "S" Then
    
        If Actualiza.ListIndex = 1 Then
        
            XEmpresa = WEmpresa
            If Val(WEmpresa) = 1 Then
                Select Case Planta.ListIndex
                    Case 1
                        WEmpresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 2
                        WEmpresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                End Select
            End If
            
            If Articulo.Text <> "  -   -   " Then
            
                XTipo = "M"
                WEntra = "N"
                WControla = 0
                Articulo.Text = UCase(Articulo.Text)
                spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                    rstArticulo.Close
                End If
            
                If WControla = 0 Then
                    WCanti = 0
                    XParam = "'" + Lote.Text + "','" _
                                + Articulo.Text + "'"
                    spLaudo = "ListaLaudoArticulo " + XParam
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        WCanti = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                        WEntra = "S"
                        rstLaudo.Close
                    End If
                
                    If WEntra = "N" Then
                        XParam = "'" + Articulo.Text + "','" _
                                + Lote.Text + "'"
                        spMovguia = "ListaMovguiaLote " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            WCanti = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                            WEntra = "S"
                            rstMovguia.Close
                        End If
                    End If
                    
                        Else
                    
                    WEntra = "S"
                    
                End If
            
                If WEntra = "S" Then
                    If WCanti >= Val(Cantidad.Text) Then
                        Rem todo ok
                            Else
                        Call Conecta_Empresa
                        m$ = Articulo.Text + " Stock Insufucuente. Cantidad:" + Str$(WCanti)
                        G% = MsgBox(m$, 0, "Grabacion de Muestras de Stock")
                        Exit Sub
                    End If
                        Else
                    Call Conecta_Empresa
                    m$ = Articulo.Text + " Articulo inexistente o Lote nro. " + Lote.Text + " inexistente"
                    G% = MsgBox(m$, 0, "Grabacion de Muestras de Stock")
                    Exit Sub
                End If
                
            End If
                
            If Producto.Text <> "  -     -   " Then
        
                XTipo = "T"
                WEntra = "N"
            
                WControla = 0
                spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                    rstTerminado.Close
                End If
            
                If WControla = 0 Then
                    XParam = "'" + Lote.Text + "','" _
                            + Producto.Text + "'"
                    spHoja = "ListaHojaProducto " + XParam
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHoja.RecordCount > 0 Then
                        WCanti = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                        WEntra = "S"
                        rstHoja.Close
                    End If
                    
                    If WEntra = "N" Then
                        XParam = "'" + Producto.Text + "','" _
                                + Lote.Text + "'"
                        spMovguia = "ListaMovguiaLote1 " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            WCanti = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                            WEntra = "S"
                            rstMovguia.Close
                        End If
                    End If
                
                        Else
                    
                    WEntra = "S"
                
                End If
                
                If WEntra = "S" Then
                    If WCanti >= Val(Cantidad.Text) Then
                        Rem todo ok
                            Else
                        Call Conecta_Empresa
                        m$ = Producto.Text + " Stock Insufucuente. Cantidad:" + Str$(WCanti)
                        G% = MsgBox(m$, 0, "Grabacion de Muestras de Stock")
                        Exit Sub
                    End If
                        Else
                    Call Conecta_Empresa
                    m$ = Producto.Text + " Producto inexistente o Lote nro. " + Lote.Text + " inexistente"
                    G% = MsgBox(m$, 0, "Grabacion de Muestras de Stock")
                    Exit Sub
                End If
                
            End If
            
            Call Conecta_Empresa
            
        End If

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

        WFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WActualiza = Str$(Actualiza.ListIndex)
    
        XParam = "'" + WMuestra + "','" _
                 + Producto.Text + "','" _
                 + Articulo.Text + "','" _
                 + Ensayo.Text + "','" _
                 + WNombre + "','" _
                 + Fecha.Text + "','" _
                 + WFechaOrd + "','" _
                 + Cantidad.Text + "','" _
                 + Lote.Text + "','" _
                 + Observaciones.Text + "','" _
                 + WActualiza + "'"
                 
        Set rstMuestra = db.OpenRecordset("ModificaMuestraII " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        XEmpresa = WEmpresa
        If Val(WEmpresa) = 1 Then
            Select Case Planta.ListIndex
                Case 1
                    WEmpresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2
                    WEmpresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        End If
        
        If Actualiza.ListIndex = 1 Then
        
            WMovlab = ""
            
            spMovlab = "ListamovlabNumero"
            Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovlab.RecordCount > 0 Then
                With rstMovlab
                    .MoveLast
                    WMovlab = Str$(rstMovlab!Codigo + 1)
                End With
                rstMovlab.Close
            End If
        
            Renglon = 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = WMovlab
            Call Ceros(Auxi1, 6)
                
            WCodigo = WMovlab
            WRenglon = Str$(Renglon)
            WFecha = Fecha.Text
            WFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            WTipo = XTipo
            WArticulo = Articulo.Text
            WTerminado = Producto.Text
            WCantidad = Cantidad.Text
            WMovi = "S"
            WTipoMov = "1"
            Call Ceros(WTipoMov, 1)
            Wobservaciones = Observaciones.Text
            Wobservaciones = "Muestra de Laboratorio"
            WClave = Auxi1 + Auxi
            WDate = Date$
            WMarca = ""
            WLote = Lote.Text
                
            XParam = "'" + WClave + "','" _
                         + WCodigo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WTipo + "','" _
                         + WArticulo + "','" _
                         + WTerminado + "','" _
                         + WCantidad + "','" _
                         + WFechaOrd + "','" _
                         + WMovi + "','" _
                         + WTipoMov + "','" _
                         + Wobservaciones + "','" _
                         + WDate + "','" _
                         + WMarca + "','" _
                         + WLote + "'"
                         
            spMovlab = "Altamovlab " + XParam
            Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
            
            Sql1 = "UPDATE Muestra SET "
            Sql2 = " ClaveStock = " + "'" + WClave + "'"
            Sql3 = " Where Codigo = " + "'" + WMuestra + "'"
            spMuestra = Sql1 + Sql2 + Sql3
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            
            Select Case XTipo
                Case "M"
                    WControla = 0
                    spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
        
                        WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                        WCodigo = Articulo.Text
                        WSalidas = Str$(rstArticulo!Salidas + Val(Cantidad.Text))
                        WEntradas = Str$(rstArticulo!Entradas)
                        WDate = Date$
                        rstArticulo.Close
                
                        XParam = "'" + WCodigo + "','" _
                            + WEntradas + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
                        spArticulo = "ModificaArticuloMovimientos " + XParam
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                        If WControla = 0 And Val(Lote.Text) <> 0 Then
                            XParam = "'" + Lote.Text + "','" _
                                    + Articulo.Text + "'"
                            spLaudo = "ListaLaudoArticulo " + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                                WClave = rstLaudo!Clave
                                WSaldo = Str$(rstLaudo!Saldo - Val(Cantidad.Text))
                                WDate = Date$
                                rstLaudo.Close
                                
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spLaudo = "ModificaLaudoSaldo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                                    Else
                                    
                                XParam = "'" + Articulo.Text + "','" _
                                        + Lote.Text + "'"
                                spMovguia = "ListaMovguiaLote " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                If rstMovguia.RecordCount > 0 Then
                                    WClave = rstMovguia!Clave
                                    WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad.Text))
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
                
                
                Case "T"
                    WControla = 0
                    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
        
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        WCodigo = Producto.Text
                        WSalidas = Str$(rstTerminado!Salidas + Val(Cantidad.Text))
                        WEntradas = Str$(rstTerminado!Entradas)
                        WDate = Date$
                        rstTerminado.Close
                
                        XParam = "'" + WCodigo + "','" _
                                + WEntradas + "','" _
                                + WSalidas + "','" _
                                + WDate + "'"
                                           
                        spTerminado = "ModificaTerminadoMovimientos " + XParam
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
                        If WControla = 0 And Val(Lote.Text) <> 0 Then
                            XParam = "'" + Lote.Text + "','" _
                                    + Producto.Text + "'"
                            spHoja = "ListaHojaProducto " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                                WClave = rstHoja!Clave
                                WSaldo = Str$(rstHoja!Saldo - Val(Cantidad.Text))
                                WDate = Date$
                                rstHoja.Close
                            
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spHoja = "ModificaHojaSaldo " + XParam
                                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            
                                    Else
                                
                                XParam = "'" + Producto.Text + "','" _
                                    + Lote.Text + "'"
                                spMovguia = "ListaMovguiaLote1 " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                If rstMovguia.RecordCount > 0 Then
                                    WClave = rstMovguia!Clave
                                    WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad.Text))
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
            
                Case Else
            End Select
                
        End If
        
        Call Conecta_Empresa
        
        Call cmdClose_Click
    
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
    Lote.Text = ""
    Observaciones.Text = ""
    Actualiza.ListIndex = 1
    Planta.ListIndex = 0
    Producto.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgMuestraLabo.Hide
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
        Lote.SetFocus
    End If
    If KeyAscii = 27 Then
        Cantidad.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Lote_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
    If KeyAscii = 27 Then
        Lote.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
            spTerminado = "ConsultaTerminado " + "'" + ClavePro$ + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                Producto.Text = rstTerminado!Codigo
                DesProducto.Caption = rstTerminado!Descripcion
                rstTerminado.Close
                    Else
                Producto.Text = "  -     -   "
                DesProducto.Caption = ""
            End If
            Producto.SetFocus
            
        Case 1
            Indice = pantalla.ListIndex
            ClavePro$ = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + ClavePro$ + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Articulo.Text = rstArticulo!Codigo
                DesArticulo.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                    Else
                Articulo.Text = "  -   -   "
                DesArticulo.Caption = ""
            End If
            Articulo.SetFocus
        
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
        Case Else
    End Select
    
    End If

End Sub



Private Sub Form_Load()

    Actualiza.Clear
    
    Actualiza.AddItem "No Actualiza Stock"
    Actualiza.AddItem "Actualiza Stock"
    
    Planta.Clear
    
    Planta.AddItem "Planta I (CO/PG)"
    Planta.AddItem "Planta III (FA)"
    Planta.AddItem "Planta V (PT/BI)"
    
    Planta.ListIndex = 0

    Producto.Text = "  -     -   "
    DesProducto.Caption = ""
    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
    Ensayo.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cantidad.Text = ""
    Observaciones.Text = ""
    Lote.Text = ""
    Actualiza.ListIndex = 1
    
    spMuestra = "ConsultaMuestra " + "'" + WMuestra + "'"
    Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
    If rstMuestra.RecordCount > 0 Then
        WFecha = IIf(IsNull(rstMuestra!fecha2), "  /  /    ", rstMuestra!fecha2)
        If WFecha = Space(10) Then
            WFecha = "  /  /    "
        End If
        Fecha.Text = WFecha
        If Fecha.Text <> "  /  /    " Then
            WGraba = "N"
            WProducto = IIf(IsNull(rstMuestra!Producto2), "  -     -   ", rstMuestra!Producto2)
            WArticulo = IIf(IsNull(rstMuestra!Articulo2), "  -   -   ", rstMuestra!Articulo2)
            If WProducto <> "" Then
                Producto.Text = IIf(IsNull(rstMuestra!Producto2), "  -     -   ", rstMuestra!Producto2)
            End If
            If WArticulo <> "" Then
                Articulo.Text = IIf(IsNull(rstMuestra!Articulo2), "  -   -   ", rstMuestra!Articulo2)
            End If
            Ensayo.Text = IIf(IsNull(rstMuestra!ensayo2), "  -     -   ", rstMuestra!ensayo2)
            Cantidad.Text = IIf(IsNull(rstMuestra!Cantidad2), "", rstMuestra!Cantidad2)
            Lote.Text = IIf(IsNull(rstMuestra!lote2), "", rstMuestra!lote2)
            Observaciones.Text = IIf(IsNull(rstMuestra!Observaciones2), "", rstMuestra!Observaciones2)
            WActualiza = IIf(IsNull(rstMuestra!Stock2), "0", rstMuestra!Stock2)
            Actualiza.ListIndex = Val(WActualiza)
                Else
            WGraba = "S"
            Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            WProducto = Trim(IIf(IsNull(rstMuestra!Producto), "  -     -   ", rstMuestra!Producto))
            WArticulo = Trim(IIf(IsNull(rstMuestra!Articulo), "  -   -   ", rstMuestra!Articulo))
            If WProducto <> "" And Len(WProducto) = 12 Then
                Producto.Text = IIf(IsNull(rstMuestra!Producto), "  -     -   ", rstMuestra!Producto)
            End If
            If WArticulo <> "" And Len(WArticulo) = 10 Then
                Articulo.Text = Trim(IIf(IsNull(rstMuestra!Articulo), "  -   -   ", rstMuestra!Articulo))
            End If
            Ensayo.Text = rstMuestra!Ensayo
            Cantidad.Text = ""
            Lote.Text = ""
            Observaciones.Text = ""
            Actualiza.ListIndex = 1
        End If
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

Private Sub Observaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Lote_KeyDown(KeyCode As Integer, Shift As Integer)
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





Private Sub Conecta_Empresa()

    Select Case Val(XEmpresa)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select

End Sub









