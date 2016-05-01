VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsultaCompras 
   Caption         =   "Consulta Datos Factura"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   5
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   3
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   4
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox codigo 
      Height          =   495
      Left            =   7920
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame PantaHistoria 
      Caption         =   "Historial"
      Height          =   3375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox HistoriaOrden 
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
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   " "
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox HistoriaInforme 
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
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   " "
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox HistoriaRemito 
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
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   " "
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox HistoriaFactura 
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
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   " "
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton PantaHistoriaCierra 
         Caption         =   "Cierra"
         Height          =   495
         Left            =   2760
         TabIndex        =   2
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox HistoriaCarpeta 
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
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   " "
         Top             =   1800
         Width           =   1335
      End
      Begin MSMask.MaskEdBox HistoriaFechaOrden 
         Height          =   285
         Left            =   3720
         TabIndex        =   7
         Top             =   360
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
      Begin MSMask.MaskEdBox HistoriaFechaInforme 
         Height          =   285
         Left            =   3720
         TabIndex        =   8
         Top             =   720
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
      Begin MSMask.MaskEdBox HistoriaFechaFactura 
         Height          =   285
         Left            =   3720
         TabIndex        =   9
         Top             =   1440
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
      Begin VB.Label Desproveedor 
         BackColor       =   &H00FFFFC0&
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
         Left            =   2280
         TabIndex        =   16
         Top             =   2160
         Width           =   3855
      End
      Begin VB.Label Label60 
         Caption         =   "Proveedor"
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
         TabIndex        =   15
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label59 
         Caption         =   "Orden de Compra"
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
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label58 
         Caption         =   "Informe de Recepcion"
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
         TabIndex        =   13
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label57 
         Caption         =   "Remito"
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
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label52 
         Caption         =   "Factura"
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
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label51 
         Caption         =   "Carpeta"
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
         Top             =   1800
         Width           =   1695
      End
   End
   Begin MSFlexGridLib.MSFlexGrid WGrilla 
      Height          =   3015
      Left            =   240
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5318
      _Version        =   327680
      BackColor       =   16777152
   End
End
Attribute VB_Name = "PrgConsultaCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CargaEmpresa(100, 10) As String



Private Sub Muestra_Historial()

    ZZZZEmpresa = WEmpresa
     
     Call Limpia_Grilla
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM IvaComp"
    ZSql = ZSql + " Where IvaComp.NroInterno = " + "'" + WPasaNroInterno + "'"
    spIvaComp = ZSql
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
        HistoriaFactura.Text = Str$(Val(rstIvaComp!Numero))
        HistoriaFechaFactura.Text = rstIvaComp!Fecha
        HistoriaRemito.Text = rstIvaComp!Remito
        ZZProveedor = rstIvaComp!Proveedor
        rstIvaComp.Close
    End If

    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            CargaEmpresa(1, 1) = "0001"
            CargaEmpresa(1, 2) = "Empresa01"
            CargaEmpresa(2, 1) = "0003"
            CargaEmpresa(2, 2) = "Empresa03"
            CargaEmpresa(3, 1) = "0005"
            CargaEmpresa(3, 2) = "Empresa05"
            CargaEmpresa(4, 1) = "0006"
            CargaEmpresa(4, 2) = "Empresa06"
            CargaEmpresa(5, 1) = "0007"
            CargaEmpresa(5, 2) = "Empresa07"
            CargaEmpresa(6, 1) = "0010"
            CargaEmpresa(6, 2) = "Empresa10"
            CargaEmpresa(7, 1) = "0011"
            CargaEmpresa(7, 2) = "Empresa11"
            ZHasta = 7
                    
        Case Else
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
            ZHasta = 4
                    
    End Select
    
    If Trim(HistoriaRemito.Text) <> "" Then
                    
        For Cicla = 1 To ZHasta
        
            If CargaEmpresa(Cicla, 1) <> "" Then
                
                ZZSalida = "N"
        
                WEmpresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Informe"
                ZSql = ZSql + " Where Informe.Remito = " + "'" + HistoriaRemito.Text + "'"
                ZSql = ZSql + " and Informe.Proveedor = " + "'" + ZZProveedor + "'"
                spInforme = ZSql
                Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                If rstInforme.RecordCount > 0 Then
                    HistoriaInforme.Text = rstInforme!Informe
                    HistoriaFechaInforme.Text = rstInforme!Fecha
                    HistoriaOrden.Text = rstInforme!Orden
                    ZZSalida = "S"
                    rstInforme.Close
                End If
                
                If ZZSalida = "S" Then
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Informe"
                    ZSql = ZSql + " Where Informe.Remito = " + "'" + HistoriaRemito.Text + "'"
                    ZSql = ZSql + " and Informe.Proveedor = " + "'" + ZZProveedor + "'"
                    spInforme = ZSql
                    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                    If rstInforme.RecordCount > 0 Then
                        HistoriaInforme.Text = rstInforme!Informe
                        HistoriaFechaInforme.Text = rstInforme!Fecha
                        HistoriaOrden.Text = rstInforme!Orden
                        rstInforme.Close
                    End If
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Proveedor"
                    ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + ZZProveedor + "'"
                    spProveedor = ZSql
                    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    If RstProveedor.RecordCount > 0 Then
                        Desproveedor.Caption = RstProveedor!Nombre
                        RstProveedor.Close
                    End If
                    
                    If Trim(HistoriaOrden.Text) <> "" Then
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden.Orden = " + "'" + HistoriaOrden.Text + "'"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            HistoriaFechaOrden.Text = rstOrden!Fecha
                            HistoriaCarpeta.Text = rstOrden!Carpeta
                            rstOrden.Close
                        End If
                        
                    End If
                    
    
                    Renglon = 0
    
                    spInforme = "ListaInforme " + "'" + HistoriaInforme.Text + "'"
                    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                        
                    If rstInforme.RecordCount > 0 Then
                        With rstInforme
                            .MoveFirst
                            Do
                                If .EOF = False Then
                                
                                    Renglon = Renglon + 1
                                    WGrilla.TextMatrix(Renglon, 1) = rstInforme!Orden
                                    WGrilla.TextMatrix(Renglon, 2) = rstInforme!Articulo
                                    WGrilla.TextMatrix(Renglon, 4) = Pusing("###,###.##", rstInforme!Cantidad)
                                    
                                    .MoveNext
                                        Else
                                    Exit Do
                                End If
                            Loop
                        End With
                        rstInforme.Close
                    End If
                    
                    WRenglon = Renglon
                    Renglon = 0
                    
                    For da = 1 To WRenglon
                    
                        Renglon = Renglon + 1
                                
                        spArticulo = "ConsultaArticulo " + "'" + WGrilla.TextMatrix(Renglon, 2) + "'"
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            WGrilla.TextMatrix(Renglon, 3) = rstArticulo!Descripcion
                            rstArticulo.Close
                        End If
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden.Orden = " + "'" + HistoriaOrden.Text + "'"
                        ZSql = ZSql + " and Orden.Articulo = " + "'" + WGrilla.TextMatrix(Renglon, 2) + "'"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            WGrilla.TextMatrix(Renglon, 5) = Pusing("###,###.##", Str$(rstOrden!Precio))
                            rstOrden.Close
                        End If
                        
                    
                    Next da
                    
                    Exit For
                    
                End If
    
            End If
            
        Next Cicla

    End If

    XEmpresa = ZZZZEmpresa
    Call Conecta_Empresa
    
    
End Sub

Private Sub Form_Load()
    codigo.Text = WPasaNroInterno
    Call Muestra_Historial
End Sub


Private Sub Limpia_Grilla()

    WGrilla.Clear
    WGrilla.Font.Bold = True
    
    WGrilla.FixedCols = 1
    WGrilla.Cols = 6
    WGrilla.FixedRows = 1
    WGrilla.Rows = 101
    
    WGrilla.ColWidth(0) = 200
    WGrilla.Row = 0
    For Ciclo = 1 To WGrilla.Cols - 1
        WGrilla.Col = Ciclo
        Select Case Ciclo
            Case 1
                WGrilla.Text = "Orden"
                WGrilla.ColWidth(Ciclo) = 1000
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                WGrilla.Text = "Producto"
                WGrilla.ColWidth(Ciclo) = 1300
                WGrilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                WGrilla.Text = "Descripcion"
                WGrilla.ColWidth(Ciclo) = 2500
                WGrilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                WGrilla.Text = "Cantidad"
                WGrilla.ColWidth(Ciclo) = 1100
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WGrilla.Text = "Precio"
                WGrilla.ColWidth(Ciclo) = 1100
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WGrilla.Row = 0
    For Ciclo = 1 To WGrilla.Cols - 1
        WGrilla.Col = Ciclo
        WTitulo(Ciclo).Text = WGrilla.Text
        WTitulo(Ciclo).Left = WGrilla.CellLeft + WGrilla.Left
        WTitulo(Ciclo).Top = WGrilla.CellTop + WGrilla.Top
        WTitulo(Ciclo).Width = WGrilla.CellWidth
        WTitulo(Ciclo).Height = WGrilla.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WGrilla.Cols - 1
        WAncho = WAncho + WGrilla.ColWidth(Ciclo)
    Next Ciclo
    WGrilla.Width = WAncho

    ' Size the columns.
    Font.Name = WGrilla.Font.Name
    Font.Size = WGrilla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WGrilla.AllowUserResizing = flexResizeBoth
    
    WGrilla.Visible = True
    
    WGrilla.TopRow = 1
    WGrilla.Col = 1
    WGrilla.Row = 1
    
End Sub

