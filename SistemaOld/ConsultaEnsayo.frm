VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsultaEnsayo 
   Caption         =   "Ingreso de Ensayos"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   11775
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
      Left            =   0
      TabIndex        =   21
      Top             =   7200
      Visible         =   0   'False
      Width           =   1935
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
      Height          =   300
      Left            =   2040
      TabIndex        =   20
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   1920
      TabIndex        =   19
      Top             =   7560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "ConsultaEnsayo.frx":0000
      Left            =   0
      List            =   "ConsultaEnsayo.frx":0007
      TabIndex        =   18
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
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
      TabIndex        =   8
      Text            =   " "
      Top             =   600
      Width           =   1095
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
      TabIndex        =   7
      Text            =   " "
      Top             =   960
      Width           =   5895
   End
   Begin VB.TextBox WTexto24 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
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
      Left            =   1440
      TabIndex        =   4
      Top             =   2280
      Width           =   375
   End
   Begin VB.ComboBox WCombo14 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto14 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Width           =   375
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin MSMask.MaskEdBox Orden 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "AA-#####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox WTexto34 
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Top             =   2280
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
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
   Begin MSFlexGridLib.MSFlexGrid WVector4 
      Height          =   5535
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9763
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Top             =   240
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
   Begin MSMask.MaskEdBox FechaEntrega 
      Height          =   285
      Left            =   8640
      TabIndex        =   10
      Top             =   240
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
   Begin VB.Image BusquedaEnsayo 
      Height          =   480
      Left            =   5040
      Picture         =   "ConsultaEnsayo.frx":0015
      ToolTipText     =   "Busqueda de Ensayos"
      Top             =   7200
      Width           =   480
   End
   Begin VB.Image BusquedaEnsayoII 
      Height          =   480
      Left            =   7080
      Picture         =   "ConsultaEnsayo.frx":0457
      ToolTipText     =   "Busqueda de Ensayos por CLiente"
      Top             =   7200
      Width           =   480
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "x Ensayo"
      Height          =   255
      Left            =   4800
      TabIndex        =   17
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "x Cliente"
      Height          =   255
      Left            =   6960
      TabIndex        =   16
      Top             =   7680
      Width           =   855
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
      Left            =   3720
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Comrpometida"
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
      Left            =   6600
      TabIndex        =   14
      Top             =   240
      Width           =   1935
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
      Left            =   3240
      TabIndex        =   13
      Top             =   600
      Width           =   4695
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
      Left            =   360
      TabIndex        =   12
      Top             =   600
      Width           =   1575
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
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   1815
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   9120
      MouseIcon       =   "ConsultaEnsayo.frx":0899
      MousePointer    =   99  'Custom
      Picture         =   "ConsultaEnsayo.frx":0BA3
      ToolTipText     =   "Salida"
      Top             =   7200
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   3120
      MouseIcon       =   "ConsultaEnsayo.frx":13E5
      MousePointer    =   99  'Custom
      Picture         =   "ConsultaEnsayo.frx":16EF
      ToolTipText     =   "Limpia la pantalla"
      Top             =   7200
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Orden de Trabajo"
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "PrgConsultaEnsayo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOrdenTrabajo As Recordset
Dim spOrdenTrabajo As String
Dim rstOrdenTrabajoII As Recordset
Dim spOrdenTrabajoII As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstCargaEnsayo As Recordset
Dim spCargaEnsayo As String
Dim rstCargaEnsayoII As Recordset
Dim spCargaEnsayoII As String
Dim rstCargaEnsayoIII As Recordset
Dim spCargaEnsayoIII As String
Dim rstCargaEnsayoIV As Recordset
Dim spCargaEnsayoIV As String
Dim rstCargaEnsayoV As Recordset
Dim spCargaEnsayoV As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String

Dim CargaEmpresa(12, 2) As String
Dim ZCarga(10000, 3) As String
Dim ZClienteII As String
Dim ZProceso As Integer

Private ZAuxiliar(100, 7) As String
Dim Producto As String
Dim XCosto1 As Double
Dim XCosto2 As Double
Dim XCosto3 As Double

Dim XParam As String
Dim EmpresaActual As String
Dim WVersion As String

Rem para el vector IV

Dim WBorraIV(1000, 20) As String
Dim WParametrosIV(10, 20) As Double
Dim WFormatoIV(20) As String
Dim WControlIV As String

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Sub Imprime_Datos()

    On Error GoTo WError

    ZOrden = Orden.Text
    
    Call CmdLimpiar_Click
    
    ZExiste = "N"
    
    Orden.Text = ZOrden
    
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesCliente.Caption = rstCliente!Razon
        rstCliente.Close
            Else
        DesCliente.Caption = ""
    End If
        
    Call Limpia_VectorIV
    WRenglon = 0
    
    XEmpresa = WEmpresa
        
    WEmpresa = "0003"
    txtOdbc = "Empresa03"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaEnsayoV"
    ZSql = ZSql + " Where CargaEnsayoV.Orden = " + "'" + Orden.Text + "'"
    ZSql = ZSql + " Order by CargaEnsayoV.Clave"
    spCargaEnsayoV = ZSql
    Set rstCargaEnsayoV = db.OpenRecordset(spCargaEnsayoV, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaEnsayoV.RecordCount > 0 Then
        With rstCargaEnsayoV
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    WVector4.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector4.Col = 1
                    WVector4.Text = Trim(rstCargaEnsayoV!Version)
                    
                    WVector4.Col = 2
                    WVector4.Text = Trim(rstCargaEnsayoV!Etapa)
            
                    WVector4.Col = 3
                    WVector4.Text = Trim(rstCargaEnsayoV!Fecha)
            
                    WVector4.Col = 4
                    WVector4.Text = Trim(rstCargaEnsayoV!Participantes)
                    
                    WVector4.Col = 5
                    WVector4.Text = Trim(rstCargaEnsayoV!Resultados)
                    
                    WVector4.Col = 6
                    WVector4.Text = Trim(rstCargaEnsayoV!Acciones)
                    
                    WVector4.Col = 7
                    WVector4.Text = Trim(rstCargaEnsayoV!Responsables)
            
                    WVector4.Col = 8
                    WVector4.Text = Trim(rstCargaEnsayoV!Estado)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaEnsayoV.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM OrdenTrabajo"
    ZSql = ZSql + " Where OrdenTrabajo.Orden = " + "'" + Orden.Text + "'"
    spOrdenTrabajo = ZSql
    Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrdenTrabajo.RecordCount > 0 Then
        Fecha.Text = rstOrdenTrabajo!Fecha
        FechaEntrega.Text = rstOrdenTrabajo!FechaEntrega
        Cliente.Text = rstOrdenTrabajo!Cliente
        Observaciones.Text = Trim(rstOrdenTrabajo!Observaciones)
        rstOrdenTrabajo.Close
    End If
    
    Call Conecta_Empresa
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub CmdLimpiar_Click()
    
    Call Limpia_VectorIV

    Orden.Text = "  -     "
    Fecha.Text = "  /  /    "
    FechaEntrega.Text = "  /  /    "
    Cliente.Text = ""
    Observaciones.Text = ""
    DesCliente.Caption = ""
    
    Orden.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call CmdLimpiar_Click
    PrgConsultaEnsayo.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Orden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Orden.Text <> "" Then
        
            Orden.Text = UCase(Orden.Text)
            
            XEmpresa = WEmpresa
        
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM OrdenTrabajo"
            ZSql = ZSql + " Where OrdenTrabajo.Orden = " + "'" + Orden.Text + "'"
            spOrdenTrabajo = ZSql
            Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrdenTrabajo.RecordCount > 0 Then
                Fecha.Text = rstOrdenTrabajo!Fecha
                FechaEntrega.Text = rstOrdenTrabajo!FechaEntrega
                Cliente.Text = rstOrdenTrabajo!Cliente
                Observaciones.Text = Trim(rstOrdenTrabajo!Observaciones)
                rstOrdenTrabajo.Close
                Call Conecta_Empresa
                Call Imprime_Datos
                    Else
                Call Conecta_Empresa
            End If
            
        End If
    End If
    If KeyAscii = 27 Then
        Orden.Text = "  -     "
    End If
End Sub

Sub Form_Load()

    Call Limpia_VectorIV
    
    WVector4.Col = 1
    WVector4.Row = 1
    
    Orden.Text = "  -     "
    Fecha.Text = "  /  /    "
    FechaEntrega.Text = "  /  /    "
    Cliente.Text = ""
    Observaciones.Text = ""
    DesCliente.Caption = ""
    
End Sub


Rem
Rem Controles de la WVector4
Rem

Private Sub GridEditTextIV(ByVal KeyAscii As Integer)

    XColumna = WVector4.Col
    XTipoDato = WParametrosIV(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto14.Left = WVector4.CellLeft + WVector4.Left
            WTexto14.Top = WVector4.CellTop + WVector4.Top
            WTexto14.Width = WVector4.CellWidth
            WTexto14.Height = WVector4.CellHeight
            WTexto14.MaxLength = WParametrosIV(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto14.Text = WVector4.Text
                    WTexto14.SelStart = Len(WTexto14.Text)
                Case Else
                    WTexto14.Text = Chr$(KeyAscii)
                    WTexto14.SelStart = 1
            End Select
            WTexto14.Visible = True
            WTexto14.SetFocus
        Case 1
            WTexto24.Left = WVector4.CellLeft + WVector4.Left
            WTexto24.Top = WVector4.CellTop + WVector4.Top
            WTexto24.Width = WVector4.CellWidth
            WTexto24.Height = WVector4.CellHeight
            WTexto24.MaxLength = WParametrosIV(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto24.Text = WVector4.Text
                    Rem WTexto24.SelStart = Len(WTexto24.Text)
                    WTexto24.SelStart = 0
                Case Else
                    WTexto24.Text = Chr$(KeyAscii)
                    WTexto24.SelStart = 1
            End Select
            WTexto24.Visible = True
            WTexto24.SetFocus
        Case 2
            WTexto34.Left = WVector4.CellLeft + WVector4.Left
            WTexto34.Top = WVector4.CellTop + WVector4.Top
            WTexto34.Width = WVector4.CellWidth
            WTexto34.Height = WVector4.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector4.Text) = 10 Then
                        WTexto34.Text = WVector4.Text
                            Else
                        WTexto34.Text = "  /  /    "
                    End If
                    WTexto34.SelStart = 0
                Case Else
                    WTexto34.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto34.SelStart = 1
            End Select
            WTexto34.Visible = True
            WTexto34.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditIV()
    Pasa = 0
    If WCombo14.Visible Then
        Pasa = 0
        WVector4.Text = WCombo14.Text
        WCombo14.Visible = False
            Else
        If WTexto14.Visible Then
            Pasa = 1
            WVector4.Text = WTexto14.Text
            WTexto14.Visible = False
                Else
            If WTexto24.Visible Then
                Pasa = 1
                WVector4.Text = WTexto24.Text
                WTexto24.Visible = False
                    Else
                If WTexto34.Visible Then
                    Pasa = 1
                    WVector4.Text = WTexto34.Text
                    WTexto34.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoIV(WVector4.Col) <> "" Then
            WVector4.Text = Pusing(WFormatoIV(WVector4.Col), WVector4.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboIV()
    ' Position the ComboBox over the cell.
    WCombo14.Left = WVector4.CellLeft + WVector4.Left
    WCombo14.Top = WVector4.CellTop + WVector4.Top
    WCombo14.Width = WVector4.CellWidth
    WCombo14.Visible = True
    WCombo14.SetFocus
End Sub

Private Sub WTexto14_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto14.Text = ""
            
        Rem F1
        Case 113
            WTexto14.Text = WVector4.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector4.SetFocus
            DoEvents
            Call Control_CampoIV
            If WControlIV = "S" Then
                Call Control_WVectorIV
            End If
            Call StartEditIV

        Case vbKeyDown
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row < WVector4.Rows - 1 Then
                Call Control_CampoIV
                If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row + 1
                End If
            End If
            Call StartEditIV

        Case vbKeyUp
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row > WVector4.FixedRows Then
                Call Control_CampoIV
                If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row - 1
                End If
            End If
            Call StartEditIV
        Case 34
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow < WVector4.Rows - 12 Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow + 12
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV
            
        Case 33
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow - 12 > WVector4.FixedRows Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow - 12
                    WVector4.Row = WVector4.TopRow
                        Else
                    WVector4.TopRow = 1
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV
            
        Case 123
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Col > 1 Then
                WVector4.Col = WVector4.Col - 1
            End If
            Call StartEditIV

    End Select
End Sub

Private Sub WTexto24_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto24.Text = ""
            
        Rem F1
        Case 113
            WTexto24.Text = WVector4.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector4.SetFocus
            DoEvents
            Call Control_CampoIV
            If WControlIV = "S" Then
                Call Control_WVectorIV
            End If
            Call StartEditIV
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row < WVector4.Rows - 1 Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row + 1
                Rem End If
            End If
            Call StartEditIV

        Case vbKeyUp
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row > WVector4.FixedRows Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row - 1
                Rem End If
            End If
            Call StartEditIV
        Case 34
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow < WVector4.Rows - 12 Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow + 12
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV
            
        Case 33
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow - 12 > WVector4.FixedRows Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow - 12
                    WVector4.Row = WVector4.TopRow
                        Else
                    WVector4.TopRow = 1
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV

    End Select
End Sub

Private Sub WTexto34_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto34.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto34.Text = WVector4.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector4.SetFocus
            Call Control_CampoIV
            If WControlIV = "S" Then
                Call Control_WVectorIV
            End If
            Call StartEditIV

        Case vbKeyDown
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row < WVector4.Rows - 1 Then
                Call Control_CampoIV
                If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row + 1
                End If
            End If
            Call StartEditIV

        Case vbKeyUp
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row > WVector4.FixedRows Then
                Call Control_CampoIV
                If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row - 1
                End If
            End If
            Call StartEditIV
        Case 34
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow < WVector4.Rows - 12 Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow + 12
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV
            
        Case 33
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow - 12 > WVector4.FixedRows Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow - 12
                    WVector4.Row = WVector4.TopRow
                        Else
                    WVector4.TopRow = 1
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto14_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto24_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto34_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo14_Click()
    WVector4.SetFocus
End Sub


Private Sub WVector4_Click()
    StartEditIV
End Sub

Private Sub WVector4_LeaveCell()
    EndEditIV
End Sub

Private Sub WVector4_GotFocus()
    EndEditIV
End Sub

Private Sub WVector4_KeyPress(KeyAscii As Integer)
    XColumna = WVector4.Col
    Select Case WParametrosIV(4, WVector4.Col)
        Case 1
        Case Else
            If WParametrosIV(2, XColumna) = 0 Then
                GridEditTextIV KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditIV()
    Select Case WParametrosIV(4, WVector4.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo14.Clear
            WCombo14.AddItem "Campo1"
            WCombo14.AddItem "Campo2"
            On Error Resume Next
            WCombo14.Text = WVector4.Text
            On Error GoTo 0
            GridEditComboIV
        Case Else
            If WParametrosIV(2, WVector4.Col) = 0 Then
                GridEditTextIV Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVectorIV()
    Select Case WVector4.Col
        Case 8
            If WVector4.Row < WVector4.Rows - 1 Then
                WVector4.Row = WVector4.Row + 1
            End If
            WVector4.Col = 1
        Case Else
            If WVector4.Col < WVector4.Cols - 1 Then
                WVector4.Col = WVector4.Col + 1
            End If
    End Select
    WVector4.SetFocus
    GridEditTextIV KeyAscii
End Sub

Private Sub Control_CampoIV()
    XColumna = WVector4.Col
    XFila = WVector4.Row
    WControlIV = "S"
End Sub



Private Sub WVector4_DblClick()

    If WVector4.Col = 0 Or WVector4.Col = 1 Then
    
    WTexto14.Visible = False
    WTexto24.Visible = False
    WTexto34.Visible = False
    
    RenglonAuxiliar = WVector4.Row

    For Ciclo = 1 To WVector4.Cols - 1
        WVector4.Col = Ciclo
        WVector4.Text = ""
    Next Ciclo
    
    Erase WBorraIV
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 99 To 1 Step -1
        
        ZEtapa = WVector4.TextMatrix(iRow, 1)
        ZFecha = WVector4.TextMatrix(iRow, 2)
        ZParticipantes = WVector4.TextMatrix(iRow, 3)
        ZResultados = WVector4.TextMatrix(iRow, 4)
        ZAcciones = WVector4.TextMatrix(iRow, 5)
        ZResponsables = WVector4.TextMatrix(iRow, 6)
        ZEstado = WVector4.TextMatrix(iRow, 7)
            
        If ZEtapa <> "" Or ZFecha <> "" Or ZParticipantes <> "" Or ZResultados <> "" Or ZAcciones <> "" Or ZResponsables <> "" Or ZEstado <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector4.Row = Ciclo
        WVector4.Col = 1
        WAuxi1 = WVector4.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector4.Cols - 1
                WVector4.Col = Ciclo1
                WBorraIV(EntraVector, Ciclo1) = WVector4.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_VectorIV
    
    For Ciclo = 1 To EntraVector
        WVector4.Row = Ciclo
        For DA = 0 To WVector4.Cols - 1
            WVector4.Col = DA
            WVector4.Text = WBorraIV(Ciclo, DA)
        Next DA
    Next Ciclo
    
    End If
    
End Sub

Private Sub AgregaRenglonIV_Click()

    Hasta = WVector4.Row

    For iRow = 1000 To Hasta Step -1
        WVector4.TextMatrix(iRow, 0) = WVector4.TextMatrix(iRow - 1, 0)
        WVector4.TextMatrix(iRow, 1) = WVector4.TextMatrix(iRow - 1, 1)
        WVector4.TextMatrix(iRow, 2) = WVector4.TextMatrix(iRow - 1, 2)
        WVector4.TextMatrix(iRow, 3) = WVector4.TextMatrix(iRow - 1, 3)
        WVector4.TextMatrix(iRow, 4) = WVector4.TextMatrix(iRow - 1, 4)
        WVector4.TextMatrix(iRow, 5) = WVector4.TextMatrix(iRow - 1, 5)
        WVector4.TextMatrix(iRow, 6) = WVector4.TextMatrix(iRow - 1, 6)
        WVector4.TextMatrix(iRow, 7) = WVector4.TextMatrix(iRow - 1, 7)
        WVector4.TextMatrix(iRow, 8) = WVector4.TextMatrix(iRow - 1, 8)
    Next iRow

    WVector4.TextMatrix(Hasta, 0) = ""
    WVector4.TextMatrix(Hasta, 1) = ""
    WVector4.TextMatrix(Hasta, 2) = ""
    WVector4.TextMatrix(Hasta, 3) = ""
    WVector4.TextMatrix(Hasta, 4) = ""
    WVector4.TextMatrix(Hasta, 5) = ""
    WVector4.TextMatrix(Hasta, 6) = ""
    WVector4.TextMatrix(Hasta, 7) = ""
    WVector4.TextMatrix(Hasta, 8) = ""
    
    WTexto14.Text = ""
    WTexto24.Text = ""

End Sub




Private Sub Limpia_VectorIV()

    WVector4.Clear

    Rem ponga la WVector4 en negritas
    WVector4.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto14.FontName = WVector4.FontName
    WTexto14.FontSize = WVector4.FontSize
    WTexto14.Visible = False
    WTexto24.FontName = WVector4.FontName
    WTexto24.FontSize = WVector4.FontSize
    WTexto24.Visible = False
    WTexto34.FontName = WVector4.FontName
    WTexto34.FontSize = WVector4.FontSize
    WTexto34.Visible = False
    WCombo14.FontName = WVector4.FontName
    WCombo14.FontSize = WVector4.FontSize
    WCombo14.Visible = False

    ' Establesco loa Valores de la WVector4
    
    WVector4.FixedCols = 1
    WVector4.Cols = 9
    WVector4.FixedRows = 1
    WVector4.Rows = 1001
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector4.Text = "Articulo"
    
    Rem Longitud
    Rem WVector4.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector4.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametrosIV(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametrosIV(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametrosIV(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametrosIV(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector4.ColWidth(0) = 200
    WVector4.Row = 0
    For Ciclo = 1 To WVector4.Cols - 1
        WVector4.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector4.Text = "Version"
                WVector4.ColWidth(Ciclo) = 800
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 10
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 2
                WVector4.Text = "Etapa"
                WVector4.ColWidth(Ciclo) = 900
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 20
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 3
                WVector4.Text = "Fecha"
                WVector4.ColWidth(Ciclo) = 1000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 10
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 4
                WVector4.Text = "Participantes"
                WVector4.ColWidth(Ciclo) = 2000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 50
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 5
                WVector4.Text = "Resultados"
                WVector4.ColWidth(Ciclo) = 4000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 50
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 6
                WVector4.Text = "Acciones"
                WVector4.ColWidth(Ciclo) = 4000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 50
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 7
                WVector4.Text = "Responsables"
                WVector4.ColWidth(Ciclo) = 2000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 20
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 8
                WVector4.Text = "Estado"
                WVector4.ColWidth(Ciclo) = 1200
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 20
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
                
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector4.Row = 0
    For Ciclo = 1 To WVector4.Cols - 1
        WVector4.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector4.Text
        Rem WTitulo(Ciclo).Left = WVector4.CellLeft + WVector4.Left
        Rem WTitulo(Ciclo).Top = WVector4.CellTop + WVector4.Top
        Rem WTitulo(Ciclo).Width = WVector4.CellWidth
        Rem WTitulo(Ciclo).Height = WVector4.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector4
    
    WAncho = 400
    For Ciclo = 0 To WVector4.Cols - 1
        WAncho = WAncho + WVector4.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector4.Width = WAncho

    ' Size the columns.
    Font.Name = WVector4.Font.Name
    Font.Size = WVector4.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector4.AllowUserResizing = flexResizeBoth
    
    WVector4.Col = 1
    WVector4.Row = 1
    
End Sub

Private Sub WVector4_Scroll()
    WTexto14.Visible = False
    WTexto24.Visible = False
    WTexto34.Visible = False
End Sub


Private Sub BusquedaEnsayo_Click()
    ZClienteII = ""
    Call Busca_Ensayo
End Sub

Private Sub BusquedaEnsayoII_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear
    ZProceso = 2
    
    Pantalla.Height = 5340
    Pantalla.Left = 480
    Pantalla.Top = 1680
    Pantalla.Width = 10575
    
    Erase ZCarga
    ZLugar = 0
    
    Pasa = 0
    
    XEmpresa = WEmpresa
        
    WEmpresa = "0003"
    txtOdbc = "Empresa03"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM OrdenTrabajo"
    ZSql = ZSql + " Order by OrdenTRabajo.Cliente"
    spOrdenTrabajo = ZSql
    Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrdenTrabajo.RecordCount > 0 Then
        With rstOrdenTrabajo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    aa = rstOrdenTrabajo!Cliente
                    AA1 = WEmpresa
                
                    If Pasa = 0 Then
                        Pasa = 1
                        ZCliente = rstOrdenTrabajo!Cliente
                    End If
                    
                    If ZCliente <> rstOrdenTrabajo!Cliente Then
                        ZLugar = ZLugar + 1
                        ZCarga(ZLugar, 1) = ZCliente
                        ZCliente = rstOrdenTrabajo!Cliente
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrdenTrabajo.Close
    End If
    
    Call Conecta_Empresa

    If Pasa <> 0 Then
        ZLugar = ZLugar + 1
        ZCarga(ZLugar, 1) = ZCliente
    End If
    
    For Ciclo = 1 To ZLugar
        
        ZCliente = ZCarga(Ciclo, 1)
        
        spCliente = "ConsultaCliente " + "'" + ZCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZDesCliente = Trim(rstCliente!Razon)
            rstCliente.Close
                Else
            ZDesCliente = ""
        End If
        
        IngresaItem = ZCliente + "  " + ZDesCliente
        Pantalla.AddItem IngresaItem
        IngresaItem = ZCliente
        WIndice.AddItem IngresaItem
        
    Next Ciclo
    
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case ZProceso
        Case 1
            Indice = Pantalla.ListIndex
            Orden.Text = Left$(WIndice.List(Indice), 8)
            Call Orden_KeyPress(13)
            
        Case 2
            Indice = Pantalla.ListIndex
            ZClienteII = WIndice.List(Indice)
            Call Busca_Ensayo
            
        Case Else
    End Select
End Sub

Private Sub Busca_Ensayo()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear
    ZProceso = 1
    
    Pantalla.Height = 5340
    Pantalla.Left = 480
    Pantalla.Top = 1680
    Pantalla.Width = 10575
    
    Erase ZCarga
    ZLugar = 0
    
    Pasa = 0
    
    XEmpresa = WEmpresa
        
    WEmpresa = "0003"
    txtOdbc = "Empresa03"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaEnsayo"
    ZSql = ZSql + " Order by CargaEnsayo.Clave"
    spCargaEnsayo = ZSql
    Set rstCargaEnsayo = db.OpenRecordset(spCargaEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaEnsayo.RecordCount > 0 Then
        With rstCargaEnsayo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    aa = rstCargaEnsayo!Orden
                    AA1 = WEmpresa
                
                    If Pasa = 0 Then
                        Pasa = 1
                        ZOrden = rstCargaEnsayo!Orden
                    End If
                    
                    If ZOrden <> rstCargaEnsayo!Orden Then
                        ZLugar = ZLugar + 1
                        ZCarga(ZLugar, 1) = ZOrden
                        ZCarga(ZLugar, 2) = ZVersion
                        ZCarga(ZLugar, 3) = ZClave
                        ZOrden = rstCargaEnsayo!Orden
                    End If
                    
                    ZVersion = rstCargaEnsayo!Version
                    ZClave = rstCargaEnsayo!Clave
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaEnsayo.Close
    End If

    If Pasa <> 0 Then
        ZLugar = ZLugar + 1
        ZCarga(ZLugar, 1) = ZOrden
        ZCarga(ZLugar, 2) = ZVersion
        ZCarga(ZLugar, 3) = ZClave
    End If
    
    Call Conecta_Empresa
    
    For Ciclo = 1 To ZLugar
        
        ZOrden = ZCarga(Ciclo, 1)
        ZVersion = ZCarga(Ciclo, 2)
        ZClave = ZCarga(Ciclo, 3)
        
        XEmpresa = WEmpresa
        
        WEmpresa = "0003"
        txtOdbc = "Empresa03"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM OrdenTrabajo"
        ZSql = ZSql + " Where OrdenTrabajo.Orden = " + "'" + ZOrden + "'"
        spOrdenTrabajo = ZSql
        Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrdenTrabajo.RecordCount > 0 Then
            ZObservaciones = Trim(rstOrdenTrabajo!Observaciones)
            ZCliente = Trim(rstOrdenTrabajo!Cliente)
            rstOrdenTrabajo.Close
                Else
            ZObservaciones = ""
            ZCliente = ""
        End If
        
        Call Conecta_Empresa
            
        spCliente = "ConsultaCliente " + "'" + ZCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZDesCliente = Trim(rstCliente!Razon)
            rstCliente.Close
                Else
            ZDesCliente = ""
        End If
        
        If ZClienteII = "" Or ZClienteII = ZCliente Then
        
            IngresaItem = ZOrden + "/" + Str$(ZVersion) + "  " + ZObservaciones + "  (" + ZDesCliente + ")"
            Pantalla.AddItem IngresaItem
            IngresaItem = ZClave
            WIndice.AddItem IngresaItem
        
        End If
        
    Next Ciclo
    
    Pantalla.Visible = True

End Sub


