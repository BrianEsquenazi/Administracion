VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgSolic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitudes de Compra"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11760
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8625
   ScaleWidth      =   11760
   Visible         =   0   'False
   Begin VB.CommandButton ReImpresion 
      Caption         =   "ReImpresion"
      Height          =   500
      Left            =   3360
      TabIndex        =   42
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Frame SeleccionPuerto 
      Caption         =   "Seleccione el puerto de Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3000
      TabIndex        =   39
      Top             =   3000
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton WPuerto2 
         Caption         =   "LPT2"
         Height          =   375
         Left            =   2040
         TabIndex        =   41
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton WPuerto1 
         Caption         =   "LPT1"
         Height          =   375
         Left            =   600
         TabIndex        =   40
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Clave 
      Caption         =   "  Ingreso de Clave de Seguridad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3120
      TabIndex        =   33
      Top             =   1560
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   35
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Cancelagraba 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   960
         TabIndex        =   34
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Ingrese su Password"
         Height          =   255
         Left            =   1080
         TabIndex        =   36
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.TextBox Solicitante 
      Height          =   285
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   30
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox Planta 
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   28
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   4800
      TabIndex        =   26
      Top             =   6360
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   20
      Text            =   " "
      Top             =   840
      Width           =   9855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4080
      Top             =   8400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "orden.rpt"
      Destination     =   3
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   17
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   500
      Left            =   2280
      TabIndex        =   18
      Top             =   6960
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   9120
      TabIndex        =   17
      Top             =   6720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   15
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Solicitud 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   13
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   500
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      Height          =   500
      Left            =   1200
      TabIndex        =   10
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   500
      Left            =   2280
      TabIndex        =   8
      Top             =   6360
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   5040
      Width           =   11655
      Begin VB.TextBox WEntregado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   37
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox WObser 
         Height          =   285
         Left            =   7920
         MaxLength       =   50
         TabIndex        =   32
         Top             =   720
         Width           =   3495
      End
      Begin MSMask.MaskEdBox WFecha 
         Height          =   300
         Left            =   6840
         TabIndex        =   21
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.TextBox WCantidad 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   19
         Text            =   " "
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Text            =   " "
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Entregado"
         Height          =   255
         Left            =   5880
         TabIndex        =   38
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   7920
         TabIndex        =   31
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F.Entrega"
         Height          =   255
         Left            =   6840
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   4800
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Materia Prima"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label WDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   1800
         TabIndex        =   6
         Top             =   720
         Width           =   3015
      End
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   500
      Left            =   120
      TabIndex        =   4
      Top             =   6960
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3735
      Left            =   240
      OleObjectBlob   =   "solic.frx":0000
      TabIndex        =   3
      Top             =   1200
      Width           =   11655
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   8160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1620
      ItemData        =   "solic.frx":09EA
      Left            =   4800
      List            =   "solic.frx":09F1
      TabIndex        =   1
      Top             =   6720
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   500
      Left            =   1200
      TabIndex        =   0
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Solicitante"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Planta"
      Height          =   255
      Left            =   5760
      TabIndex        =   27
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Observaciones"
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro de Solicitud"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgSolic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 6 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private WAnterior As Integer
Private Cantidad As Single
Private XCantidad As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstSolic As Recordset
Dim spSolic As String
Dim XParam As String
Dim Vector(100, 2) As String
Private TipoConsulta As String
Private XVector(3, 4) As String
Private Auxi As String
Private WAuxi As String
Private WSaldo As Double
Private Desdelugar As Integer
Private WGraba As String
Private WLpt1 As String
Private WLpt2 As String
Private WSolicitud As String

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
    
    DBGrid1.Col = 5
    DBGrid1.Text = ""
    
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WEntregado.Text = ""
    WFecha.Text = "  /  /    "
    WObser.Text = ""
    WLinea.Text = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstAuxiliar
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    PrgSolic.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Consulta_Click()

    TipoConsulta = "0"

     Opcion.Clear

     Opcion.AddItem "Articulos"

     Opcion.Visible = True
     
 End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Reimpresion_Click()
    T$ = "Solicitud de Orden de Compra"
    m$ = "Desea Imprimir la Solicitud de Compra"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Rem SeleccionPuerto.Visible = True
        Call Impresion
    End If
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
            Ayuda.Visible = False
            spArticulo = "ListaArticulo"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstArticulo!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstArticulo.Close
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 10 Then
        WLinea.Text = DBGrid1.Row + 1
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -   -   "
        WLinea.Text = ""
    End If
    
    DBGrid1.Col = 1
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 2
    If Val(DBGrid1.Text) <> 0 Then
        WCantidad.Text = DBGrid1.Text
            Else
        WCantidad.Text = ""
    End If
    
    DBGrid1.Col = 3
    If Val(DBGrid1.Text) <> 0 Then
        WEntregado.Text = DBGrid1.Text
            Else
        WEntregado.Text = ""
    End If
    
    DBGrid1.Col = 4
    If DBGrid1.Text <> "" Then
        WFecha.Text = DBGrid1.Text
    End If
    
    DBGrid1.Col = 5
    WObser.Text = DBGrid1.Text
    
    WArticulo.SetFocus

End Sub

Private Sub Graba_Click()

    On Error GoTo WError
    
    If WGraba <> "S" Then
        Call Ingresa_clave
            Else
            
        WGraba = ""
    
        If Val(Solicitud.Text) = 0 Then
            spSolic = "ListaSolicitudNumero"
            Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
            If rstSolic.RecordCount > 0 Then
                With rstSolic
                    .MoveLast
                    Solicitud.Text = rstSolic!Solicitud + 1
                End With
                rstSolic.Close
            End If
        End If
        
        If Val(Solicitud.Text) = 0 Then
            Solicitud.Text = "1"
        End If
    
        Renglon = Renglon + 1
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
    
        DBGrid1.Col = 0
        DBGrid1.Text = ""

        Rem Borra la solicitud original
    
        spSolic = "BorrarSolicitudTotal " + "'" + Solicitud.Text + "'"
        Set rstSolic = db.OpenRecordset(spSolic, dbOpenDynaset, dbSQLPassThrough)
        
        Renglon = 0
        
        DBGrid1.Refresh
        
        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Articulo = UCase(DBGrid1.Text)
            
                DBGrid1.Col = 2
                Cantidad = Val(DBGrid1.Text)
                XCantidad = DBGrid1.Text
                
                DBGrid1.Col = 3
                Entregado = Val(DBGrid1.Text)
                XEntregado = DBGrid1.Text
                    
                DBGrid1.Col = 4
                Fecha1 = DBGrid1.Text
            
                DBGrid1.Col = 5
                Obser = DBGrid1.Text
                    
                If Articulo <> "" Then
            
                    Renglon = Renglon + 1
            
                    WSolicitud = Solicitud.Text
                    WRenglon = Str$(Renglon)
                    WFecha = Fecha.Text
                    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    WObservaciones = Observaciones.Text
                    WArticulo = Articulo
                    WCantidad = XCantidad
                    WEntrega = Fecha1
                    WObser = Obser
                    WOrdEntrega = Right$(Fecha1, 4) + Mid$(Fecha1, 4, 2) + Left$(Fecha1, 2)
                    WPlanta = Planta.Text
                    WSolicitante = Solicitante.Text
                    WDate = Date$
                    WMarca = ""
                    WEntregado = XEntregado
                
                    Auxi1 = WSolicitud
                    Auxi = WRenglon
                    Call Ceros(Auxi1, 6)
                    Call Ceros(Auxi, 2)
                    WClave = Auxi1 + Auxi
                         
                    XParam = "'" + WClave + "','" _
                         + WSolicitud + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WFechaord + "','" _
                         + WObservaciones + "','" _
                         + WArticulo + "','" _
                         + WCantidad + "','" _
                         + WEntrega + "','" _
                         + WOrdEntrega + "','" _
                         + WPlanta + "','" _
                         + WSolicitante + "','" _
                         + WDate + "','" _
                         + WMarca + "','" _
                         + WObser + "','" _
                         + WEntregado + "'"
                         
                    spSolic = "AltaSolicitud " + XParam
                    Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                        
            Next iRow
            
        Next a
                
        T$ = "Solicitud de Orden de Compra"
        m$ = "Desea Imprimir la Solicitud de Compra"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            Rem SeleccionPuerto.Visible = True
            Call Impresion
        End If
        
        Call Limpia_Click

        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
    
        Solicitud.SetFocus
    
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WEntregado.Text = ""
    WFecha.Text = "  /  /    "
    WObser.Text = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub Limpia_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WEntregado.Text = ""
    WFecha.Text = "  /  /    "
    WObser.Text = ""

    Solicitud.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Observaciones.Text = ""
    Select Case Val(WEmpresa)
        Case 1
            Planta.Text = "SI"
        Case 2
            Planta.Text = "PI"
        Case 3
            Planta.Text = "SII"
        Case 4
            Planta.Text = "PII"
        Case 5
            Planta.Text = "SIII"
        Case 6
            Planta.Text = "SIV"
        Case 7
            Planta.Text = "SV"
        Case 8
            Planta.Text = "PV"
        Case 9
            Planta.Text = "PVI"
        Case 10
            Planta.Text = "SVI"
        Case 11
            Planta.Text = "SVII"
        Case Else
            Planta.Text = "SI"
    End Select
    Solicitante.Text = ""

    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 5
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    Solicitud.Text = ""
    
    Rem spSolic = "ListaSolicitudNumero"
    Rem Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstSolic.RecordCount > 0 Then
    Rem     With rstSolic
    Rem         .MoveLast
    Rem         Solicitud.Text = rstSolic!Solicitud + 1
    Rem     End With
    Rem     rstSolic.Close
    Rem End If
    
    DBGrid1.FirstRow = 0
    Renglon = 0

    Solicitante.SetFocus

End Sub



Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ingre = "N"
        WArticulo.Text = UCase(WArticulo.Text)
        spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WDescripcion.Caption = rstArticulo!Descripcion
            Ingre = "S"
            rstArticulo.Close
                Else
            WArticulo.SetFocus
        End If
        If Ingre = "S" Then
            WCantidad.SetFocus
        End If
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        WEntregado.Text = "0"
        WFecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Auxi = "S" Then
            WObser.SetFocus
                Else
            WFecha.SetFocus
        End If
    End If
End Sub


Private Sub WObser_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Alta_Vector
        Call Ingresa_Click
        WArticulo.SetFocus
    End If
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
            
            If WPasa = "S" Then
                WArticulo.Text = rstArticulo!Codigo
                WDescripcion.Caption = rstArticulo!Descripcion
                    
                DBGrid1.Col = 0
                DBGrid1.Text = rstArticulo!Codigo
                DBGrid1.Col = 1
                DBGrid1.Text = rstArticulo!Descripcion
                    
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
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4, 5, 6, 7
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
ReDim UserData(0 To 5, 0 To 40)

mTotalRows& = 40

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
For i = 0 To 5
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1100
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Comprado"
             DBGrid1.Columns(newcnt).Width = 1100
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 4
             DBGrid1.Columns(newcnt).Caption = "F.Entrega"
             DBGrid1.Columns(newcnt).Width = 1100
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Observaciones"
             DBGrid1.Columns(newcnt).Width = 3000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    WLinea.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WEntregado.Text = ""
    WFecha.Text = "  /  /    "
    WObser.Text = ""

    Solicitud.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Observaciones.Text = ""
    Select Case Val(WEmpresa)
        Case 1
            Planta.Text = "SI"
        Case 2
            Planta.Text = "PI"
        Case 3
            Planta.Text = "SII"
        Case 4
            Planta.Text = "PII"
        Case 5
            Planta.Text = "SIII"
        Case 6
            Planta.Text = "SIV"
        Case 7
            Planta.Text = "SV"
        Case 8
            Planta.Text = "PV"
        Case 9
            Planta.Text = "PVI"
        Case 10
            Planta.Text = "SVI"
        Case 11
            Planta.Text = "SVII"
        Case Else
            Planta.Text = "SI"
    End Select
    
    Solicitante.Text = ""
    Solicitud.Text = ""
 
    Rem spSolic = "ListaSolicitudNumero"
    Rem Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstSolic.RecordCount > 0 Then
    Rem     With rstSolic
    Rem         .MoveLast
    Rem         Solicitud.Text = rstSolic!Solicitud + 1
    Rem     End With
    Rem     rstSolic.Close
    Rem End If

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgSolic.Caption = "Ingreso de Solicitud de Compras :  " + !Nombre
        End If
    End With
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Solicitud.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 5
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    Erase Vector
    
    spSolic = "ListaSolicitud " + "'" + Solicitud.Text + "'"
    Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
    If rstSolic.RecordCount > 0 Then
            
        With rstSolic
            .MoveFirst
            Do
                If .EOF = False Then
            
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstSolic!Articulo
                    Auxi1 = rstSolic!Articulo
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.##", rstSolic!Cantidad)
                    
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", rstSolic!Entregado)
                
                    DBGrid1.Col = 4
                    DBGrid1.Text = rstSolic!Entrega
                    
                    DBGrid1.Col = 5
                    DBGrid1.Text = rstSolic!Obser
                
                    Vector(Renglon, 1) = Auxi1
                
            
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstSolic.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
        
        Auxi1 = Vector(Renglon, 1)
    
        spArticulo = "ConsultaArticulo " + "'" + Auxi1 + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DBGrid1.Col = 1
            DBGrid1.Text = rstArticulo!Descripcion
            WArticulo.SetFocus
            rstArticulo.Close
        End If
    Next Da
    
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
            DBGrid1.Text = Pusing("###,###.##", WEntregado.Text)
                
            DBGrid1.Col = 4
            DBGrid1.Text = WFecha.Text
            
            DBGrid1.Col = 5
            DBGrid1.Text = WObser.Text
            
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
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
            DBGrid1.Text = Pusing("###,###.##", WEntregado.Text)
            
            DBGrid1.Col = 4
            DBGrid1.Text = WFecha.Text
            
            DBGrid1.Col = 5
            DBGrid1.Text = WObser.Text
            
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Solicitud_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi = Solicitud.Text
        Call Ceros(Auxi, 6)
        WClave = Auxi + "01"
            
        Entra = "N"
        spSolic = "ConsultaSolicitud1 " + "'" + Solicitud.Text + "'"
        Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
        If rstSolic.RecordCount > 0 Then
            Fecha.Text = rstSolic!Fecha
            Observaciones.Text = rstSolic!Observaciones
            Planta.Text = rstSolic!Planta
            Solicitante.Text = rstSolic!Solicitante
            rstSolic.Close
            Entra = "S"
                Else
            WSolicitud = Solicitud.Text
            Call Limpia_Click
            Solicitud.Text = WSolicitud
            Fecha.SetFocus
        End If
        
        If Entra = "S" Then
            Call Proceso_Click
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Solicitante.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Solicitante_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
End Sub

Private Sub Observaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WArticulo.SetFocus
    End If
End Sub

Private Sub ImpresionLpt()

        Solicitud.Text = WSolicitud
        Auxi = WSolicitud
        Call Ceros(Auxi, 6)
        WClave = Auxi + "01"
            
        Entra = "N"
        spSolic = "ConsultaSolicitud1 " + "'" + WSolicitud + "'"
        Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
        If rstSolic.RecordCount > 0 Then
            Fecha.Text = rstSolic!Fecha
            Observaciones.Text = rstSolic!Observaciones
            Planta.Text = rstSolic!Planta
            Solicitante.Text = rstSolic!Solicitante
            rstSolic.Close
            Entra = "S"
        End If
        
        If Entra = "S" Then
            Call Proceso_Click
        End If

        If WLpt1 = "S" Then
            Rem Open "dada.txt" For Output As #1
            Open "lpt1" For Output As #1
        End If
        
        If WLpt2 = "S" Then
            Open "lpt2" For Output As #1
        End If
        
        WObservaciones = Left$(Observaciones.Text + Space$(100), 100)
        
        With rstEmpresa
            .Index = "Empresa"
            Claveven$ = WEmpresa
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Impretit = !Nombre
                    Else
                Impretit = ""
            End If
        End With
    
        For Ci = 1 To 2

        '  Copia 1
        
        Print #1, Chr$(18)
        Print #1, ""
        Print #1, ""

        Print #1, Tab(1); "--------------------------------------------------------------------------------"
        
        Print #1, Tab(1); "|";
        Print #1, Impretit;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|                                                                              |"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Solicitud......: ";
        Print #1, Tab(25); Alinea("######", Solicitud.Text);
        Print #1, Tab(50); "Fecha : "; Fecha.Text;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Observaciones..:"; Tab(25); Left$(WObservaciones, 50);
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(25); Right$(WObservaciones, 50);
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Planta.........:"; Tab(25); Planta.Text;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Solicitante....:"; Tab(25); Solicitante.Text;
        Print #1, Tab(80); "|"
        
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, "|Producto  |  Descripcion      |Cantidad|Fecha Ent.|  Observaciones            |"
        Print #1, "--------------------------------------------------------------------------------"

        WCantidad = 0
        Valor = 0
        
        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Articulo = UCase(DBGrid1.Text)
                    
                DBGrid1.Col = 2
                Cantidad = Val(DBGrid1.Text)
                    
                If Left$(Articulo, 2) <> "" And Left$(Articulo, 2) <> Space$(2) And Cantidad <> 0 Then
                
                        DBGrid1.Col = 1
                        WDescripcion = DBGrid1.Text
                    
                        DBGrid1.Col = 4
                        Fecha = DBGrid1.Text
            
                        DBGrid1.Col = 5
                        Obser = DBGrid1.Text

                        WCantidad = WCantidad + 1

                        Print #1, Tab(1); "|"; Articulo;
                        Print #1, Tab(12); "|"; Left$(WDescripcion, 15);
                        Print #1, Tab(32); "|"; Alinea("###,###", Str$(Cantidad));
                        Print #1, Tab(41); "|"; Fecha;
                        Print #1, Tab(52); "|"; Left$(Obser, 25);
                        Print #1, Tab(80); "|"

                End If
                                        
            Next iRow
        Next a

        For Ciclo = WCantidad To 15
            Print #1, "|          |                   |        |          |                           |"
        Next Ciclo

        Print #1, "--------------------------------------------------------------------------------"
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""

        Next Ci
        
        Close #1

        Call Limpia_Click

        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
    
        Solicitud.SetFocus

End Sub

Private Sub Impresion()

    Listado.WindowTitle = "Solicitud de Compra"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Solic.Solicitud} in " + Solicitud.Text + " to " + Solicitud.Text
    Listado.SelectionFormula = "{Solic.Solicitud} in " + Solicitud.Text + " to " + Solicitud.Text
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Solic.Solicitud, Solic.Renglon, Solic.Fecha, Solic.Observaciones, Solic.Articulo, Solic.Cantidad, Solic.Entrega, Solic.Planta, Solic.Solicitante, Solic.Obser, " _
                + "Articulo.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Solic Solic, " _
                + DSQ + ".dbo.Articulo Articulo " _
                + "Where " _
                + "Solic.Articulo = Articulo.Codigo AND " _
                + "Solic.Solicitud >= " + Solicitud.Text + " AND " _
                + "Solic.Solicitud <= " + Solicitud.Text
    
    Listado.Connect = Connect()
    
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
        Listado.ReportFileName = "ImpreSolicSurfa.rpt"
        Listado.CopiesToPrinter = 1
        Listado.Action = 1
        Listado.CopiesToPrinter = 1
            Else
        Listado.ReportFileName = "ImpreSolicPelli.rpt"
        Listado.CopiesToPrinter = 2
        Listado.Action = 1
        Listado.CopiesToPrinter = 1
    End If
    


End Sub
 
Private Sub aYUDA_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spProveedor = "ListaProveedoresOrd"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
    With RstProveedor
        .MoveFirst
        Do
            If .EOF = False Then
            
                Da = Len(RstProveedor!Nombre) - WEspacios
                
                For aa = 1 To Da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                        Auxi = Str$(RstProveedor!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + "    " + RstProveedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
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
    RstProveedor.Close
    
    End If

End Sub

Sub Ingresa_clave()

    WClave.Text = ""
    Clave.Visible = True
    WClave.SetFocus
    
End Sub

Private Sub CancelaGraba_Click()

    Clave.Visible = False
    Solicitud.SetFocus

End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WGraba = "N"
        If WClave.Text = "SOL" Then
            WGraba = "S"
            Clave.Visible = False
            Call Graba_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Solicitud de Orden de Compra")
            WClave.SetFocus
        End If
    End If

End Sub

Private Sub WPuerto1_Click()
    SeleccionPuerto.Visible = False
    WLpt1 = "S"
    WLpt2 = "N"
    Call Impresion
End Sub

Private Sub WPuerto2_Click()
    SeleccionPuerto.Visible = False
    WLpt1 = "N"
    WLpt2 = "S"
    Call Impresion
End Sub
