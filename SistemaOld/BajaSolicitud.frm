VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgBajaSolicitud 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizacion de Solicitud de Produccion"
   ClientHeight    =   7410
   ClientLeft      =   465
   ClientTop       =   570
   ClientWidth     =   11160
   LinkTopic       =   "Form2"
   ScaleHeight     =   7410
   ScaleWidth      =   11160
   Visible         =   0   'False
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
      Index           =   6
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   " "
      Top             =   2040
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
      Index           =   5
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   " "
      Top             =   2040
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   " "
      Top             =   2040
      Width           =   375
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
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   15
      Top             =   480
      Width           =   5295
   End
   Begin VB.TextBox Codigo 
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   1095
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   " "
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox WTexto1 
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
      Left            =   1800
      TabIndex        =   10
      Top             =   1320
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto2 
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
      Left            =   1200
      TabIndex        =   8
      Top             =   1320
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1920
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1920
      Width           =   375
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
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   6855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10560
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
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
      Height          =   1260
      Left            =   2280
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   5640
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
      Height          =   2460
      ItemData        =   "BajaSolicitud.frx":0000
      Left            =   120
      List            =   "BajaSolicitud.frx":0007
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   1320
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
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   3375
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5953
      _Version        =   393216
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   4800
      TabIndex        =   16
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
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
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   1335
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   10440
      MouseIcon       =   "BajaSolicitud.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "BajaSolicitud.frx":031F
      ToolTipText     =   "Salida"
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   10440
      MouseIcon       =   "BajaSolicitud.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "BajaSolicitud.frx":0E6B
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   10440
      MouseIcon       =   "BajaSolicitud.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "BajaSolicitud.frx":19B7
      ToolTipText     =   "Consulta de Datos"
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   10440
      MouseIcon       =   "BajaSolicitud.frx":21F9
      MousePointer    =   99  'Custom
      Picture         =   "BajaSolicitud.frx":2503
      ToolTipText     =   "Limpia la pantalla"
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Nro. Ingreso"
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
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgBajaSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstBajaSolicitud As Recordset
Dim spBajaSolicitud As String
Dim rstCargaSolicitud As Recordset
Dim spCargaSolicitud As String

Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Cantidad As Double
Dim XPaso As String
Dim Renglon As Integer
Dim ZVector(100, 3) As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Private Sub cmdClose1_Click()

    Call Limpia_Click
    PrgBajaSolicitud.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()

    Erase ZVector
    ZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM BajaSolicitud"
    ZSql = ZSql + " Where BajaSolicitud.Codigo = " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by BajaSolicitud.Clave"
    
    spBajaSolicitud = ZSql
    Set rstBajaSolicitud = db.OpenRecordset(spBajaSolicitud, dbOpenSnapshot, dbSQLPassThrough)
    If rstBajaSolicitud.RecordCount > 0 Then
        With rstBajaSolicitud
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZLugar = ZLugar + 1
                    
                    ZVector(ZLugar, 1) = rstBajaSolicitud!Articulo
                    ZVector(ZLugar, 2) = Str$(rstBajaSolicitud!Cantidad)
                    ZVector(ZLugar, 3) = rstBajaSolicitud!ClaveCarga
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstBajaSolicitud.Close
    End If
    
    For Ciclo = 1 To ZLugar

        ZArticulo = ZVector(Ciclo, 1)
        ZCantidad = ZVector(Ciclo, 2)
        ZClaveCarga = ZVector(Ciclo, 3)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CargaSolicitud SET "
        ZSql = ZSql + " Entregado = Entregado - " + "'" + ZCantidad + "'"
        ZSql = ZSql + " Where Clave = " + "'" + ZClaveCarga + "'"
            
        spCargaSolicitud = ZSql
        Set rstCargaSolicitud = db.OpenRecordset(spCargaSolicitud, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo







    Sql1 = "DELETE BajaSolicitud"
    Sql2 = " Where Codigo = " + "'" + Codigo.Text + "'"
    spBajaSolicitud = Sql1 + Sql2
    Set rstBajaSolicitud = db.OpenRecordset(spBajaSolicitud, dbOpenSnapshot, dbSQLPassThrough)
    
    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
        
        WVector1.Row = iRow
            
        WVector1.Col = 1
        PTerminado = WVector1.Text
        
        WVector1.Col = 3
        XSolicitud = WVector1.Text
        
        WVector1.Col = 5
        XCantidad = WVector1.Text
            
        If PTerminado <> "" Or XCantidad <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    WRenglon = 0
    For iRow = 1 To HastaRenglon
        
        WVector1.Row = iRow
            
        WVector1.Col = 1
        ZArticulo = WVector1.Text
        
        WVector1.Col = 3
        ZSolicitud = WVector1.Text
        
        WVector1.Col = 5
        ZCantidad = WVector1.Text
        
        WVector1.Col = 6
        ZClaveCarga = WVector1.Text
        
        If Val(ZCantidad) <> 0 Then
        
            WRenglon = WRenglon + 1
            Auxi = Str$(WRenglon)
            Call Ceros(Auxi, 2)
        
            Auxi1 = Str$(Codigo.Text)
            Call Ceros(Auxi1, 6)
        
            WClave = Auxi1 + Auxi
            ZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO BajaSolicitud ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Solicitud ,"
            ZSql = ZSql + "ClaveCarga )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + Str$(WRenglon) + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + ZOrdFecha + "',"
            ZSql = ZSql + "'" + Observaciones.Text + "',"
            ZSql = ZSql + "'" + ZArticulo + "',"
            ZSql = ZSql + "'" + ZCantidad + "',"
            ZSql = ZSql + "'" + ZSolicitud + "',"
            ZSql = ZSql + "'" + ZClaveCarga + "')"
            
            spBajaSolicitud = ZSql
            Set rstBajaSolicitud = db.OpenRecordset(spBajaSolicitud, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaSolicitud SET "
            ZSql = ZSql + " Entregado = Entregado + " + "'" + ZCantidad + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZClaveCarga + "'"
            
            spCargaSolicitud = ZSql
            Set rstCargaSolicitud = db.OpenRecordset(spCargaSolicitud, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
            
    Next iRow
    
    Call Limpia_Click

    WVector1.Col = 1
    WVector1.Row = 1
        
    Codigo.SetFocus
        
End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector

    Codigo.Text = ""
    Observaciones.Text = ""
    Fecha.Text = "  /  /    "
    
    Renglon = 0
    Graba.Enabled = True
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM BajaSolicitud"
    spBajaSolicitud = Sql1 + Sql2
    Set rstBajaSolicitud = db.OpenRecordset(spBajaSolicitud, dbOpenSnapshot, dbSQLPassThrough)
    If rstBajaSolicitud.RecordCount > 0 Then
        rstBajaSolicitud.MoveLast
        WCodigo = IIf(IsNull(rstBajaSolicitud!CodigoMayor), "0", rstBajaSolicitud!CodigoMayor)
        Codigo.Text = Str$(Val(WCodigo) + 1)
        rstBajaSolicitud.Close
            Else
        Codigo.Text = "1"
    End If
    
    WVector1.Col = 1
    WVector1.Row = 1
    
    Codigo.SetFocus

End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    WVector1.Col = 1
    WVector1.Row = 1

    Codigo.Text = ""
    Observaciones.Text = ""
    Fecha.Text = "  /  /    "
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM BajaSolicitud"
    spBajaSolicitud = Sql1 + Sql2
    Set rstBajaSolicitud = db.OpenRecordset(spBajaSolicitud, dbOpenSnapshot, dbSQLPassThrough)
    If rstBajaSolicitud.RecordCount > 0 Then
        rstBajaSolicitud.MoveLast
        WCodigo = IIf(IsNull(rstBajaSolicitud!CodigoMayor), "0", rstBajaSolicitud!CodigoMayor)
        Codigo.Text = Str$(Val(WCodigo) + 1)
        rstBajaSolicitud.Close
            Else
        Codigo.Text = "1"
    End If
    
    Renglon = 0
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


Private Sub Proceso_Click()

    Call Limpia_Vector
    WRenglon = 0
    
    Sql1 = "Select BajaSolicitud.Clave, BajaSolicitud.Codigo, BajaSolicitud.Articulo, BajaSolicitud.Cantidad, BajaSolicitud.Solicitud, BajaSolicitud.ClaveCarga, Terminado.Descripcion"
    Sql2 = " FROM BajaSolicitud, Terminado"
    Sql3 = " Where BajaSolicitud.Codigo = " + "'" + Codigo.Text + "'"
    Sql4 = " and BajaSolicitud.Articulo = Terminado.Codigo"
    Sql5 = " Order by BajaSolicitud.Clave"
    
    spBajaSolicitud = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstBajaSolicitud = db.OpenRecordset(spBajaSolicitud, dbOpenSnapshot, dbSQLPassThrough)
    If rstBajaSolicitud.RecordCount > 0 Then
        With rstBajaSolicitud
            .MoveFirst
            Do
                If .EOF = False Then
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = Trim(rstBajaSolicitud!Articulo)
                    
                    WVector1.Col = 2
                    WVector1.Text = Trim(rstBajaSolicitud!Descripcion)
                    
                    WVector1.Col = 3
                    WVector1.Text = rstBajaSolicitud!Solicitud
                    
                    WVector1.Col = 4
                    WVector1.Text = ""
            
                    WVector1.Col = 5
                    WVector1.Text = rstBajaSolicitud!Cantidad
                    WVector1.Text = Pusing("######.##", WVector1.Text)
                    
                    WVector1.Col = 6
                    WVector1.Text = rstBajaSolicitud!ClaveCarga
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstBajaSolicitud.Close
    End If
    
    Graba.Enabled = True

End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM BajaSolicitud"
        Sql3 = " Where BajaSolicitud.Codigo = " + "'" + Codigo.Text + "'"
        spBajaSolicitud = Sql1 + Sql2 + Sql3
        Set rstBajaSolicitud = db.OpenRecordset(spBajaSolicitud, dbOpenSnapshot, dbSQLPassThrough)
        If rstBajaSolicitud.RecordCount > 0 Then
            Fecha.Text = rstBajaSolicitud!Fecha
            Observaciones.Text = rstBajaSolicitud!Observaciones
            rstBajaSolicitud.Close
            Call Proceso_Click
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                Else
            Graba.Enabled = True
            WCodigo = Codigo.Text
            Call Limpia_Click
            Codigo.Text = WCodigo
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Observaciones.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Observaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Lee_Solicitud
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub Lee_Solicitud()

    Sql1 = "UPDATE CargaSolicitud SET "
    Sql2 = " Saldo =  0"
    spCargaSolicitud = Sql1 + Sql2
    Set rstCargaSolicitud = db.OpenRecordset(spCargaSolicitud, dbOpenSnapshot, dbSQLPassThrough)
    
    Sql1 = "UPDATE CargaSolicitud SET "
    Sql2 = " Saldo = Cantidad - Entregado"
    spCargaSolicitud = Sql1 + Sql2
    Set rstCargaSolicitud = db.OpenRecordset(spCargaSolicitud, dbOpenSnapshot, dbSQLPassThrough)

    Call Limpia_Vector
    WRenglon = 0

    Sql1 = "Select CargaSolicitud.Clave, CargaSolicitud.Solicitud, CargaSolicitud.Articulo, CargaSolicitud.Cantidad,CargaSolicitud.Saldo, Terminado.Descripcion"
    Sql2 = " FROM CargaSolicitud, Terminado"
    Sql3 = " Where Saldo > 0 "
    Sql4 = " and CargaSolicitud.Articulo = Terminado.Codigo"
    Sql5 = " Order by CargaSolicitud.Articulo"
    
    spCargaSolicitud = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstCargaSolicitud = db.OpenRecordset(spCargaSolicitud, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSolicitud.RecordCount > 0 Then
        With rstCargaSolicitud
            .MoveFirst
            Do
                If .EOF = False Then
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = Trim(rstCargaSolicitud!Articulo)
                    
                    WVector1.Col = 2
                    WVector1.Text = Trim(rstCargaSolicitud!Descripcion)
                    
                    WVector1.Col = 3
                    WVector1.Text = rstCargaSolicitud!Solicitud
            
                    WVector1.Col = 4
                    WVector1.Text = Str$(rstCargaSolicitud!Saldo)
                    WVector1.Text = Pusing("######.##", WVector1.Text)
            
                    WVector1.Col = 6
                    WVector1.Text = rstCargaSolicitud!Clave
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaSolicitud.Close
    End If

End Sub




Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1
        Case 113
            WTexto1.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1
        Case 113
            WTexto2.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto3.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 5
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM Terminado"
            Sql3 = " Where Terminado.Codigo = " + "'" + WVector1.Text + "'"
            spTerminado = Sql1 + Sql2 + Sql3
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WVector1.Col = 2
                If WVector1.Text = "" Then
                    WVector1.Text = Trim(rstTerminado!Descripcion)
                End If
                WVector1.Col = 4
                rstTerminado.Close
            End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
    
    RenglonAuxiliar = WVector1.Row

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
        
        WVector1.Row = iRow
            
        WVector1.Col = 1
        PTerminado = WVector1.Text
        
        WVector1.Col = 5
        XCantidad = WVector1.Text
            
        If PTerminado <> "" Or XCantidad <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For Da = 1 To WVector1.Cols - 1
            WVector1.Col = Da
            WVector1.Text = WBorra(Ciclo, Da)
        Next Da
    Next Ciclo
    
    End If
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 7
    WVector1.FixedRows = 1
    WVector1.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "P.Terminado"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 3500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 70
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Solicitud"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 6
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Pedida"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "######.##"
            Case 5
                WVector1.Text = "Recibida"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "######.##"
            Case 6
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tama�o de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub



