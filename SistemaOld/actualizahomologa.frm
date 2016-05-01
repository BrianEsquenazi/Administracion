VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgActualizaHomologa 
   Caption         =   "Respuesta de Laboratorio"
   ClientHeight    =   5865
   ClientLeft      =   1665
   ClientTop       =   405
   ClientWidth     =   9165
   LinkTopic       =   "Form2"
   ScaleHeight     =   5865
   ScaleWidth      =   9165
   Begin VB.ComboBox TipoMp 
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
      Left            =   4320
      TabIndex        =   32
      Top             =   120
      Width           =   2055
   End
   Begin VB.ComboBox Trazabilidad 
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
      Left            =   2760
      TabIndex        =   28
      Top             =   3000
      Width           =   2055
   End
   Begin VB.ComboBox Senasa 
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
      Left            =   2760
      TabIndex        =   27
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   3000
      TabIndex        =   23
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   25
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   26
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.ComboBox Estado 
      Height          =   315
      Left            =   5880
      TabIndex        =   21
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox ComentariosII 
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
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   19
      Text            =   " "
      Top             =   2640
      Width           =   6015
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
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   17
      Text            =   " "
      Top             =   1200
      Width           =   6015
   End
   Begin VB.TextBox ResultadoEntrega 
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
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   12
      Text            =   " "
      Top             =   2280
      Width           =   6015
   End
   Begin VB.TextBox Responsable 
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
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   10
      Text            =   " "
      Top             =   1560
      Width           =   3135
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
      Left            =   2880
      TabIndex        =   9
      Top             =   4560
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
      Left            =   5400
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Unidad 
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
      Left            =   2760
      MaxLength       =   15
      TabIndex        =   6
      Text            =   " "
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Resultado 
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
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   4
      Text            =   " "
      Top             =   840
      Width           =   6015
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   120
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
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8280
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox Articulo 
      Height          =   285
      Left            =   2760
      TabIndex        =   14
      Top             =   1920
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
   Begin MSMask.MaskEdBox VtoSenasa 
      Height          =   285
      Left            =   4920
      TabIndex        =   29
      Top             =   3360
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
   Begin VB.Label Label13 
      Caption         =   "Insc. Senasa"
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
      Height          =   495
      Left            =   120
      TabIndex        =   31
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label14 
      Caption         =   "Trazabilidad"
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
      Left            =   120
      TabIndex        =   30
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Resultado"
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
      Left            =   4320
      TabIndex        =   22
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Comentarios Laboratorio"
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
      TabIndex        =   20
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
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
      TabIndex        =   18
      Top             =   1200
      Width           =   2055
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
      Top             =   1920
      Width           =   1815
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
      Left            =   4440
      TabIndex        =   15
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label5 
      Caption         =   "Resultado 1ra Entrega"
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
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Responsable"
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
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Unidad de Negocio"
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
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Resultado"
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
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label15 
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   2
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "PrgActualizaHomologa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstHomologa As Recordset
Dim spHomologa As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim XParam As String
Dim WActualiza As String
Dim WGraba As String
Dim WTipomov As String
Dim XIndice As Integer
Dim ZEstado As Integer


Private Sub cmdGraba_Click()

    If Estado.ListIndex = 0 Then
        m$ = "Se debe informar el codigo de estado"
        G% = MsgBox(m$, 0, "Actualizacion de Homologacion de Muestras")
        Exit Sub
    End If

    If Estado.ListIndex = 1 Then
        Articulo.Text = UCase(Articulo.Text)
        spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            rstArticulo.Close
                Else
            m$ = "Se debe informar el codigo de Materia Prima"
            G% = MsgBox(m$, 0, "Actualizacion de Homologacion de Muestras")
            Exit Sub
        End If
    End If
    
    If TipoMp.ListIndex = 0 Then
        m$ = "Se debe informar el si es M.P. o Envases"
        G% = MsgBox(m$, 0, "Homologacion de Muestras")
        Exit Sub
    End If
    
    If Estado.ListIndex = 1 Then
        Select Case TipoMp.ListIndex
            Case 1
                If Left$(UCase(Articulo.Text), 2) = "ZE" Then
                    m$ = "Se debe informar un codigo de materia prima y se informa un codigo de Envase"
                    G% = MsgBox(m$, 0, "Homologacion de Muestras")
                    Exit Sub
                End If
            Case 2
                If Left$(UCase(Articulo.Text), 2) <> "ZE" Then
                    m$ = "Se debe informar un codigo de envase y se informa un codigo de Materia Prima"
                    G% = MsgBox(m$, 0, "Homologacion de Muestras")
                    Exit Sub
                End If
            Case Else
        End Select
    End If
    
    If Estado.ListIndex = 1 Then
        If TipoMp.ListIndex = 2 Then
            If Trazabilidad.ListIndex = 0 Then
                m$ = "Se debe informar si el envase posee trazabilidad"
                G% = MsgBox(m$, 0, "Homologacion de Muestras")
                Exit Sub
            End If
        End If
    End If
    
    If Estado.ListIndex = 1 Then
        If TipoMp.ListIndex = 2 Then
            If Senasa.ListIndex = 0 Then
                m$ = "Se debe informar si el proveedor esta inscripto en senasa"
                G% = MsgBox(m$, 0, "Homologacion de Muestras")
                Exit Sub
            End If
        End If
    End If
    
    If Estado.ListIndex = 1 Then
        If TipoMp.ListIndex = 2 Then
            If Senasa.ListIndex = 1 Then
                 Call Valida_fecha(VtoSenasa.Text, Auxi)
                 If Auxi = "N" Or VtoSenasa = "  /  /    " Or VtoSenasa.Text = "00/00/0000" Then
                    m$ = "Se debe informar la fecha de vencimiento en el senasa"
                    G% = MsgBox(m$, 0, "Homologacion de Muestras")
                    Exit Sub
                End If
            End If
        End If
    End If
    
    
    If WGraba <> "S" Then
    
        Call Ingresa_clave

               Else
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Homologa SET "
        ZSql = ZSql + "FechaII =  " + "'" + Fecha.Text + "',"
        ZSql = ZSql + "Unidad =  " + "'" + Unidad.Text + "',"
        ZSql = ZSql + "Resultado =  " + "'" + Resultado.Text + "',"
        ZSql = ZSql + "Observaciones =  " + "'" + Observaciones.Text + "',"
        ZSql = ZSql + "Responsable =  " + "'" + Responsable.Text + "',"
        ZSql = ZSql + "CodigoMp =  " + "'" + Articulo.Text + "',"
        ZSql = ZSql + "ResultadoEntrega =  " + "'" + ResultadoEntrega.Text + "',"
        ZSql = ZSql + "ComentariosII =  " + "'" + ComentariosII.Text + "',"
        ZSql = ZSql + "Estado =  " + "'" + Str$(Estado.ListIndex) + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WMuestra + "'"
            
        spHomologa = ZSql
        Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
            
        Call Conecta_Empresa
        
        Call cmdClose_Click
    
    End If
        
End Sub

Private Sub CmdLimpiar_Click()

    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Unidad.Text = ""
    Resultado.Text = ""
    Observaciones.Text = ""
    Responsable.Text = ""
    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
    ResultadoEntrega.Text = ""
    ComentariosII.Text = ""
    Estado.ListIndex = 0
    
    Fecha.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgActualizaHomologa.Hide
    Unload Me
    PrgHomologaProve.Show
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Unidad.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Unidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Resultado.SetFocus
    End If
    If KeyAscii = 27 Then
        Unidad.Text = ""
    End If
End Sub

Private Sub Resultado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
    If KeyAscii = 27 Then
        Resultado.Text = ""
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub Responsable_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Articulo.SetFocus
    End If
    If KeyAscii = 27 Then
        Responsable.Text = ""
    End If
End Sub

Sub Articulo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Articulo.Text = UCase(Articulo.Text)
        spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DesArticulo.Caption = rstArticulo!Descripcion
            rstArticulo.Close
            ResultadoEntrega.SetFocus
                Else
            Articulo.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Articulo.Text = "  -   -   "
        DesArticulo.Caption = ""
    End If
End Sub

Private Sub ResultadoEntrega_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ComentariosII.SetFocus
    End If
    If KeyAscii = 27 Then
        ResultadoEntrega.Text = ""
    End If
End Sub

Private Sub ComentariosII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentariosII.Text = ""
    End If
End Sub


Private Sub Form_Load()

    Estado.Clear
    
    Estado.AddItem ""
    Estado.AddItem "Aprobado"
    Estado.AddItem "Rechazado"
    
    Estado.ListIndex = 0

    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Unidad.Text = ""
    Resultado.Text = ""
    Observaciones.Text = ""
    Responsable.Text = ""
    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
    ResultadoEntrega.Text = ""
    ComentariosII.Text = ""
    VtoSenasa.Text = "  /  /    "
    
    TipoMp.Clear
    
    TipoMp.AddItem ""
    TipoMp.AddItem "M.P."
    TipoMp.AddItem "Envases"
    
    TipoMp.ListIndex = 0
    
    Trazabilidad.Clear
    
    Trazabilidad.AddItem ""
    Trazabilidad.AddItem "Si"
    Trazabilidad.AddItem "No"
    
    Trazabilidad.ListIndex = 0
    
    Senasa.Clear
    
    Senasa.AddItem ""
    Senasa.AddItem "Si"
    Senasa.AddItem "No"
    
    Senasa.ListIndex = 0
    
    WGraba = ""
    
    If Val(WMuestra) <> 0 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        Sql1 = "Select *"
        Sql2 = " FROM Homologa"
        Sql3 = " Where Homologa.Codigo = " + "'" + WMuestra + "'"
        spHomologa = Sql1 + Sql2 + Sql3
        Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
        If rstHomologa.RecordCount > 0 Then
            Fecha.Text = rstHomologa!Fecha
            Unidad.Text = Trim(rstHomologa!Unidad)
            Resultado.Text = Trim(rstHomologa!Resultado)
            Observaciones.Text = Trim(rstHomologa!Observaciones)
            Responsable.Text = Trim(rstHomologa!Responsable)
            Articulo.Text = rstHomologa!codigomp
            ResultadoEntrega.Text = Trim(rstHomologa!ResultadoEntrega)
            ComentariosII.Text = Trim(rstHomologa!ComentariosII)
            ZEstado = IIf(IsNull(rstHomologa!Estado), "0", rstHomologa!Estado)
            Estado.ListIndex = ZEstado
            
            ZTipoMp = IIf(IsNull(rstHomologa!TipoMp), "0", rstHomologa!TipoMp)
            ZTrazabilidad = IIf(IsNull(rstHomologa!Trazabilidad), "0", rstHomologa!Trazabilidad)
            ZSenasa = IIf(IsNull(rstHomologa!Senasa), "0", rstHomologa!Senasa)
            ZVtoSenasa = IIf(IsNull(rstHomologa!VtoSenasa), "  /  /    ", rstHomologa!VtoSenasa)
            
            TipoMp.ListIndex = ZTipoMp
            Trazabilidad.ListIndex = ZTrazabilidad
            Senasa.ListIndex = ZSenasa
            VtoSenasa.Text = ZVtoSenasa
            
            rstHomologa.Close
        End If
        
        Call Conecta_Empresa
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DesArticulo.Caption = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
    End If
        
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Unidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Resultado_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Responsable_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CodigoMp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub ResultadoEntrega_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub ComentariosII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call cmdGraba_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub

Sub Ingresa_clave()

    WClave.Text = ""
    XClave.Visible = True
    WClave.SetFocus
    
End Sub

Private Sub CancelaGraba_Click()

    XClave.Visible = False

End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WGraba = "N"
        WGrabaI = ""
        
        If TipoMp.ListIndex = 1 Then
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Operador"
            ZSql = ZSql + " Where Operador.Clave = " + "'" + WClave.Text + "'"
            spOperador = ZSql
            Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
            If rstOperador.RecordCount > 0 Then
                ZOperador = rstOperador!Operador
                WGrabaI = IIf(IsNull(rstOperador!GrabaI), "", rstOperador!GrabaI)
                rstOperador.Close
            End If
            
                Else
                
            If WClave.Text = "JUJUY" Then
                WGrabaI = "S"
            End If
            
        End If
        
        If WGrabaI = "S" Then
            WGraba = "S"
            XClave.Visible = False
            Call cmdGraba_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Actualizacion de Homologacion de Materias Primas")
            WClave.SetFocus
        End If
        
    End If
End Sub



