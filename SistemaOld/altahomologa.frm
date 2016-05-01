VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAltaHomologa 
   Caption         =   "Solicitud de Homologacion de Muestras de Materias Pruimas"
   ClientHeight    =   8415
   ClientLeft      =   135
   ClientTop       =   465
   ClientWidth     =   11640
   LinkTopic       =   "Form2"
   ScaleHeight     =   8415
   ScaleWidth      =   11640
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
      Left            =   2280
      TabIndex        =   38
      Top             =   3720
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
      Left            =   2280
      TabIndex        =   37
      Top             =   3360
      Width           =   2055
   End
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
      Left            =   8520
      TabIndex        =   34
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Ct 
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
      TabIndex        =   32
      Text            =   " "
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox Entregado 
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
      MaxLength       =   20
      TabIndex        =   30
      Text            =   " "
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox Comentarios 
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
      TabIndex        =   28
      Text            =   " "
      Top             =   2640
      Width           =   5895
   End
   Begin VB.TextBox Nombre 
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
      TabIndex        =   26
      Text            =   " "
      Top             =   2280
      Width           =   5895
   End
   Begin VB.TextBox Origen 
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
      TabIndex        =   24
      Text            =   " "
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Certificado 
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
      TabIndex        =   22
      Text            =   " "
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox EspecificacionesProve 
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
      MaxLength       =   30
      TabIndex        =   20
      Text            =   " "
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox Precio 
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
      TabIndex        =   18
      Text            =   " "
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Solicita 
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
      MaxLength       =   10
      TabIndex        =   16
      Text            =   " "
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox Material 
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
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   6015
   End
   Begin VB.TextBox DesProveedor 
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   14
      Top             =   840
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
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.TextBox Proveedor 
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
      MaxLength       =   11
      TabIndex        =   11
      Text            =   " "
      Top             =   840
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   8040
      TabIndex        =   10
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
      Mask            =   "##/##/####"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10440
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
      Height          =   2700
      ItemData        =   "altahomologa.frx":0000
      Left            =   120
      List            =   "altahomologa.frx":0007
      TabIndex        =   5
      Top             =   5520
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
      Left            =   4560
      TabIndex        =   4
      Top             =   4320
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
      Left            =   3240
      TabIndex        =   3
      Top             =   4320
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
      Left            =   5880
      TabIndex        =   2
      Top             =   4320
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
      Left            =   1920
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin MSMask.MaskEdBox VtoSenasa 
      Height          =   285
      Left            =   4440
      TabIndex        =   39
      Top             =   3720
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
      TabIndex        =   36
      Top             =   3360
      Width           =   2055
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
      TabIndex        =   35
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "C  /  T"
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
      TabIndex        =   33
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Entregado a"
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
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Comentarios"
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
      TabIndex        =   29
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Denominacion Comercial"
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
      TabIndex        =   27
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Origen"
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
      Left            =   5760
      TabIndex        =   25
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Certificado de Analisis"
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
      Left            =   5760
      TabIndex        =   23
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Especif. del Proveedor"
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
      TabIndex        =   21
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Precio "
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
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Solicita"
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
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Material"
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
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label4 
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
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1575
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
      Left            =   5880
      TabIndex        =   9
      Top             =   480
      Width           =   2415
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
Attribute VB_Name = "PrgAltaHomologa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstHomologa As Recordset
Dim spHomologa As String
Dim XParam As String
Dim XIndice As Integer

Dim ZTipoMp As Integer
Dim ZTrazabilidad As Integer
Dim ZSenasa As Integer
Dim ZVtoSenasa As String


Private Sub cmdGraba_Click()

    Sql1 = "Select *"
    Sql2 = " FROM Proveedor"
    Sql3 = " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
    spProveedor = Sql1 + Sql2 + Sql3
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        RstProveedor.Close
            Else
        m$ = "Codigo de Proveedor invalido"
        G% = MsgBox(m$, 0, "Homologacion de Muestras")
        Exit Sub
    End If
    
    If TipoMp.ListIndex = 0 Then
        m$ = "Se debe informar el si es M.P. o Envases"
        G% = MsgBox(m$, 0, "Homologacion de Muestras")
        Exit Sub
    End If

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

    If Val(WMuestra) <> 0 Then
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Homologa SET "
        ZSql = ZSql + "TipoMp =  " + "'" + Str$(TipoMp.ListIndex) + "',"
        ZSql = ZSql + "Trazabilidad =  " + "'" + Str$(Trazabilidad.ListIndex) + "',"
        ZSql = ZSql + "Senasa =  " + "'" + Str$(Senasa.ListIndex) + "',"
        ZSql = ZSql + "VtoSenasa =  " + "'" + VtoSenasa.Text + "',"
        ZSql = ZSql + "Material =  " + "'" + Material.Text + "',"
        ZSql = ZSql + "Solicita =  " + "'" + Solicita.Text + "',"
        ZSql = ZSql + "Fecha =  " + "'" + Fecha.Text + "',"
        ZSql = ZSql + "Proveedor =  " + "'" + Proveedor.Text + "',"
        ZSql = ZSql + "DesProveedor =  " + "'" + DesProveedor.Text + "',"
        ZSql = ZSql + "EspecificacionesProve =  " + "'" + EspecificacionesProve.Text + "',"
        ZSql = ZSql + "Certificado =  " + "'" + Certificado.Text + "',"
        ZSql = ZSql + "Precio =  " + "'" + Precio.Text + "',"
        ZSql = ZSql + "Origen =  " + "'" + Origen.Text + "',"
        ZSql = ZSql + "Ct =  " + "'" + Ct.Text + "',"
        ZSql = ZSql + "Nombre =  " + "'" + Nombre.Text + "',"
        ZSql = ZSql + "Comentarios =  " + "'" + Comentarios.Text + "',"
        ZSql = ZSql + "Entregado =  " + "'" + Entregado.Text + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WMuestra + "'"
        spHomologa = ZSql
        Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
        
        Call Conecta_Empresa
        
        Call cmdClose_Click

            Else

        WCodigo = 1
        
        Sql1 = "Select Max(Codigo) as [CodigoMayor]"
        Sql2 = " FROM Homologa"
        spHomologa = Sql1 + Sql2
        Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
        If rstHomologa.RecordCount > 0 Then
            rstHomologa.MoveLast
            ZUltimo = IIf(IsNull(rstHomologa!CodigoMayor), "0", rstHomologa!CodigoMayor)
            WCodigo = ZUltimo + 1
            rstHomologa.Close
        End If
    
        XCodigo = Str$(WCodigo)
        ZBlanco = ""
        ZBlancoII = "  -   -   "
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Homologa ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Material ,"
        ZSql = ZSql + "Solicita ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Proveedor ,"
        ZSql = ZSql + "DesProveedor ,"
        ZSql = ZSql + "EspecificacionesProve ,"
        ZSql = ZSql + "Certificado ,"
        ZSql = ZSql + "Precio ,"
        ZSql = ZSql + "Origen ,"
        ZSql = ZSql + "Ct ,"
        ZSql = ZSql + "Nombre ,"
        ZSql = ZSql + "Comentarios ,"
        ZSql = ZSql + "Entregado ,"
        ZSql = ZSql + "FechaII ,"
        ZSql = ZSql + "Unidad ,"
        ZSql = ZSql + "Resultado ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Responsable ,"
        ZSql = ZSql + "CodigoMp ,"
        ZSql = ZSql + "ResultadoEntrega ,"
        ZSql = ZSql + "TipoMp ,"
        ZSql = ZSql + "Trazabilidad ,"
        ZSql = ZSql + "Senasa ,"
        ZSql = ZSql + "VtoSenasa ,"
        ZSql = ZSql + "ComentariosII) "
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + XCodigo + "',"
        ZSql = ZSql + "'" + Material.Text + "',"
        ZSql = ZSql + "'" + Solicita.Text + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + Proveedor.Text + "',"
        ZSql = ZSql + "'" + DesProveedor.Text + "',"
        ZSql = ZSql + "'" + EspecificacionesProve.Text + "',"
        ZSql = ZSql + "'" + Certificado.Text + "',"
        ZSql = ZSql + "'" + Precio.Text + "',"
        ZSql = ZSql + "'" + Origen.Text + "',"
        ZSql = ZSql + "'" + Ct.Text + "',"
        ZSql = ZSql + "'" + Nombre.Text + "',"
        ZSql = ZSql + "'" + Comentarios.Text + "',"
        ZSql = ZSql + "'" + Entregado.Text + "',"
        ZSql = ZSql + "'" + ZBlanco + "',"
        ZSql = ZSql + "'" + ZBlanco + "',"
        ZSql = ZSql + "'" + ZBlanco + "',"
        ZSql = ZSql + "'" + ZBlanco + "',"
        ZSql = ZSql + "'" + ZBlanco + "',"
        ZSql = ZSql + "'" + ZBlancoII + "',"
        ZSql = ZSql + "'" + ZBlanco + "',"
        ZSql = ZSql + "'" + Str$(TipoMp.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Trazabilidad.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Senasa.ListIndex) + "',"
        ZSql = ZSql + "'" + VtoSenasa.Text + "',"
        ZSql = ZSql + "'" + ZBlanco + "')"
      
        spHomologa = ZSql
        Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
    
        Call Conecta_Empresa
        
        Call CmdLimpiar_Click
        Material.SetFocus
        
    End If
        
End Sub

Private Sub CmdLimpiar_Click()

    Material.Text = ""
    Solicita.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Proveedor.Text = ""
    DesProveedor.Text = ""
    EspecificacionesProve.Text = ""
    Certificado.Text = ""
    Precio.Text = ""
    Origen.Text = ""
    Ct.Text = ""
    Nombre.Text = ""
    Comentarios.Text = ""
    Entregado.Text = ""
    TipoMp.ListIndex = 0
    Senasa.ListIndex = 0
    Trazabilidad.ListIndex = 0
    VtoSenasa.Text = "  /  /    "
    
    Material.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgAltaHomologa.Hide
    Unload Me
    PrgHomologaProve.Show
End Sub

Sub Material_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Solicita.SetFocus
    End If
    If KeyAscii = 27 Then
        Material.Text = ""
    End If
End Sub

Private Sub Solicita_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Solicita.Text = ""
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Proveedor.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM Proveedor"
        Sql3 = " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = Sql1 + Sql2 + Sql3
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            DesProveedor.Text = RstProveedor!Nombre
            RstProveedor.Close
            EspecificacionesProve.SetFocus
                Else
            Proveedor.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Proveedor.Text = ""
    End If
End Sub

Private Sub DesProveedor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EspecificacionesProve.SetFocus
    End If
    If KeyAscii = 27 Then
        DesProveedor.Text = ""
    End If
End Sub

Private Sub EspecificacionesProve_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Certificado.SetFocus
    End If
    If KeyAscii = 27 Then
        EspecificacionesProve.Text = ""
    End If
End Sub

Private Sub Certificado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Precio.SetFocus
    End If
    If KeyAscii = 27 Then
        Certificado.Text = ""
    End If
End Sub

Private Sub Precio_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Origen.SetFocus
    End If
    If KeyAscii = 27 Then
        Precio.Text = ""
    End If
End Sub

Private Sub Origen_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ct.SetFocus
    End If
    If KeyAscii = 27 Then
        Origen.Text = ""
    End If
End Sub

Private Sub Ct_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Nombre.SetFocus
    End If
    If KeyAscii = 27 Then
        Ct.Text = ""
    End If
End Sub

Private Sub Nombre_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentarios.SetFocus
    End If
    If KeyAscii = 27 Then
        Nombre.Text = ""
    End If
End Sub

Private Sub Comentarios_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Entregado.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentarios.Text = ""
    End If
End Sub

Private Sub Entregado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Material.SetFocus
    End If
    If KeyAscii = 27 Then
        Entregado.Text = ""
    End If
End Sub

Private Sub Consulta_Click()
    Opcion.Clear
    Opcion.AddItem "Proveedores"
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
            Ayuda.Visible = True
            Ayuda.Text = ""
            
            spProveedor = "ListaProveedoresordConsulta"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
            
                With RstProveedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Auxi$ = Mascara("###########", Str$(RstProveedor!Proveedor))
                            Call Ceros(Auxi, 11)
                            IngresaItem = Auxi + "      " + RstProveedor!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = RstProveedor!Proveedor
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                RstProveedor.Close
            
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Proveedor.Text = WIndice.List(Indice)
            Call Proveedor_KeyPress(13)
        
        Case Else
    End Select
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
    Opcion.Visible = False
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    WEspacios = Len(Ayuda.Text)
    
    Select Case XIndice
        Case 0
            If Ayuda.Text <> "" Then
                spProveedor = "ListaProveedoresOrdConsultaII " + "'" + Ayuda.Text + "'"
                    Else
                spProveedor = "ListaProveedoresOrdConsulta"
            End If
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
    
            With RstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
            
                        Da = Len(!Nombre) - WEspacios
                
                        For aa = 1 To Da
                            If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                                Auxi = Str$(!Proveedor)
                                Call Ceros(Auxi, 11)
                                IngresaItem = Auxi + "    " + !Nombre
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
    
        Case Else
    End Select
    
    End If

End Sub

Private Sub Form_Load()

    Material.Text = ""
    Solicita.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Proveedor.Text = ""
    DesProveedor.Text = ""
    EspecificacionesProve.Text = ""
    Certificado.Text = ""
    Precio.Text = ""
    Origen.Text = ""
    Ct.Text = ""
    Nombre.Text = ""
    Comentarios.Text = ""
    Entregado.Text = ""
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
            Material.Text = Trim(rstHomologa!Material)
            Solicita.Text = Trim(rstHomologa!Solicita)
            Fecha.Text = rstHomologa!Fecha
            Proveedor.Text = rstHomologa!Proveedor
            DesProveedor.Text = Trim(rstHomologa!DesProveedor)
            EspecificacionesProve.Text = Trim(rstHomologa!EspecificacionesProve)
            Certificado.Text = Trim(rstHomologa!Certificado)
            Precio.Text = rstHomologa!Precio
            Origen.Text = Trim(rstHomologa!Origen)
            Ct.Text = Trim(rstHomologa!Ct)
            Nombre.Text = Trim(rstHomologa!Nombre)
            Comentarios.Text = Trim(rstHomologa!Comentarios)
            Entregado.Text = Trim(rstHomologa!Entregado)
            
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
        
    End If
        
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Material_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Solicita_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub DesProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub EspecificacionesProve_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Certificado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Precio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Origen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ct_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Nombre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Comentarios_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Entregado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub TipoMP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Trazabilidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Senasa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub VtoSenasa_KeyDown(KeyCode As Integer, Shift As Integer)
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







