VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAtrasoExpedicion 
   Caption         =   "Ingreso de Aviso de No entrega"
   ClientHeight    =   4410
   ClientLeft      =   1800
   ClientTop       =   840
   ClientWidth     =   8580
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4410
   ScaleWidth      =   8580
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
      Left            =   2280
      TabIndex        =   16
      Top             =   2280
      Width           =   3495
   End
   Begin VB.ComboBox Concepto 
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
      TabIndex        =   1
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox Pedido 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      MaxLength       =   65
      TabIndex        =   2
      Text            =   " "
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Cliente 
      Enabled         =   0   'False
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
      MaxLength       =   6
      TabIndex        =   9
      Text            =   " "
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Problema 
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
      Top             =   1560
      Width           =   6015
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
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
      Left            =   7320
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdGraba 
      Caption         =   "    Graba           (F1)"
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
      Left            =   3960
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin MSMask.MaskEdBox FecEntrega 
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
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
   Begin VB.Label Label5 
      Caption         =   "Planta"
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
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha del Entrega"
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
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de Retraso"
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
      Top             =   1920
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
      Left            =   3720
      TabIndex        =   12
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Numero de Pedido"
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
      Top             =   480
      Width           =   1935
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
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "Fecha del Aviso"
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
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Problema"
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
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   5
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "PrgAtrasoExpedicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstAtraso As Recordset
Dim spAtraso As String
Dim rstVendedor As Recordset
Dim spVendedor As String
Dim rstSolic As Recordset
Dim spSolic As String
Dim XParam As String
Dim EmpresaActual As String
Dim XIndice As Integer
Dim EmailAddress As String
Dim WEmail(100) As String
Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
Dim MSubject As String
Dim MBody As String
Dim AllPath As String

Private Sub cmdGraba_Click()

    Rem On Error GoTo WError
    
    If Concepto.ListIndex <= 0 Then
        m$ = "Se debe informar el concpeto del atraso"
        a% = MsgBox(m$, 0, "Aviso de Atraso")
        Exit Sub
    End If
    
    If Trim(Problema.Text) = "" Then
        m$ = "Se debe informar el problema del atraso"
        a% = MsgBox(m$, 0, "Aviso de Atraso")
        Exit Sub
    End If
    
    If Concepto.ListIndex <= 0 Then
        m$ = "Se debe informar la Planta"
        a% = MsgBox(m$, 0, "Aviso de Atraso")
        Exit Sub
    End If
    
    WAtraso = "1"
    
    Sql1 = "Select Max(Numero) as [NumeroMayor]"
    Sql2 = " FROM Atraso"
    spAtraso = Sql1 + Sql2
    Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
    If rstAtraso.RecordCount > 0 Then
        WAtraso = Str$(rstAtraso!Numeromayor + 1)
        rstAtraso.Close
    End If

    WFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WFechaEntregaord = Right$(FecEntrega.Text, 4) + Mid$(FecEntrega.Text, 4, 2) + Left$(FecEntrega.Text, 2)
    
    If Planta.ListIndex = 1 Then
        ZZOrigen = "3"
            Else
        ZZOrigen = "4"
    End If
    
    ZSql = ""
    ZSql = ZSql + "INSERT INTO Atraso ("
    ZSql = ZSql + "Numero ,"
    ZSql = ZSql + "Fecha ,"
    ZSql = ZSql + "OrdFecha ,"
    ZSql = ZSql + "Pedido ,"
    ZSql = ZSql + "Cliente ,"
    ZSql = ZSql + "Terminado ,"
    ZSql = ZSql + "Problema ,"
    ZSql = ZSql + "Articulo ,"
    ZSql = ZSql + "FechaEntrega ,"
    ZSql = ZSql + "OrdFechaEntrega ,"
    ZSql = ZSql + "DesCliente ,"
    ZSql = ZSql + "DesTerminado ,"
    ZSql = ZSql + "DesArticulo ,"
    ZSql = ZSql + "Concepto ,"
    ZSql = ZSql + "Solicitud ,"
    ZSql = ZSql + "Origen ,"
    ZSql = ZSql + "VersionPedido)"
    ZSql = ZSql + "Values ("
    ZSql = ZSql + "'" + WAtraso + "',"
    ZSql = ZSql + "'" + Fecha.Text + "',"
    ZSql = ZSql + "'" + WOrdFecha + "',"
    ZSql = ZSql + "'" + Pedido.Text + "',"
    ZSql = ZSql + "'" + Cliente.Text + "',"
    ZSql = ZSql + "'" + "  -     -   " + "',"
    ZSql = ZSql + "'" + Problema.Text + "',"
    ZSql = ZSql + "'" + "  -   -   " + "',"
    ZSql = ZSql + "'" + FecEntrega.Text + "',"
    ZSql = ZSql + "'" + WOrdFechaEntrega + "',"
    ZSql = ZSql + "'" + DesCliente.Caption + "',"
    ZSql = ZSql + "'" + "" + "',"
    ZSql = ZSql + "'" + "" + "',"
    ZSql = ZSql + "'" + Str$(Concepto.ListIndex + 4) + "',"
    ZSql = ZSql + "'" + "" + "',"
    ZSql = ZSql + "'" + ZZOrigen + "',"
    ZSql = ZSql + "'" + "" + "')"
    
    spAtraso = ZSql
    Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
    ZLLave = "N"
    
    Call cmdClose_Click
        
    Exit Sub

WError:
    Resume Next
        
End Sub


Private Sub cmdClose_Click()
    PrgAtrasoExpedicion.Hide
    Unload Me
    PrgHojaRuta.Show
End Sub

Private Sub Problema_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Solicitud.SetFocus
    End If
    If KeyAscii = 27 Then
        Problema.Text = ""
    End If
End Sub

Private Sub Concepto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
End Sub

Private Sub Form_Load()

    Concepto.Clear
    
    Concepto.AddItem ""
    Concepto.AddItem "Error del Sistema"
    Concepto.AddItem "Varios"
    Concepto.AddItem "Problemas Vehiculos"
    Concepto.AddItem "Problemas Logistica"
    Concepto.AddItem "Problemas Recepcion Cliente"
    Concepto.AddItem "Varios"
    Concepto.AddItem "Corte de Luz"
    Concepto.AddItem "Pedido por el Cliente"
    Concepto.AddItem "Falta de Pago"
    Concepto.AddItem "Confirmacion Pedido Parcial"
    Concepto.AddItem "Envases"
    
    
    
    Concepto.ListIndex = 0
    
    Planta.Clear
    
    Planta.AddItem ""
    Planta.AddItem "Planta I"
    Planta.AddItem "Planta V"

    Planta.ListIndex = 2

    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Pedido.Text = ZAtraso(ZLugarAtraso, 1)
    Cliente.Text = ZAtraso(ZLugarAtraso, 2)
    Problema.Text = ""
    DesCliente.Caption = ZAtraso(ZLugarAtraso, 3)
    FecEntrega.Text = ZAtraso(ZLugarAtraso, 4)
        
End Sub

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Pedido_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Terminado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Problema_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub FechaEntrega_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call cmdGraba_Click
        Case Else
    End Select
End Sub






