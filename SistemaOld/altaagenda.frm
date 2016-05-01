VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAltaAgenda 
   Caption         =   "Alta de Registro en la Agenda de Clientes"
   ClientHeight    =   5520
   ClientLeft      =   1800
   ClientTop       =   570
   ClientWidth     =   8565
   LinkTopic       =   "Form2"
   ScaleHeight     =   5520
   ScaleWidth      =   8565
   Begin VB.CommandButton ZMinuta 
      Caption         =   "Minuta"
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
      Left            =   6120
      TabIndex        =   20
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton CtaCte 
      Caption         =   "Cuenta Corriente"
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
      Left            =   4920
      TabIndex        =   19
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton DatosCliente 
      Caption         =   "Datos Cliente"
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
      Left            =   3720
      TabIndex        =   18
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Baja 
      Caption         =   "Elimina"
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
      Left            =   2520
      TabIndex        =   17
      Top             =   1320
      Width           =   1095
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
      TabIndex        =   0
      Text            =   " "
      Top             =   120
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
      TabIndex        =   3
      Text            =   " "
      Top             =   840
      Width           =   5895
   End
   Begin VB.TextBox Hora 
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
      Left            =   4800
      MaxLength       =   15
      TabIndex        =   2
      Text            =   " "
      Top             =   480
      Width           =   1335
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
      Top             =   2160
      Visible         =   0   'False
      Width           =   8175
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
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
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7320
      TabIndex        =   8
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
      ItemData        =   "altaagenda.frx":0000
      Left            =   120
      List            =   "altaagenda.frx":0007
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Ayuda"
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
      Left            =   1320
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Fin de  Ingreso "
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
      Left            =   7320
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdGraba 
      Caption         =   "Graba  "
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
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label DesCliente 
      BackColor       =   &H00FFFF00&
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
      Height          =   285
      Left            =   3720
      TabIndex        =   16
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label8 
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
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Hora"
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
      TabIndex        =   14
      Top             =   480
      Width           =   1095
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
      TabIndex        =   12
      Top             =   120
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
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   9
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "PrgAltaAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCliente As Recordset
Dim spCliente As String
Dim XParam As String
Dim EmpresaActual As String
Dim XIndice As Integer

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


Private Sub Baja_Click()

    T$ = "Borrar Registro"
    m$ = "Desea eliminar el registro"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Cliente SET "
        ZSql = ZSql + "Fecha =  " + "'" + "  /  /    " + "',"
        ZSql = ZSql + "OrdFecha =  " + "'" + "" + "',"
        ZSql = ZSql + "Anotacion =  " + "'" + "" + "',"
        ZSql = ZSql + "Hora =  " + "'" + "" + "'"
        ZSql = ZSql + " Where Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        
        Rem T$ = "Borrar Registro"
        Rem m$ = "Desea generar la Minuta"
        Rem Respuesta% = MsgBox(m$, 32 + 4, T$)
        Rem If Respuesta% = 6 Then
        Rem     PrgAltaAgenda.Hide
        Rem     Unload Me
        Rem     PrgAltaMinuta.Show
        Rem         Else
        Rem     PrgAltaAgenda.Hide
        Rem     Unload Me
        Rem     PrgMiraAgenda.Show
        Rem End If
        
        PrgAltaAgenda.Hide
        Unload Me
        PrgMiraAgenda.Show
    
    End If
    
End Sub

Private Sub cmdGraba_Click()

    Sql1 = "Select *"
    Sql2 = " FROM Cliente"
    Sql3 = " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = Sql1 + Sql2 + Sql3
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        rstCliente.Close
            Else
        m$ = "Codigo de Cliente invalido"
        G% = MsgBox(m$, 0, "Alta de Registro en la Agenda")
        Exit Sub
    End If
    
    ZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    ZCompara = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    ZOrdCompara = Right$(ZCompara, 4) + Mid$(ZCompara, 4, 2) + Left$(ZCompara, 2)
    If ZOrdCompara > ZOrdFecha Then
        m$ = "La fecha a ingresar debe ser mayor a " + ZCompara
        a% = MsgBox(m$, 0, "Alta de Registro en la Agenda")
        Exit Sub
    End If

    ZSql = ""
    ZSql = ZSql + "UPDATE Cliente SET "
    ZSql = ZSql + "Fecha =  " + "'" + Fecha.Text + "',"
    ZSql = ZSql + "OrdFecha =  " + "'" + ZOrdFecha + "',"
    ZSql = ZSql + "Anotacion =  " + "'" + Observaciones.Text + "',"
    ZSql = ZSql + "Hora =  " + "'" + Hora.Text + "'"
    ZSql = ZSql + " Where Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    
    If Trim(WMuestra) <> "" Then
        ZZPasaProcesoFechaAgenda = ""
        ZZPasaProcesoAltaAgenda = 0
        PrgAltaAgenda.Hide
        Unload Me
        PrgMiraAgenda.Show
            Else
        ZZPasaProcesoFechaAgenda = Fecha.Text
        ZZPasaProcesoAltaAgenda = 1
        PrgAltaAgenda.Hide
        Unload Me
        PrgMiraAgenda.Show
    End If
        
End Sub

Private Sub cmdClose_Click()
    ZZPasaProcesoFechaAgenda = ""
    ZZPasaProcesoAltaAgenda = 0
    PrgAltaAgenda.Hide
    Unload Me
    PrgMiraAgenda.Show
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente.Caption = rstCliente!Razon
            ZOrdFecha = IIf(IsNull(rstCliente!OrdFecha), "", rstCliente!OrdFecha)
            ZOrdFecha = Trim(ZOrdFecha)
            If ZOrdFecha <> "" Then
                Fecha.Text = rstCliente!Fecha
                Hora.Text = Str$(rstCliente!Hora)
                Hora.Text = Pusing("###,###.##", Hora.Text)
                Observaciones.Text = Trim(rstCliente!Anotacion)
            End If
            rstCliente.Close
            Fecha.SetFocus
                Else
            Cliente.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
        DesCliente.Caption = ""
    End If
End Sub

Private Sub CtaCte_Click()
    PCliente = Cliente.Text
    PrgCtaCteAgenda.Show
End Sub

Private Sub DatosCliente_Click()
    PCliente = Cliente.Text
    prgcliagenda.Show
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Hora.SetFocus
                Else
            G$ = "Formato de fecha invalido"
            a% = MsgBox(G$, 0, "Alta de Registros en la Agenda de Clientes")
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Sub Hora_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
    If KeyAscii = 27 Then
        Hora.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub Consulta_Click()
    Ayuda.Visible = True
    Ayuda.Text = ""
    Pantalla.Clear
    WIndice.Clear
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Order by Razon"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
    
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = rstCliente!Cliente + "      " + rstCliente!Razon
                    Pantalla.AddItem IngresaItem
                    IngresaItem = rstCliente!Cliente
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCliente.Close
    
    End If
            
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
            Cliente.Text = WIndice.List(Indice)
            Call Cliente_KeyPress(13)
        
        Case Else
    End Select
    
End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Pantalla.Clear
        WIndice.Clear
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Razon LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
        ZSql = ZSql + " Order by Razon"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
        
            With rstCliente
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstCliente!Cliente + "      " + rstCliente!Razon
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstCliente!Cliente
                        WIndice.AddItem IngresaItem
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

Private Sub Form_Load()

    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = "  /  /    "
    Hora.Text = ""
    Observaciones.Text = ""
    Baja.Visible = False
    CtaCte.Visible = False
    DatosCliente.Visible = False
    
    If ZZPasaProcesoFechaAgenda <> "" And Trim(WMuestra) = "" Then
        Fecha.Text = ZZPasaProcesoFechaAgenda
    End If
    
    If Trim(WMuestra) <> "" Then
    
        Cliente.Text = WMuestra
        Baja.Visible = True
        CtaCte.Visible = True
        DatosCliente.Visible = True
    
        Sql1 = "Select *"
        Sql2 = " FROM Cliente"
        Sql3 = " Where Cliente.cliente = " + "'" + WMuestra + "'"
        spCliente = Sql1 + Sql2 + Sql3
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Fecha.Text = rstCliente!Fecha
            Hora.Text = Str$(rstCliente!Hora)
            Hora.Text = Pusing("###,###.##", Hora.Text)
            Observaciones.Text = Trim(rstCliente!Anotacion)
            DesCliente.Caption = Trim(rstCliente!Razon)
            rstCliente.Close
        End If
        
    End If
    
        
End Sub

Private Sub ZMinuta_Click()
    WMuestra = Cliente.Text
    PrgAltaAgenda.Hide
    Unload Me
    PrgAltaMinuta.Show
End Sub
