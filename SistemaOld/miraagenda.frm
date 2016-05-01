VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMiraAgenda 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Agenda de Clientes"
   ClientHeight    =   8325
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11790
   LinkTopic       =   "Form2"
   ScaleHeight     =   8325
   ScaleWidth      =   11790
   Begin VB.Frame CargaFecha 
      Height          =   1215
      Left            =   2880
      TabIndex        =   19
      Top             =   2640
      Visible         =   0   'False
      Width           =   3975
      Begin MSMask.MaskEdBox FechaAlta 
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Top             =   360
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
      Begin MSMask.MaskEdBox FechaMinuta 
         Height          =   285
         Left            =   2040
         TabIndex        =   22
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
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
         Caption         =   "Fecha de Minuta"
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
         TabIndex        =   23
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de Agenda"
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
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame PantaAltaCliente 
      Height          =   1935
      Left            =   1200
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton CierraPantalla 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   17
         Top             =   7560
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid VectorCtaCte 
         Height          =   6855
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12091
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin VB.TextBox Desde 
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
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   13
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Hasta 
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
         Left            =   4920
         MaxLength       =   6
         TabIndex        =   12
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3360
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton MiraCtaCte 
      Caption         =   "Cta.Cte."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9600
      TabIndex        =   18
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton ZBajaII 
      Caption         =   "Baja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3000
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame PantaBaja 
      BackColor       =   &H00FFFFC0&
      Height          =   2295
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton ZCancela 
         Caption         =   "Cancela"
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
         Left            =   5760
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
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
         Left            =   3600
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton ZModifica 
         Caption         =   "Modifica"
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
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport ListaGRilla 
      Left            =   11160
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "muestra.rpt"
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   7335
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12938
      _Version        =   327680
      BackColor       =   16777215
      ForeColor       =   4210752
      FocusRect       =   2
      GridLines       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Alta 
      Caption         =   "Alta "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Impresion 
      Caption         =   "Impresion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Fin "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   7560
      TabIndex        =   4
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
      Left            =   6480
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "PrgMiraAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Dim rstCliente As Recordset
Dim spCliente As String
Dim XParam As String
Dim WFecha As String
Dim WFecha2 As String
Dim ColumnaOpcion As Integer
Dim Seleccion As String
Dim ZFecha As String
Dim Importe1 As Double
Dim Importe2 As Double
Dim Importe3 As Double
Dim ZAltaCliente As String

Private Sub Alta_Click()
    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    Fila = Muestra.Row
    WMuestra = ""
    ZZPasaProcesoFechaAgenda = ""
    PrgAltaAgenda.Show
End Sub

Private Sub CierraPantalla_Click()
    PantaAltaCliente.Visible = False
    Call Fecha_Keypress(13)
End Sub

Private Sub cmdClose_Click()
    PrgMiraAgenda.Hide
    Unload Me
    Close
    End
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Fecha.Text <> "  /  /    " Then
            Call Proceso_Click
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub


Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Call LeeCtaCte
    End If
End Sub

Private Sub Impresion_Click()

    ListaGRilla.WindowTitle = "Agenda de Clientes"
    ListaGRilla.WindowTop = 0
    ListaGRilla.WindowLeft = 0
    ListaGRilla.WindowWidth = Screen.Width
    ListaGRilla.WindowHeight = Screen.Height

    ListaGRilla.ReportFileName = "ListaAgenda.rpt"

    Uno = "{Cliente.Fecha} in " + Chr$(34) + Fecha.Text + Chr$(34) + " to " + Chr$(34) + Fecha.Text + Chr$(34)
    
    ListaGRilla.GroupSelectionFormula = Uno
    ListaGRilla.SelectionFormula = Uno
    
    ListaGRilla.Destination = 1
    ListaGRilla.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    ListaGRilla.SQLQuery = "SELECT Cliente.Cliente, Cliente.Razon, Cliente.Telefono, Cliente.Fecha, Cliente.Hora, Cliente.Anotacion " _
            + "From " _
            + DSQ + ".dbo.Cliente Cliente " _
            + "Where " _
            + "Cliente.Fecha >= '" + Fecha.Text + "' AND " _
            + "Cliente.Fecha <= '" + Fecha.Text + "'"
            
    ListaGRilla.Connect = Connect()
    
    ListaGRilla.Action = 1
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Fecha = " + "'" + Fecha.Text + "'"
    ZSql = ZSql + " Order by Cliente.Hora, Cliente.Razon"
    spCliente = ZSql
            
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
    
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    WLugar = WLugar + 1
                    
                    Muestra.TextMatrix(WLugar, 1) = rstCliente!Cliente
                    Muestra.TextMatrix(WLugar, 2) = rstCliente!Razon
                    Muestra.TextMatrix(WLugar, 3) = rstCliente!Anotacion
                    Muestra.TextMatrix(WLugar, 4) = rstCliente!Hora
                    
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
        
        End With
        rstCliente.Close
    End If
    
    Muestra.Visible = True
    Muestra.SetFocus
    
End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    If Fecha.Text = "  /  /    " Then
        ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        Fecha.Text = ZFecha
    End If
  Rem********* BY NAN ME CONECTO A LA EMPRESA DE LABURO 0012*****
 Rem   XEmpresa = WEmpresa
 Rem   ZZEmpre = "12"
    
 Rem   txtOdbc = "Empresa" + ZZEmpre
  Rem  WEmpresa = "00" + ZZEmpre
  Rem  strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
  Rem  Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
  Rem BY NAN
    
    
    Call Proceso_Click
    If ZZPasaProcesoAltaAgenda = 1 Then
        ZZPasaProcesoAltaAgenda = 0
        WMuestra = ""
        PrgAltaAgenda.Show
    End If

    
     Rem BY NAN ME CONECTO A LA EMPRESA QE ESTABA
     Rem  WEmpresa = XEmpresa
     Rem   txtOdbc = "Empresa" + Right$(XEmpresa, 2)
   
     Rem   strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
     Rem   Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

Rem EN B

End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Alta_Click
        Case 120
            Call Impresion_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector()

    Muestra.Clear

    Rem ponga la muestra en negritas
    Rem Muestra.Font.Bold = True

    ' Establesco loa Valores de la muestra
    
    Muestra.FixedCols = 1
    Muestra.Cols = 5
    Muestra.FixedRows = 1
    Muestra.Rows = 1000
    
    Muestra.ColWidth(0) = 200
    Muestra.Row = 0
    
    For Ciclo = 1 To Muestra.Cols - 1
        Muestra.Col = Ciclo
        Select Case Ciclo
            Case 1
                Muestra.Text = "Codigo"
                Muestra.ColWidth(Ciclo) = 1000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                Muestra.Text = "Razon"
                Muestra.ColWidth(Ciclo) = 3500
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                Muestra.Text = "Observaciones"
                Muestra.ColWidth(Ciclo) = 5700
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                Muestra.Text = "Hora"
                Muestra.ColWidth(Ciclo) = 1000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Muestra.AllowUserResizing = flexResizeBoth
    
    Muestra.Col = 1
    Muestra.Row = 1
    
End Sub

Private Sub Limpia_VectorII()

    VectorCtaCte.Clear

    Rem ponga la vectorctacte en negritas
    Rem vectorctacte.Font.Bold = True

    ' Establesco loa Valores de la vectorctacte
    
    VectorCtaCte.FixedCols = 1
    VectorCtaCte.Cols = 11
    VectorCtaCte.FixedRows = 1
    VectorCtaCte.Rows = 10000
    
    VectorCtaCte.ColWidth(0) = 200
    VectorCtaCte.Row = 0
    
    For Ciclo = 1 To VectorCtaCte.Cols - 1
        VectorCtaCte.Col = Ciclo
        Select Case Ciclo
            Case 1
                VectorCtaCte.Text = "Cliente"
                VectorCtaCte.ColWidth(Ciclo) = 800
                VectorCtaCte.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                VectorCtaCte.Text = "Razon"
                VectorCtaCte.ColWidth(Ciclo) = 2000
                VectorCtaCte.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                VectorCtaCte.Text = "Tipo"
                VectorCtaCte.ColWidth(Ciclo) = 600
                VectorCtaCte.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                VectorCtaCte.Text = "Numero"
                VectorCtaCte.ColWidth(Ciclo) = 800
                VectorCtaCte.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                VectorCtaCte.Text = "Fecha"
                VectorCtaCte.ColWidth(Ciclo) = 1150
                VectorCtaCte.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 6
                VectorCtaCte.Text = "Debito"
                VectorCtaCte.ColWidth(Ciclo) = 1100
                VectorCtaCte.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                VectorCtaCte.Text = "Credito"
                VectorCtaCte.ColWidth(Ciclo) = 1100
                VectorCtaCte.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 8
                VectorCtaCte.Text = "Saldo"
                VectorCtaCte.ColWidth(Ciclo) = 1100
                VectorCtaCte.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 9
                VectorCtaCte.Text = "Vto"
                VectorCtaCte.ColWidth(Ciclo) = 1150
                VectorCtaCte.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 10
                VectorCtaCte.Text = "Acumulado"
                VectorCtaCte.ColWidth(Ciclo) = 1100
                VectorCtaCte.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    VectorCtaCte.AllowUserResizing = flexResizeBoth
    
    VectorCtaCte.Col = 1
    VectorCtaCte.Row = 1
    
End Sub

Private Sub FechaAlta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaAlta.Text, Auxi)
        If Auxi = "S" Then
        
            ZOrdFecha = Right$(FechaAlta.Text, 4) + Mid$(FechaAlta.Text, 4, 2) + Left$(FechaAlta.Text, 2)
            
            ZCompara = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZOrdCompara = Right$(ZCompara, 4) + Mid$(ZCompara, 4, 2) + Left$(ZCompara, 2)
            If ZOrdCompara > ZOrdFecha Then
                m$ = "La fecha a ingresar debe ser mayor a " + ZCompara
                a% = MsgBox(m$, 0, "Alta de Registro en la Agenda")
                Exit Sub
            End If
        
            ZOrdFecha = Right$(FechaAlta.Text, 4) + Mid$(FechaAlta.Text, 4, 2) + Left$(FechaAlta.Text, 2)
            ZSql = ""
            ZSql = ZSql + "UPDATE Cliente SET "
            ZSql = ZSql + "Fecha =  " + "'" + FechaAlta.Text + "',"
            ZSql = ZSql + "OrdFecha =  " + "'" + ZOrdFecha + "',"
            ZSql = ZSql + "Anotacion =  " + "'" + "" + "',"
            ZSql = ZSql + "Hora =  " + "'" + "" + "'"
            ZSql = ZSql + " Where Cliente = " + "'" + ZAltaCliente + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            CargaFecha.Visible = False
        End If
    End If
    If KeyAscii = 27 Then
        FechaAlta.Text = "  /  /    "
    End If
End Sub

Private Sub MiraCtaCte_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Call Limpia_VectorII
    PantaAltaCliente.Height = 8055
    PantaAltaCliente.Left = 120
    PantaAltaCliente.Top = 120
    PantaAltaCliente.Width = 11655
    PantaAltaCliente.Visible = True
    Desde.SetFocus
End Sub

Private Sub Muestra_DblClick()

    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    Fila = Muestra.Row
    WMuestra = Muestra.TextMatrix(Muestra.Row, 1)
    Rem If Trim(WMuestra) <> "" Then
    Rem     PrgAltaAgenda.Show
    Rem End If
    If Trim(WMuestra) <> "" Then
        Rem PantaBaja.Visible = True
        PrgAltaAgenda.Show
    End If

End Sub

Private Sub MuestraAnterior_dblClick()

    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    Fila = Muestra.Row
    WMuestra = Muestra.TextMatrix(Muestra.Row, 1)
    If Trim(WMuestra) <> "" Then
        PrgAltaAgenda.Show
    End If

End Sub

Private Sub VectorCtaCte_dblClick()

    ZAltaCliente = VectorCtaCte.TextMatrix(VectorCtaCte.Row, 1)
    FechaAlta.Text = "  /  /    "
    FechaMinuta.Text = "  /  /    "
    
    Sql1 = "Select *"
    Sql2 = " FROM Cliente"
    Sql3 = " Where Cliente.cliente = " + "'" + ZAltaCliente + "'"
    spCliente = Sql1 + Sql2 + Sql3
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
    
        FechaAlta.Text = IIf(IsNull(rstCliente!Fecha), "  /  /    ", rstCliente!Fecha)
        FechaMinuta.Text = IIf(IsNull(rstCliente!FechaMinuta), "00/00/0000", rstCliente!FechaMinuta)
        ZFechaDia = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        
        ZComparaI = Right$(FechaMinuta.Text, 4) + Mid$(FechaMinuta.Text, 4, 2) + Left$(FechaMinuta.Text, 2)
        ZComparaII = Right$(ZFechaDia, 4) + Mid$(ZFechaDia, 4, 2) + Left$(ZFechaDia, 2)
        
        If ZComparaI < ZComparaII Then
            FechaMinuta.Text = "  /  /    "
        End If
        
        rstCliente.Close
    End If
    
    CargaFecha.Visible = True
    FechaAlta.SetFocus
    
End Sub

Private Sub ZBajaII_Click()
    rowini = Muestra.Row
    RowFin = Muestra.RowSel
    For Ciclo = rowini To RowFin
        ZZCliente = Muestra.TextMatrix(Ciclo, 1)
        ZSql = ""
        ZSql = ZSql + "UPDATE Cliente SET "
        ZSql = ZSql + "Fecha =  " + "'" + "  /  /    " + "',"
        ZSql = ZSql + "OrdFecha =  " + "'" + "" + "',"
        ZSql = ZSql + "Anotacion =  " + "'" + "" + "',"
        ZSql = ZSql + "Hora =  " + "'" + "" + "'"
        ZSql = ZSql + " Where Cliente = " + "'" + ZZCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    Next Ciclo
    Call Fecha_Keypress(13)
End Sub

Private Sub ZCancela_Click()
    PantaBaja.Visible = False
End Sub

Private Sub ZMinuta_Click()
    PantaBaja.Visible = False
    PrgAltaMinuta.Show
End Sub

Private Sub ZModifica_Click()
    PantaBaja.Visible = False
    If Trim(WMuestra) <> "" Then
        PrgAltaAgenda.Show
    End If
End Sub

Private Sub LeeCtaCte()

    Call Limpia_VectorII
    Renglon = 0
    Pasa = 0
    
    ZSql = ""
    ZSql = ZSql + "Select CtaCte.Cliente, CtaCte.Saldo, CtaCte.Tipo, CtaCte.OrdFecha, CtaCte.Numero, CtaCte.Total, CtaCte.Impre, CtaCte.Fecha, CtaCte.vencimiento, Cliente.Razon as [WRazon]"
    ZSql = ZSql + " FROM CtaCte, Cliente"
    ZSql = ZSql + " Where Ctacte.Cliente = Cliente.Cliente"
    ZSql = ZSql + " and CtaCte.Cliente >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and CtaCte.Cliente <= " + "'" + Hasta.Text + "'"
    ZSql = ZSql + " and CtaCte.Saldo <> 0 "
    ZSql = ZSql + " and CtaCte.Tipo < " + "'" + "50" + "'"
    ZSql = ZSql + " Order by Ctacte.cliente, Ctacte.OrdFecha, CtaCte.numero"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
    
        With rstCtacte
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    If !Total > 0 Then
                        Importe1 = !Total
                        Importe2 = 0
                            Else
                        Importe1 = 0
                        Importe2 = !Total
                    End If
                    Importe3 = !Saldo
                        
                    Call Redondeo(Importe3)
                
                    If Importe3 <> 0 Then
                    
                        If Pasa = 0 Then
                            Pasa = 1
                            ZCorte = !Cliente
                            WSaldo = 0
                        End If
                        
                        If ZCorte <> !Cliente Then
                            Renglon = Renglon + 1
                            ZCorte = !Cliente
                            WSaldo = 0
                        End If
                
                        Renglon = Renglon + 1
                        
                        VectorCtaCte.TextMatrix(Renglon, 1) = !Cliente
                        VectorCtaCte.TextMatrix(Renglon, 2) = !WRazon
                
                        Select Case !Tipo
                            Case 1
                                VectorCtaCte.TextMatrix(Renglon, 3) = "Fac"
                            Case 2
                                VectorCtaCte.TextMatrix(Renglon, 3) = "Dev"
                            Case 3
                                VectorCtaCte.TextMatrix(Renglon, 3) = "Fac"
                            Case 4
                                Select Case Left$(!Impre, 2)
                                    Case "DC"
                                        VectorCtaCte.TextMatrix(Renglon, 3) = "D.C"
                                    Case "CH"
                                        VectorCtaCte.TextMatrix(Renglon, 3) = "CHR"
                                    Case Else
                                        VectorCtaCte.TextMatrix(Renglon, 3) = "N/D"
                                End Select
                            Case 5
                                Select Case Left$(!Impre, 2)
                                    Case "DC"
                                        VectorCtaCte.TextMatrix(Renglon, 3) = "D.C"
                                    Case "CH"
                                        VectorCtaCte.TextMatrix(Renglon, 3) = "CHR"
                                    Case Else
                                        VectorCtaCte.TextMatrix(Renglon, 3) = "N/C"
                                End Select
                            Case 6
                                VectorCtaCte.TextMatrix(Renglon, 3) = "Rec"
                            Case 7
                                VectorCtaCte.TextMatrix(Renglon, 3) = "Ant"
                            Case 10
                                VectorCtaCte.TextMatrix(Renglon, 3) = "FCR"
                            Case 50
                                VectorCtaCte.TextMatrix(Renglon, 3) = "Doc"
                            Case Else
                        End Select
                        
                        VectorCtaCte.TextMatrix(Renglon, 4) = Pusing("######", Str$(!Numero))
                        VectorCtaCte.TextMatrix(Renglon, 5) = !Fecha
                
                        If Importe1 <> 0 Then
                            VectorCtaCte.TextMatrix(Renglon, 6) = Pusing("###,###,###.##", Str$(Importe1))
                                Else
                            VectorCtaCte.TextMatrix(Renglon, 6) = ""
                        End If
                
                        If Importe2 <> 0 Then
                            VectorCtaCte.TextMatrix(Renglon, 7) = Pusing("###,###,###.##", Str$(Importe2))
                                Else
                            VectorCtaCte.TextMatrix(Renglon, 7) = ""
                        End If
                
                        If Importe3 <> 0 Then
                            VectorCtaCte.TextMatrix(Renglon, 8) = Pusing("###,###,###.##", Str$(Importe3))
                                Else
                            VectorCtaCte.TextMatrix(Renglon, 8) = ""
                        End If
                        
                        WSaldo = WSaldo + Importe3
                
                        VectorCtaCte.TextMatrix(Renglon, 9) = !Vencimiento
                        VectorCtaCte.TextMatrix(Renglon, 10) = Pusing("###,###,###.##", Str$(WSaldo))
                    
                    End If
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            End If
        
        End With
        
        rstCtacte.Close
        
    End If
    
    VectorCtaCte.Col = 1
    VectorCtaCte.Row = 1
    VectorCtaCte.SetFocus

End Sub

