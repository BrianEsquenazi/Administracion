VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCuenta 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Cuentas Contables"
   ClientHeight    =   6330
   ClientLeft      =   3555
   ClientTop       =   2100
   ClientWidth     =   6150
   LinkTopic       =   "Form2"
   ScaleHeight     =   6330
   ScaleWidth      =   6150
   Begin VB.TextBox Cuenta 
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1455
      Left            =   240
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   25
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   24
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   23
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4440
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wcuentas.rpt"
      Destination     =   1
      WindowTitle     =   "Listados de Cuentas Contables"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "cuenta.frx":0000
      Left            =   480
      List            =   "cuenta.frx":0007
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton lista 
      Caption         =   "Listado"
      Height          =   300
      Left            =   1800
      TabIndex        =   14
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   600
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles"
      Height          =   1335
      Left            =   4320
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   3000
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   3000
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Descripcion 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion de la Cuenta"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Cuenta Contable"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   2055
   End
End
Attribute VB_Name = "PrgCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCuenta As Recordset
Dim spCuenta As String
Dim XParam As String

Sub Verifica_datos()
    Rem If Val(Nivel.text) = 0 Then
    Rem     Nivel.text = "0"
    Rem End If
End Sub
Sub Format_datos()
    Rem Comision.text = PUsing("###,###.##", Comision.text)
End Sub

Sub Imprime_Datos()
    spCuenta = "ConsultaCuentas " + "'" + Cuenta.Text + "'"
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        Cuenta.Text = rstCuenta!Cuenta
        Descripcion.Text = rstCuenta!Descripcion
        rstCuenta.Close
        Call Format_datos
    End If
End Sub

Private Sub Acepta_Click()
    
    Rem listados.Report1.ReportFileName = PATH_PROG + "LEYENDAS.RPT"
    Rem listados.Report1.SelectionFormula = "{" + TABLA_LEYENDAS + "." + CODIGO_LEYENDA + "} > ''"
    Rem listados.Report1.SortFields(0) = "+{" + TABLA_LEYENDAS + "." + CODIGO_LEYENDA + "}"
    
    listado.WindowTitle = "Listado de Cuentas Contables"
    listado.WindowTop = 0
    listado.WindowLeft = 0
    listado.WindowWidth = Screen.Width
    listado.WindowHeight = Screen.Height

    listado.GroupSelectionFormula = "{Cuenta.Cuenta} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        listado.Destination = 1
            Else
        listado.Destination = 0
    End If
    Cuenta.SetFocus
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    listado.SQLQuery = "SELECT Cuenta.Cuenta , Cuenta.Descripcion " _
                          + "From " + DSQ + ".dbo.Cuenta Cuenta " _
                          + "Where Cuenta.Cuenta >= ' ' AND Cuenta.Cuenta <= 'ZZZZZZZZZZ'"
    listado.DataFiles(1) = WEmpresa + "Auxi.mdb"
    listado.Connect = Connect()
    
    listado.Action = 1
    Frame2.Visible = False
End Sub


Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Cuenta.Text <> "" Then
    
        Call Verifica_datos
        
        XParam = "'" + Cuenta.Text + "','" _
                        + Descripcion.Text + "','" _
                        + "0" + "','" _
                        + "1" + "'"
        
        spCuenta = "ConsultaCuentas " + "'" + Cuenta.Text + "'"
        Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
        If rstCuenta.RecordCount > 0 Then
            Set rstCuenta = db.OpenRecordset("ModificaCuenta " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                Else
            Set rstCuenta = db.OpenRecordset("AltaCuenta " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        End If
        Call CmdLimpiar_Click
        Cuenta.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Cuenta.Text <> "" Then
        spCuenta = "ConsultaCuentas " + "'" + Cuenta.Text + "'"
        Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
        If rstCuenta.RecordCount > 0 Then
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spCuenta = "BorrarCuenta " + "'" + Cuenta.Text + "'"
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    End If
    Cuenta.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Cuenta.Text = ""
    Descripcion.Text = ""
    Cuenta.SetFocus
End Sub

Private Sub cmdClose_Click()

    Call CmdLimpiar_Click
    
    With rstEmpresa
        .Close
    End With
    
    Cuenta.SetFocus
    PrgCuenta.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Lista_Click()
    Desde.Text = "0"
    Hasta.Text = "9999999999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cuenta.SetFocus
    End If
End Sub

Private Sub Cuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cuenta.Text <> "" Then
            WCuenta = Cuenta.Text
            spCuenta = "ConsultaCuentas " + "'" + Cuenta.Text + "'"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                Cuenta.Text = rstCuenta!Cuenta
                Descripcion.Text = rstCuenta!Descripcion
                rstCuenta.Close
                Call Imprime_Datos
                    Else
                WCuenta = Cuenta.Text
                CmdLimpiar_Click
                Cuenta.Text = WCuenta
            End If
        End If
        Descripcion.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    Rem  Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

Rem     Opcion.Clear
Rem
Rem     Opcion.AddItem "Productos"
Rem     Opcion.AddItem "Ensayos"
Rem
Rem     Opcion.Visible = True
Rem End Sub
Rem
Rem Private Sub Opcion_Click()
Rem
Rem     Opcion.Visible = False

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    Rem XIndice = Opcion.ListIndex
    XIndice = 0
    
    Select Case XIndice
        Case 0
            spCuenta = "ListaCuentas"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstCuenta
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstCuenta!Cuenta + " " + rstCuenta!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstCuenta!Cuenta
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCuenta.Close
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WCuenta = WIndice.List(Indice)
            spCuenta = "ConsultaCuentas " + "'" + WCuenta + "'"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                Cuenta.Text = rstCuenta!Cuenta
                rstCuenta.Close
                Call Imprime_Datos
                        Else
                CmdLimpiar_Click
                Cuenta.Text = WCuenta
            End If
            Cuenta.SetFocus
            
        Case Else
    End Select
    
End Sub

Sub Form_Load()
    Cuenta.Text = ""
    Descripcion.Text = ""
End Sub

Private Sub Primer_Click()

    On Error GoTo WError
    
    spCuenta = "ListaCuentas"
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        With rstCuenta
            .MoveFirst
            Cuenta.Text = rstCuenta!Cuenta
            rstCuenta.Close
            Call Imprime_Datos
        End With
    End If
    Cuenta.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Cuenta", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Cuenta.SetFocus
 End Sub

Private Sub Ultimo_Click()

   On Error GoTo Error_ultimo
    
    spCuenta = "ListaCuentas"
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        With rstCuenta
            .MoveLast
            Cuenta.Text = rstCuenta!Cuenta
            rstCuenta.Close
            Call Imprime_Datos
        End With
    End If
    Cuenta.SetFocus
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Cuenta", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Cuenta.SetFocus
 End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    spCuenta = "AnteriorCuenta " + "'" + Cuenta.Text + "'"
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        With rstCuenta
            .MoveLast
            Cuenta.Text = rstCuenta!Cuenta
            rstCuenta.Close
            Call Imprime_Datos
        End With
    End If
    
    Cuenta.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Cuenta", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Cuenta.SetFocus
    
End Sub


Private Sub Siguiente_Click()

    On Error GoTo WError
    
    spCuenta = "PosteriorCuenta " + "'" + Cuenta.Text + "'"
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        With rstCuenta
            .MoveFirst
            Cuenta.Text = rstCuenta!Cuenta
            rstCuenta.Close
            Call Imprime_Datos
        End With
    End If
    
    Cuenta.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Cuenta", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Cuenta.SetFocus
    
End Sub




