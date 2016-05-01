VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgVendedor 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Vendedores"
   ClientHeight    =   4560
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   7245
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   7245
   Begin VB.TextBox Vendedor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1455
      Left            =   360
      TabIndex        =   17
      Top             =   2760
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   25
         Text            =   " "
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   24
         Text            =   " "
         Top             =   240
         Width           =   855
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
      ReportFileName  =   "wvendedor.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Rubros"
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
      ItemData        =   "Vende.frx":0000
      Left            =   480
      List            =   "Vende.frx":0007
      TabIndex        =   15
      Top             =   3000
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
      TabIndex        =   1
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
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Nombre 
      Height          =   285
      Left            =   2640
      MaxLength       =   50
      TabIndex        =   4
      Top             =   600
      Width           =   3375
   End
   Begin VB.ListBox Opcion 
      Height          =   1230
      Left            =   840
      TabIndex        =   26
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Vendedor"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "PrgVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstVendedor As Recordset
Dim spVendedor As String
Dim XParam As String

Sub Verifica_datos()
    Rem If Val(Cuenta.text) = 0 Then
    Rem     Cuenta.text = "0"
    Rem End If
End Sub

Sub Format_datos()
    Rem Comision.text = PUsing("###,###.##", Comision.text)
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


Sub Imprime_Datos()
    spVendedor = "ConsultaVendedor " + "'" + vendedor.Text + "'"
    Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstVendedor.RecordCount > 0 Then
        WPasa = "S"
            Else
        WPasa = "N"
    End If
    If WPasa = "S" Then
        vendedor.Text = rstVendedor!vendedor
        Nombre.Text = rstVendedor!Nombre
    End If
End Sub

Private Sub Acepta_Click()
    
    Listado.WindowTitle = "Listado de Vendedores"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Vendedor.Vendedor} in " + Desde.Text + " to " + Hasta.Text
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    vendedor.SetFocus
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Vendedor.Vendedor , Vendedor.Nombre " _
                       + "From " + DSQ + ".dbo.Vendedor Vendedor " _
                       + "Where Vendedor.Vendedor >= 0 AND Vendedor.Vendedor <= 9999"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If vendedor.Text <> "" Then
    
        spVendedor = "ConsultaVendedor " + "'" + vendedor.Text + "'"
        Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstVendedor.RecordCount > 0 Then
            WPasa = "S"
                Else
            WPasa = "N"
        End If
        
        If WPasa = "N" Then
            XParam = "'" + vendedor.Text + "','" + Nombre.Text + "'"
            Set rstVendedor = db.OpenRecordset("AltaVendedor " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                Else
            XParam = "'" + vendedor.Text + "','" + Nombre.Text + "'"
            Set rstVendedor = db.OpenRecordset("ModificaVendedor " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        Call CmdLimpiar_Click
        vendedor.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If vendedor.Text <> "" Then
    
        spVendedor = "ConsultaVendedor " + "'" + vendedor.Text + "'"
        Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstVendedor.RecordCount > 0 Then
            WPasa = "S"
                Else
            WPasa = "N"
        End If
        
        If WPasa = "S" Then
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spVendedor = "BorrarVendedor " + "'" + vendedor.Text + "'"
                Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    
    End If
    vendedor.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    vendedor.Text = ""
    Nombre.Text = ""
    vendedor.SetFocus
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Rem With rstVendedores
    Rem     .Close
    Rem End With
    Rem With rstEmpresa
    Rem     .Close
    Rem End With
    Rem DbsVentas.Close
    vendedor.SetFocus
    PrgVendedor.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Lista_Click()
    Desde.Text = "0"
    Hasta.Text = "9999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Nombre_Keypress(KeyAscii As Integer)
    'If KeyAscii = 13 Then
    '      Cuenta.SetFocus
    'End If
End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If vendedor.Text <> "" Then
            WVendedor = vendedor.Text
            spVendedor = "ConsultaVendedor " + "'" + vendedor.Text + "'"
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
                
            If WPasa = "S" Then
                vendedor.Text = rstVendedor!vendedor
                Call Imprime_Datos
                    Else
                WVendedor = vendedor.Text
                CmdLimpiar_Click
                vendedor.Text = WVendedor
            End If
        
        End If
        Nombre.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
     Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     'Opcion.Clear
     '
     'Opcion.AddItem "Vendedors"
     'Opcion.AddItem "Cuentas Contables"

     'Opcion.Visible = True
     
'End Sub

'' Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    'XIndice = Opcion.ListIndex
    XIndice = 0
    
    Select Case XIndice
        Case 0
            spVendedor = "ListaVendedor"
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstVendedor
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstVendedor!vendedor) + " " + rstVendedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstVendedor!vendedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstVendedor.Close
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WVendedor = WIndice.List(Indice)
            spVendedor = "ConsultaVendedor " + "'" + Str$(WVendedor) + "'"
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
            
            If WPasa = "S" Then
                vendedor.Text = rstVendedor!vendedor
                Call Imprime_Datos
                        Else
                CmdLimpiar_Click
                vendedor.Text = WVendedor
            End If
            
            vendedor.SetFocus
            
        Case Else
    End Select
    
End Sub

Sub Form_Load()
    vendedor.Text = ""
    Nombre.Text = ""
End Sub


Private Sub Primer_Click()

    On Error GoTo WError
    
    spVendedor = "ListaVendedor"
    Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstVendedor
        .MoveFirst
        vendedor.Text = rstVendedor!vendedor
        Nombre.Text = rstVendedor!Nombre
    End With
    
    rstVendedor.Close
    vendedor.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Vendedor", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     vendedor.SetFocus
 End Sub

Private Sub Ultimo_Click()

   On Error GoTo Error_ultimo
    
    spVendedor = "ListaVendedor"
    Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstVendedor
        .MoveLast
        vendedor.Text = rstVendedor!vendedor
        Nombre.Text = rstVendedor!Nombre
        vendedor.SetFocus
    End With
    rstVendedor.Close
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Vendedor", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     vendedor.SetFocus
 End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    spVendedor = "AnteriorVendedor " + "'" + vendedor.Text + "'"
    Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstVendedor
        .MoveLast
        vendedor.Text = rstVendedor!vendedor
        Nombre.Text = rstVendedor!Nombre
    End With
    
    rstVendedor.Close
    vendedor.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Vendedor", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     vendedor.SetFocus
    
End Sub


Private Sub Siguiente_Click()

    On Error GoTo WError
    
    spVendedor = "PosteriorVendedor " + "'" + vendedor.Text + "'"
    Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstVendedor
        .MoveFirst
        vendedor.Text = rstVendedor!vendedor
        Nombre.Text = rstVendedor!Nombre
    End With
    
    rstVendedor.Close
    vendedor.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Vendedor", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     vendedor.SetFocus
    
End Sub




