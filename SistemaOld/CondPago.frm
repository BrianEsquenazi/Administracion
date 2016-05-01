VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCondPago 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Condiciones de Pago"
   ClientHeight    =   6345
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   7245
   LinkTopic       =   "Form2"
   ScaleHeight     =   6345
   ScaleWidth      =   7245
   Begin VB.TextBox Descuento 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      MaxLength       =   6
      TabIndex        =   34
      Text            =   " "
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Tasa 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      MaxLength       =   6
      TabIndex        =   33
      Text            =   " "
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Plazo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   32
      Text            =   " "
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Dias 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   31
      Text            =   " "
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Pago 
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
      Top             =   4200
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
      Left            =   5640
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WCondPago.rpt"
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
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "CondPago.frx":0000
      Left            =   480
      List            =   "CondPago.frx":0007
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton lista 
      Caption         =   "Listado"
      Height          =   300
      Left            =   1800
      TabIndex        =   14
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   600
      TabIndex        =   13
      Top             =   3600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles"
      Height          =   1335
      Left            =   4320
      TabIndex        =   8
      Top             =   3120
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
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   3000
      TabIndex        =   7
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   1800
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   600
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Nombre 
      Height          =   285
      Left            =   2640
      MaxLength       =   50
      TabIndex        =   4
      Top             =   480
      Width           =   4215
   End
   Begin VB.ListBox Opcion 
      Height          =   1230
      Left            =   840
      TabIndex        =   26
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Descuento"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "Tasa"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Plazo"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Dias"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Rubro"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "PrgCondPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPago As Recordset
Dim spPago As String
Dim XParam As String

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Sub Verifica_datos()
    If Val(Dias.Text) = 0 Then
         Dias.Text = "0"
    End If
    If Val(Plazo.Text) = 0 Then
         Plazo.Text = "0"
    End If
    If Val(Tasa.Text) = 0 Then
         Tasa.Text = "0"
    End If
    If Val(Descuento.Text) = 0 Then
         Descuento.Text = "0"
    End If
End Sub
Sub Format_datos()
    Tasa.Text = Pusing("###.##", Tasa.Text)
    Descuento.Text = Pusing("###.##", Descuento.Text)
End Sub

Sub Imprime_Datos()
    WPago = Pago.Text
    spPago = "ConsultaPago " + "'" + Pago.Text + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WPasa = "S"
            Else
        WPasa = "N"
    End If
                
    If WPasa = "S" Then
        Pago.Text = rstPago!Pago
        Nombre.Text = rstPago!Nombre
        Dias.Text = rstPago!Dias
        Plazo.Text = rstPago!Plazo
        Tasa.Text = rstPago!Tasa
        Descuento.Text = rstPago!Descuento
    End If
    
End Sub

Private Sub Acepta_Click()

    
    Listado.WindowTitle = "Listado de Pagos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Pago.Pago} in " + Desde.Text + " to " + Hasta.Text
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Pago.SetFocus
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Pago.Pago , Pago.Nombre, Pago.Dias, Pago.Plazo, Pago.Tasa, Pago.Descuento " _
                       + "From " + DSQ + ".dbo.Pago Pago " _
                       + "Where Pago.Pago >= 0 AND Pago.Pago <= 9999"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Pago.Text <> "" Then
    
        spPago = "ConsultaPago " + "'" + Pago.Text + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            WPasa = "S"
                Else
            WPasa = "N"
        End If
        
        Call Verifica_datos
        If WPasa = "N" Then
            XParam = "'" + Pago.Text + "','" + Nombre.Text + "','" + Dias.Text + "','" + Plazo.Text + "','" + Tasa.Text + "','" + Descuento.Text + "'"
            Set rstPago = db.OpenRecordset("AltaPago " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                Else
            XParam = "'" + Pago.Text + "','" + Nombre.Text + "','" + Dias.Text + "','" + Plazo.Text + "','" + Tasa.Text + "','" + Descuento.Text + "'"
            Set rstPago = db.OpenRecordset("ModificaPago " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        Call CmdLimpiar_Click
        Pago.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Pago.Text <> "" Then
        
        spPago = "ConsultaPago " + "'" + Pago.Text + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            WPasa = "S"
                Else
            WPasa = "N"
        End If
        
        If WPasa = "S" Then
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spPago = "BorrarPago " + "'" + Pago.Text + "'"
                Set rstPago = db.OpenRecordset(spPago, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
        
    End If
    Pago.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Pago.Text = ""
    Nombre.Text = ""
    Dias.Text = ""
    Plazo.Text = ""
    Tasa.Text = ""
    Descuento.Text = ""
    Pago.SetFocus
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Rem With rstPago
    Rem     .Close
    Rem End With
    Rem With rstEmpresa
    Rem     .Close
    Rem End With
    Rem With rstAuxiliar
    Rem     .Close
    Rem End With
    Rem DbsVentas.Close
    Pago.SetFocus
    PrgCondPago.Hide
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
    If KeyAscii = 13 Then
          Dias.SetFocus
    End If
End Sub

Private Sub Dias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          Plazo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tasa.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Tasa_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Format_datos
        Descuento.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descuento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Format_datos
        Nombre.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Pago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Pago.Text <> "" Then
            WPago = Pago.Text
            spPago = "ConsultaPago " + "'" + Str$(WPago) + "'"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
                
            If WPasa = "S" Then
                Pago.Text = rstPago!Pago
                Call Imprime_Datos
                    Else
                WPago = Pago.Text
                CmdLimpiar_Click
                Pago.Text = WPago
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
     'Opcion.AddItem "Pagos"
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
            spPago = "ListaPago"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstPago
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstPago!Pago) + " " + rstPago!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstPago!Pago
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPago.Close
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WPago = WIndice.List(Indice)
            spPago = "ConsultaPago " + "'" + Str$(WPago) + "'"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
            
            If WPasa = "S" Then
                Pago.Text = rstPago!Pago
                Call Imprime_Datos
                        Else
                CmdLimpiar_Click
                Pago.Text = WPago
            End If
            
            Pago.SetFocus
            
        Case Else
    End Select
    
End Sub

Sub Form_Load()
    Pago.Text = ""
    Nombre.Text = ""
    Dias.Text = ""
    Plazo.Text = ""
    Tasa.Text = ""
    Descuento.Text = ""
End Sub

Private Sub Primer_Click()

    On Error GoTo WError
    
    spPago = "ListaPago"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstPago
        .MoveFirst
        Pago.Text = rstPago!Pago
    End With
    
    rstPago.Close
    Call Imprime_Datos
    Pago.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Pago", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Pago.SetFocus
 End Sub

Private Sub Ultimo_Click()

   On Error GoTo Error_ultimo
    
    spPago = "ListaPago"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstPago
        .MoveLast
        Pago.Text = rstPago!Pago
        Pago.SetFocus
    End With
    rstPago.Close
    Call Imprime_Datos
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Pago", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Pago.SetFocus
 End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    spPago = "AnteriorPago " + "'" + Pago.Text + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstPago
        .MoveLast
        Pago.Text = rstPago!Pago
    End With
    
    rstPago.Close
    Call Imprime_Datos
    Pago.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Pago", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Pago.SetFocus
    
End Sub


Private Sub Siguiente_Click()

    On Error GoTo WError
    
    spPago = "PosteriorPago " + "'" + Pago.Text + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstPago
        .MoveFirst
        Pago.Text = rstPago!Pago
    End With
    
    rstPago.Close
    Call Imprime_Datos
    Pago.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Pago", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Pago.SetFocus
    
End Sub


