VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgBanco 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Bancos"
   ClientHeight    =   4560
   ClientLeft      =   3555
   ClientTop       =   2175
   ClientWidth     =   5835
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   5835
   Begin VB.TextBox Banco 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   30
      Text            =   " "
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4560
      TabIndex        =   29
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Cuenta 
      Height          =   285
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   1
      Text            =   " "
      Top             =   720
      Width           =   1335
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
      Left            =   4560
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wbancos.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Bancos"
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
      ItemData        =   "banco.frx":0000
      Left            =   480
      List            =   "banco.frx":0007
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
      Top             =   1320
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
   Begin VB.TextBox Nombre 
      Height          =   285
      Left            =   2640
      MaxLength       =   50
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.ListBox Opcion 
      Height          =   1230
      Left            =   840
      TabIndex        =   28
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label DesCuenta 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   4200
      TabIndex        =   27
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Cuenta Contable"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nombre del Banco"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Bancos"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   2295
   End
End
Attribute VB_Name = "PrgBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstBanco As Recordset
Dim spBanco As String
Dim rstCuenta As Recordset
Dim spCuenta As String
Dim XParam As String
Dim x As Printer
Dim rstDada As Recordset
Dim spDada As String

Sub Imprime_Nombre()
    spCuenta = "ConsultaCuentas " + "'" + Cuenta.Text + "'"
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        DesCuenta.Caption = rstCuenta!Descripcion
        rstCuenta.Close
            Else
        DesCuenta.Caption = ""
    End If
End Sub

Sub Verifica_datos()
    Rem If Val(Cuenta.text) = 0 Then
    Rem     Cuenta.text = "0"
    Rem End If
End Sub
Sub Format_datos()
    Rem Comision.text = PUsing("###,###.##", Comision.text)
End Sub

Sub Imprime_Datos()
    spBanco = "ConsultaBanco " + "'" + Banco.Text + "'"
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        Banco.Text = rstBanco!Banco
        Nombre.Text = rstBanco!Nombre
        Cuenta.Text = rstBanco!Cuenta
        rstBanco.Close
        Call Format_datos
        Call Imprime_Nombre
    End If
End Sub

Private Sub Acepta_Click()
    Listado.GroupSelectionFormula = "{Banco.Banco} in " + Desde.Text + " to " + Hasta.Text
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Banco.SetFocus
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT Banco.Banco, Banco.Nombre, Banco.Cuenta, Cuenta.Descripcion " _
                       + "From " + DSQ + ".dbo.Banco Banco, " _
                       + DSQ + ".dbo.Cuenta Cuenta " _
                       + "Where Banco.Cuenta = Cuenta.Cuenta AND Banco.Banco >= 0 AND Banco.Banco <= 9999"
    Listado.DataFiles(2) = WEmpresa + "Auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Banco.Text <> "" Then
    
        spBanco = "ConsultaBanco " + "'" + Banco.Text + "'"
        Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
        If rstBanco.RecordCount > 0 Then
            XParam = "'" + Banco.Text + "','" _
                         + Nombre.Text + "','" _
                         + Cuenta.Text + "','" _
                         + "1" + "'"
            Set rstBanco = db.OpenRecordset("ModificaBanco " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                Else
            XParam = "'" + Banco.Text + "','" _
                         + Nombre.Text + "','" _
                         + Cuenta.Text + "','" _
                         + "1" + "'"
            Set rstBanco = db.OpenRecordset("AltaBanco " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        Call CmdLimpiar_Click
        Banco.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Banco.Text <> "" Then
    
        spBanco = "ConsultaBanco " + "'" + Banco.Text + "'"
        Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
        If rstBanco.RecordCount > 0 Then
            WPasa = "S"
                Else
            WPasa = "N"
        End If
        
        If WPasa = "S" Then
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spBanco = "BorrarBanco " + "'" + Banco.Text + "'"
                Set rstBanco = db.OpenRecordset(spBanco, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    
    End If
    Banco.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Banco.Text = ""
    Nombre.Text = ""
    Cuenta.Text = ""
    DesCuenta = ""
    Banco.SetFocus
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Banco.SetFocus
    PrgBanco.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Command1_Click()
    aa = "'"
    cc = Asc(aa)
    Stop
End Sub

Private Sub Command144_Click()

End Sub

Private Sub Command23_Click()

    Rem Open "c:\reactor1.jpg" For Binary Access Read Lock Read As #1
        
    Dim MiCadena, ZGraba
    ZGraba = ""
    Open "c:\reactor1.jpg" For Binary Access Read As #1   ' Abre el archivo para recibir los datos.
    Do While Not EOF(1) ' Repite el bucle hasta el final del archivo.
        Input #1, MiCadena
        ZGraba = ZGraba + MiCadena
        Debug.Print MiCadena, ZGraba
    Loop
    Close #1    ' Cierra el archivo.


    ZSql = ""
    ZSql = ZSql + "INSERT INTO Dada ("
    ZSql = ZSql + "Codigo ,"
    ZSql = ZSql + "Imagen )"
    ZSql = ZSql + "Values ("
    ZSql = ZSql + "'" + "4" + "',"
    ZSql = ZSql + "'" + ZGraba + "')"
        
    spDada = ZSql
    Set rstDada = db.OpenRecordset(spDada, dbOpenSnapshot, dbSQLPassThrough)

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
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
        Cuenta.SetFocus
    End If
End Sub

Private Sub Cuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spCuenta = "ConsultaCuentas " + "'" + Cuenta.Text + "'"
        Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
        If rstCuenta.RecordCount > 0 Then
            DesCuenta.Caption = rstCuenta!Descripcion
            rstCuenta.Close
            Nombre.SetFocus
                Else
            Cuenta.SetFocus
        End If
    End If
End Sub

Private Sub Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Banco.Text <> "" Then
            WBanco = Banco.Text
            spBanco = "ConsultaBanco " + "'" + Banco.Text + "'"
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                Banco.Text = rstBanco!Banco
                rstBanco.Close
                Call Imprime_Datos
                    Else
                WBanco = Banco.Text
                CmdLimpiar_Click
                Banco.Text = WBanco
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

     Opcion.Clear

     Opcion.AddItem "Bancos"
     Opcion.AddItem "Cuentas Contables"

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
            spBanco = "ListaBancos"
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstBanco
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstBanco!Banco) + " " + rstBanco!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstBanco!Banco
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstBanco.Close
            
        Case 1
            spCuenta = "ListaCuentas"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstCuenta
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstCuenta!Cuenta) + " " + rstCuenta!Descripcion
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
            WBanco = WIndice.List(Indice)
            spBanco = "ConsultaBanco " + "'" + Str$(WBanco) + "'"
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                Banco.Text = rstBanco!Banco
                rstBanco.Close
                Call Imprime_Datos
                        Else
                CmdLimpiar_Click
                Banco.Text = WBanco
            End If
            Banco.SetFocus
        Case 1
        
            Indice = Pantalla.ListIndex
            WCuenta = WIndice.List(Indice)
            spCuenta = "ConsultaCuentas " + "'" + WCuenta + "'"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                Cuenta.Text = rstCuenta!Cuenta
                rstCuenta.Close
                Call Imprime_Nombre
                       Else
                Cuenta.Text = WCuenta
            End If
            Cuenta.SetFocus
            
        Case Else
    End Select
    
End Sub

Sub Form_Load()
    Banco.Text = ""
    Nombre.Text = ""
    Cuenta.Text = ""
    DesCuenta = ""
End Sub

Private Sub Primer_Click()

    On Error GoTo WError
    
    spBanco = "ListaBancos"
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        With rstBanco
            .MoveFirst
            Banco.Text = rstBanco!Banco
            rstBanco.Close
            Call Imprime_Datos
        End With
    End If
    Banco.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Banco", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Banco.SetFocus
 End Sub

Private Sub Ultimo_Click()

   On Error GoTo Error_ultimo
    
    spBanco = "ListaBancos"
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        With rstBanco
            .MoveLast
            Banco.Text = rstBanco!Banco
            rstBanco.Close
            Call Imprime_Datos
        End With
    End If
    Banco.SetFocus
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Banco", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Banco.SetFocus
 End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    spBanco = "AnteriorBanco " + "'" + Banco.Text + "'"
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        With rstBanco
            .MoveLast
            Banco.Text = rstBanco!Banco
            rstBanco.Close
            Call Imprime_Datos
        End With
    End If
    
    Banco.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Banco", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Banco.SetFocus
    
End Sub


Private Sub Siguiente_Click()

    On Error GoTo WError
    
    spBanco = "PosteriorBanco " + "'" + Banco.Text + "'"
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        With rstBanco
            .MoveFirst
            Banco.Text = rstBanco!Banco
            rstBanco.Close
            Call Imprime_Datos
        End With
    End If
    
    Banco.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Banco", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Banco.SetFocus
    
End Sub



