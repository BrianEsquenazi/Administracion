VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgModif1 
   AutoRedraw      =   -1  'True
   Caption         =   "Modificacion de Precios"
   ClientHeight    =   5310
   ClientLeft      =   2040
   ClientTop       =   2100
   ClientWidth     =   8400
   LinkTopic       =   "Form2"
   ScaleHeight     =   5310
   ScaleWidth      =   8400
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   -120
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   4935
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton Consulta 
         Caption         =   "Consulta"
         Height          =   375
         Left            =   6240
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.ListBox Pantalla 
         Height          =   3570
         Left            =   2880
         TabIndex        =   18
         Top             =   1080
         Width           =   4575
      End
      Begin VB.ListBox Opcion 
         Height          =   1230
         Left            =   3480
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox Hastacliente 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   14
         Text            =   " "
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox DesdeCliente 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Porcentaje 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   13
         Text            =   " "
         Top             =   2640
         Width           =   1095
      End
      Begin MSMask.MaskEdBox HastaArti 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DesdeArti 
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label HastaDescri 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   2520
         TabIndex        =   16
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label DesdeDescri 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   2520
         TabIndex        =   15
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Porcentaje"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Articulo"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Articulo"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Cliente"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Cliente"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   120
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "listsol.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Solicitudes de Conpras Realizadas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgModif1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String

Dim XParam As String

Dim Vector(10000)

Sub Imprime_Descripcion()

    WCliente = DesdeCliente.Text
    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesdeDescri.Caption = rstCliente!Razon
            Else
        DesdeDescri.Caption = ""
    End If
    
    WCliente = Hastacliente.Text
    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        HastaDescri.Caption = rstCliente!Razon
            Else
        HastaDescri.Caption = ""
    End If
    
End Sub

Private Sub Acepta_Click()
    
    DesdeCliente.Text = UCase(DesdeCliente.Text)
    Hastacliente.Text = UCase(Hastacliente.Text)
    DesdeArti.Text = UCase(DesdeArti.Text)
    HastaArti.Text = UCase(HastaArti.Text)
                
    spPrecios = "ListaPrecios"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    
    Renglon = 0
    Erase Vector
    
    With rstPrecios
        .MoveFirst
        If .NoMatch = False Then
            Do
                If DesdeArti.Text <= rstPrecios!Terminado And HastaArti.Text >= rstPrecios!Terminado Then
                    If DesdeCliente.Text <= rstPrecios!Cliente And Hastacliente.Text >= rstPrecios!Cliente Then
                        Renglon = Renglon + 1
                        Vector(Renglon) = rstPrecios!Clave
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    For XX = 1 To 10000
        If Vector(XX) <> "" Then
            spPrecios = "ConsultaPrecios " + "'" + Vector(XX) + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                WPrecio = Str$(rstPrecios!Precio + (rstPrecios!Precio * Val(Porcentaje.Text) / 100))
                WClave = rstPrecios!Clave
                WCliente = rstPrecios!Cliente
                WTerminado = rstPrecios!Terminado
                WDescripcion = rstPrecios!Descripcion
                WDate = Date$
                                     
                XParam = "'" + WClave + "','" + WPrecio + "','" + WDate + "'"
                Set rstPrecios = db.OpenRecordset("ModificaPrecios3 " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    Next XX
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()

    DesdeCliente.Text = ""
    Hastacliente.Text = ""
    DesdeArti.Text = "  -     -   "
    HastaArti.Text = "  -     -   "
    Porcentaje.Text = ""
    DesdeDescri.Caption = ""
    HastaDescri.Caption = ""
    Frame2.Visible = True
    Opcion.Visible = False
    Pantalla.Visible = False

    Rem With rstPrecios
    Rem    .Close
    Rem End With
    Rem With rstClientes
    Rem     .Close
    Rem End With
    Rem With rstTerminado
    Rem     .Close
    Rem End With
    Rem DbsVentas.Close
    DesdeCliente.SetFocus
    PrgModif.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub DesdeCliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeCliente.Text = UCase(DesdeCliente.Text)
        Call Imprime_Descripcion
        Hastacliente.SetFocus
    End If
End Sub

Private Sub HastaCliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hastacliente.Text = UCase(Hastacliente.Text)
        Call Imprime_Descripcion
        DesdeArti.SetFocus
    End If
End Sub

Private Sub DesdeArti_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeArti.Text = UCase(DesdeArti.Text)
        HastaArti.SetFocus
    End If
End Sub

Private Sub HastaArti_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaArti.Text = UCase(HastaArti.Text)
        Porcentaje.SetFocus
    End If
End Sub

Private Sub Porcentaje_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeCliente.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()
    DesdeCliente.Text = ""
    Hastacliente.Text = ""
    DesdeArti.Text = "  -     -   "
    HastaArti.Text = "  -     -   "
    Porcentaje.Text = ""
    DesdeDescri.Caption = ""
    HastaDescri.Caption = ""
    Frame2.Visible = True
    Opcion.Visible = False
    Pantalla.Visible = False
End Sub

Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False

     Opcion.Clear

     Opcion.AddItem "Clientes"

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
            spCliente = "ListaClienteConsulta"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstCliente
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
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
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WCliente = WIndice.List(Indice)
            spCliente = "ConsultaCliente " + "'" + WCliente + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesdeCliente.Text = rstCliente!Cliente
                Hastacliente.Text = rstCliente!Cliente
                Call Imprime_Descripcion
                DesdeCliente.SetFocus
                        Else
                DesdeCliente.Text = WCliente
                Hastacliente.Text = WCliente
                Call Imprime_Descripcion
                Hastacliente.SetFocus
            End If
            DesdeCliente.SetFocus
            
        Case Else
    End Select
    
End Sub

