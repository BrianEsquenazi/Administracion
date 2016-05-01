VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgConce 
   Caption         =   "Ingreso de Conceptos de Gastos"
   ClientHeight    =   4560
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   5835
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   5835
   Begin VB.TextBox Cuenta 
      Height          =   285
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   27
      Text            =   " "
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Concepto 
      Height          =   285
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   0
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
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   25
         Text            =   " "
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Desde 
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
      ReportFileName  =   "conceptos.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Conceptos"
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
      ItemData        =   "Concepto.frx":0000
      Left            =   480
      List            =   "Concepto.frx":0007
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
   Begin VB.TextBox Descripcion 
      Height          =   285
      Left            =   2640
      MaxLength       =   50
      TabIndex        =   4
      Top             =   360
      Width           =   3375
   End
   Begin VB.ListBox Opcion 
      Height          =   1230
      Left            =   840
      TabIndex        =   29
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label DesCuenta 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   4200
      TabIndex        =   28
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
      Caption         =   "Descripcion de Conceptos"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Concepto de gastos"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   2295
   End
End
Attribute VB_Name = "PrgConce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Imprime_Descripcion()
    With rstCuenta
        .Index = "Cuenta"
        .Seek "=", Cuenta.text
        If .NoMatch = False Then
            DesCuenta.Caption = !Descripcion
                Else
            DesCuenta.Caption = ""
        End If
    End With
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
    With rstConcepto
        .Index = "Concepto"
        .Seek "=", Val(Concepto.text)
        If .NoMatch = False Then
            Concepto.text = !Concepto
            Descripcion.text = !Descripcion
            Cuenta.text = !Cuenta
            Call Format_datos
            Call Imprime_Descripcion
        End If
    End With
End Sub

Private Sub Acepta_Click()
    Listado.GroupSelectionFormula = "{Concepto.Concepto} in " + Desde.text + " to " + Hasta.text
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Concepto.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Concepto.text <> "" Then
        With rstConcepto
            .Index = "Concepto"
            .Seek "=", Val(Concepto.text)
            If .NoMatch Then
                .AddNew
                Call Verifica_datos
                !Concepto = Val(Concepto.text)
                !Descripcion = Descripcion.text
                !Cuenta = Cuenta.text
                Rem !Comision = CDbl(Comision.text)
                .Update
                .Bookmark = .LastModified
                    Else
                .Edit
                Call Verifica_datos
                !Concepto = Val(Concepto.text)
                !Descripcion = Descripcion.text
                !Cuenta = Cuenta.text
                Rem !Comision = CDbl(Comision.text)
                .Update
                .Bookmark = .LastModified
            End If
        End With
        Call CmdLimpiar_Click
        Concepto.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Concepto.text <> "" Then
        With rstConcepto
            .Index = "Concepto"
            .Seek "=", Val(Concepto.text)
            If .NoMatch = False Then
                T$ = "Borrar Registro"
                M$ = "Desea Borrar el Registro "
                Respuesta% = MsgBox(M$, 32 + 4, T$)
                If Respuesta% = 6 Then
                    .Delete
                    Call CmdLimpiar_Click
                End If
            End If
        End With
    End If
    Concepto.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Concepto.text = ""
    Descripcion.text = ""
    Cuenta.text = ""
    DesCuenta = ""
    Concepto.SetFocus
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    With rstConcepto
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstCuenta
        .Close
    End With
    DbsAdminis.Close
    Concepto.SetFocus
    PrgConce.Hide
    Menu.SetFocus
End Sub

Private Sub Anterior_Click()
    With rstConcepto
        .Index = "Concepto"
        .Seek "=", Val(Concepto.text)
        If .NoMatch = False Then
            .MovePrevious
            If .BOF = True Then
                .MoveFirst
                M$ = "No exsite registro Anterior"
                A% = MsgBox(M$, 0, "Archivo de Concepto")
                .MoveFirst
            End If
            Concepto.text = !Concepto
            Call Imprime_Datos
            Concepto.SetFocus
        End If
    End With
End Sub

Private Sub Lista_Click()
    Desde.text = "0"
    Hasta.text = "9999"
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

Private Sub Cuenta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstCuenta
            .Index = "Cuenta"
            .Seek "=", Cuenta.text
            If .NoMatch = False Then
                DesCuenta.Caption = !Descripcion
                Descripcion.SetFocus
                    Else
                Cuenta.SetFocus
            End If
        End With
    End If
End Sub

Private Sub Concepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Concepto.text <> "" Then
            With rstConcepto
                .Index = "Concepto"
                ClaveVen$ = Concepto.text
                .Seek "=", Val(Concepto.text)
                If .NoMatch Then
                    CmdLimpiar_Click
                    Concepto.text = ClaveVen$
                        Else
                    Concepto.text = !Concepto
                    Call Imprime_Datos
                End If
            End With
        End If
        Descripcion.SetFocus
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

     Opcion.AddItem "Conceptos"
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
            With rstConcepto
                .Index = "Concepto"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(!Concepto) + " " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Concepto
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
        Case 1
            With rstCuenta
                .Index = "Cuenta"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Cuenta + " " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Cuenta
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            With rstConcepto

                Indice = Pantalla.ListIndex
                ClaveVen$ = WIndice.List(Indice)
                Concepto.text = Val(ClaveVen$)
                .Index = "Concepto"
                ClaveVen$ = Concepto.text
                .Seek "=", Val(ClaveVen$)
                If .NoMatch = False Then
                    Concepto.text = !Concepto
                    Call Imprime_Datos
                        Else
                    CmdLimpiar_Click
                    Concepto.text = ClaveVen$
                End If
            End With
            Concepto.SetFocus
        Case 1
            With rstCuenta

                Indice = Pantalla.ListIndex
                ClaveVen$ = WIndice.List(Indice)
                Cuenta.text = ClaveVen$
                .Index = "Cuenta"
                ClaveVen$ = Cuenta.text
                .Seek "=", ClaveVen$
                If .NoMatch = False Then
                    Cuenta.text = !Cuenta
                    Call Imprime_Descripcion
                        Else
                    CmdLimpiar_Click
                    Cuenta.text = ClaveVen$
                End If
            End With
            Cuenta.SetFocus
            
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()
    Rem On Error GoTo Error_primer
    With rstConcepto
        .Index = "Concepto"
        .MoveFirst
        Concepto.text = !Concepto
        Call Imprime_Datos
        Concepto.SetFocus
    End With
    Exit Sub

Error_primer:
     coderr = Err
     Call Errores(coderr, "Concepto", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Concepto.SetFocus
 End Sub

Private Sub Ultimo_Click()
    Rem On Error GoTo Error_ultimo
    With rstConcepto
        .Index = "Concepto"
        .MoveLast
        Concepto.text = !Concepto
        Call Imprime_Datos
        Concepto.SetFocus
    End With
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Concepto", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Concepto.SetFocus
 End Sub

Private Sub Siguiente_Click()
    With rstConcepto
        .Index = "Concepto"
        ClaveVen$ = Val(Concepto.text)
        .Seek "=", ClaveVen$
        If .NoMatch = False Then
            .MoveNext
            If .EOF = True Then
                M$ = "No exsite registro Posterior"
                A% = MsgBox(M$, 0, "Archivo de Concepto")
                Call Ultimo_Click
            End If
            Concepto.text = !Concepto
            Call Imprime_Datos
            Concepto.SetFocus
        End If
    End With
End Sub


Sub Form_Load()
    Concepto.text = ""
    Descripcion.text = ""
    Cuenta.text = ""
    DesCuenta = ""
End Sub
