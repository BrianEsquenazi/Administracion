VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgLinea 
   Caption         =   "Articulos"
   ClientHeight    =   4020
   ClientLeft      =   1755
   ClientTop       =   285
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4020
   ScaleWidth      =   5895
   Begin VB.TextBox Linea 
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   23
      Text            =   " "
      Top             =   0
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1455
      Left            =   360
      TabIndex        =   16
      Top             =   2280
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
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5160
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Archivos de programa\DevStudio\VB\materia.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "PrgLin.frx":0000
      Left            =   480
      List            =   "PrgLin.frx":0007
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton lista 
      Caption         =   "Listado"
      Height          =   300
      Left            =   1800
      TabIndex        =   13
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   600
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles"
      Height          =   1335
      Left            =   4320
      TabIndex        =   7
      Top             =   840
      Width           =   1575
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   3000
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Descripcion 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "PrgLinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()
    Listado.GroupSelectionFormula = "{Lineas.Lineas} in " + Desde.text + " to " + Hasta.text
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    lineas.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Linea.text <> "" Then
        With rstLineas
            .Index = "LINEAS"
            .Seek "=", Linea.text
            If .NoMatch Then
                .AddNew
                !lineas = Linea.text
                !Descripcion = Descripcion.text
                .Update
                .Bookmark = .LastModified
                    Else
                .Edit
                !lineas = Linea.text
                !Descripcion = Descripcion.text
                !Rs = Rs.text
                .Bookmark = .LastModified
            End If
        End With
        Call CmdLimpiar_Click
        Linea.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Linea.text <> "" Then
        With rstLineas
            .Index = "LINEAS"
            .Seek "=", Linea.text
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
    Linea.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Linea.text = "  -   -   "
    Descripcion.text = ""
End Sub

Private Sub cmdClose_Click()
    PrgLinea.Hide
    Menu.SetFocus
End Sub

Private Sub Anterior_Click()
    With rstLineas
        .Index = "LINEAS"
        .Seek "=", Linea.text
        If .NoMatch = False Then
            .MovePrevious
            If .BOF = True Then
                .MoveFirst
                M$ = "No exsite registro Anterior"
                A% = MsgBox(M$, 0, "Archivo de Lineas")
                .MoveFirst
            End If
            Linea.text = !lineas
            Descripcion.text = !Descripcion
            Linea.SetFocus
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
        Linea.SetFocus
    End If
End Sub

Sub Linea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Linea.text) <> 0 Then
            With rstLineas
                .Index = "LINEAS"
                ClaveLin$ = Linea.text
                .Seek "=", Linea.text
                If .NoMatch Then
                    CmdLimpiar_Click
                    Linea.text = ClaveLin$
                        Else
                    Linea.text = !lineas
                    Descripcion.text = !Descripcion
                End If
            End With
        End If
        Descripcion.SetFocus
    End If
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
            With rstLineas
                .Index = "LINEAS"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !lineas + " " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !lineas
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
            With rstLineas

                Indice = Pantalla.ListIndex
                ClaveLin$ = WIndice.List(Indice)
                Linea.text = ClaveLin$
                .Index = "LINEAS"
                ClaveLin$ = Linea.text
                .Seek "=", ClaveLin$
                If .NoMatch = False Then
                    Linea.text = !lineas
                    Descripcion.text = !Descripcion
                        Else
                    CmdLimpiar_Click
                    Linea.text = ClaveLin$
                End If
            End With
            Linea.SetFocus
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()
    Rem On Error GoTo Error_primer
    With rstLineas
        .Index = "LINEAS"
        .MoveFirst
        Linea.text = !lineas
        Descripcion.text = !Descripcion
        Linea.SetFocus
    End With
    Exit Sub

Error_primer:
     coderr = Err
     Call Errores(coderr, "Productos", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Linea.SetFocus
 End Sub

Private Sub Ultimo_Click()
    Rem On Error GoTo Error_ultimo
    With rstLineas
        .Index = "LINEAS"
        .MoveLast
        Linea.text = !lineas
        Descripcion.text = !Descripcion
        Linea.SetFocus
    End With
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Prodcuto", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Linea.SetFocus
 End Sub

Private Sub Siguiente_Click()
    With rstLineas
        .Index = "LINEAS"
        ClaveLin$ = Linea.text
        .Seek "=", ClaveLin$
        If .NoMatch = False Then
            .MoveNext
            If .EOF = True Then
                M$ = "No exsite registro Posterior"
                A% = MsgBox(M$, 0, "Archivo de Lineas")
                Call Ultimo_Click
            End If
            Linea.text = !lineas
            Descripcion.text = !Descripcion
            Linea.SetFocus
        End If
    End With
End Sub

