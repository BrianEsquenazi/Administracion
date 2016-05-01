VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEspecifi 
   Caption         =   "Ensayos"
   ClientHeight    =   6255
   ClientLeft      =   1875
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form2"
   ScaleHeight     =   6255
   ScaleWidth      =   5925
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   1200
      TabIndex        =   61
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin Crystal.CrystalReport lista 
      Left            =   4560
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Archivos de programa\DevStudio\VB\especi1.rpt"
      GroupSelectionFormula=   " "
      DiscardSavedData=   -1  'True
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control de Listado"
      Height          =   1575
      Left            =   0
      TabIndex        =   51
      Top             =   4800
      Visible         =   0   'False
      Width           =   3135
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1320
         TabIndex        =   63
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1320
         TabIndex        =   62
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton ImpreListado 
         Caption         =   "Option2"
         Height          =   195
         Left            =   1920
         TabIndex        =   57
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton ImprePantalla 
         Caption         =   "Option1"
         Height          =   195
         Left            =   1920
         TabIndex        =   56
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   960
         TabIndex        =   55
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Impresora"
         Height          =   255
         Left            =   2160
         TabIndex        =   59
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Pantalla"
         Height          =   255
         Left            =   2160
         TabIndex        =   58
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta  Codigo"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ListBox Opcion 
      Height          =   1035
      Left            =   960
      TabIndex        =   50
      Top             =   4800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   5160
      TabIndex        =   49
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox valor10 
      Height          =   285
      Left            =   3480
      TabIndex        =   47
      Text            =   " "
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox valor9 
      Height          =   285
      Left            =   3480
      TabIndex        =   46
      Text            =   " "
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox valor8 
      Height          =   285
      Left            =   3480
      TabIndex        =   45
      Text            =   " "
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox valor7 
      Height          =   285
      Left            =   3480
      TabIndex        =   44
      Text            =   " "
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox valor6 
      Height          =   285
      Left            =   3480
      TabIndex        =   43
      Text            =   " "
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox valor5 
      Height          =   285
      Left            =   3480
      TabIndex        =   42
      Text            =   " "
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox valor4 
      Height          =   285
      Left            =   3480
      TabIndex        =   41
      Text            =   " "
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Valor3 
      Height          =   285
      Left            =   3480
      TabIndex        =   40
      Text            =   " "
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox valor2 
      Height          =   285
      Left            =   3480
      TabIndex        =   39
      Text            =   " "
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Valor1 
      Height          =   285
      Left            =   3480
      TabIndex        =   38
      Text            =   " "
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Ensayo10 
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   37
      Text            =   " "
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Ensayo9 
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   36
      Text            =   " "
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox Ensayo8 
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   35
      Text            =   " "
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Ensayo7 
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   34
      Text            =   " "
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox Ensayo6 
      Height          =   315
      Left            =   120
      MaxLength       =   4
      TabIndex        =   33
      Text            =   " "
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Ensayo5 
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   32
      Text            =   " "
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Ensayo4 
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   31
      Text            =   " "
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Ensayo3 
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   30
      Text            =   " "
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Ensayo2 
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   29
      Text            =   " "
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Ensayo1 
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   28
      Text            =   " "
      Top             =   1080
      Width           =   735
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox pantalla 
      Height          =   1425
      ItemData        =   "PrgEspe.frx":0000
      Left            =   0
      List            =   "PrgEspe.frx":0007
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Listado 
      Caption         =   "Listado"
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   5400
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   3120
      TabIndex        =   5
      Top             =   4680
      Width           =   1575
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   4800
      TabIndex        =   4
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4800
      TabIndex        =   3
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   4800
      TabIndex        =   2
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   4800
      TabIndex        =   1
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   120
      TabIndex        =   60
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Descriprod 
      Caption         =   " "
      Height          =   255
      Left            =   1080
      TabIndex        =   48
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Descri10 
      Caption         =   " "
      Height          =   375
      Left            =   1080
      TabIndex        =   27
      Top             =   4320
      Width           =   2100
   End
   Begin VB.Label Descri9 
      Caption         =   " "
      Height          =   255
      Left            =   1080
      TabIndex        =   26
      Top             =   3960
      Width           =   2100
   End
   Begin VB.Label Descri8 
      Caption         =   " "
      Height          =   255
      Left            =   1080
      TabIndex        =   25
      Top             =   3600
      Width           =   2100
   End
   Begin VB.Label Descri7 
      Caption         =   " "
      Height          =   255
      Left            =   1080
      TabIndex        =   24
      Top             =   3240
      Width           =   2100
   End
   Begin VB.Label Descri6 
      Caption         =   " "
      Height          =   255
      Left            =   1080
      TabIndex        =   23
      Top             =   2880
      Width           =   2100
   End
   Begin VB.Label Descri5 
      Caption         =   " "
      Height          =   255
      Left            =   1080
      TabIndex        =   22
      Top             =   2520
      Width           =   2100
   End
   Begin VB.Label Descri4 
      Caption         =   " "
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   2160
      Width           =   2100
   End
   Begin VB.Label Descri3 
      Caption         =   " "
      Height          =   375
      Left            =   1080
      TabIndex        =   20
      Top             =   1800
      Width           =   2100
   End
   Begin VB.Label descri2 
      Caption         =   " "
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   1440
      Width           =   2100
   End
   Begin VB.Label Descri1 
      Caption         =   " "
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   1080
      Width           =   2100
   End
   Begin VB.Label lblresultado 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor Standard"
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblDescri 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblensayo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ensayo"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   14
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "PrgEspecifi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Acepta_Click()
    lista.GroupSelectionFormula = "{Especificaciones.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If ImpreListado.Value = True Then
        lista.Destination = 1
            Else
        lista.Destination = 0
    End If
    lista.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Producto.Text <> "" Then
        With rstEspecificaciones
            .Index = "Producto"
            ClaveProd$ = Producto.Text
            .Seek "=", ClaveProd$
            If .NoMatch Then
                .AddNew
                TipoImpre = "1"
                Call imprime_Click
                .Update
                .Bookmark = .LastModified
                    Else
                .Edit
                TipoImpre = "1"
                Call imprime_Click
                .Update
                .Bookmark = .LastModified
            End If
        End With
        Call CmdLimpiar_Click
        Producto.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Producto.Text <> "" Then
        With rstEspecificaciones
            .Index = "Producto"
            ClaveProd$ = Producto.Text
            .Seek "=", ClaveProd$
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
    Producto.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Producto.Text = "  -   -   "
    Ensayo1.Text = ""
    Valor1.Text = ""
    Ensayo2.Text = ""
    valor2.Text = ""
    Ensayo3.Text = ""
    Valor3.Text = ""
    Ensayo4.Text = ""
    valor4.Text = ""
    Ensayo5.Text = ""
    valor5.Text = ""
    Ensayo6.Text = ""
    valor6.Text = ""
    Ensayo7.Text = ""
    valor7.Text = ""
    Ensayo8.Text = ""
    valor8.Text = ""
    Ensayo9.Text = ""
    valor9.Text = ""
    Ensayo10.Text = ""
    valor10.Text = ""
    Descriprod.Caption = ""
    Descri1.Caption = ""
    descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    Descri6.Caption = ""
    Descri7.Caption = ""
    Descri8.Caption = ""
    Descri9.Caption = ""
    Descri10.Caption = ""
    Producto.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgEspecifi.Hide
    Menu.SetFocus
End Sub

Private Sub Anterior_Click()
    With rstEspecificaciones
        .Index = "Producto"
        ClaveProd$ = Producto.Text
        .Seek "=", ClaveProd$
        If .NoMatch = False Then
            .MovePrevious
            If .BOF = True Then
                .MoveFirst
                M$ = "No exsite registro Anterior"
                A% = MsgBox(M$, 0, "Archivo de Ensayos")
                .MoveFirst
            End If
            TipoImpre = "2"
            Call imprime_Click
            Producto.SetFocus
        End If
    End With
End Sub





Private Sub Ensayo1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstEnsayos
            .Index = "Codigo"
            .Seek "=", Val(Ensayo1.Text)
            If .NoMatch Then
                Descri1.Caption = ""
                        Else
                Descri1.Caption = !Descripcion
                Valor1.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstEnsayos
            .Index = "Codigo"
            .Seek "=", Val(Ensayo2.Text)
            If .NoMatch Then
                descri2.Caption = ""
                        Else
                descri2.Caption = !Descripcion
                valor2.SetFocus
            End If
        End With
    End If
End Sub

Private Sub Ensayo3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstEnsayos
            .Index = "Codigo"
            .Seek "=", Val(Ensayo3.Text)
            If .NoMatch Then
                Descri3.Caption = ""
                        Else
                Descri3.Caption = !Descripcion
                Valor3.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstEnsayos
            .Index = "Codigo"
            .Seek "=", Val(Ensayo4.Text)
            If .NoMatch Then
                Descri4.Caption = ""
                        Else
                Descri4.Caption = !Descripcion
                valor4.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstEnsayos
            .Index = "Codigo"
            .Seek "=", Val(Ensayo5.Text)
            If .NoMatch Then
                Descri5.Caption = ""
                           Else
                Descri5.Caption = !Descripcion
                valor5.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstEnsayos
            .Index = "Codigo"
            .Seek "=", Val(Ensayo6.Text)
            If .NoMatch Then
                Descri6.Caption = ""
                        Else
                Descri6.Caption = !Descripcion
                valor6.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo7_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstEnsayos
            .Index = "Codigo"
            .Seek "=", Val(Ensayo7.Text)
            If .NoMatch Then
                Descri7.Caption = ""
                        Else
                Descri7.Caption = !Descripcion
                valor7.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo8_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstEnsayos
            .Index = "Codigo"
            .Seek "=", Val(Ensayo8.Text)
            If .NoMatch Then
                Descri8.Caption = ""
                        Else
                Descri8.Caption = !Descripcion
                valor8.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo9_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstEnsayos
            .Index = "Codigo"
            .Seek "=", Val(Ensayo9.Text)
            If .NoMatch Then
                Descri9.Caption = ""
                        Else
                Descri9.Caption = !Descripcion
                valor9.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo10_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstEnsayos
            .Index = "Codigo"
            .Seek "=", Val(Ensayo10.Text)
            If .NoMatch Then
                Descri10.Caption = ""
                        Else
                Descri10.Caption = !Descripcion
                valor10.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub Listado_Click()
    Desde.Text = "AA-000-000"
    Hasta.Text = "ZZ-999-999"
    ImprePantalla.Value = False
    ImpreListado.Value = True
    Frame2.Visible = True
End Sub



Private Sub Valor1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo2.SetFocus
    End If
End Sub
Private Sub Valor2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo3.SetFocus
    End If
End Sub
Private Sub Valor3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo4.SetFocus
    End If
End Sub
Private Sub Valor4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo5.SetFocus
    End If
End Sub
Private Sub Valor5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo6.SetFocus
    End If
End Sub
Private Sub Valor6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo7.SetFocus
    End If
End Sub
Private Sub Valor7_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo8.SetFocus
    End If
End Sub
Private Sub Valor8_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo9.SetFocus
    End If
End Sub
Private Sub Valor9_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo10.SetFocus
    End If
End Sub
Private Sub Valor10_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo1.SetFocus
    End If
End Sub

Private Sub imprime_Click()
    Select Case TipoImpre
        Case "1"
            With rstEspecificaciones
                !Producto = Producto.Text
                !Ensayo1 = Val(Ensayo1.Text)
                !Ensayo2 = Val(Ensayo2.Text)
                !Ensayo3 = Val(Ensayo3.Text)
                !Ensayo4 = Val(Ensayo4.Text)
                !Ensayo5 = Val(Ensayo5.Text)
                !Ensayo6 = Val(Ensayo6.Text)
                !Ensayo7 = Val(Ensayo7.Text)
                !Ensayo8 = Val(Ensayo8.Text)
                !Ensayo9 = Val(Ensayo9.Text)
                !Ensayo10 = Val(Ensayo10.Text)
                !Valor1 = Valor1.Text
                !valor2 = valor2.Text
                !Valor3 = Valor3.Text
                !valor4 = valor4.Text
                !valor5 = valor5.Text
                !valor6 = valor6.Text
                !valor7 = valor7.Text
                !valor8 = valor8.Text
                !valor9 = valor9.Text
                !valor10 = valor10.Text
            End With
        Case "2"
            With rstEspecificaciones
                Producto.Text = !Producto
                Ensayo1.Text = !Ensayo1
                Ensayo2.Text = !Ensayo2
                Ensayo3.Text = !Ensayo3
                Ensayo4.Text = !Ensayo4
                Ensayo5.Text = !Ensayo5
                Ensayo6.Text = !Ensayo6
                Ensayo7.Text = !Ensayo7
                Ensayo8.Text = !Ensayo8
                Ensayo9.Text = !Ensayo9
                Ensayo10.Text = !Ensayo10
                Valor1.Text = !Valor1
                valor2.Text = !valor2
                Valor3.Text = !Valor3
                valor4.Text = !valor4
                valor5.Text = !valor5
                valor6.Text = !valor6
                valor7.Text = !valor7
                valor8.Text = !valor8
                valor9.Text = !valor9
                valor10.Text = !valor10
            End With
            With rstEnsayos
                .Index = "Codigo"
                .Seek "=", Val(Ensayo1.Text)
                If .NoMatch Then
                    Descri1.Caption = ""
                        Else
                    Descri1.Caption = !Descripcion
                End If
                .Seek "=", Val(Ensayo2.Text)
                If .NoMatch Then
                    descri2.Caption = ""
                        Else
                    descri2.Caption = !Descripcion
                End If
                .Seek "=", Val(Ensayo3.Text)
                If .NoMatch Then
                    Descri3.Caption = ""
                        Else
                    Descri3.Caption = !Descripcion
                End If
                .Seek "=", Val(Ensayo4.Text)
                If .NoMatch Then
                    Descri4.Caption = ""
                        Else
                    Descri4.Caption = !Descripcion
                End If
                .Seek "=", Val(Ensayo5.Text)
                If .NoMatch Then
                    Descri5.Caption = ""
                        Else
                    Descri5.Caption = !Descripcion
                End If
                .Seek "=", Val(Ensayo6.Text)
                If .NoMatch Then
                    Descri6.Caption = ""
                        Else
                    Descri6.Caption = !Descripcion
                End If
                .Seek "=", Val(Ensayo7.Text)
                If .NoMatch Then
                    Descri7.Caption = ""
                        Else
                    Descri7.Caption = !Descripcion
                End If
                .Seek "=", Val(Ensayo8.Text)
                If .NoMatch Then
                    Descri8.Caption = ""
                        Else
                    Descri8.Caption = !Descripcion
                End If
                .Seek "=", Val(Ensayo9.Text)
                If .NoMatch Then
                    Descri9.Caption = ""
                        Else
                    Descri9.Caption = !Descripcion
                End If
                .Seek "=", Val(Ensayo10.Text)
                If .NoMatch Then
                    Descri10.Caption = ""
                        Else
                    Descri10.Caption = !Descripcion
                End If
            End With
            
            With rstProductos
                .Index = "Producto"
                ClaveProd$ = Producto.Text
                .Seek "=", ClaveProd$
                If .NoMatch = True Then
                    Producto.SetFocus
                        Else
                    Descriprod.Caption = !Descripcion
                End If
            End With

        Case Else
    End Select
End Sub
Sub Producto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "" Then
            With rstEspecificaciones
                .Index = "Producto"
                ClaveProd$ = Producto.Text
                .Seek "=", ClaveProd$
                If .NoMatch Then
                    CmdLimpiar_Click
                    Producto.Text = ClaveProd$
                        Else
                    TipoImpre = "2"
                    Call imprime_Click
                End If
            End With
            With rstProductos
                .Index = "Producto"
                ClaveProd$ = Producto.Text
                .Seek "=", ClaveProd$
                If .NoMatch = True Then
                    Producto.SetFocus
                    Exit Sub
                        Else
                    Descriprod.Caption = !Descripcion
                End If
            End With
        End If
        Ensayo1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()
    Opcion.Clear
    
    Opcion.AddItem "Productos"
    Opcion.AddItem "Ensayos"
    
    Opcion.Visible = True
End Sub

Private Sub Opcion_Click()
    Opcion.Visible = False
    Dim IngresaItem As String

    pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            With rstProductos
                .Index = "Producto"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Producto + " " + !Descripcion
                        pantalla.AddItem IngresaItem
                        IngresaItem = !Producto
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case 1
            With rstEnsayos
                .Index = "Codigo"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(!Codigo) + " " + !Descripcion
                        pantalla.AddItem IngresaItem
                        IngresaItem = Str$(!Codigo)
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
        Case Else
    End Select
            
    pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    pantalla.Visible = False
    Select Case XIndice
        Case 0
            With rstEspecificaciones

                Indice = pantalla.ListIndex
                ClaveProd$ = WIndice.List(Indice)
                Producto.Text = ClaveProd$
                .Index = "Producto"
                ClaveProd$ = Producto.Text
                .Seek "=", ClaveProd$
                If .NoMatch = False Then
                    TipoImpre = "2"
                    Call imprime_Click
                        Else
                    CmdLimpiar_Click
                    Producto.Text = ClaveProd$
                    With rstProductos
                        .Index = "Producto"
                        ClaveProd$ = Producto.Text
                        .Seek "=", ClaveProd$
                        If .NoMatch = True Then
                            Producto.SetFocus
                                Else
                            Descriprod.Caption = !Descripcion
                        End If
                    End With
                End If
            End With
            Producto.SetFocus
        Case 1
            Entra$ = "S"
            If Val(Ensayo1.Text) = 0 And Entra$ = "S" Then
                    Indice = pantalla.ListIndex
                    Ensayo1.Text = Val(WIndice.List(Indice))
                    Valor1.SetFocus
                    Entra$ = "N"
                    With rstEnsayos
                        .Index = "Codigo"
                        .Seek "=", Val(Ensayo1.Text)
                        If .NoMatch = False Then
                            Descri1.Caption = !Descripcion
                        End If
                    End With
            End If
            
            If Val(Ensayo2.Text) = 0 And Entra$ = "S" Then
                    Indice = pantalla.ListIndex
                    Ensayo2.Text = Val(WIndice.List(Indice))
                    valor2.SetFocus
                    Entra$ = "N"
                    With rstEnsayos
                        .Index = "Codigo"
                        .Seek "=", Val(Ensayo2.Text)
                        If .NoMatch = False Then
                            descri2.Caption = !Descripcion
                        End If
                    End With
            End If
            
            If Val(Ensayo3.Text) = 0 And Entra$ = "S" Then
                    Indice = pantalla.ListIndex
                    Ensayo3.Text = Val(WIndice.List(Indice))
                    Valor3.SetFocus
                    Entra$ = "N"
                    With rstEnsayos
                        .Index = "Codigo"
                        .Seek "=", Val(Ensayo3.Text)
                        If .NoMatch = False Then
                            Descri3.Caption = !Descripcion
                        End If
                    End With
            End If
            If Val(Ensayo4.Text) = 0 And Entra$ = "S" Then
                    Indice = pantalla.ListIndex
                    Ensayo4.Text = Val(WIndice.List(Indice))
                    valor4.SetFocus
                    Entra$ = "N"
                    With rstEnsayos
                        .Index = "Codigo"
                        .Seek "=", Val(Ensayo4.Text)
                        If .NoMatch = False Then
                            Descri4.Caption = !Descripcion
                        End If
                    End With
            End If
            If Val(Ensayo5.Text) = 0 And Entra$ = "S" Then
                    Indice = pantalla.ListIndex
                    Ensayo5.Text = Val(WIndice.List(Indice))
                    valor5.SetFocus
                    Entra$ = "N"
                    With rstEnsayos
                        .Index = "Codigo"
                        .Seek "=", Val(Ensayo5.Text)
                        If .NoMatch = False Then
                            Descri5.Caption = !Descripcion
                        End If
                    End With
            End If
            If Val(Ensayo6.Text) = 0 And Entra$ = "S" Then
                    Indice = pantalla.ListIndex
                    Ensayo6.Text = Val(WIndice.List(Indice))
                    valor6.SetFocus
                    Entra$ = "N"
                    With rstEnsayos
                        .Index = "Codigo"
                        .Seek "=", Val(Ensayo6.Text)
                        If .NoMatch = False Then
                            Descri6.Caption = !Descripcion
                        End If
                    End With
            End If
            If Val(Ensayo7.Text) = 0 And Entra$ = "S" Then
                    Indice = pantalla.ListIndex
                    Ensayo7.Text = Val(WIndice.List(Indice))
                    valor7.SetFocus
                    Entra$ = "N"
                    With rstEnsayos
                        .Index = "Codigo"
                        .Seek "=", Val(Ensayo7.Text)
                        If .NoMatch = False Then
                            Descri7.Caption = !Descripcion
                        End If
                    End With
            End If
            If Val(Ensayo8.Text) = 0 And Entra$ = "S" Then
                    Indice = pantalla.ListIndex
                    Ensayo8.Text = Val(WIndice.List(Indice))
                    valor8.SetFocus
                    Entra$ = "N"
                    With rstEnsayos
                        .Index = "Codigo"
                        .Seek "=", Val(Ensayo8.Text)
                        If .NoMatch = False Then
                            Descri8.Caption = !Descripcion
                        End If
                    End With
            End If
            If Val(Ensayo9.Text) = 0 And Entra$ = "S" Then
                    Indice = pantalla.ListIndex
                    Ensayo9.Text = Val(WIndice.List(Indice))
                    valor9.SetFocus
                    Entra$ = "N"
                    With rstEnsayos
                        .Index = "Codigo"
                        .Seek "=", Val(Ensayo9.Text)
                        If .NoMatch = False Then
                            Descri9.Caption = !Descripcion
                        End If
                    End With
            End If
            If Val(Ensayo10.Text) = 0 And Entra$ = "S" Then
                    Indice = pantalla.ListIndex
                    Ensayo10.Text = Val(WIndice.List(Indice))
                    valor10.SetFocus
                    Entra$ = "N"
                    With rstEnsayos
                        .Index = "Codigo"
                        .Seek "=", Val(Ensayo10.Text)
                        If .NoMatch = False Then
                            Descri10.Caption = !Descripcion
                        End If
                    End With
            End If
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()
    On Error GoTo Error_primer
    With rstEspecificaciones
        .Index = "Producto"
        .MoveFirst
        TipoImpre = "2"
        Call imprime_Click
        Producto.SetFocus
    End With
    Exit Sub

Error_primer:
     coderr = Err
     Call Errores(coderr, "Especificaciones", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Producto.SetFocus
 End Sub

Private Sub Ultimo_Click()
    On Error GoTo Error_ultimo
    With rstEspecificaciones
        .Index = "Producto"
        .MoveLast
        TipoImpre = "2"
        Call imprime_Click
        Producto.SetFocus
    End With
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Ensayos", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Producto.SetFocus
 End Sub

Private Sub Siguiente_Click()
    With rstEspecificaciones
        .Index = "Producto"
        ClaveProd$ = Producto.Text
        .Seek "=", ClaveProd$
        If .NoMatch = False Then
            .MoveNext
            If .EOF = True Then
                M$ = "No exsite registro Posterior"
                A% = MsgBox(M$, 0, "Archivo de Ensayos")
                Call Ultimo_Click
            End If
            TipoImpre = "2"
            Call imprime_Click
            Producto.SetFocus
        End If
    End With
End Sub


