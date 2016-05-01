VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPruedev 
   Caption         =   "Ingreso de Devolucion de Re o Nk"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   540
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   ScaleHeight     =   9480
   ScaleWidth      =   15240
   Begin VB.TextBox NroDevol 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5280
      TabIndex        =   67
      Text            =   " "
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Cantidad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   66
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox LoteOriginal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7920
      MaxLength       =   6
      TabIndex        =   64
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Cliente 
      Height          =   285
      Left            =   7920
      MaxLength       =   6
      TabIndex        =   61
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Partida 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5280
      TabIndex        =   59
      Text            =   " "
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Impensayo 
      Caption         =   "Impresion Prueba"
      Height          =   495
      Left            =   12360
      TabIndex        =   57
      Top             =   7800
      Width           =   1455
   End
   Begin MSMask.MaskEdBox fecha 
      Height          =   285
      Left            =   3120
      TabIndex        =   45
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Confecciono 
      Height          =   285
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   43
      Text            =   " "
      Top             =   7920
      Width           =   3975
   End
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   42
      Text            =   " "
      Top             =   7680
      Width           =   3975
   End
   Begin VB.TextBox Aspecto 
      Height          =   285
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   41
      Text            =   " "
      Top             =   7440
      Width           =   3975
   End
   Begin VB.TextBox Ensayo 
      Height          =   285
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   40
      Text            =   " "
      Top             =   7200
      Width           =   3975
   End
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "AA-#####-###"
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
      Height          =   1020
      Left            =   5280
      TabIndex        =   34
      Top             =   7320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   9480
      TabIndex        =   33
      Top             =   6960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox valor10 
      Height          =   285
      Left            =   11160
      MaxLength       =   50
      TabIndex        =   31
      Text            =   " "
      Top             =   6600
      Width           =   3855
   End
   Begin VB.TextBox valor9 
      Height          =   285
      Left            =   11160
      MaxLength       =   50
      TabIndex        =   30
      Text            =   " "
      Top             =   6000
      Width           =   3855
   End
   Begin VB.TextBox valor8 
      Height          =   285
      Left            =   11160
      MaxLength       =   50
      TabIndex        =   29
      Text            =   " "
      Top             =   5400
      Width           =   3855
   End
   Begin VB.TextBox valor7 
      Height          =   285
      Left            =   11160
      MaxLength       =   50
      TabIndex        =   28
      Text            =   " "
      Top             =   4800
      Width           =   3855
   End
   Begin VB.TextBox valor6 
      Height          =   285
      Left            =   11160
      MaxLength       =   50
      TabIndex        =   27
      Text            =   " "
      Top             =   4200
      Width           =   3855
   End
   Begin VB.TextBox valor5 
      Height          =   285
      Left            =   11160
      MaxLength       =   50
      TabIndex        =   26
      Text            =   " "
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox valor4 
      Height          =   285
      Left            =   11160
      MaxLength       =   50
      TabIndex        =   25
      Text            =   " "
      Top             =   3000
      Width           =   3855
   End
   Begin VB.TextBox Valor3 
      Height          =   285
      Left            =   11160
      MaxLength       =   50
      TabIndex        =   24
      Text            =   " "
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox valor2 
      Height          =   285
      Left            =   11160
      MaxLength       =   50
      TabIndex        =   23
      Text            =   " "
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox Valor1 
      Height          =   285
      Left            =   11160
      MaxLength       =   50
      TabIndex        =   22
      Text            =   " "
      Top             =   1200
      Width           =   3855
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   6960
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
      Height          =   1980
      ItemData        =   "pruedev.frx":0000
      Left            =   960
      List            =   "pruedev.frx":0007
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   11160
      TabIndex        =   5
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   300
      Left            =   11160
      TabIndex        =   4
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   14160
      TabIndex        =   3
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdAddlote 
      Caption         =   "Graba   Prueba"
      Height          =   540
      Left            =   12360
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   2
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Std1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   88
      Top             =   1080
      Width           =   5355
   End
   Begin VB.Label Std2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   87
      Top             =   1680
      Width           =   5355
   End
   Begin VB.Label Std3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   86
      Top             =   2280
      Width           =   5355
   End
   Begin VB.Label Std4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   85
      Top             =   2880
      Width           =   5355
   End
   Begin VB.Label Std5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   84
      Top             =   3480
      Width           =   5355
   End
   Begin VB.Label Std6 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   83
      Top             =   4080
      Width           =   5355
   End
   Begin VB.Label Std7 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   82
      Top             =   4680
      Width           =   5355
   End
   Begin VB.Label Std8 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   81
      Top             =   5280
      Width           =   5355
   End
   Begin VB.Label Std9 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   80
      Top             =   5880
      Width           =   5355
   End
   Begin VB.Label Std10 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   79
      Top             =   6480
      Width           =   5355
   End
   Begin VB.Label Std11 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   78
      Top             =   1320
      Width           =   5355
   End
   Begin VB.Label Std22 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   77
      Top             =   1920
      Width           =   5355
   End
   Begin VB.Label Std33 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   76
      Top             =   2520
      Width           =   5355
   End
   Begin VB.Label Std44 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   75
      Top             =   3120
      Width           =   5355
   End
   Begin VB.Label Std55 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   74
      Top             =   3720
      Width           =   5355
   End
   Begin VB.Label Std66 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   73
      Top             =   4320
      Width           =   5355
   End
   Begin VB.Label Std77 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   72
      Top             =   4920
      Width           =   5355
   End
   Begin VB.Label Std88 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   71
      Top             =   5520
      Width           =   5355
   End
   Begin VB.Label Std99 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   70
      Top             =   6120
      Width           =   5355
   End
   Begin VB.Label Std1010 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   69
      Top             =   6720
      Width           =   5355
   End
   Begin VB.Label Label4 
      Caption         =   "Nro.Ent.Dev."
      Height          =   255
      Left            =   4200
      TabIndex        =   68
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   9480
      TabIndex        =   65
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Lote Original"
      Height          =   255
      Left            =   6720
      TabIndex        =   63
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9720
      TabIndex        =   62
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   6840
      TabIndex        =   60
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Partida"
      Height          =   255
      Left            =   4440
      TabIndex        =   58
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Ensayo10 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   56
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Ensayo9 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   55
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Ensayo8 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   54
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Ensayo7 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   53
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Ensayo6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   52
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Ensayo5 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   51
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Ensayo4 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   50
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Ensayo3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   49
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Ensayo2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   48
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Ensayo1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   47
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor Obtenido"
      Height          =   255
      Left            =   11160
      TabIndex        =   46
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label15 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   2400
      TabIndex        =   44
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Confecciono"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   7920
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Descriprod 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   1320
      TabIndex        =   32
      Top             =   360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Descri10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2040
      TabIndex        =   21
      Top             =   6480
      Width           =   3180
   End
   Begin VB.Label Descri9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   5880
      Width           =   3180
   End
   Begin VB.Label Descri8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   5280
      Width           =   3180
   End
   Begin VB.Label Descri7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   4680
      Width           =   3180
   End
   Begin VB.Label Descri6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2040
      TabIndex        =   17
      Top             =   4080
      Width           =   3180
   End
   Begin VB.Label Descri5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   3480
      Width           =   3180
   End
   Begin VB.Label Descri4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   2880
      Width           =   3180
   End
   Begin VB.Label Descri3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   2280
      Width           =   3180
   End
   Begin VB.Label descri2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   1680
      Width           =   3180
   End
   Begin VB.Label Descri1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   1080
      Width           =   3180
   End
   Begin VB.Label lblresultado 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor Standard"
      Height          =   255
      Left            =   6600
      TabIndex        =   11
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label lblDescri 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label lblensayo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ensayo"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   8
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "PrgPruedev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPrueter As Recordset
Dim spPrueter As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstEntdev As Recordset
Dim spEntdev As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstLiberaTerminado As Recordset
Dim spLiberaTerminado As String
Dim XParam As String
Dim WProducto As String

Dim EmpresaActual As String

Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
Dim MSubject As String
Dim MBody As String
Dim AllPath As String

Dim WDireccionEmail As String
Dim EmailAddress As String
Dim CopiaAddress As String
Dim WNombreEmail As String
Dim MAttach As String

Private Sub CancelaLote_Click()
    panLote.Visible = False
    Producto.SetFocus
End Sub

Private Sub cmdAddlote_Click()

    WPasa = "S"
    
    Producto.Text = UCase(Producto.Text)
    WProducto = "PT" + Mid$(Producto.Text, 3, 10)
    XPro = "NK" + Mid$(Producto.Text, 3, 10)
    
    Rem verifica que el producto sea NK o RE
    
    If Left$(Producto.Text, 2) <> "NK" And Left$(Producto.Text, 2) <> "RE" Then
        m$ = "El codigo de producto debe ser NK o RE"
        A% = MsgBox(m$, 0, "Grabacion de Devolucion de NK o RE")
        WPasa = "N"
    End If
    
    Rem verifica que el producto exista
    
    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        rstTerminado.Close
                    Else
        m$ = "Codigo de Producto inexistente"
        A% = MsgBox(m$, 0, "Grabacion de Devolucion de NK o RE")
        WPasa = "N"
    End If
    
    Rem verifica que la partida asignada sea <> 0
    
    If Val(Partida.Text) = 0 Then
        m$ = "Codigo de Partida invalido. Partida = 0"
        A% = MsgBox(m$, 0, "Grabacion de Devolucion de NK o RE")
        WPasa = "N"
    End If
    
    Rem verifica que la partida NO exista
    
    spHoja = "ListaHoja " + "'" + Partida.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        m$ = "El numero de partida ya existe"
        A% = MsgBox(m$, 0, "Grabacion de Devolucion de NK o RE")
        WPasa = "N"
        rstHoja.Close
    End If
    
    Rem verifica que la cantidad verificada sea <> 0
    
    If Val(Cantidad.Text) = 0 Then
        m$ = "Cantidad incorrecta, Cantidad = 0"
        A% = MsgBox(m$, 0, "Grabacion de Devolucion de NK o RE")
        WPasa = "N"
    End If
    
    Rem verifica que el lote original sea <> 0
    
    If Val(LoteOriginal.Text) = 0 Then
        m$ = "Lote Original incorrecto, Lote = 0"
        A% = MsgBox(m$, 0, "Grabacion de Devolucion de NK o RE")
        WPasa = "N"
    End If
    
    Rem verifica que el lote original exista y que pertenresca a ese producto
    
    WSaldo = 0
    WEntra = "N"
    XParam = "'" + LoteOriginal.Text + "','" _
                    + WProducto + "'"
    spHoja = "ListaHojaProducto " + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        WSaldo = rstHoja!Saldo
        WEntra = "S"
        rstHoja.Close
    End If
    If WEntra = "N" Then
        XParam = "'" + WProducto + "','" _
                    + LoteOriginal.Text + "'"
        spMovguia = "ListaMovguiaLote1 " + XParam
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
            WSaldo = rstMovguia!Saldo
            WEntra = "S"
            rstMovguia.Close
        End If
    End If
                
    If WEntra = "N" Then
        m$ = WProducto + " Producto inexistente o Lote nro. " + LoteOriginal.Text + " inexistente"
        G% = MsgBox(m$, 0, "Grabacion de Devolucion de NK o RE")
        WPasa = "N"
    End If
    
    Rem verifica que el cliente existe
    
    If Cliente.Text <> "" Then
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            rstCliente.Close
                Else
            m$ = "Codigo de CLiente incorrecto"
            A% = MsgBox(m$, 0, "Grabacion de Devolucion de NK o RE")
            WPasa = "N"
        End If
            Else
        m$ = "Np se ha informado codigo de cliente"
        A% = MsgBox(m$, 0, "Grabacion de Devolucion de NK o RE")
        WPasa = "N"
    End If
    
    Rem verifica que exista un ingreso de mercaderia de devolucion
    
    If Cliente.Text <> "" Then
    
        Rem XParam = "'" + Cliente.Text + "','" _
        rem         + LoteOriginal.Text + "','" _
        rem         + "NK" + Mid$(WProducto, 3, 10) + "'"
        Rem spEntdev = "ConsultaEntdev2 " + XParam
        Rem Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstEntdev.RecordCount > 0 Then
        Rem     WSaldo = rstEntdev!Saldo
        Rem     NroDevol.Text = rstEntdev!Codigo
        Rem     rstEntdev.Close
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM EntDev"
        ZSql = ZSql + " Where EntDev.Cliente = " + "'" + Cliente.Text + "'"
        ZSql = ZSql + " and EntDev.Lote = " + "'" + LoteOriginal.Text + "'"
        ZSql = ZSql + " and EntDev.Terminado = " + "'" + "NK" + Mid$(WProducto, 3, 10) + "'"
        ZSql = ZSql + " and EntDev.Saldo <> 0"
        spEntdev = ZSql
        Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
        If rstEntdev.RecordCount > 0 Then
            WSaldo = rstEntdev!Saldo
            NroDevol.Text = rstEntdev!Codigo
            rstEntdev.Close
            If WSaldo < Val(Cantidad.Text) Then
                m$ = "La cantidad informada supera el saldo disponible de PT de este producto. Saldo: " + Str$(WSaldo)
                G% = MsgBox(m$, 0, "Movimientos Varios de Stock")
                WPasa = "N"
            End If
                Else
            m$ = "No se encontro datos de entrada de devolucion que coincidan con los informados"
            G% = MsgBox(m$, 0, "Grabacion de Devolucion de NK o RE")
            WPasa = "N"
        End If
        
            Else
   
        If WSaldo < Val(Cantidad.Text) Then
            m$ = "La cantidad informada supera el saldo disponible de PT de este producto. Saldo: " + Str$(WSaldo)
            G% = MsgBox(m$, 0, "Movimientos Varios de Stock")
            WPasa = "N"
        End If
        
    End If
    
    If WPasa = "S" Then
    
        ZSql = ""
        ZSql = ZSql + "Select PrueTer.Prueba, PrueTer.Lote"
        ZSql = ZSql + " FROM PrueTer"
        ZSql = ZSql + " Where PrueTer.Prueba <= " + "'" + "1199999" + "'"
        ZSql = ZSql + " Order by PrueTer.Prueba"
        spPrueter = ZSql
        Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrueter.RecordCount > 0 Then
            With rstPrueter
                .MoveLast
                Lote = Str$(rstPrueter!Lote + 1)
            End With
            rstPrueter.Close
                Else
            Lote = "1"
        End If
    
    
        Rem spPrueter = "ConsultaPrueterMenor " + "'" + "1999999" + "'"
        Rem Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstPrueter.RecordCount > 0 Then
        Rem     Lote = Str$(rstPrueter!Lote + 1)
        Rem     rstPrueter.Close
        Rem         Else
        Rem     Lote = "1"
        Rem End If
       
        If Val(Partida.Text) <> 0 Then
            Auxi1 = Partida.Text
            Call Ceros(Auxi1, 6)
            Lote = Auxi1
                Else
            Auxi1 = Lote
            Call Ceros(Auxi1, 6)
            Lote = Auxi1
        End If
        
        Auxi = "1"
        
        WPrueba = Auxi + Lote
        WProducto = Producto.Text
        WFecha = fecha.Text
        WValor1 = Valor1.Text
        WValor2 = valor2.Text
        WValor3 = Valor3.Text
        WValor4 = valor4.Text
        WValor5 = valor5.Text
        WValor6 = valor6.Text
        WValor7 = valor7.Text
        WValor8 = valor8.Text
        WValor9 = valor9.Text
        WValor10 = valor10.Text
        WEnsayo = Ensayo.Text
        WAspecto = Aspecto.Text
        WObservaciones = Observaciones.Text
        WConfecciono = Confecciono.Text
        WLiberada = ""
        WLote = Lote
        WRechazo = Lote
        WDate = Date$
        WFechaord = Right$(fecha.Text, 4) + Mid$(fecha.Text, 4, 2) + Left$(fecha.Text, 2)
    
        XParam = "'" + WPrueba + "','" _
                + WProducto + "','" _
                + WFecha + "','" _
                + WValor1 + "','" _
                + WValor2 + "','" _
                + WValor3 + "','" _
                + WValor4 + "','" _
                + WValor5 + "','" _
                + WValor6 + "','" _
                + WValor7 + "','" _
                + WValor8 + "','" _
                + WValor9 + "','" _
                + WValor10 + "','" _
                + WEnsayo + "','" _
                + WAspecto + "','" _
                + WObservaciones + "','" _
                + WConfecciono + "','" _
                + WLiberada + "','" _
                + WLote + "','" _
                + WRechazo + "','" _
                + WFechaord + "','" _
                + WDate + "'"
                
        Set rstPrueter = db.OpenRecordset("AltaPrueter " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        Auxi = "01"
        Auxi1 = Str$(Partida.Text)
        Call Ceros(Auxi1, 6)
        
        WClave = Auxi1 + Auxi
        WHoja = Partida.Text
        WRenglon = "1"
        WFecha = fecha.Text
        WProducto = Producto.Text
        WTeorico = Cantidad.Text
        WReal = Cantidad.Text
        WFechaing = fecha.Text
        WFechaingord = Right$(WFechaing, 4) + Mid$(WFechaing, 4, 2) + Left$(WFechaing, 2)
        WTipo = "T"
        WArticulo = "  -   -   "
        If Cliente.Text = "" Then
            WTerminado = "PT" + Mid$(WProducto, 3, 10)
                Else
            WTerminado = "NK" + Mid$(WProducto, 3, 10)
        End If
        WCantidad = Cantidad.Text
        WLote = ""
        WDate = Date$
        WImporte = ""
        WMarca = ""
        WSaldo = Cantidad.Text
        WLote1 = LoteOriginal.Text
        WLote2 = "0"
        WLote3 = "0"
        WCanti1 = Cantidad.Text
        WCanti2 = "0"
        WCanti3 = "0"
        WCosto1 = "0"
        WCosto2 = "0"
        WCosto3 = "0"
        XParam = "'" + WClave + "','" _
                    + WHoja + "','" _
                    + WRenglon + "','" _
                    + WFecha + "','" _
                    + WProducto + "','" _
                    + WCantidad + "','" _
                    + WTipo + "','" _
                    + WLote + "','" _
                    + WArticulo + "','" _
                    + WTerminado + "','" _
                    + WTeorico + "','" _
                    + WReal + "','" _
                    + WFechaing + "','" _
                    + WFechaingord + "','" _
                    + WDate + "','" _
                    + WImporte + "','" _
                    + WMarca + "','" _
                    + WSaldo + "','" _
                    + WLote1 + "','" + WCanti1 + "','" _
                    + WLote2 + "','" + WCanti2 + "','" _
                    + WLote3 + "','" + WLote3 + "','" _
                    + WCosto1 + "','" _
                    + WCosto2 + "','" _
                    + WCosto3 + "'"
                                           
        spHoja = "AltaHoja " + XParam
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        XParam = "'" + WHoja + "','" _
                 + WFechaord + "'"
        Set rstHoja = db.OpenRecordset("ModificaHojaFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WCodigo = Producto.Text
            WEntradas = Str$(rstTerminado!Entradas + Val(Cantidad.Text))
            WLinea = rstTerminado!Linea
            WDate = Date$
            rstTerminado.Close
                        
            XParam = "'" + WCodigo + "','" _
                         + WEntradas + "','" _
                         + WDate + "'"
            
            spTerminado = "ModificaTerminadoEntradas " + XParam
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            If Cliente.Text <> "" Then
            
                WClaveEntDev = ""
                Sql1 = "Select *"
                Sql2 = " FROM EntDev"
                Sql3 = " Where EntDev.Terminado = " + "'" + XPro + "'"
                Sql4 = " and EntDev.Lote = " + "'" + LoteOriginal.Text + "'"
                spEntdev = Sql1 + Sql2 + Sql3 + Sql4
                Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                If rstEntdev.RecordCount > 0 Then
                    WClaveEntDev = rstEntdev!Clave
                    rstEntdev.Close
                End If
                
                If WClaveEntDev <> "" Then
                    ZSql = ""
                    ZSql = ZSql + "UPDATE EntDev SET "
                    ZSql = ZSql + "Laboratorio = Laboratorio + " + "'" + Cantidad.Text + "',"
                    ZSql = ZSql + "Saldo = Saldo - " + "'" + Cantidad.Text + "'"
                    ZSql = ZSql + " Where Clave = " + "'" + WClaveEntDev + "'"
                    spEntdev = ZSql
                    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
                    Else
                    
                ZTerminado = "PT" + Mid$(WProducto, 3, 10)
                    
                XParam = "'" + LoteOriginal.Text + "','" _
                             + ZTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WClave = rstHoja!Clave
                    WSaldo = Str$(rstHoja!Saldo - Val(Cantidad.Text))
                    WDate = Date$
                    rstHoja.Close
                            
                    XParam = "'" + WClave + "','" _
                                 + WDate + "','" _
                                 + WSaldo + "'"
                    spHoja = "ModificaHojaSaldo " + XParam
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            
                        Else
                                                       
                    XParam = "'" + ZTerminado + "','" _
                                 + LoteOriginal.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WClave = rstMovguia!Clave
                        WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad.Text))
                        WDate = Date$
                        rstMovguia.Close
                            
                        XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                        spMovguia = "ModificaMovguiaSaldo " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                            
                End If
                
            End If
            
            Rem XParam = "'" + Cliente.Text + "','" _
            rem              + LoteOriginal.Text + "','" _
            rem              + "NK" + Mid$(WProducto, 3, 10) + "'"
            Rem spEntdev = "ConsultaEntdev2 " + XParam
            Rem Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstEntdev.RecordCount > 0 Then
            Rem     XLaboratorio = Str$(rstEntdev!Laboratorio + Val(Cantidad.Text))
            Rem     XSaldo = Str$(rstEntdev!Cantidad - Val(XLaboratorio))
            Rem     rstEntdev.Close
            Rem     XParam = "'" + Cliente.Text + "','" _
            rem         + LoteOriginal.Text + "','" _
            rem         + "NK" + Mid$(WProducto, 3, 10) + "','" _
            rem         + XSaldo + "','" _
            rem         + XLaboratorio + "'"
            Rem     spEntdev = "ModificaEntdev2 " + XParam
            Rem     Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
            Rem End If
                
        End If
        
        If Cliente.Text = "" Then
            ZTerminado = "PT" + Mid$(WProducto, 3, 10)
                Else
            ZTerminado = "NK" + Mid$(WProducto, 3, 10)
        End If
        
        spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WCodigo = XPro
            WEntradas = Str$(rstTerminado!Entradas)
            WSalidas = Str$(rstTerminado!Salidas + Val(Cantidad.Text))
            rstTerminado.Close
            XParam = "'" + ZTerminado + "','" _
                    + WEntradas + "','" _
                    + WSalidas + "','" _
                    + WDate + "'"
            spTerminado = "ModificaTerminadoMovimientos " + XParam
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
        
        Sql1 = "Select Max(Codigo) as [CodigoMayor]"
        Sql2 = " FROM LiberaTerminado"
        spLiberaTerminado = Sql1 + Sql2
        Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstLiberaTerminado.RecordCount > 0 Then
            rstLiberaTerminado.MoveLast
            WCodigoMayor = IIf(IsNull(rstLiberaTerminado!CodigoMayor), "0", rstLiberaTerminado!CodigoMayor)
            Lote = Str$(WCodigoMayor)
            rstLiberaTerminado.Close
                Else
            Lote = "0"
        End If
        
        WCodigo = Str$(Val(Lote) + 1)
        WProducto = "PT" + Mid$(Producto.Text, 3, 10)
        WFecha = fecha.Text
        WFechaord = Right$(fecha.Text, 4) + Mid$(fecha.Text, 4, 2) + Left$(fecha.Text, 2)
        WPartida = LoteOriginal.Text
        WPartiOri = Partida.Text
        WValor1 = Valor1.Text
        WValor2 = valor2.Text
        WValor3 = Valor3.Text
        WValor4 = valor4.Text
        WValor5 = valor5.Text
        WValor6 = valor6.Text
        WValor7 = valor7.Text
        WValor8 = valor8.Text
        WValor9 = valor9.Text
        WValor10 = valor10.Text
        WEnsayo = Ensayo.Text
        WAspecto = Aspecto.Text
        WObservaciones = Observaciones.Text
        WConfecciono = Confecciono.Text
        WMarca = "N"
        WCliente = Cliente.Text
        WCantidad = Cantidad.Text
        WFacturado = "0"
        WOrigen = "R"
        If Left$(Producto.Text, 2) = "NK" Then
            WObserva = "Estado del Producto NK"
            WTipo = "NK"
                Else
            WObserva = "Estado del Producto RE"
            WTipo = "RE"
        End If
        
        WImpreProdI = "N"
        WImpreProdII = "N"
        WImpreProdIII = "N"
        WImpreVentas = "N"
        WTipopro = ""
            
        XTipoPro = ""
        XCodigo = Val(Mid$(Producto.Text, 4, 5))
        If Left$(Producto.Text, 2) = "DY" Or Left$(Producto.Text, 2) = "DW" Or Left$(Producto.Text, 2) = "DS" Then
            XTipoPro = "CO"
                Else
            If XCodigo >= 0 And XCodigo <= 999 Then
                XTipoPro = "CO"
                    Else
                If XCodigo >= 11000 And XCodigo <= 11999 Then
                    XTipoPro = "CO"
                        Else
                    If XCodigo >= 25000 And XCodigo <= 25999 Then
                        XTipoPro = "FA"
                            Else
                        If XCodigo >= 2300 And XCodigo <= 2399 Then
                            XTipoPro = "BI"
                                Else
                            XTipoPro = "PT"
                        End If
                    End If
                End If
            End If
        End If
                
        ZLinea = 0
        spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            ZLinea = rstTerminado!Linea
            rstTerminado.Close
        End If
                
        Select Case ZLinea
            Case 8
                XTipoPro = "PG"
            Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                XTipoPro = "FA"
            Case Else
        End Select
                
        WTipopro = XTipoPro
            
        ZSql = ""
        ZSql = ZSql & "INSERT INTO LiberaTerminado ("
        ZSql = ZSql & "Codigo, "
        ZSql = ZSql & "Producto, "
        ZSql = ZSql & "Fecha, "
        ZSql = ZSql & "OrdFecha, "
        ZSql = ZSql & "Partida, "
        ZSql = ZSql & "PartiOri, "
        ZSql = ZSql & "PedidoDevol, "
        ZSql = ZSql & "Valor1, "
        ZSql = ZSql & "Valor2, "
        ZSql = ZSql & "Valor3, "
        ZSql = ZSql & "Valor4, "
        ZSql = ZSql & "Valor5, "
        ZSql = ZSql & "Valor6, "
        ZSql = ZSql & "Valor7, "
        ZSql = ZSql & "Valor8, "
        ZSql = ZSql & "Valor9, "
        ZSql = ZSql & "Valor10, "
        ZSql = ZSql & "Ensayo, "
        ZSql = ZSql & "Aspecto, "
        ZSql = ZSql & "Observaciones, "
        ZSql = ZSql & "Confecciono, "
        ZSql = ZSql & "Marca, "
        ZSql = ZSql & "Cliente, "
        ZSql = ZSql & "Cantidad, "
        ZSql = ZSql & "Facturado, "
        ZSql = ZSql & "Observa, "
        ZSql = ZSql & "Origen, "
        ZSql = ZSql & "Tipo, "
        ZSql = ZSql & "ImpreProdI, "
        ZSql = ZSql & "ImpreProdII, "
        ZSql = ZSql & "ImpreProdIII, "
        ZSql = ZSql & "ImpreVentas, "
        ZSql = ZSql & "TipoPro) "
        ZSql = ZSql & "Values ("
        ZSql = ZSql & "'" + WCodigo + "',"
        ZSql = ZSql & "'" + WProducto + "',"
        ZSql = ZSql & "'" + WFecha + "',"
        ZSql = ZSql & "'" + WOrdFecha + "',"
        ZSql = ZSql & "'" + WPartida + "',"
        ZSql = ZSql & "'" + WPartiOri + "',"
        ZSql = ZSql & "'" + NroDevol.Text + "',"
        ZSql = ZSql & "'" + WValor1 + "',"
        ZSql = ZSql & "'" + WValor2 + "',"
        ZSql = ZSql & "'" + WValor3 + "',"
        ZSql = ZSql & "'" + WValor4 + "',"
        ZSql = ZSql & "'" + WValor5 + "',"
        ZSql = ZSql & "'" + WValor6 + "',"
        ZSql = ZSql & "'" + WValor7 + "',"
        ZSql = ZSql & "'" + WValor8 + "',"
        ZSql = ZSql & "'" + WValor9 + "',"
        ZSql = ZSql & "'" + WValor10 + "',"
        ZSql = ZSql & "'" + WEnsayo + "',"
        ZSql = ZSql & "'" + WAspecto + "',"
        ZSql = ZSql & "'" + WObservaciones + "',"
        ZSql = ZSql & "'" + WConfecciono + "',"
        ZSql = ZSql & "'" + WMarca + "',"
        ZSql = ZSql & "'" + WCliente + "',"
        ZSql = ZSql & "'" + WCantidad + "',"
        ZSql = ZSql & "'" + WFacturado + "',"
        ZSql = ZSql & "'" + WObserva + "',"
        ZSql = ZSql & "'" + WOrigen + "',"
        ZSql = ZSql & "'" + WTipo + "',"
        ZSql = ZSql & "'" + WImpreProdI + "',"
        ZSql = ZSql & "'" + WImpreProdII + "',"
        ZSql = ZSql & "'" + WImpreProdIII + "',"
        ZSql = ZSql & "'" + WImpreVentas + "',"
        ZSql = ZSql & "'" + WTipopro + "')"
          
        spLiberaTerminado = ZSql
        Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
        WEntra = "N"
        XProducto = "PT" + Mid$(Producto.Text, 3, 10)
        XParam = "'" + LoteOriginal.Text + "','" _
                    + XProducto + "'"
        spHoja = "ListaHojaProducto " + XParam
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            WSaldo = rstHoja!Saldo
            WEntra = "S"
            rstHoja.Close
        End If
        If WEntra = "N" Then
            XParam = "'" + XProducto + "','" _
                        + LoteOriginal.Text + "'"
            spMovguia = "ListaMovguiaLote1 " + XParam
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                WSaldo = rstMovguia!Saldo
                WEntra = "S"
                rstMovguia.Close
            End If
        End If
    
        If WSaldo > 0 Then
            
            T$ = "Liberacion de Partidas de Producto Terminado"
            m$ = "Desea Liberar el resto de la partida. Saldo:" + Str$(WSaldo)
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
        
                WEntra = "N"
                XParam = "'" + LoteOriginal.Text + "','" _
                            + WProducto + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    rstHoja.Close
                    WEntra = "S"
                    WMarcaEstado = ""
                    Sql1 = "UPDATE Hoja SET "
                    Sql2 = "Estado  = " + "'" + WMarcaEstado + "'"
                    Sql3 = " Where Hoja.Producto = " + "'" + WProducto + "'"
                    Sql4 = " and Hoja.Hoja = " + "'" + LoteOriginal.Text + "'"
                    spHoja = Sql1 + Sql2 + Sql3 + Sql4
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
                    
                If WEntra = "N" Then
                    XParam = "'" + WProducto + "','" _
                                + LoteOriginal.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        rstMovguia.Close
                        WMarcaEstado = ""
                        Sql1 = "UPDATE Guia SET "
                        Sql2 = "Estado  = " + "'" + WMarcaEstado + "'"
                        Sql3 = " Where Guia.Terminado = " + "'" + WProducto + "'"
                        Sql4 = " and Guia.Lote = " + "'" + LoteOriginal.Text + "'"
                        spMovguia = Sql1 + Sql2 + Sql3 + Sql4
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                End If
                
                Sql1 = "Select Max(Codigo) as [CodigoMayor]"
                Sql2 = " FROM LiberaTerminado"
                spLiberaTerminado = Sql1 + Sql2
                Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstLiberaTerminado.RecordCount > 0 Then
                    rstLiberaTerminado.MoveLast
                    WCodigoMayor = IIf(IsNull(rstLiberaTerminado!CodigoMayor), "0", rstLiberaTerminado!CodigoMayor)
                    Lote = Str$(WCodigoMayor)
                    rstLiberaTerminado.Close
                        Else
                    Lote = "0"
                End If
        
                WCodigo = Str$(Val(Lote) + 1)
                WProducto = "PT" + Mid$(Producto.Text, 3, 10)
                WFecha = fecha.Text
                WFechaord = Right$(fecha.Text, 4) + Mid$(fecha.Text, 4, 2) + Left$(fecha.Text, 2)
                WPartida = LoteOriginal.Text
                WPartiOri = ""
                WValor1 = Valor1.Text
                WValor2 = valor2.Text
                WValor3 = Valor3.Text
                WValor4 = valor4.Text
                WValor5 = valor5.Text
                WValor6 = valor6.Text
                WValor7 = valor7.Text
                WValor8 = valor8.Text
                WValor9 = valor9.Text
                WValor10 = valor10.Text
                WEnsayo = Ensayo.Text
                WAspecto = Aspecto.Text
                WObservaciones = Observaciones.Text
                WConfecciono = Confecciono.Text
                WMarca = "N"
                WCliente = ""
                WCantidad = Str$(WSaldo)
                WObserva = "Liberacion de Producto"
                WOrigen = "R"
                WTipo = "PT"
                
                WImpreProdI = "N"
                WImpreProdII = "N"
                WImpreProdIII = "N"
                WImpreVentas = "N"
                WTipopro = ""
            
                XTipoPro = ""
                XCodigo = Val(Mid$(Producto.Text, 4, 5))
                If Left$(Producto.Text, 2) = "DY" Or Left$(Producto.Text, 2) = "DW" Or Left$(Producto.Text, 2) = "DS" Then
                    XTipoPro = "CO"
                        Else
                    If XCodigo >= 0 And XCodigo <= 999 Then
                        XTipoPro = "CO"
                            Else
                        If XCodigo >= 11000 And XCodigo <= 11999 Then
                            XTipoPro = "CO"
                                Else
                            If XCodigo >= 25000 And XCodigo <= 25999 Then
                                XTipoPro = "FA"
                                    Else
                                If XCodigo >= 2300 And XCodigo <= 2399 Then
                                    XTipoPro = "BI"
                                        Else
                                    XTipoPro = "PT"
                                End If
                            End If
                        End If
                    End If
                End If
                
                ZLinea = 0
                spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZLinea = rstTerminado!Linea
                    rstTerminado.Close
                End If
                
                Select Case ZLinea
                    Case 8
                        XTipoPro = "PG"
                    Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                        XTipoPro = "FA"
                    Case Else
                End Select
                
                WTipopro = XTipoPro
                
                Select Case WTipopro
                    Case "CO", "PG"
                        WImpreProdI = "S"
                    Case "BI", "PT"
                        WImpreProdII = "S"
                    Case "FA"
                        WImpreProdIII = "S"
                    Case Else
                End Select
                
                ZSql = ""
                ZSql = ZSql & "INSERT INTO LiberaTerminado ("
                ZSql = ZSql & "Codigo, "
                ZSql = ZSql & "Producto, "
                ZSql = ZSql & "Fecha, "
                ZSql = ZSql & "OrdFecha, "
                ZSql = ZSql & "Partida, "
                ZSql = ZSql & "PartiOri, "
                ZSql = ZSql & "PedidoDevol, "
                ZSql = ZSql & "Valor1, "
                ZSql = ZSql & "Valor2, "
                ZSql = ZSql & "Valor3, "
                ZSql = ZSql & "Valor4, "
                ZSql = ZSql & "Valor5, "
                ZSql = ZSql & "Valor6, "
                ZSql = ZSql & "Valor7, "
                ZSql = ZSql & "Valor8, "
                ZSql = ZSql & "Valor9, "
                ZSql = ZSql & "Valor10, "
                ZSql = ZSql & "Ensayo, "
                ZSql = ZSql & "Aspecto, "
                ZSql = ZSql & "Observaciones, "
                ZSql = ZSql & "Confecciono, "
                ZSql = ZSql & "Marca, "
                ZSql = ZSql & "Cliente, "
                ZSql = ZSql & "Cantidad, "
                ZSql = ZSql & "Facturado, "
                ZSql = ZSql & "Observa, "
                ZSql = ZSql & "Origen, "
                ZSql = ZSql & "Tipo, "
                ZSql = ZSql & "ImpreProdI, "
                ZSql = ZSql & "ImpreProdII, "
                ZSql = ZSql & "ImpreProdIII, "
                ZSql = ZSql & "ImpreVentas, "
                ZSql = ZSql & "TipoPro) "
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + WCodigo + "',"
                ZSql = ZSql & "'" + WProducto + "',"
                ZSql = ZSql & "'" + WFecha + "',"
                ZSql = ZSql & "'" + WOrdFecha + "',"
                ZSql = ZSql & "'" + WPartida + "',"
                ZSql = ZSql & "'" + WPartiOri + "',"
                ZSql = ZSql & "'" + NroDevol.Text + "',"
                ZSql = ZSql & "'" + WValor1 + "',"
                ZSql = ZSql & "'" + WValor2 + "',"
                ZSql = ZSql & "'" + WValor3 + "',"
                ZSql = ZSql & "'" + WValor4 + "',"
                ZSql = ZSql & "'" + WValor5 + "',"
                ZSql = ZSql & "'" + WValor6 + "',"
                ZSql = ZSql & "'" + WValor7 + "',"
                ZSql = ZSql & "'" + WValor8 + "',"
                ZSql = ZSql & "'" + WValor9 + "',"
                ZSql = ZSql & "'" + WValor10 + "',"
                ZSql = ZSql & "'" + WEnsayo + "',"
                ZSql = ZSql & "'" + WAspecto + "',"
                ZSql = ZSql & "'" + WObservaciones + "',"
                ZSql = ZSql & "'" + WConfecciono + "',"
                ZSql = ZSql & "'" + WMarca + "',"
                ZSql = ZSql & "'" + WCliente + "',"
                ZSql = ZSql & "'" + WCantidad + "',"
                ZSql = ZSql & "'" + WFacturado + "',"
                ZSql = ZSql & "'" + WObserva + "',"
                ZSql = ZSql & "'" + WOrigen + "',"
                ZSql = ZSql & "'" + WTipo + "',"
                ZSql = ZSql & "'" + WImpreProdI + "',"
                ZSql = ZSql & "'" + WImpreProdII + "',"
                ZSql = ZSql & "'" + WImpreProdIII + "',"
                ZSql = ZSql & "'" + WImpreVentas + "',"
                ZSql = ZSql & "'" + WTipopro + "')"
            
                spLiberaTerminado = ZSql
                Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        End If
        
        
        Rem
        Rem
        Rem envio email de aviso
        Rem
        Rem
        
        sTo = "lsantos@surfactan.com.ar; dsuarez@surfactan.com.ar; ebiglieri@surfactan.com.ar"
        sCC = ""
        sBCC = ""
        sSubject = "DEVOLUCION DE NK O RE " + Terminado.Text
    
        sBody = "Se rechazo la partida " + LoteOriginal.Text + " - " + _
                "del " + Producto.Text + "  " + Descriprod.Caption + " - " + _
                "del Cliente " + Cliente.Text + _
                "la cantidad  de " + Cantidad.Text + " Kgs."
            
        ret = Shell("Start.exe " _
                & "mailto:" & """" & sTo & """" _
                & "?Subject=" & """" & sSubject & """" _
                & "&cc=" & """" & sCC & """" _
                & "&bcc=" & """" & sBCC & """" _
                & "&Body=" & """" & sBody & """" _
                & "&File=" & """" & "c:\autoexec.bat" & """" _
                , 0)
                
        Rem sTo = "hgutierrez@pellital.com.ar"
        Rem sCC = ""
        Rem sBCC = ""
        Rem sSubject = "CAMBIO DE FORMULA DEL " + Terminado.Text
        Rem sBody = "Fecha:" + XFecha + " - " + _
        rem         "Costo Anterior : " + ImpreCosto1 + " - " + _
        rem         "Costo Actual : " + ImpreCosto2
        Rem SFile = ""

        Rem EmailAddress = sTo
        Rem CopiaAddress = sCC
        Rem MSubject = sSubject
        Rem MBody = sBody
        Rem MAttach = ""
        Rem MAttachI = ""
        Rem MAttachII = ""
        Rem MAttachIII = ""
        Rem MAttachIV = ""
        Rem MAttachVI = ""
        Rem MAttachVII = ""
        Rem MAttachVIII = ""
        
        Rem SendEmail
        
        
        Call CmdLimpiar_Click
        Producto.SetFocus
        
    End If
        
End Sub

Private Sub CmdLimpiar_Click()
    Opcion.Visible = False
    pantalla.Visible = False
   
    Ensayo.Visible = True
    Aspecto.Visible = True
    Observaciones.Visible = True
    Confecciono.Visible = True
    
    
    
    Producto.Text = "  -     -   "
    fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    LoteOriginal.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Cantidad.Text = ""
    Ensayo1.Caption = ""
    Valor1.Text = ""
    Ensayo2.Caption = ""
    valor2.Text = ""
    Ensayo3.Caption = ""
    Valor3.Text = ""
    Ensayo4.Caption = ""
    valor4.Text = ""
    Ensayo5.Caption = ""
    valor5.Text = ""
    Ensayo6.Caption = ""
    valor6.Text = ""
    Ensayo7.Caption = ""
    valor7.Text = ""
    Ensayo8.Caption = ""
    valor8.Text = ""
    Ensayo9.Caption = ""
    valor9.Text = ""
    Ensayo10.Caption = ""
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
    Ensayo.Text = ""
    Aspecto.Text = ""
    Observaciones.Text = ""
    Confecciono.Text = ""
    Std1.Caption = ""
    Std2.Caption = ""
    Std3.Caption = ""
    Std4.Caption = ""
    Std5.Caption = ""
    Std6.Caption = ""
    Std7.Caption = ""
    Std8.Caption = ""
    Std9.Caption = ""
    Std10.Caption = ""
    Partida.Text = ""
    NroDevol.Text = ""
    Producto.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgPruedev.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(fecha.Text, Auxi)
        If Auxi = "S" Then
        
            spHoja = "ListaHojaNumero"
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                With rstHoja
                    .MoveLast
                    Partida.Text = rstHoja!Hoja + 1
                End With
                rstHoja.Close
            End If
            
            If Val(Wempresa) = 11 And Val(Partida.Text) = 0 Then
                Partida.Text = "395000"
            End If
        
            Partida.SetFocus
            
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Form_Activate()
    Select Case Val(EmpresaActual)
        Case 1
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            Wempresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            Wempresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            Wempresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            Wempresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            Wempresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            Wempresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 11
            Wempresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    OPEN_FILE_Empresa
End Sub

Private Sub Impensayo_Click()

    If Val(Auxi) = 0 Then
        Auxi = "0"
    End If
    
    If Val(Lote) = 0 Then
        Lote = "000000"
    End If

    Rem lista.ReportFileName = "Ensayoter.rpt"
    Rem lista.GroupSelectionFormula = "{Prueter.Prueba} in " + Chr$(34) + Auxi + Lote + Chr$(34) + " to " + Chr$(34) + Auxi + Lote + Chr$(34)
    Rem lista.Destination = 1
    Rem lista.Action = 1
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    Printer.Font = "Times New Roman"
    Printer.FontSize = "12"
    Printer.Print Tab(1); ""
    Printer.FontSize = "10"
    
    Printer.Print Tab(1); "Empresa : " + WAuxiliar
    Printer.Print Tab(1); ""
    Printer.Print Tab(20); "ENSAYO DE PRODUCTO TERMINADO"
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "Prueba"; Tab(15); Lote
    Printer.Print Tab(1); "Producto"; Tab(15); Producto.Text; Tab(40); Descriprod.Caption
    Printer.Print Tab(1); "Fecha"; Tab(15); fecha.Text
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo1.Caption; Tab(25); Descri1.Caption; Tab(80); Std1.Caption; Tab(105); Valor1.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo2.Caption; Tab(25); descri2.Caption; Tab(80); Std2.Caption; Tab(105); valor2.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo3.Caption; Tab(25); Descri3.Caption; Tab(80); Std3.Caption; Tab(105); Valor3.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo4.Caption; Tab(25); Descri4.Caption; Tab(80); Std4.Caption; Tab(105); valor4.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo5.Caption; Tab(25); Descri5.Caption; Tab(80); Std5.Caption; Tab(105); valor5.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo6.Caption; Tab(25); Descri6.Caption; Tab(80); Std6.Caption; Tab(105); valor6.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo7.Caption; Tab(25); Descri7.Caption; Tab(80); Std7.Caption; Tab(105); valor7.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo8.Caption; Tab(25); Descri8.Caption; Tab(80); Std8.Caption; Tab(105); valor8.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo9.Caption; Tab(25); Descri9.Caption; Tab(80); Std9.Caption; Tab(105); valor9.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo10.Caption; Tab(25); Descri10.Caption; Tab(80); Std10.Caption; Tab(105); valor10.Text
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "Observaciones"; Tab(20); Ensayo.Text
    Printer.Print Tab(1); "Observaciones"; Tab(20); Aspecto.Text
    Printer.Print Tab(1); "Observaciones"; Tab(20); Observaciones.Text
    Printer.Print Tab(1); "Confecciono"; Tab(20); Confecciono.Text
    Printer.Print Tab(1); ""
    
    Printer.EndDoc

End Sub



Private Sub Label12_Click()

End Sub

Private Sub Valor1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor2.SetFocus
    End If
End Sub
Private Sub Valor2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor3.SetFocus
    End If
End Sub
Private Sub Valor3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor4.SetFocus
    End If
End Sub
Private Sub Valor4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor5.SetFocus
    End If
End Sub
Private Sub Valor5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor6.SetFocus
    End If
End Sub
Private Sub Valor6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor7.SetFocus
    End If
End Sub
Private Sub Valor7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor8.SetFocus
    End If
End Sub
Private Sub Valor8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor9.SetFocus
    End If
End Sub
Private Sub Valor9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor10.SetFocus
    End If
End Sub
Private Sub Valor10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo.SetFocus
    End If
End Sub
Private Sub Ensayo_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Aspecto.SetFocus
    End If
End Sub
Private Sub Aspecto_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
End Sub
Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Confecciono.SetFocus
    End If
End Sub
Private Sub Confecciono_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor1.SetFocus
    End If
End Sub

Private Sub imprime_Click()

    XEmpresa = Wempresa
    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select

    Producto.Text = UCase(Producto.Text)
    WProducto = "PT" + Mid$(Producto.Text, 3, 10)
    
    
    Sql1 = "Select *"
    Sql2 = " FROM EspecifUnifica"
    Sql3 = " Where EspecifUnifica.Producto = " + "'" + WProducto + "'"
    spEspecifUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
    
        Ensayo1.Caption = rstEspecifUnifica!Ensayo1
        Ensayo2.Caption = rstEspecifUnifica!Ensayo2
        Ensayo3.Caption = rstEspecifUnifica!Ensayo3
        Ensayo4.Caption = rstEspecifUnifica!Ensayo4
        Ensayo5.Caption = rstEspecifUnifica!Ensayo5
        Ensayo6.Caption = rstEspecifUnifica!Ensayo6
        Ensayo7.Caption = rstEspecifUnifica!Ensayo7
        Ensayo8.Caption = rstEspecifUnifica!Ensayo8
        Ensayo9.Caption = rstEspecifUnifica!Ensayo9
        Ensayo10.Caption = rstEspecifUnifica!Ensayo10
        
        Std1.Caption = rstEspecifUnifica!Valor1
        Std2.Caption = rstEspecifUnifica!valor2
        Std3.Caption = rstEspecifUnifica!Valor3
        Std4.Caption = rstEspecifUnifica!valor4
        Std5.Caption = rstEspecifUnifica!valor5
        Std6.Caption = rstEspecifUnifica!valor6
        Std7.Caption = rstEspecifUnifica!valor7
        Std8.Caption = rstEspecifUnifica!valor8
        Std9.Caption = rstEspecifUnifica!valor9
        Std10.Caption = rstEspecifUnifica!valor10
        
        Rem by nan
        Std11.Caption = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
            Std22.Caption = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
            Std33.Caption = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
            Std44.Caption = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
            Std55.Caption = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
            Std66.Caption = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
            Std77.Caption = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
            Std88.Caption = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
            Std99.Caption = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
            Std1010.Caption = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
        
        
        Rem
        
        
        
        
        
        
        
        
        
        rstEspecifUnifica.Close
    End If
    
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri1.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
            Descri1.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        descri2.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        descri2.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri3.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri3.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri4.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri4.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri5.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri5.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri6.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri6.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri7.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri7.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri8.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri8.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri9.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri9.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri10.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri10.Caption = ""
    End If
    
    Call Conecta_Empresa

End Sub

Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Producto.Text = UCase(Producto.Text)
        If Producto.Text <> "" Then
            If Left$(Producto.Text, 2) <> "NK" And Left$(Producto.Text, 2) <> "RE" Then
                m$ = "El codigo de producto debe ser NK o RE"
                A% = MsgBox(m$, 0, "Grabacion de Devolucion de NK o RE")
                Producto.SetFocus
                    Else
                    
                XEmpresa = Wempresa
                Select Case Val(Wempresa)
                    Case 1, 3, 5, 6, 7, 10, 11
                        Wempresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        Wempresa = "0004"
                        txtOdbc = "Empresa04"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
                    
                    
                Producto.Text = UCase(Producto.Text)
                XProducto = Producto.Text
                WProducto = "PT" + Mid$(Producto.Text, 3, 10)
                
                Sql1 = "Select *"
                Sql2 = " FROM  EspecifUnifica"
                Sql3 = " Where EspecifUnifica.Producto = " + "'" + WProducto + "'"
                spEspecifUnifica = Sql1 + Sql2 + Sql3
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEspecifUnifica.RecordCount > 0 Then
                    rstEspecifUnifica.Close
                    Call Conecta_Empresa
                    Call imprime_Click
                        Else
                    Call Conecta_Empresa
                    CmdLimpiar_Click
                    Producto.Text = XProducto
                End If
                
                spTerminado = "ConsultaTerminado " + "'" + WProducto + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    Rem Descriprod.Caption = rstTerminado!Descripcion
                    rstTerminado.Close
                        Else
                    Producto.SetFocus
                    Exit Sub
                End If
                fecha.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Consulta_Click()
    Ensayo.Visible = False
    Aspecto.Visible = False
    Observaciones.Visible = False
    Confecciono.Visible = False
    Opcion.Clear
    
    Opcion.AddItem "Productos"
    Rem Opcion.AddItem "Pruebas"
    
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
            spTerminado = "ListaTerminadoConsulta"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
            
            With rstTerminado
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Left$(rstTerminado!Codigo, 2) = "NK" Or Left$(rstTerminado!Codigo, 2) = "RE" Then
                            Rem IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                            IngresaItem = rstTerminado!Codigo
                            pantalla.AddItem IngresaItem
                            IngresaItem = rstTerminado!Codigo
                            WIndice.AddItem IngresaItem
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstTerminado.Close
            
            End If
        
        Case 1
            spPrueter = "ListaPrueterConsulta"
            Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueter.RecordCount > 0 Then
            
            With rstPrueter
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = "Tipo:" + Left$(rstPrueter!Prueba, 1) + " Prueba:" + Str$(rstPrueter!Lote) + " Producto:" + rstPrueter!Producto + " Fecha : " + rstPrueter!fecha
                        pantalla.AddItem IngresaItem
                        IngresaItem = rstPrueter!Prueba
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPrueter.Close
            
            End If
        
        Case Else
    End Select
            
    pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = pantalla.ListIndex
            Clavepro$ = WIndice.List(Indice)
            Clavepro$ = UCase(Clavepro$)
            WProducto = "PT" + Mid$(Clavepro$, 3, 10)
            spTerminado = "ConsultaTerminado " + "'" + WProducto + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                Producto.Text = Clavepro$
                Rem Descriprod.Caption = rstTerminado!Descripcion
                rstTerminado.Close
                Call imprime_Click
                    Else
                CmdLimpiar_Click
                Producto.Text = "  -     -   "
                Descriprod.Caption = ""
            End If
            Producto.SetFocus
            
        Case 1
            Indice = pantalla.ListIndex
            ClavePrue$ = WIndice.List(Indice)
            spPrueter = "ConsultaPrueter " + "'" + ClavePrue$ + "'"
            Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueter.RecordCount > 0 Then
                Partida.Text = rstPrueter!Lote
                Producto.Text = rstPrueter!Producto
                fecha.Text = rstPrueter!fecha
                Valor1.Text = rstPrueter!Valor1
                valor2.Text = rstPrueter!valor2
                Valor3.Text = rstPrueter!Valor3
                valor4.Text = rstPrueter!valor4
                valor5.Text = rstPrueter!valor5
                valor6.Text = rstPrueter!valor6
                valor7.Text = rstPrueter!valor7
                valor8.Text = rstPrueter!valor8
                valor9.Text = rstPrueter!valor9
                valor10.Text = rstPrueter!valor10
                Ensayo.Text = rstPrueter!Ensayo
                Aspecto.Text = rstPrueter!Aspecto
                Observaciones.Text = rstPrueter!Observaciones
                Confecciono.Text = rstPrueter!Confecciono
                Auxi = Left$(rstPrueter!Prueba, 1)
                Lote = rstPrueter!Lote
                
                rstPrueter.Close
                
                LlamaImprime = "N"
                
                If Left$(Producto.Text, 2) = "DW" Then
                    WProducto = "DW" + Mid$(Producto.Text, 3, 10)
                        Else
                    WProducto = "PT" + Mid$(Producto.Text, 3, 10)
                End If
                
                Sql1 = "Select *"
                Sql2 = " FROM EspecifUnifica"
                Sql3 = " Where EspecifUnifica.Producto = " + "'" + WProducto + "'"
                spEspecifUnifica = Sql1 + Sql2 + Sql3
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEspecifUnifica.RecordCount > 0 Then
                    rstEspecifUnifica.Close
                    LlamaImprime = "S"
                End If
                
                Call Conecta_Empresa
                
                If LlamaImprime = "S" Then
                    Call imprime_Click
                End If
                
                spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    Rem Descriprod.Caption = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
                    
                    Else
                    
                Call CmdLimpiar_Click
                
            End If
            Producto.SetFocus
        
        Case Else
    End Select
    Ensayo.Visible = True
    Aspecto.Visible = True
    Observaciones.Visible = True
    Confecciono.Visible = True
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgPruedev.Caption = "Ingreso de Devoluciuon de NK o RE :  " + !Nombre
        End If
    End With
    EmpresaActual = Wempresa
    fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
End Sub

Private Sub Partida_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spHoja = "ListaHoja " + "'" + Partida.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            m$ = "El numero de partida ya existe"
            A% = MsgBox(m$, 0, "Grabacion de Devolucion de NK o RE")
            Partida.Text = ""
            WPasa = "N"
            rstHoja.Close
                Else
            LoteOriginal.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub LoteOriginal_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WEntra = "N"
        XProducto = "PT" + Mid$(Producto.Text, 3, 10)
        XParam = "'" + LoteOriginal.Text + "','" _
                    + XProducto + "'"
        spHoja = "ListaHojaProducto " + XParam
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            WEntra = "S"
            rstHoja.Close
        End If
        If WEntra = "N" Then
            XParam = "'" + XProducto + "','" _
                        + LoteOriginal.Text + "'"
            spMovguia = "ListaMovguiaLote1 " + XParam
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                WEntra = "S"
                rstMovguia.Close
            End If
        End If
                
        If WEntra = "N" Then
            m$ = Producto.Text + " Producto inexistente o Lote nro. " + LoteOriginal.Text + " inexistente"
            G% = MsgBox(m$, 0, "Movimientos Varios de Stock")
            LoteOriginal.SetFocus
                Else
            Cantidad.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.SetFocus
    End If
End Sub

Private Sub Cliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cliente.Text <> "" Then
            Cliente.Text = UCase(Cliente.Text)
            If Cliente.Text <> "" Then
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    Cliente.Text = rstCliente!Cliente
                    DesCliente.Caption = rstCliente!Razon
                    rstCliente.Close
                    Valor1.SetFocus
                        Else
                    Cliente.Text = Claveven$
                    Cliente.SetFocus
                End If
            End If
                Else
            Valor1.SetFocus
        End If
    End If
End Sub









