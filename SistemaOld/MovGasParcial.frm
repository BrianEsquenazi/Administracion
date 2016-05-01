VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMovgasParcial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Nacionalizacion de Mercaderia"
   ClientHeight    =   8370
   ClientLeft      =   225
   ClientTop       =   480
   ClientWidth     =   11490
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8370
   ScaleWidth      =   11490
   Visible         =   0   'False
   Begin VB.Frame IngresaDerechos 
      Caption         =   "Ingresos de Cantidades a Nacionalizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   960
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   7095
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Index           =   3
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Index           =   2
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox WTexto2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
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
         Left            =   2880
         TabIndex        =   30
         Top             =   1920
         Width           =   375
      End
      Begin VB.ComboBox WCombo1 
         Height          =   315
         Left            =   2040
         TabIndex        =   29
         Top             =   2400
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto1 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2040
         TabIndex        =   28
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton FinIngreso 
         Caption         =   "Fin de Ingreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   23
         Top             =   5160
         Width           =   2175
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   3480
         TabIndex        =   34
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         _Version        =   327680
         BackColor       =   16776960
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
      Begin MSFlexGridLib.MSFlexGrid WVector1 
         Height          =   4575
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8070
         _Version        =   327680
         BackColor       =   16777152
      End
   End
   Begin VB.CommandButton Derechos 
      Caption         =   "Articulos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6480
      TabIndex        =   41
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox Empresa 
      Alignment       =   1  'Right Justify
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
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   40
      Text            =   " "
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Orden 
      Alignment       =   1  'Right Justify
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
      Left            =   6000
      MaxLength       =   6
      TabIndex        =   38
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Carpeta 
      Alignment       =   1  'Right Justify
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
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   36
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame PantaFecha 
      Height          =   2055
      Left            =   4080
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton CancelaFecha 
         Caption         =   "   Cancela Actualizacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   25
         Top             =   1320
         Width           =   1935
      End
      Begin MSMask.MaskEdBox FechaLLegada 
         Height          =   285
         Left            =   600
         TabIndex        =   26
         Top             =   840
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
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Indique Fecha Prevista de Arribo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1080
      TabIndex        =   17
      Top             =   6480
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9000
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impreord.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5400
      TabIndex        =   15
      Top             =   6480
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   7800
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   480
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
   Begin VB.TextBox Codigo 
      Alignment       =   1  'Right Justify
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
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   0
      TabIndex        =   10
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4320
      TabIndex        =   9
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2160
      TabIndex        =   7
      Top             =   6480
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   6975
      Begin VB.TextBox WImporte 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   20
         Text            =   " "
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox WConcepto 
         Alignment       =   1  'Right Justify
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   16
         Text            =   " "
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox WLinea 
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
         Left            =   0
         TabIndex        =   8
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.Label WDescripcion 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   3735
      End
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3240
      TabIndex        =   4
      Top             =   6480
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4215
      Left            =   0
      OleObjectBlob   =   "MovGasParcial.frx":0000
      TabIndex        =   3
      Top             =   960
      Width           =   7215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7080
      ItemData        =   "MovGasParcial.frx":09EA
      Left            =   7800
      List            =   "MovGasParcial.frx":09F1
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label6 
      Caption         =   "Orden"
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
      Left            =   4080
      TabIndex        =   39
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Carpeta"
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
      Left            =   4080
      TabIndex        =   37
      Top             =   120
      Width           =   1095
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
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro. de Movimiento"
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
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgMovgasParcial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 3 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Auxiliar(100, 5) As String
Dim rstMovGasParcial As Recordset
Dim spMovGasParcial As String
Dim rstMovGasParcialArti As Recordset
Dim spMovGasParcialArti As String
Dim rstMovgas As Recordset
Dim spMovgas As String
Dim rstGasimpo As Recordset
Dim spGasimpo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim XParam As String
Dim WProveedor As String
Dim Vector(100) As String
Dim WCalcula(100, 10) As String
Dim EmpresaAnterior As String
Dim EmpresaOrden As Integer
Dim XOrden As Integer
Dim CargaEmpresa(12, 2) As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String

Dim WControl As String

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    WConcepto.Text = ""
    WDescripcion.Caption = ""
    WImporte.Text = ""
    WLinea.Text = ""
    
    WConcepto.SetFocus
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub cmdClose_Click()
    PrgMovgasParcial.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()
     Opcion.Clear
     Opcion.AddItem "Conceptos de Gastos de Importacion"
     Opcion.Visible = True
 End Sub

Private Sub Opcion_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Rem XIndice = 0
    
    Select Case XIndice
        Case 0
            spGasimpo = "ListaGasimpo"
            Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstGasimpo.RecordCount > 0 Then
                With rstGasimpo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstGasimpo!Codigo) + " " + rstGasimpo!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstGasimpo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstGasimpo.Close
            End If
            
        Case Else
    End Select
    Pantalla.Visible = True
            
End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Val(DBGrid1.Text) <> 0 Then
        WLinea.Text = DBGrid1.Row + 1
        WConcepto.Text = DBGrid1.Text
            Else
        WConcepto.Text = ""
        WLinea.Text = ""
    End If
    
    DBGrid1.Col = 1
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 2
    If Val(DBGrid1.Text) <> 0 Then
        WImporte.Text = DBGrid1.Text
            Else
        WImporte.Text = ""
    End If
    
    WConcepto.SetFocus

End Sub

Private Sub Graba_Click()

    Sql1 = "Select *"
    Sql2 = " FROM MovGasParcial"
    Sql3 = " Where MovGasParcial.Codigo = " + "'" + Codigo.Text + "'"
    spMovGasParcial = Sql1 + Sql2 + Sql3
    Set rstMovGasParcial = db.OpenRecordset(spMovGasParcial, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovGasParcial.RecordCount > 0 Then
        WMarca = IIf(IsNull(rstMovGasParcial!Marca), "", rstMovGasParcial!Marca)
        rstMovGasParcial.Close
        If WMarca = "X" Then
            m$ = "Esta carpeta ya fue actualizada"
            a% = MsgBox(m$, 0, "Nacionalizacion de Mercaderia")
            Exit Sub
        End If
    End If
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                    
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Sql1 = "DELETE MovGasParcial"
    Sql2 = " Where Codigo = " + "'" + Codigo.Text + "'"
    spMovGasParcial = Sql1 + Sql2
    Set rstMovGasParcial = db.OpenRecordset(spMovGasParcial, dbOpenSnapshot, dbSQLPassThrough)
    
    Renglon = 0
    DBGrid1.Refresh
        
    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Concepto = DBGrid1.Text
                    
            DBGrid1.Col = 2
            Importe = DBGrid1.Text
                    
            If Val(Concepto) <> 0 Then
                        
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Codigo.Text)
                Call Ceros(Auxi1, 6)
                
                WClave = Auxi1 + Auxi
                WCodigo = Codigo.Text
                WRenglon = Str$(Renglon)
                WFecha = Fecha.Text
                WCarpeta = Carpeta.Text
                WConcepto = Concepto
                WImporte = Importe
                WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WMarca = ""
                WOrden = Orden.Text
                ZEmpresa = Empresa.Text
                
                Sql1 = "INSERT INTO MovGasParcial ("
                Sql2 = "Clave ,"
                Sql3 = "Codigo ,"
                Sql4 = "Renglon ,"
                Sql5 = "Fecha ,"
                Sql6 = "Carpeta ,"
                Sql7 = "Concepto ,"
                Sql8 = "Importe ,"
                Sql9 = "OrdFecha ,"
                Sql10 = "Marca ,"
                Sql11 = "Empresa ,"
                Sql12 = "Orden )"
                Sql13 = "Values ("
                Sql14 = "'" + WClave + "',"
                Sql15 = "'" + WCodigo + "',"
                Sql16 = "'" + WRenglon + "',"
                Sql17 = "'" + WFecha + "',"
                Sql18 = "'" + WCarpeta + "',"
                Sql19 = "'" + WConcepto + "',"
                Sql20 = "'" + WImporte + "',"
                Sql21 = "'" + WOrdFecha + "',"
                Sql22 = "'" + WMarca + "',"
                Sql23 = "'" + ZEmpresa + "',"
                Sql24 = "'" + WOrden + "')"
                spMovGasParcial = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 _
                        + Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 _
                        + Sql21 + Sql22 + Sql23 + Sql24
                Set rstMovGasParcial = db.OpenRecordset(spMovGasParcial, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                
        Next iRow
            
    Next a
    
    Renglon = 0
           
    Sql1 = "DELETE MovGasParcialArti"
    Sql2 = " Where Codigo = " + "'" + Codigo.Text + "'"
    spMovGasParcialArti = Sql1 + Sql2
    Set rstMovGasParcialArti = db.OpenRecordset(spMovGasParcialArti, dbOpenSnapshot, dbSQLPassThrough)
            
    For a = 1 To 100
        
        Articulo = WVector1.TextMatrix(a, 1)
        Cantidad = WVector1.TextMatrix(a, 3)
                    
        If Articulo <> "" Or Val(Cantidad) <> 0 Then
                        
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = Str$(Codigo.Text)
            Call Ceros(Auxi1, 6)
                
            WClave = Auxi1 + Auxi
            WCodigo = Codigo.Text
            WRenglon = Str$(Renglon)
            WFecha = Fecha.Text
            WCarpeta = Carpeta.Text
            WArticulo = Articulo
            WCantidad = Cantidad
            WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            WMarca = ""
            WOrden = Orden.Text
            ZEmpresa = Empresa.Text
                
            Sql1 = "INSERT INTO MovGasParcialArti ("
            Sql2 = "Clave ,"
            Sql3 = "Codigo ,"
            Sql4 = "Renglon ,"
            Sql5 = "Fecha ,"
            Sql6 = "Carpeta ,"
            Sql7 = "Articulo ,"
            Sql8 = "Cantidad ,"
            Sql9 = "OrdFecha ,"
            Sql10 = "Marca ,"
            Sql11 = "Empresa ,"
            Sql12 = "Orden )"
            Sql13 = "Values ("
            Sql14 = "'" + WClave + "',"
            Sql15 = "'" + WCodigo + "',"
            Sql16 = "'" + WRenglon + "',"
            Sql17 = "'" + WFecha + "',"
            Sql18 = "'" + WCarpeta + "',"
            Sql19 = "'" + WArticulo + "',"
            Sql20 = "'" + WCantidad + "',"
            Sql21 = "'" + WOrdFecha + "',"
            Sql22 = "'" + WMarca + "',"
            Sql23 = "'" + ZEmpresa + "',"
            Sql24 = "'" + WOrden + "')"
            spMovGasParcialArti = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 _
                                + Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 _
                                + Sql21 + Sql22 + Sql23 + Sql24
            Set rstMovGasParcialArti = db.OpenRecordset(spMovGasParcialArti, dbOpenSnapshot, dbSQLPassThrough)
                
        End If
                
    Next a
            
    Call Limpia_Click
        
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Codigo.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WConcepto.Text = ""
    WDescripcion.Caption = ""
    WImporte.Text = ""
    
    WConcepto.SetFocus
    
End Sub

Private Sub Limpia_Click()

    Pantalla.Visible = False
    
    WLinea.Text = ""
    WConcepto.Text = ""
    WDescripcion.Caption = ""
    WImporte.Text = ""

    Codigo.Text = ""
    Carpeta.Text = ""
    Orden.Text = ""
    Empresa.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 2
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM MovGasParcial"
    spMovGasParcial = Sql1 + Sql2
    Set rstMovGasParcial = db.OpenRecordset(spMovGasParcial, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovGasParcial.RecordCount > 0 Then
        rstMovGasParcial.MoveLast
        ZCodigo1 = IIf(IsNull(rstMovGasParcial!CodigoMayor), "0", rstMovGasParcial!CodigoMayor)
        rstMovGasParcial.Close
    End If
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM MovGasParcialArti"
    spMovGasParcialArti = Sql1 + Sql2
    Set rstMovGasParcialArti = db.OpenRecordset(spMovGasParcialArti, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovGasParcialArti.RecordCount > 0 Then
        rstMovGasParcialArti.MoveLast
        ZCodigo2 = IIf(IsNull(rstMovGasParcialArti!CodigoMayor), "0", rstMovGasParcialArti!CodigoMayor)
        rstMovGasParcialArti.Close
    End If
    
    If ZCodigo1 > ZCodigo2 Then
        Codigo.Text = ZCodigo1 + 1
            Else
        Codigo.Text = ZCodigo1 + 1
    End If
    
    DBGrid1.FirstRow = 0
    Renglon = 0

    Codigo.SetFocus

End Sub

Private Sub WConcepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spGasimpo = "ConsultaGasimpo " + "'" + WConcepto.Text + "'"
        Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
        If rstGasimpo.RecordCount > 0 Then
            WDescripcion.Caption = rstGasimpo!Nombre
            WImporte.SetFocus
                Else
            WConcepto.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WImporte_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WImporte.Text = Pusing("###,###.##", WImporte.Text)
        Call Alta_Vector
        Call Ingresa_Click
        WConcepto.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
        
            spGasimpo = "ConsultaGasimpo " + "'" + Claveven$ + "'"
            Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
            If rstGasimpo.RecordCount > 0 Then
                    WConcepto.Text = rstGasimpo!Codigo
                    WDescripcion.Caption = rstGasimpo!Nombre
                    
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstGasimpo!Codigo
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstGasimpo!Nombre
                    
                    Call Alta_Vector
                    WLinea.Text = WAnterior + 1
                    If Val(WLinea.Text) > 0 Then
                        DBGrid1.Row = Val(WLinea.Text) - 1
                    End If
                    
                    Call DBGrid1.SetFocus
                    WImporte.SetFocus
                    
            End If

        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 40 Then
                            DBGrid1.Row = DBGrid1.Row + 1
                            DBGrid1.Col = 0
                            KeyCode = 0
                        End If
                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
    End Select

    
End Sub


' Cuando el usuario hace clic en el icono Agregar, esta subrutina agrega una
' nueva fila a la variable RowBuf y un marcador a la variable NewRowBookmark
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows = mTotalRows + 1
ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
NewRowBookmark = mTotalRows - 1 'Establece el marcador a la última fila.

' El bucle siguiente agrega un nuevo registro a la base de datos.
For iCol = 0 To UBound(UserData, 1)
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
    Else
        ' Si no se establece ningún valor para la columna, usa DefaultValue
        UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
    End If
Next iCol

End Sub

' Esta subrutina elimina una fila basándose en su marcador.
Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

' Mueve todas las filas encima de la fila eliminada de
' la matriz.

For iRow = Bookmark + 1 To mTotalRows - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows = mTotalRows - 1

End Sub

' Se llama a esta subrutina cada vez que DBGrid quiere mostrar
' datos nuevos.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid está solicitando filas, así que se las damos

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' Si StartLocation es Null, empieza a leer por el final
' o por el principio del conjunto de datos.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Busca la posición para empezar a leer, basándose en el marcador
    ' StartLocation y en la variable iIncr
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfiere datos de nuestra matriz de conjunto de datos al objeto RowBuf
' que DBGrid utiliza para presentar los datos
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
    For iCol = 0 To UBound(UserData, 1)
        RowBuf.Value(iRow, iCol) = UserData(iCol, CurRow&)
    Next iCol
    ' Establece el marcador mediante CurRow&, que es también
    ' nuestro índice de matriz
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

' Esta subrutina actualiza los datos de la matriz después de
' haberse modificado.

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Se están actualizando los datos

' Actualiza cada columna de la matriz de conjuntos de datos
For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub


Private Sub Form_Load()

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 2, 0 To 40)

mTotalRows& = 40

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DBGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DBGrid1.Columns.Count - 1 To 0 Step -1
      DBGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 2
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Concepto"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Importe"
             DBGrid1.Columns(newcnt).Width = 1800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
 DBGrid1.Font.Bold = True
 
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Rem With rstMovGasParcial
    Rem     .Index = "Clave"
    Rem Claveven$ = "99999999"
    Rem     .Seek "<=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Carpeta.text = !Carpeta + 1
    Rem             Else
    Rem         Carpeta.text = ""
    Rem     End If
    Rem End With
    
    Carpeta.Text = ""
    Codigo.Text = ""
    Orden.Text = ""
    Empresa.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM MovGasParcial"
    spMovGasParcial = Sql1 + Sql2
    Set rstMovGasParcial = db.OpenRecordset(spMovGasParcial, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovGasParcial.RecordCount > 0 Then
        rstMovGasParcial.MoveLast
        ZCodigo1 = IIf(IsNull(rstMovGasParcial!CodigoMayor), "0", rstMovGasParcial!CodigoMayor)
        rstMovGasParcial.Close
    End If
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM MovGasParcialArti"
    spMovGasParcialArti = Sql1 + Sql2
    Set rstMovGasParcialArti = db.OpenRecordset(spMovGasParcialArti, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovGasParcialArti.RecordCount > 0 Then
        rstMovGasParcialArti.MoveLast
        ZCodigo2 = IIf(IsNull(rstMovGasParcialArti!CodigoMayor), "0", rstMovGasParcialArti!CodigoMayor)
        rstMovGasParcialArti.Close
    End If
    
    If ZCodigo1 > ZCodigo2 Then
        Codigo.Text = ZCodigo1 + 1
            Else
        Codigo.Text = ZCodigo1 + 1
    End If
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgMovgasParcial.Caption = "Ingreso de Gastos de Despacho :  " + !Nombre
        End If
    End With
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Codigo.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 2
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    Erase Auxiliar
    
    Sql1 = "Select *"
    Sql2 = " FROM MovGasParcial"
    Sql3 = " Where MovGasParcial.Codigo = " + "'" + Codigo.Text + "'"
    Sql4 = " Order by MovGasParcial.Clave"
    spMovGasParcial = Sql1 + Sql2 + Sql3 + Sql4
    Set rstMovGasParcial = db.OpenRecordset(spMovGasParcial, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovGasParcial.RecordCount > 0 Then
    
        With rstMovGasParcial
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstMovGasParcial!Concepto
                    Auxi1 = rstMovGasParcial!Concepto
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.##", rstMovGasParcial!Importe)
                
                    Auxiliar(Renglon, 1) = Auxi1
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovGasParcial.Close
                
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Renglon = 1 To WRenglon
    
        Auxi1 = Auxiliar(Renglon, 1)
    
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
        
        spGasimpo = "ConsultaGAsimpo " + "'" + Auxi1 + "'"
        Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
        If rstGasimpo.RecordCount > 0 Then
            DBGrid1.Col = 1
            DBGrid1.Text = rstGasimpo!Nombre
            WConcepto.SetFocus
        End If
        
    Next Renglon
    
    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    Call Limpia_Vector
    
    
    Erase Vector
    Entra = 0

    Sql1 = "Select *"
    Sql2 = " FROM MovGasParcialArti"
    Sql3 = " Where MovGasParcialArti.Codigo = " + "'" + Codigo.Text + "'"
    Sql4 = " Order by MovGasParcialArti.Clave"
    spMovGasParcialArti = Sql1 + Sql2 + Sql3 + Sql4
    Set rstMovGasParcialArti = db.OpenRecordset(spMovGasParcialArti, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovGasParcialArti.RecordCount > 0 Then
        With rstMovGasParcialArti
            .MoveFirst
            Do
                If .EOF = False Then
                    Entra = Entra + 1
                    WVector1.Row = Entra
                    WVector1.Col = 1
                    WVector1.Text = rstMovGasParcialArti!Articulo
                    WVector1.Col = 2
                    WVector1.Text = ""
                    WVector1.Col = 3
                    WVector1.Text = Str$(rstMovGasParcialArti!Cantidad)
                    WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovGasParcialArti.Close
    End If
    
    For Cicla = 1 To 100
        WArticulo = WVector1.TextMatrix(Cicla, 1)
        spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WVector1.Row = Cicla
            WVector1.Col = 2
            WVector1.Text = rstArticulo!Descripcion
            rstArticulo.Close
        End If
    Next Cicla
    
    WConcepto.SetFocus

End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) = 0 Then

            Renglon = Renglon + 1
            
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            
            DBGrid1.Text = WConcepto.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WImporte.Text)
            
            DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WConcepto.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WImporte.Text)
            
            DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Codigo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Sql1 = "Select *"
        Sql2 = " FROM MovGasParcial"
        Sql3 = " Where MovGasParcial.Codigo = " + "'" + Codigo.Text + "'"
        spMovGasParcial = Sql1 + Sql2 + Sql3
        Set rstMovGasParcial = db.OpenRecordset(spMovGasParcial, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovGasParcial.RecordCount > 0 Then
        
            Fecha.Text = rstMovGasParcial!Fecha
            Carpeta.Text = rstMovGasParcial!Carpeta
            Orden.Text = rstMovGasParcial!Orden
            Empresa.Text = rstMovGasParcial!Empresa
            rstMovGasParcial.Close
            Call Proceso_Click
            WConcepto.SetFocus
            
                Else
            
            Sql1 = "Select *"
            Sql2 = " FROM MovGasParcialArti"
            Sql3 = " Where MovGasParcialArti.Codigo = " + "'" + Codigo.Text + "'"
            spMovGasParcialArti = Sql1 + Sql2 + Sql3
            Set rstMovGasParcialArti = db.OpenRecordset(spMovGasParcialArti, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovGasParcialArti.RecordCount > 0 Then
        
                Fecha.Text = rstMovGasParcialArti!Fecha
                Carpeta.Text = rstMovGasParcialArti!Carpeta
                Orden.Text = rstMovGasParcialArti!Orden
                Empresa.Text = rstMovGasParcialArti!Empresa
                rstMovGasParcialArti.Close
                Call Proceso_Click
                WConcepto.SetFocus
            
                    Else
                
                WCodigo = Codigo.Text
                Call Limpia_Click
                Codigo.Text = WCodigo
                Fecha.SetFocus
                
            End If
            
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Carpeta.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Carpeta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 And Val(Carpeta.Text) <> 0 Then
        Sql1 = "Select *"
        Sql2 = " FROM MovGas"
        Sql3 = " Where MovGas.Carpeta = " + "'" + Carpeta.Text + "'"
        spMovgas = Sql1 + Sql2 + Sql3
        Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovgas.RecordCount > 0 Then
            Orden.Text = rstMovgas!Orden
            Empresa.Text = rstMovgas!Empresa
            Call Lee_Articulos
            rstMovgas.Close
            WConcepto.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1
        Case 113
            WTexto1.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1
        Case 113
            WTexto2.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto3.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 3
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 3
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            WVector1.Col = XColumna
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 4
    WVector1.FixedRows = 1
    WVector1.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Articulo"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 3400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "######"
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub








Private Sub Derechos_Click()

    IngresaDerechos.Height = 6015
    IngresaDerechos.Left = 1080
    IngresaDerechos.Top = 960
    IngresaDerechos.Width = 7095

    
    IngresaDerechos.Visible = True
    
    WVector1.Col = 3
    WVector1.Row = 1
    WVector1.TopRow = 1
    Call StartEdit
    
End Sub


Private Sub FinIngreso_Click()
    IngresaDerechos.Visible = False
End Sub

Private Sub Lee_Articulos()
                
    EmpresaAnterior = WEmpresa
    Select Case Val(Empresa.Text)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 11
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    
    Call Limpia_Vector
    
    Erase Vector
    Entre = 0

    spOrden = "ListaOrden " + "'" + Orden.Text + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                    Entra = Entra + 1
                    WVector1.Row = Entra
                    WVector1.Col = 1
                    WVector1.Text = rstOrden!Articulo
                    WVector1.Col = 2
                    WVector1.Text = ""
                    WVector1.Col = 3
                    WVector1.Text = ""
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If
    
    For Cicla = 1 To 100
        WArticulo = WVector1.TextMatrix(Cicla, 1)
        spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WVector1.Row = Cicla
            WVector1.Col = 2
            WVector1.Text = rstArticulo!Descripcion
            rstArticulo.Close
        End If
    Next Cicla
    
    Select Case Val(EmpresaAnterior)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 11
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select

End Sub

