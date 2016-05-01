VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCargaIV 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Instrucciones de Produccion de P.T."
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   495
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin VB.CommandButton david 
      Caption         =   "david"
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
      Left            =   9360
      TabIndex        =   38
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Salva 
      Caption         =   "salva"
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
      Left            =   7680
      TabIndex        =   37
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox ControlCambio 
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
      Left            =   2160
      MaxLength       =   100
      TabIndex        =   36
      Top             =   840
      Width           =   9615
   End
   Begin VB.TextBox Metodo 
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
      Left            =   9480
      MaxLength       =   2
      TabIndex        =   33
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame IngresaBase 
      Height          =   1215
      Left            =   6960
      TabIndex        =   29
      Top             =   2520
      Visible         =   0   'False
      Width           =   3015
      Begin MSMask.MaskEdBox ProductoBase 
         Height          =   285
         Left            =   720
         TabIndex        =   30
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
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
   End
   Begin VB.CommandButton Base 
      Caption         =   "Instrucciones Base"
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
      Left            =   8520
      TabIndex        =   28
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Frame XClaveII 
      Height          =   1935
      Left            =   3120
      TabIndex        =   24
      Top             =   1920
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton CancelaGrabaII 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   26
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox WClaveII 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   25
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   27
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton Revalida 
      Caption         =   "Revalida"
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
      Left            =   7320
      TabIndex        =   23
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   3480
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   22
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton AgregaRenglon 
      Caption         =   "Agrega Renglon"
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
      Left            =   10560
      TabIndex        =   18
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox Autorizado 
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
      Left            =   6840
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Version 
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
      Left            =   4200
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   15
      Top             =   480
      Width           =   1095
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
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   390
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
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Ayuda 
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
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   6855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10560
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
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
      Height          =   1260
      Left            =   2280
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
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
      ItemData        =   "CargaIV.frx":0000
      Left            =   120
      List            =   "CargaIV.frx":0007
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   2280
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
      Height          =   4455
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7858
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox Terminado 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   1560
      TabIndex        =   13
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   327680
      Enabled         =   0   'False
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
   Begin VB.Label Label4 
      Caption         =   "Control de Cambios"
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
      Left            =   240
      TabIndex        =   35
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label40 
      Caption         =   "Metodo Lavado"
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
      Left            =   7920
      TabIndex        =   34
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Responsable"
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
      Left            =   7920
      TabIndex        =   32
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label DesOperador 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9480
      TabIndex        =   31
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Autorizado"
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
      Left            =   5520
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Version"
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
      Left            =   3240
      TabIndex        =   14
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label DesTerminado 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3240
      TabIndex        =   11
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   9720
      MouseIcon       =   "CargaIV.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "CargaIV.frx":031F
      ToolTipText     =   "Salida"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7320
      MouseIcon       =   "CargaIV.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "CargaIV.frx":0E6B
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   8160
      MouseIcon       =   "CargaIV.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "CargaIV.frx":19B7
      ToolTipText     =   "Consulta de Datos"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   9000
      MouseIcon       =   "CargaIV.frx":21F9
      MousePointer    =   99  'Custom
      Picture         =   "CargaIV.frx":2503
      ToolTipText     =   "Limpia la pantalla"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Producto"
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
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgCargaIV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEquipoFabrica As Recordset
Dim spEquipoFabrica As String
Dim rstCargaIV As Recordset
Dim rsCargaIV As String
Dim rstOperador As Recordset
Dim spOperador As String

Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Cantidad As Double
Dim XPaso As String
Dim Renglon As Integer
Dim ZCodigo As String
Dim ZOperador As String
Dim ZLugar2(700) As Integer

Dim ZLugar(600) As Integer
Dim ZDescri(1000, 700) As String

Dim ZTraspaso(1000, 20) As String
Dim WVersion As String
Dim WRenglon As String

Dim ZVector(10000) As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Private WGraba As String
Private WGrabaII As String

Dim CargaEmpresa(10, 2) As String

Private Sub Base_Click()

    IngresaBase.Visible = True
    
    ProductoBase.Text = "  -     -   "
    ProductoBase.SetFocus

End Sub

Private Sub Command1_Click()
    
    Erase ZLugar
    Erase ZDescri
    
    Sql1 = "Select *"
    Sql2 = " FROM EquipoFabrica"
    Sql3 = " Order by Codigo"
    spEquipoFabrica = Sql1 + Sql2 + Sql3
    Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipoFabrica.RecordCount > 0 Then
        With rstEquipoFabrica
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZDescripcion = IIf(IsNull(rstEquipoFabrica!Descripcion), "", rstEquipoFabrica!Descripcion)
                    ZDescripcionII = IIf(IsNull(rstEquipoFabrica!DescripcionII), "", rstEquipoFabrica!DescripcionII)
                    ZDescripcionIII = IIf(IsNull(rstEquipoFabrica!DescripcionIII), "", rstEquipoFabrica!DescripcionIII)
                    
                    WDescripcion = Trim(ZDescripcion) + " " + Trim(ZDescripcionII) + " " + Trim(ZDescripcionIII)
                    ZCodigo = rstEquipoFabrica!codigo
                    
                    ZHasta = Len(WDescripcion)
                    Desde = 1
                    
                    Do
                        Hasta = Desde + 15
                        If Hasta > ZHasta Then
                            Hasta = ZHasta
                        End If
                        ZLugar(ZCodigo) = ZLugar(ZCodigo) + 1
                        aa = Mid(WDescripcion, Desde, Hasta)
                        ZDescri(ZCodigo, ZLugar(ZCodigo)) = Mid(WDescripcion, Desde, Hasta)
                        For Cicla = Hasta To Desde Step -1
                            aa = Mid(WDescripcion, Cicla, 1)
                            If Mid(WDescripcion, Cicla, 1) = Space(1) Then
                                aa = Mid(WDescripcion, Desde, Cicla - Desde)
                                ZDescri(ZCodigo, ZLugar(ZCodigo)) = Mid(WDescripcion, Desde, Cicla - Desde)
                                Desde = Cicla + 1
                                Exit For
                            End If
                        Next Cicla
                        
                        If Hasta >= ZHasta Then
                            Exit Do
                        End If
                    Loop
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEquipoFabrica.Close
    End If
    
    Erase ZVector
    Lugar = 0
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstTerminado
        .MoveFirst
            Do
            If .EOF = False Then
                If rstTerminado!codigo >= "PT-11038-000" And rstTerminado!codigo <= "PT-11930-999" Then
                    Lugar = Lugar + 1
                    ZVector(Lugar) = rstTerminado!codigo
                End If
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstTerminado.Close
    
    For Ciclo = 1 To Lugar
    
        ZTerminado = ZVector(Ciclo)
    
        Sql1 = "Select *"
        Sql2 = " FROM CargaIV"
        Sql3 = " Where CargaIV.Terminado = " + "'" + ZTerminado + "'"
        rsCargaIV = Sql1 + Sql2 + Sql3
        Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaIV.RecordCount > 0 Then
            rstCargaIV.Close
            ZGraba = "N"
                Else
            ZGraba = "S"
        End If
        
        If ZGraba = "S" Then
        
            HastaRenglon = 0
            For iRow = 100 To 1 Step -1
        
                Etapa = WVector1.TextMatrix(iRow, 1)
                LetraInstrucciones = WVector1.TextMatrix(iRow, 2)
                Instrucciones = WVector1.TextMatrix(iRow, 3)
                Equipo = WVector1.TextMatrix(iRow, 4)
                LetraTemperatura = WVector1.TextMatrix(iRow, 5)
                Temperatura = WVector1.TextMatrix(iRow, 6)
                LetraTiempo = WVector1.TextMatrix(iRow, 7)
                Tiempo = WVector1.TextMatrix(iRow, 8)
                LetraControl = WVector1.TextMatrix(iRow, 9)
                Control = WVector1.TextMatrix(iRow, 10)
                Seguridad = WVector1.TextMatrix(iRow, 11)
                
                If Etapa <> "" Or Instrucciones <> "" Or Equipo <> "" Or Temperatura <> "" Or Tiempo <> "" Or Control <> "" Or Seguridad <> "" Then
                    HastaRenglon = iRow
                    Exit For
                End If
            
            Next iRow
    
            WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            Erase ZLugar
            XEquipo = ""
            XControl = ""
            XSeguridad = ""

            WRenglon = 0
            For iRow = 1 To HastaRenglon
    
                ZLote = ""
        
                Etapa = WVector1.TextMatrix(iRow, 1)
                LetraInstrucciones = WVector1.TextMatrix(iRow, 2)
                Instrucciones = WVector1.TextMatrix(iRow, 3)
                Equipo = WVector1.TextMatrix(iRow, 4)
                LetraTemperatura = WVector1.TextMatrix(iRow, 5)
                Temperatura = WVector1.TextMatrix(iRow, 6)
                LetraTiempo = WVector1.TextMatrix(iRow, 7)
                Tiempo = WVector1.TextMatrix(iRow, 8)
                LetraControl = WVector1.TextMatrix(iRow, 9)
                Control = WVector1.TextMatrix(iRow, 10)
                Seguridad = WVector1.TextMatrix(iRow, 11)
        
                IMPRE11 = ""
                Impre12 = ""
                Impre13 = ""
                Impre14 = ""
                Impre15 = ""
                Impre16 = ""
                Impre17 = ""
                Impre18 = ""
                Impre19 = ""
            
                Impre31 = ""
                Impre32 = ""
                Impre33 = ""
                Impre34 = ""
                Impre35 = ""
                Impre36 = ""
                Impre37 = ""
                Impre38 = ""
                Impre39 = ""
        
                Impre41 = ""
                Impre42 = ""
                Impre43 = ""
                Impre44 = ""
                Impre45 = ""
                Impre46 = ""
                Impre47 = ""
                Impre48 = ""
                Impre49 = ""
        
                Impre51 = ""
                Impre52 = ""
                Impre53 = ""
                Impre54 = ""
                Impre55 = ""
                Impre56 = ""
                Impre57 = ""
                Impre58 = ""
                Impre59 = ""
        
                Select Case LetraInstrucciones
                    Case "8"
                        Impre12 = Instrucciones
                    Case "12"
                        Impre13 = Instrucciones
                    Case "N8"
                        Impre14 = Instrucciones
                    Case "N10"
                        Impre15 = Instrucciones
                    Case "N12"
                        Impre16 = Instrucciones
                    Case "FS"
                        Impre17 = Instrucciones
                    Case "FO"
                        Impre18 = Instrucciones
                    Case "R"
                        Impre19 = Instrucciones
                    Case Else
                        IMPRE11 = Instrucciones
                End Select
            
                If Val(Equipo) <> XEquipo Then
                    ZLugar(Val(Equipo)) = 0
                End If
                If Val(Control) <> XControl Then
                    ZLugar(Val(Control)) = 0
                End If
                If Val(Seguridad) <> XSeguridad Then
                    ZLugar(Val(Seguridad)) = 0
                End If
            
                If Val(Equipo) <> 0 Then
                    ZLugar(Val(Equipo)) = ZLugar(Val(Equipo)) + 1
                    If ZDescri(Val(Equipo), ZLugar(Val(Equipo))) <> "" Then
                        Impre2 = ZDescri(Val(Equipo), ZLugar(Val(Equipo)))
                            Else
                        Impre2 = "."
                    End If
                        Else
                    Impre2 = Equipo
                End If
        
                Select Case LetraTemperatura
                    Case "8"
                        Impre32 = Temperatura
                    Case "12"
                        Impre33 = Temperatura
                    Case "N8"
                        Impre34 = Temperatura
                    Case "N10"
                        Impre35 = Temperatura
                    Case "N12"
                        Impre36 = Temperatura
                    Case "FS"
                        Impre37 = Temperatura
                    Case "FO"
                        Impre38 = Temperatura
                    Case "R"
                        Impre39 = Temperatura
                    Case Else
                        Impre31 = Temperatura
                End Select
        
                Select Case LetraTiempo
                    Case "8"
                        Impre42 = Tiempo
                    Case "12"
                        Impre43 = Tiempo
                    Case "N8"
                        Impre44 = Tiempo
                    Case "N10"
                        Impre45 = Tiempo
                    Case "N12"
                        Impre46 = Tiempo
                    Case "FS"
                        Impre47 = Tiempo
                    Case "FO"
                        Impre48 = Tiempo
                    Case "R"
                        Impre49 = Tiempo
                    Case Else
                        Impre41 = Tiempo
                End Select
        
                If Val(Control) <> 0 Then
                    ZLugar(Val(Control)) = ZLugar(Val(Control)) + 1
                    If ZDescri(Val(Control), ZLugar(Val(Control))) <> "" Then
                        LetraControl = "FS"
                        ZControl = ZDescri(Val(Control), ZLugar(Val(Control)))
                            Else
                        LetraControl = "FS"
                        ZControl = "."
                    End If
                        Else
                    ZControl = Control
                End If
        
                Select Case LetraControl
                    Case "8"
                        Impre52 = ZControl
                    Case "12"
                        Impre53 = ZControl
                    Case "N8"
                        Impre54 = ZControl
                    Case "N10"
                        Impre55 = ZControl
                    Case "N12"
                        Impre56 = ZControl
                    Case "FS"
                        Impre57 = ZControl
                    Case "FO"
                        Impre58 = ZControl
                    Case "R"
                        Impre59 = ZControl
                    Case Else
                        Impre51 = ZControl
                End Select
            
                If Val(Seguridad) <> 0 Then
                    ZLugar(Val(Seguridad)) = ZLugar(Val(Seguridad)) + 1
                    If ZDescri(Val(Seguridad), ZLugar(Val(Seguridad))) <> "" Then
                        Impre6 = ZDescri(Val(Seguridad), ZLugar(Val(Seguridad)))
                            Else
                        Impre6 = "."
                    End If
                        Else
                    Impre6 = Seguridad
                End If
                
                XEquipo = ZEquipo
                XControl = ZControl
                XSeguridad = ZSeguridad
            
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 2)
        
                WClave = ZTerminado + Auxi
        
                XXVersion = "1"
                XXFechaVersion = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                XXAutorizado = "S"
                XXOrdFecha = Right$(XXFechaVersion, 4) + Mid$(XXFechaVersion, 4, 2) + Left$(XXFechaVersion, 2)
        
                Sql1 = "INSERT INTO CargaIV ("
                Sql2 = "Clave ,"
                Sql3 = "Terminado ,"
                Sql4 = "Renglon ,"
                Sql5 = "Fecha ,"
                Sql6 = "OrdFecha ,"
                Sql7 = "Lote ,"
                Sql8 = "Version ,"
                Sql9 = "Autorizado ,"
                Sql10 = "Etapa ,"
                Sql11 = "LetraInstrucciones ,"
                Sql12 = "Instrucciones ,"
                Sql13 = "Equipo ,"
                Sql14 = "LetraTemperatura ,"
                Sql15 = "Temperatura ,"
                Sql16 = "LetraTiempo ,"
                Sql17 = "Tiempo ,"
                Sql18 = "LetraControl ,"
                Sql19 = "Control ,"
                Sql20 = "Seguridad ,"
                Sql21 = "DesTerminado )"
                Sql22 = "Values ("
                Sql23 = "'" + WClave + "',"
                Sql24 = "'" + ZTerminado + "',"
                Sql25 = "'" + Str$(WRenglon) + "',"
                Sql26 = "'" + XXFechaVersion + "',"
                Sql27 = "'" + XXOrdFecha + "',"
                Sql28 = "'" + ZLote + "',"
                Sql29 = "'" + XXVersion + "',"
                Sql30 = "'" + XXAutorizado + "',"
                Sql31 = "'" + Etapa + "',"
                Sql32 = "'" + LetraInstrucciones + "',"
                Sql33 = "'" + Instrucciones + "',"
                Sql34 = "'" + Equipo + "',"
                Sql35 = "'" + LetraTemperatura + "',"
                Sql36 = "'" + Temperatura + "',"
                Sql37 = "'" + LetraTiempo + "',"
                Sql38 = "'" + Tiempo + "',"
                Sql39 = "'" + LetraControl + "',"
                Sql40 = "'" + Control + "',"
                Sql41 = "'" + Seguridad + "',"
                Sql42 = "'" + DesTerminado.Caption + "')"
                
                rsCargaIV = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 _
                            + Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 _
                            + Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 _
                            + Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 _
                            + Sql41 + Sql42
                Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
        
                Sql1 = "UPDATE CargaIV SET "
                Sql2 = "Impre11 = " + "'" + IMPRE11 + "',"
                Sql3 = "Impre12 = " + "'" + Impre12 + "',"
                Sql4 = "Impre13 = " + "'" + Impre13 + "',"
                Sql5 = "Impre14 = " + "'" + Impre14 + "',"
                Sql6 = "Impre15 = " + "'" + Impre15 + "',"
                Sql7 = "Impre16 = " + "'" + Impre16 + "',"
                Sql8 = "Impre17 = " + "'" + Impre17 + "',"
                Sql9 = "Impre18 = " + "'" + Impre18 + "',"
                Sql10 = "Impre19 = " + "'" + Impre19 + "',"
                Sql11 = "Impre2 = " + "'" + Impre2 + "',"
                Sql12 = "Impre31 = " + "'" + Impre31 + "',"
                Sql13 = "Impre32 = " + "'" + Impre32 + "',"
                Sql14 = "Impre33 = " + "'" + Impre33 + "',"
                Sql15 = "Impre34 = " + "'" + Impre34 + "',"
                Sql16 = "Impre35 = " + "'" + Impre35 + "',"
                Sql17 = "Impre36 = " + "'" + Impre36 + "',"
                Sql18 = "Impre37 = " + "'" + Impre37 + "',"
                Sql19 = "Impre38 = " + "'" + Impre38 + "',"
                Sql20 = "Impre39 = " + "'" + Impre39 + "',"
                Sql21 = "Impre41 = " + "'" + Impre41 + "',"
                Sql22 = "Impre42 = " + "'" + Impre42 + "',"
                Sql23 = "Impre43 = " + "'" + Impre43 + "',"
                Sql24 = "Impre44 = " + "'" + Impre44 + "',"
                Sql25 = "Impre45 = " + "'" + Impre45 + "',"
                Sql26 = "Impre46 = " + "'" + Impre46 + "',"
                Sql27 = "Impre47 = " + "'" + Impre47 + "',"
                Sql28 = "Impre48 = " + "'" + Impre48 + "',"
                Sql29 = "Impre49 = " + "'" + Impre49 + "',"
                Sql30 = "Impre51 = " + "'" + Impre51 + "',"
                Sql31 = "Impre52 = " + "'" + Impre52 + "',"
                Sql32 = "Impre53 = " + "'" + Impre53 + "',"
                Sql33 = "Impre54 = " + "'" + Impre54 + "',"
                Sql34 = "Impre55 = " + "'" + Impre55 + "',"
                Sql35 = "Impre56 = " + "'" + Impre56 + "',"
                Sql36 = "Impre57 = " + "'" + Impre57 + "',"
                Sql37 = "Impre58 = " + "'" + Impre58 + "',"
                Sql38 = "Impre59 = " + "'" + Impre59 + "',"
                Sql39 = "Impre6 = " + "'" + Impre6 + "'"
                Sql40 = " Where Clave = " + "'" + WClave + "'"
    
                rsCargaIV = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 _
                            + Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 _
                            + Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 _
                            + Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40
                Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
            
            Next iRow
        
            XEmpresa = WEmpresa
            Erase CargaEmpresa
        
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7
                    CargaEmpresa(1, 1) = "0001"
                    CargaEmpresa(1, 2) = "Empresa01"
                    CargaEmpresa(2, 1) = "0003"
                    CargaEmpresa(2, 2) = "Empresa03"
                    CargaEmpresa(3, 1) = "0005"
                    CargaEmpresa(3, 2) = "Empresa05"
                    CargaEmpresa(4, 1) = "0006"
                    CargaEmpresa(4, 2) = "Empresa06"
                    CargaEmpresa(5, 1) = "0007"
                    CargaEmpresa(5, 2) = "Empresa07"
                Case 2, 4, 8, 9
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                Case 10
                    CargaEmpresa(1, 1) = "0010"
                    CargaEmpresa(1, 2) = "Empresa10"
                Case Else
            End Select
                
            For Cicla = 1 To 5
                If CargaEmpresa(Cicla, 1) <> "" Then
            
                    WEmpresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Terminado SET "
                    ZSql = ZSql + " VersionI = " + "'" + XXVersion + "',"
                    ZSql = ZSql + " FechaVersionI = " + "'" + XXFechaVersion + "',"
                    ZSql = ZSql + " EstadoI = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + ZTerminado + "'"
                    
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
                    XTipoPro = "PT"
                    XCodigo = Val(Mid$(ZTerminado, 4, 5))
                    
                    If Left$(ZTerminado, 2) = "DY" Or Left$(ZTerminado, 2) = "DW" Or Left$(ZTerminado, 2) = "DS" Then
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
                    
                    If XTipoPro <> "FA" Then
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Terminado SET "
                        ZSql = ZSql + " EstadoII = " + "'" + "N" + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + ZTerminado + "'"
                        
                        spTerminado = ZSql
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                
                End If
            Next Cicla
        
            Select Case Val(XEmpresa)
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
                Case Else
            End Select
        End If
    
    Next Ciclo

    m$ = "Provceso a finalizado correctamente"
    G% = MsgBox(m$, 0, "Ingreso de Procesos de Fabricacion")

End Sub



Private Sub david_Click()
    
    OPEN_FILE_Procesos
    
    With rstProcesos
        .Index = "Codigo"
        .Seek ">=", ""
        If .NoMatch = False Then
            
            Do
            
                Terminado.Text = !codigo
                Call Terminado_KeyPress(13)
                
                Call Salva_Click
                Call GrabaAuto_Click
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With

End Sub

Private Sub ProductoBase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        Call Limpia_Vector
        WRenglon = 0
    
        Sql1 = "Select *"
        Sql2 = " FROM CargaIV"
        Sql3 = " Where CargaIV.Terminado = " + "'" + ProductoBase.Text + "'"
        Sql4 = " Order by CargaIV.Clave"
    
        rsCargaIV = Sql1 + Sql2 + Sql3 + Sql4
        Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaIV.RecordCount > 0 Then
            With rstCargaIV
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        WRenglon = WRenglon + 1
                        WVector1.Row = WRenglon
                        Renglon = WRenglon
                
                        WVector1.Col = 0
                        WVector1.Text = Trim(rstCargaIV!Etapa)
                    
                        WVector1.Col = 1
                        WVector1.Text = Trim(rstCargaIV!Etapa)
            
                        WVector1.Col = 2
                        WVector1.Text = Trim(rstCargaIV!LetraInstrucciones)
                    
                        WVector1.Col = 3
                        WVector1.Text = Trim(rstCargaIV!Instrucciones)
            
                        WVector1.Col = 4
                        WVector1.Text = Trim(rstCargaIV!Equipo)
            
                        WVector1.Col = 5
                        WVector1.Text = Trim(rstCargaIV!LetraTemperatura)
                    
                        WVector1.Col = 6
                        WVector1.Text = Trim(rstCargaIV!Temperatura)
            
                        WVector1.Col = 7
                        WVector1.Text = Trim(rstCargaIV!LetraTiempo)
                    
                        WVector1.Col = 8
                        WVector1.Text = Trim(rstCargaIV!Tiempo)
            
                        WVector1.Col = 9
                        WVector1.Text = Trim(rstCargaIV!LetraControl)
                    
                        WVector1.Col = 10
                        WVector1.Text = Trim(rstCargaIV!Control)
            
                        WVector1.Col = 11
                        WVector1.Text = Trim(rstCargaIV!Seguridad)
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCargaIV.Close
        End If
        
        IngresaBase.Visible = False
        
    End If
    If KeyAscii = 27 Then
        ProductoBase.Text = "  -     -   "
    End If
End Sub


Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Equipos"
     Opcion.AddItem "Control"
     Opcion.AddItem "Seguridad"
     Opcion.Visible = True
     
End Sub



Private Sub Image1_Click()


        Terminado.Text = UCase(Terminado.Text)

        Sql1 = "DELETE CargaIV"
        Sql2 = " Where Terminado = " + "'" + Terminado.Text + "'"
        rsCargaIV = Sql1 + Sql2
        Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
    
        Erase ZLugar
        Erase ZDescri
    
        Sql1 = "Select *"
        Sql2 = " FROM EquipoFabrica"
        Sql4 = " Order by Codigo"
        spEquipoFabrica = Sql1 + Sql2 + Sql3
        Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEquipoFabrica.RecordCount > 0 Then
            With rstEquipoFabrica
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        ZDescripcion = IIf(IsNull(rstEquipoFabrica!Descripcion), "", rstEquipoFabrica!Descripcion)
                        ZDescripcionII = IIf(IsNull(rstEquipoFabrica!DescripcionII), "", rstEquipoFabrica!DescripcionII)
                        ZDescripcionIII = IIf(IsNull(rstEquipoFabrica!DescripcionIII), "", rstEquipoFabrica!DescripcionIII)
                    
                        WDescripcion = Trim(ZDescripcion) + " " + Trim(ZDescripcionII) + " " + Trim(ZDescripcionIII)
                        ZCodigo = rstEquipoFabrica!codigo
                    
                        ZHasta = Len(WDescripcion)
                        Desde = 1
                    
                        Do
                            Hasta = Desde + 15
                            If Hasta > ZHasta Then
                                Hasta = ZHasta
                            End If
                            ZLugar(ZCodigo) = ZLugar(ZCodigo) + 1
                            aa = Mid(WDescripcion, Desde, Hasta)
                            ZDescri(ZCodigo, ZLugar(ZCodigo)) = Mid(WDescripcion, Desde, Hasta)
                            For Cicla = Hasta To Desde Step -1
                                aa = Mid(WDescripcion, Cicla, 1)
                                If Mid(WDescripcion, Cicla, 1) = Space(1) Then
                                    aa = Mid(WDescripcion, Desde, Cicla - Desde)
                                    ZDescri(ZCodigo, ZLugar(ZCodigo)) = Mid(WDescripcion, Desde, Cicla - Desde)
                                    Desde = Cicla + 1
                                    Exit For
                                End If
                            Next Cicla
                        
                            If Hasta >= ZHasta Then
                                Exit Do
                            End If
                        Loop
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstEquipoFabrica.Close
        End If
    
        HastaRenglon = 0
        For iRow = 100 To 1 Step -1
        
            Etapa = WVector1.TextMatrix(iRow, 1)
            LetraInstrucciones = WVector1.TextMatrix(iRow, 2)
            Instrucciones = WVector1.TextMatrix(iRow, 3)
            Equipo = WVector1.TextMatrix(iRow, 4)
            LetraTemperatura = WVector1.TextMatrix(iRow, 5)
            Temperatura = WVector1.TextMatrix(iRow, 6)
            LetraTiempo = WVector1.TextMatrix(iRow, 7)
            Tiempo = WVector1.TextMatrix(iRow, 8)
            LetraControl = WVector1.TextMatrix(iRow, 9)
            Control = WVector1.TextMatrix(iRow, 10)
            Seguridad = WVector1.TextMatrix(iRow, 11)
            
            If Etapa <> "" Or Instrucciones <> "" Or Equipo <> "" Or Temperatura <> "" Or Tiempo <> "" Or Control <> "" Or Seguridad <> "" Then
                HastaRenglon = iRow
                Exit For
            End If
            
        Next iRow
    
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        Erase ZLugar

        WRenglon = 0
        For iRow = 1 To HastaRenglon
    
            ZLote = ""
        
            Etapa = WVector1.TextMatrix(iRow, 1)
            LetraInstrucciones = WVector1.TextMatrix(iRow, 2)
            Instrucciones = WVector1.TextMatrix(iRow, 3)
            Equipo = WVector1.TextMatrix(iRow, 4)
            LetraTemperatura = WVector1.TextMatrix(iRow, 5)
            Temperatura = WVector1.TextMatrix(iRow, 6)
            LetraTiempo = WVector1.TextMatrix(iRow, 7)
            Tiempo = WVector1.TextMatrix(iRow, 8)
            LetraControl = WVector1.TextMatrix(iRow, 9)
            Control = WVector1.TextMatrix(iRow, 10)
            Seguridad = WVector1.TextMatrix(iRow, 11)
        
            IMPRE11 = ""
            Impre12 = ""
            Impre13 = ""
            Impre14 = ""
            Impre15 = ""
            Impre16 = ""
            Impre17 = ""
            Impre18 = ""
            Impre19 = ""
        
            Impre31 = ""
            Impre32 = ""
            Impre33 = ""
            Impre34 = ""
            Impre35 = ""
            Impre36 = ""
            Impre37 = ""
            Impre38 = ""
            Impre39 = ""
        
            Impre41 = ""
            Impre42 = ""
            Impre43 = ""
            Impre44 = ""
            Impre45 = ""
            Impre46 = ""
            Impre47 = ""
            Impre48 = ""
            Impre49 = ""
        
            Impre51 = ""
            Impre52 = ""
            Impre53 = ""
            Impre54 = ""
            Impre55 = ""
            Impre56 = ""
            Impre57 = ""
            Impre58 = ""
            Impre59 = ""
        
            Select Case LetraInstrucciones
                Case "8"
                    Impre12 = Instrucciones
                Case "12"
                    Impre13 = Instrucciones
                Case "N8"
                    Impre14 = Instrucciones
                Case "N10"
                    Impre15 = Instrucciones
                Case "N12"
                    Impre16 = Instrucciones
                Case "FS"
                    Impre17 = Instrucciones
                Case "FO"
                    Impre18 = Instrucciones
                Case "R"
                    Impre19 = Instrucciones
                Case Else
                    IMPRE11 = Instrucciones
            End Select
            
            If Val(Equipo) <> 0 Then
                ZLugar(Val(Equipo)) = ZLugar(Val(Equipo)) + 1
                If ZDescri(Val(Equipo), ZLugar(Val(Equipo))) <> "" Then
                    Impre2 = ZDescri(Val(Equipo), ZLugar(Val(Equipo)))
                        Else
                    Impre2 = "."
                End If
                    Else
                Impre2 = Equipo
            End If
        
            Select Case LetraTemperatura
                Case "8"
                    Impre32 = Temperatura
                Case "12"
                    Impre33 = Temperatura
                Case "N8"
                    Impre34 = Temperatura
                Case "N10"
                    Impre35 = Temperatura
                Case "N12"
                    Impre36 = Temperatura
                Case "FS"
                    Impre37 = Temperatura
                Case "FO"
                    Impre38 = Temperatura
                Case "R"
                    Impre39 = Temperatura
                Case Else
                    Impre31 = Temperatura
            End Select
        
            Select Case LetraTiempo
                Case "8"
                    Impre42 = Tiempo
                Case "12"
                    Impre43 = Tiempo
                Case "N8"
                    Impre44 = Tiempo
                Case "N10"
                    Impre45 = Tiempo
                Case "N12"
                    Impre46 = Tiempo
                Case "FS"
                    Impre47 = Tiempo
                Case "FO"
                    Impre48 = Tiempo
                Case "R"
                    Impre49 = Tiempo
                Case Else
                    Impre41 = Tiempo
            End Select
        
            If Val(Control) <> 0 Then
                ZLugar(Val(Control)) = ZLugar(Val(Control)) + 1
                If ZDescri(Val(Control), ZLugar(Val(Control))) <> "" Then
                    LetraControl = "FS"
                    ZControl = ZDescri(Val(Control), ZLugar(Val(Control)))
                        Else
                    LetraControl = "FS"
                    ZControl = "."
                End If
                    Else
                ZControl = Control
            End If
        
            Select Case LetraControl
                Case "8"
                    Impre52 = ZControl
                Case "12"
                    Impre53 = ZControl
                Case "N8"
                    Impre54 = ZControl
                Case "N10"
                    Impre55 = ZControl
                Case "N12"
                    Impre56 = ZControl
                Case "FS"
                    Impre57 = ZControl
                Case "FO"
                    Impre58 = ZControl
                Case "R"
                    Impre59 = ZControl
                Case Else
                    Impre51 = ZControl
            End Select
        
            If Val(Seguridad) <> 0 Then
                ZLugar(Val(Seguridad)) = ZLugar(Val(Seguridad)) + 1
                If ZDescri(Val(Seguridad), ZLugar(Val(Seguridad))) <> "" Then
                    Impre6 = ZDescri(Val(Seguridad), ZLugar(Val(Seguridad)))
                        Else
                    Impre6 = "."
                End If
                    Else
                Impre6 = Seguridad
            End If
            
            WRenglon = WRenglon + 1
            Auxi = Str$(WRenglon)
            Call Ceros(Auxi, 2)
        
            WClave = Terminado.Text + Auxi
        
            XXVersion = Str$(Val(Version.Text) + 1)
            XXFechaVersion = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            XXAutorizado = "S"
            XXOrdFecha = Right$(XXFechaVersion, 4) + Mid$(XXFechaVersion, 4, 2) + Left$(XXFechaVersion, 2)
        
            Sql1 = "INSERT INTO CargaIV ("
            Sql2 = "Clave ,"
            Sql3 = "Terminado ,"
            Sql4 = "Renglon ,"
            Sql5 = "Fecha ,"
            Sql6 = "OrdFecha ,"
            Sql7 = "Lote ,"
            Sql8 = "Version ,"
            Sql9 = "Autorizado ,"
            Sql10 = "Etapa ,"
            Sql11 = "LetraInstrucciones ,"
            Sql12 = "Instrucciones ,"
            Sql13 = "Equipo ,"
            Sql14 = "LetraTemperatura ,"
            Sql15 = "Temperatura ,"
            Sql16 = "LetraTiempo ,"
            Sql17 = "Tiempo ,"
            Sql18 = "LetraControl ,"
            Sql19 = "Control ,"
            Sql20 = "Seguridad ,"
            Sql21 = "DesTerminado )"
            Sql22 = "Values ("
            Sql23 = "'" + WClave + "',"
            Sql24 = "'" + Terminado.Text + "',"
            Sql25 = "'" + Str$(WRenglon) + "',"
            Sql26 = "'" + XXFechaVersion + "',"
            Sql27 = "'" + XXOrdFecha + "',"
            Sql28 = "'" + ZLote + "',"
            Sql29 = "'" + XXVersion + "',"
            Sql30 = "'" + XXAutorizado + "',"
            Sql31 = "'" + Etapa + "',"
            Sql32 = "'" + LetraInstrucciones + "',"
            Sql33 = "'" + Instrucciones + "',"
            Sql34 = "'" + Equipo + "',"
            Sql35 = "'" + LetraTemperatura + "',"
            Sql36 = "'" + Temperatura + "',"
            Sql37 = "'" + LetraTiempo + "',"
            Sql38 = "'" + Tiempo + "',"
            Sql39 = "'" + LetraControl + "',"
            Sql40 = "'" + Control + "',"
            Sql41 = "'" + Seguridad + "',"
            Sql42 = "'" + DesTerminado.Caption + "')"
            
            rsCargaIV = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 _
                    + Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 _
                    + Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 _
                    + Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 _
                    + Sql41 + Sql42
            Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
        
            Sql1 = "UPDATE CargaIV SET "
            Sql2 = "Impre11 = " + "'" + IMPRE11 + "',"
            Sql3 = "Impre12 = " + "'" + Impre12 + "',"
            Sql4 = "Impre13 = " + "'" + Impre13 + "',"
            Sql5 = "Impre14 = " + "'" + Impre14 + "',"
            Sql6 = "Impre15 = " + "'" + Impre15 + "',"
            Sql7 = "Impre16 = " + "'" + Impre16 + "',"
            Sql8 = "Impre17 = " + "'" + Impre17 + "',"
            Sql9 = "Impre18 = " + "'" + Impre18 + "',"
            Sql10 = "Impre19 = " + "'" + Impre19 + "',"
            Sql11 = "Impre2 = " + "'" + Impre2 + "',"
            Sql12 = "Impre31 = " + "'" + Impre31 + "',"
            Sql13 = "Impre32 = " + "'" + Impre32 + "',"
            Sql14 = "Impre33 = " + "'" + Impre33 + "',"
            Sql15 = "Impre34 = " + "'" + Impre34 + "',"
            Sql16 = "Impre35 = " + "'" + Impre35 + "',"
            Sql17 = "Impre36 = " + "'" + Impre36 + "',"
            Sql18 = "Impre37 = " + "'" + Impre37 + "',"
            Sql19 = "Impre38 = " + "'" + Impre38 + "',"
            Sql20 = "Impre39 = " + "'" + Impre39 + "',"
            Sql21 = "Impre41 = " + "'" + Impre41 + "',"
            Sql22 = "Impre42 = " + "'" + Impre42 + "',"
            Sql23 = "Impre43 = " + "'" + Impre43 + "',"
            Sql24 = "Impre44 = " + "'" + Impre44 + "',"
            Sql25 = "Impre45 = " + "'" + Impre45 + "',"
            Sql26 = "Impre46 = " + "'" + Impre46 + "',"
            Sql27 = "Impre47 = " + "'" + Impre47 + "',"
            Sql28 = "Impre48 = " + "'" + Impre48 + "',"
            Sql29 = "Impre49 = " + "'" + Impre49 + "',"
            Sql30 = "Impre51 = " + "'" + Impre51 + "',"
            Sql31 = "Impre52 = " + "'" + Impre52 + "',"
            Sql32 = "Impre53 = " + "'" + Impre53 + "',"
            Sql33 = "Impre54 = " + "'" + Impre54 + "',"
            Sql34 = "Impre55 = " + "'" + Impre55 + "',"
            Sql35 = "Impre56 = " + "'" + Impre56 + "',"
            Sql36 = "Impre57 = " + "'" + Impre57 + "',"
            Sql37 = "Impre58 = " + "'" + Impre58 + "',"
            Sql38 = "Impre59 = " + "'" + Impre59 + "',"
            Sql39 = "Impre6 = " + "'" + Impre6 + "'"
            Sql40 = " Where Clave = " + "'" + WClave + "'"

            rsCargaIV = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 _
                   + Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 _
                   + Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 _
                   + Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40
            Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
            
        Next iRow



End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Ayuda.Visible = True
    Ayuda.Text = ""
    
    Select Case XIndice
        Case 0, 1, 2
            Sql1 = "Select *"
            Sql2 = " FROM EquipoFabrica"
            Sql3 = " Order by Codigo"
            spEquipoFabrica = Sql1 + Sql2 + Sql3
            Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEquipoFabrica.RecordCount > 0 Then
                With rstEquipoFabrica
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstEquipoFabrica!codigo) + " " + rstEquipoFabrica!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = Str$(rstEquipoFabrica!codigo)
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEquipoFabrica.Close
            End If
            
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub cmdClose1_Click()

    Call Limpia_Click
    PrgCargaIV.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()
               
    If Trim(Metodo.Text) = "" Then
        m$ = "Se debe informar el metodo de seguridad"
        A% = MsgBox(m$, 0, "Ingreso de Procesos de Fabricacion")
        Exit Sub
    End If
    
    If Trim(ControlCambio.Text) = "" Then
        m$ = "Se debe informar el campo Control de Cambios"
        A% = MsgBox(m$, 0, "Ingreso de Procesos de Fabricacion")
        Exit Sub
    End If

    If WGraba <> "S" Then
    
        Call Ingresa_clave

               Else
               
        Terminado.Text = UCase(Terminado.Text)
        
        Erase ZTraspaso
        LugarTraspaso = 0
        
        Sql1 = "Select *"
        Sql2 = " FROM CargaIV"
        Sql3 = " Where CargaIV.Terminado = " + "'" + Terminado.Text + "'"
        Sql4 = " Order by CargaIV.Clave"
    
        rsCargaIV = Sql1 + Sql2 + Sql3 + Sql4
        Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaIV.RecordCount > 0 Then
            With rstCargaIV
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        LugarTraspaso = LugarTraspaso + 1
                        
                        ZTraspaso(LugarTraspaso, 1) = rstCargaIV!Etapa
                        ZTraspaso(LugarTraspaso, 2) = rstCargaIV!LetraInstrucciones
                        ZTraspaso(LugarTraspaso, 3) = rstCargaIV!Instrucciones
                        ZTraspaso(LugarTraspaso, 4) = rstCargaIV!Equipo
                        ZTraspaso(LugarTraspaso, 5) = rstCargaIV!LetraTemperatura
                        ZTraspaso(LugarTraspaso, 6) = rstCargaIV!Temperatura
                        ZTraspaso(LugarTraspaso, 7) = rstCargaIV!LetraTiempo
                        ZTraspaso(LugarTraspaso, 8) = rstCargaIV!Tiempo
                        ZTraspaso(LugarTraspaso, 9) = rstCargaIV!LetraControl
                        ZTraspaso(LugarTraspaso, 10) = rstCargaIV!Control
                        ZTraspaso(LugarTraspaso, 11) = rstCargaIV!Seguridad
                        
                
                        ZFecha = rstCargaIV!Fecha
                        ZVersion = rstCargaIV!Version
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCargaIV.Close
        End If
        
        For Ciclo = 1 To LugarTraspaso
        
            ZEtapa = ZTraspaso(Ciclo, 1)
            ZLetraInstrucciones = ZTraspaso(Ciclo, 2)
            ZInstrucciones = ZTraspaso(Ciclo, 3)
            ZEquipo = ZTraspaso(Ciclo, 4)
            ZLetraTemperatura = ZTraspaso(Ciclo, 5)
            ZTemperatura = ZTraspaso(Ciclo, 6)
            ZLetraTiempo = ZTraspaso(Ciclo, 7)
            ZTiempo = ZTraspaso(Ciclo, 8)
            ZLetraControl = ZTraspaso(Ciclo, 9)
            ZControl = ZTraspaso(Ciclo, 10)
            ZSeguridad = ZTraspaso(Ciclo, 11)
            
            ZFechaInicio = ZFecha
            ZFechaFinal = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            
            WVersion = Str$(ZVersion)
            ZControlCambio = ControlCambio.Text
            
            WRenglon = Str$(Ciclo)
            Call Ceros(WRenglon, 2)
                    
            Call Ceros(WVersion, 4)
            ZClave = Terminado.Text + WVersion + WRenglon
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaIVVersion ("
            ZSql = ZSql + "Clave, "
            ZSql = ZSql + "Terminado, "
            ZSql = ZSql + "Version, "
            ZSql = ZSql + "Renglon, "
            ZSql = ZSql + "FechaInicio, "
            ZSql = ZSql + "FechaFinal, "
            ZSql = ZSql + "ControlCambio, "
            ZSql = ZSql + "Etapa, "
            ZSql = ZSql + "LetraInstrucciones, "
            ZSql = ZSql + "Instrucciones, "
            ZSql = ZSql + "Equipo, "
            ZSql = ZSql + "LetraTemperatura, "
            ZSql = ZSql + "Temperatura, "
            ZSql = ZSql + "LetraTiempo, "
            ZSql = ZSql + "Tiempo, "
            ZSql = ZSql + "LetraControl, "
            ZSql = ZSql + "Control, "
            ZSql = ZSql + "Seguridad) "
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZClave + "',"
            ZSql = ZSql + "'" + Terminado.Text + "',"
            ZSql = ZSql + "'" + WVersion + "',"
            ZSql = ZSql + "'" + WRenglon + "',"
            ZSql = ZSql + "'" + ZFechaInicio + "',"
            ZSql = ZSql + "'" + ZFechaFinal + "',"
            ZSql = ZSql + "'" + ZControlCambio + "',"
            ZSql = ZSql + "'" + ZEtapa + "',"
            ZSql = ZSql + "'" + ZLetraInstrucciones + "',"
            ZSql = ZSql + "'" + ZInstrucciones + "',"
            ZSql = ZSql + "'" + ZEquipo + "',"
            ZSql = ZSql + "'" + ZLetraTemperatura + "',"
            ZSql = ZSql + "'" + ZTemperatura + "',"
            ZSql = ZSql + "'" + ZLetraTiempo + "',"
            ZSql = ZSql + "'" + ZTiempo + "',"
            ZSql = ZSql + "'" + ZLetraControl + "',"
            ZSql = ZSql + "'" + ZControl + "',"
            ZSql = ZSql + "'" + ZSeguridad + "')"
            
            spCargaIVVersion = ZSql
            Set rstCargaIVVersion = db.OpenRecordset(spCargaIVVersion, dbOpenSnapshot, dbSQLPassThrough)
        
        Next Ciclo
        
        
        

        Sql1 = "DELETE CargaIV"
        Sql2 = " Where Terminado = " + "'" + Terminado.Text + "'"
        rsCargaIV = Sql1 + Sql2
        Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
    
        Erase ZLugar
        Erase ZDescri
       
      
        Sql1 = "Select *"
        Sql2 = " FROM EquipoFabrica"
        Sql3 = " Order by Codigo"
        spEquipoFabrica = Sql1 + Sql2 + Sql3
        Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEquipoFabrica.RecordCount > 0 Then
            With rstEquipoFabrica
                .MoveFirst
                Do
                    If .EOF = False Then
                        
                        ZDescripcion = IIf(IsNull(rstEquipoFabrica!Descripcion), "", rstEquipoFabrica!Descripcion)
                        ZDescripcionII = IIf(IsNull(rstEquipoFabrica!DescripcionII), "", rstEquipoFabrica!DescripcionII)
                        ZDescripcionIII = IIf(IsNull(rstEquipoFabrica!DescripcionIII), "", rstEquipoFabrica!DescripcionIII)
                    
                        WDescripcion = Trim(ZDescripcion) + " " + Trim(ZDescripcionII) + " " + Trim(ZDescripcionIII)
                        ZCodigo = rstEquipoFabrica!codigo
                    
                        ZHasta = Len(WDescripcion)
                        Desde = 1
                        
                        Do
                            Hasta = Desde + 15
                            If Hasta > ZHasta Then
                                Hasta = ZHasta
                            End If
                            
                            ZLugar(ZCodigo) = ZLugar(ZCodigo) + 1
                            aa = Mid(WDescripcion, Desde, Hasta)
                            ZDescri(ZCodigo, ZLugar(ZCodigo)) = Mid(WDescripcion, Desde, Hasta)
                            For Cicla = Hasta To Desde Step -1
                                aa = Mid(WDescripcion, Cicla, 1)
                                If Mid(WDescripcion, Cicla, 1) = Space(1) Then
                                    aa = Mid(WDescripcion, Desde, Cicla - Desde)
                                    ZDescri(ZCodigo, ZLugar(ZCodigo)) = Mid(WDescripcion, Desde, Cicla - Desde)
                                    Desde = Cicla + 1
                                    Exit For
                                End If
                            Next Cicla
                        
                            If Hasta >= ZHasta Then
                                Exit Do
                            End If
                        Loop
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstEquipoFabrica.Close
        End If
    
        HastaRenglon = 0
        For iRow = 100 To 1 Step -1
        
            Etapa = WVector1.TextMatrix(iRow, 1)
            LetraInstrucciones = WVector1.TextMatrix(iRow, 2)
            Instrucciones = WVector1.TextMatrix(iRow, 3)
            Equipo = WVector1.TextMatrix(iRow, 4)
            LetraTemperatura = WVector1.TextMatrix(iRow, 5)
            Temperatura = WVector1.TextMatrix(iRow, 6)
            LetraTiempo = WVector1.TextMatrix(iRow, 7)
            Tiempo = WVector1.TextMatrix(iRow, 8)
            LetraControl = WVector1.TextMatrix(iRow, 9)
            Control = WVector1.TextMatrix(iRow, 10)
            Seguridad = WVector1.TextMatrix(iRow, 11)
            
            If Etapa <> "" Or Instrucciones <> "" Or Equipo <> "" Or Temperatura <> "" Or Tiempo <> "" Or Control <> "" Or Seguridad <> "" Then
                HastaRenglon = iRow
                Exit For
            End If
            
        Next iRow
    
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        Erase ZLugar
        XEquipo = ""
        XControl = ""
        XSeguridad = ""

        WRenglon = 0
        For iRow = 1 To HastaRenglon
    
            ZLote = ""
        
            Etapa = WVector1.TextMatrix(iRow, 1)
            LetraInstrucciones = WVector1.TextMatrix(iRow, 2)
            Instrucciones = WVector1.TextMatrix(iRow, 3)
            Equipo = WVector1.TextMatrix(iRow, 4)
            LetraTemperatura = WVector1.TextMatrix(iRow, 5)
            Temperatura = WVector1.TextMatrix(iRow, 6)
            LetraTiempo = WVector1.TextMatrix(iRow, 7)
            Tiempo = WVector1.TextMatrix(iRow, 8)
            LetraControl = WVector1.TextMatrix(iRow, 9)
            Control = WVector1.TextMatrix(iRow, 10)
            Seguridad = WVector1.TextMatrix(iRow, 11)
        
            IMPRE11 = ""
            Impre12 = ""
            Impre13 = ""
            Impre14 = ""
            Impre15 = ""
            Impre16 = ""
            Impre17 = ""
            Impre18 = ""
            Impre19 = ""
        
            Impre31 = ""
            Impre32 = ""
            Impre33 = ""
            Impre34 = ""
            Impre35 = ""
            Impre36 = ""
            Impre37 = ""
            Impre38 = ""
            Impre39 = ""
        
            Impre41 = ""
            Impre42 = ""
            Impre43 = ""
            Impre44 = ""
            Impre45 = ""
            Impre46 = ""
            Impre47 = ""
            Impre48 = ""
            Impre49 = ""
        
            Impre51 = ""
            Impre52 = ""
            Impre53 = ""
            Impre54 = ""
            Impre55 = ""
            Impre56 = ""
            Impre57 = ""
            Impre58 = ""
            Impre59 = ""
        
            Select Case LetraInstrucciones
                Case "8"
                    Impre12 = Instrucciones
                Case "12"
                    Impre13 = Instrucciones
                Case "N8"
                    Impre14 = Instrucciones
                Case "N10"
                    Impre15 = Instrucciones
                Case "N12"
                    Impre16 = Instrucciones
                Case "FS"
                    Impre17 = Instrucciones
                Case "FO"
                    Impre18 = Instrucciones
                Case "R"
                    Impre19 = Instrucciones
                Case Else
                    IMPRE11 = Instrucciones
            End Select
            
            If Val(Equipo) <> Val(XEquipo) Then
                ZLugar(Val(Equipo)) = 0
            End If
            If Val(Control) <> Val(XControl) Then
                ZLugar(Val(Control)) = 0
            End If
            If Val(Seguridad) <> Val(XSeguridad) Then
                ZLugar(Val(Seguridad)) = 0
            End If
            
            If Val(Equipo) <> 0 Then
                ZLugar(Val(Equipo)) = ZLugar(Val(Equipo)) + 1
                If ZDescri(Val(Equipo), ZLugar(Val(Equipo))) <> "" Then
                    Impre2 = ZDescri(Val(Equipo), ZLugar(Val(Equipo)))
                        Else
                    Impre2 = "."
                End If
                    Else
                Impre2 = Equipo
            End If
            
            Select Case LetraTemperatura
                Case "8"
                    Impre32 = Temperatura
                Case "12"
                    Impre33 = Temperatura
                Case "N8"
                    Impre34 = Temperatura
                Case "N10"
                    Impre35 = Temperatura
                Case "N12"
                    Impre36 = Temperatura
                Case "FS"
                    Impre37 = Temperatura
                Case "FO"
                    Impre38 = Temperatura
                Case "R"
                    Impre39 = Temperatura
                Case Else
                    Impre31 = Temperatura
            End Select
        
            Select Case LetraTiempo
                Case "8"
                    Impre42 = Tiempo
                Case "12"
                    Impre43 = Tiempo
                Case "N8"
                    Impre44 = Tiempo
                Case "N10"
                    Impre45 = Tiempo
                Case "N12"
                    Impre46 = Tiempo
                Case "FS"
                    Impre47 = Tiempo
                Case "FO"
                    Impre48 = Tiempo
                Case "R"
                    Impre49 = Tiempo
                Case Else
                    Impre41 = Tiempo
            End Select
        
            If Val(Control) <> 0 Then
                ZLugar(Val(Control)) = ZLugar(Val(Control)) + 1
                If ZDescri(Val(Control), ZLugar(Val(Control))) <> "" Then
                    LetraControl = "FS"
                    ZControl = ZDescri(Val(Control), ZLugar(Val(Control)))
                        Else
                    LetraControl = "FS"
                    ZControl = "."
                End If
                    Else
                ZControl = Control
            End If
        
            Select Case LetraControl
                Case "8"
                    Impre52 = ZControl
                Case "12"
                    Impre53 = ZControl
                Case "N8"
                    Impre54 = ZControl
                Case "N10"
                    Impre55 = ZControl
                Case "N12"
                    Impre56 = ZControl
                Case "FS"
                    Impre57 = ZControl
                Case "FO"
                    Impre58 = ZControl
                Case "R"
                    Impre59 = ZControl
                Case Else
                    Impre51 = ZControl
            End Select
        
            If Val(Seguridad) <> 0 Then
                ZLugar(Val(Seguridad)) = ZLugar(Val(Seguridad)) + 1
                If ZDescri(Val(Seguridad), ZLugar(Val(Seguridad))) <> "" Then
                    Impre6 = ZDescri(Val(Seguridad), ZLugar(Val(Seguridad)))
                        Else
                    Impre6 = "."
                End If
                    Else
                Impre6 = Seguridad
            End If
            
                
            XEquipo = Equipo
            XControl = Control
            XSeguridad = Seguridad
            XControlCambio = ControlCambio.Text
            
            
            WRenglon = WRenglon + 1
            Auxi = Str$(WRenglon)
            Call Ceros(Auxi, 2)
        
            WClave = Terminado.Text + Auxi
        
            XXVersion = Str$(Val(Version.Text) + 1)
            XXFechaVersion = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            XXAutorizado = "S"
            XXOrdFecha = Right$(XXFechaVersion, 4) + Mid$(XXFechaVersion, 4, 2) + Left$(XXFechaVersion, 2)
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaIV ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Terminado ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "ControlCambio ,"
            ZSql = ZSql + "Lote ,"
            ZSql = ZSql + "Version ,"
            ZSql = ZSql + "Autorizado ,"
            ZSql = ZSql + "Etapa ,"
            ZSql = ZSql + "LetraInstrucciones ,"
            ZSql = ZSql + "Instrucciones ,"
            ZSql = ZSql + "Equipo ,"
            ZSql = ZSql + "LetraTemperatura ,"
            ZSql = ZSql + "Temperatura ,"
            ZSql = ZSql + "LetraTiempo ,"
            ZSql = ZSql + "Tiempo ,"
            ZSql = ZSql + "LetraControl ,"
            ZSql = ZSql + "Control ,"
            ZSql = ZSql + "Seguridad ,"
            ZSql = ZSql + "DesTerminado )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Terminado.Text + "',"
            ZSql = ZSql + "'" + Str$(WRenglon) + "',"
            ZSql = ZSql + "'" + XXFechaVersion + "',"
            ZSql = ZSql + "'" + XXOrdFecha + "',"
            ZSql = ZSql + "'" + XControlCambio + "',"
            ZSql = ZSql + "'" + ZLote + "',"
            ZSql = ZSql + "'" + XXVersion + "',"
            ZSql = ZSql + "'" + XXAutorizado + "',"
            ZSql = ZSql + "'" + Etapa + "',"
            ZSql = ZSql + "'" + LetraInstrucciones + "',"
            ZSql = ZSql + "'" + Instrucciones + "',"
            ZSql = ZSql + "'" + Equipo + "',"
            ZSql = ZSql + "'" + LetraTemperatura + "',"
            ZSql = ZSql + "'" + Temperatura + "',"
            ZSql = ZSql + "'" + LetraTiempo + "',"
            ZSql = ZSql + "'" + Tiempo + "',"
            ZSql = ZSql + "'" + LetraControl + "',"
            ZSql = ZSql + "'" + Control + "',"
            ZSql = ZSql + "'" + Seguridad + "',"
            ZSql = ZSql + "'" + DesTerminado.Caption + "')"
            
            rsCargaIV = ZSql
            Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
        
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaIV SET "
            ZSql = ZSql + "Impre11 = " + "'" + IMPRE11 + "',"
            ZSql = ZSql + "Impre12 = " + "'" + Impre12 + "',"
            ZSql = ZSql + "Impre13 = " + "'" + Impre13 + "',"
            ZSql = ZSql + "Impre14 = " + "'" + Impre14 + "',"
            ZSql = ZSql + "Impre15 = " + "'" + Impre15 + "',"
            ZSql = ZSql + "Impre16 = " + "'" + Impre16 + "',"
            ZSql = ZSql + "Impre17 = " + "'" + Impre17 + "',"
            ZSql = ZSql + "Impre18 = " + "'" + Impre18 + "',"
            ZSql = ZSql + "Impre19 = " + "'" + Impre19 + "',"
            ZSql = ZSql + "Impre2 = " + "'" + Impre2 + "',"
            ZSql = ZSql + "Impre31 = " + "'" + Impre31 + "',"
            ZSql = ZSql + "Impre32 = " + "'" + Impre32 + "',"
            ZSql = ZSql + "Impre33 = " + "'" + Impre33 + "',"
            ZSql = ZSql + "Impre34 = " + "'" + Impre34 + "',"
            ZSql = ZSql + "Impre35 = " + "'" + Impre35 + "',"
            ZSql = ZSql + "Impre36 = " + "'" + Impre36 + "',"
            ZSql = ZSql + "Impre37 = " + "'" + Impre37 + "',"
            ZSql = ZSql + "Impre38 = " + "'" + Impre38 + "',"
            ZSql = ZSql + "Impre39 = " + "'" + Impre39 + "',"
            ZSql = ZSql + "Impre41 = " + "'" + Impre41 + "',"
            ZSql = ZSql + "Impre42 = " + "'" + Impre42 + "',"
            ZSql = ZSql + "Impre43 = " + "'" + Impre43 + "',"
            ZSql = ZSql + "Impre44 = " + "'" + Impre44 + "',"
            ZSql = ZSql + "Impre45 = " + "'" + Impre45 + "',"
            ZSql = ZSql + "Impre46 = " + "'" + Impre46 + "',"
            ZSql = ZSql + "Impre47 = " + "'" + Impre47 + "',"
            ZSql = ZSql + "Impre48 = " + "'" + Impre48 + "',"
            ZSql = ZSql + "Impre49 = " + "'" + Impre49 + "',"
            ZSql = ZSql + "Impre51 = " + "'" + Impre51 + "',"
            ZSql = ZSql + "Impre52 = " + "'" + Impre52 + "',"
            ZSql = ZSql + "Impre53 = " + "'" + Impre53 + "',"
            ZSql = ZSql + "Impre54 = " + "'" + Impre54 + "',"
            ZSql = ZSql + "Impre55 = " + "'" + Impre55 + "',"
            ZSql = ZSql + "Impre56 = " + "'" + Impre56 + "',"
            ZSql = ZSql + "Impre57 = " + "'" + Impre57 + "',"
            ZSql = ZSql + "Impre58 = " + "'" + Impre58 + "',"
            ZSql = ZSql + "Impre59 = " + "'" + Impre59 + "',"
            ZSql = ZSql + "Impre6 = " + "'" + Impre6 + "'"
            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"

            rsCargaIV = ZSql
            Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
            
        Next iRow
        
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CargaIV SET "
        ZSql = ZSql + " Operador = " + "'" + ZOperador + "'"
        ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
                            
        spCargaIV = ZSql
        Set rstCargaIV = db.OpenRecordset(spCargaIV, dbOpenSnapshot, dbSQLPassThrough)
        
    
        XEmpresa = WEmpresa
        Erase CargaEmpresa
        
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7
                CargaEmpresa(1, 1) = "0001"
                CargaEmpresa(1, 2) = "Empresa01"
                CargaEmpresa(2, 1) = "0003"
                CargaEmpresa(2, 2) = "Empresa03"
                CargaEmpresa(3, 1) = "0005"
                CargaEmpresa(3, 2) = "Empresa05"
                CargaEmpresa(4, 1) = "0006"
                CargaEmpresa(4, 2) = "Empresa06"
                CargaEmpresa(5, 1) = "0007"
                CargaEmpresa(5, 2) = "Empresa07"
            Case 2, 4, 8, 9
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
                CargaEmpresa(4, 1) = "0009"
                CargaEmpresa(4, 2) = "Empresa09"
            Case 10
                CargaEmpresa(1, 1) = "0010"
                CargaEmpresa(1, 2) = "Empresa10"
            Case Else
        End Select
                
        For Cicla = 1 To 5
            If CargaEmpresa(Cicla, 1) <> "" Then
            
                WEmpresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                ZSql = ""
                ZSql = ZSql + "UPDATE Terminado SET "
                ZSql = ZSql + " Metodo = " + "'" + Metodo.Text + "',"
                ZSql = ZSql + " VersionI = " + "'" + XXVersion + "',"
                ZSql = ZSql + " FechaVersionI = " + "'" + XXFechaVersion + "',"
                ZSql = ZSql + " EstadoI = " + "'" + "S" + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + Terminado.Text + "'"
                    
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
        Next Cicla
        
        Select Case Val(XEmpresa)
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
            Case Else
        End Select
    
        Call Limpia_Click

        WVector1.Col = 1
        WVector1.Row = 1
        
        Terminado.SetFocus
        
    End If
        
End Sub



Private Sub GrabaAuto_Click()
               
    If Trim(Metodo.Text) = "" Then
        m$ = "Se debe informar el metodo de seguridad"
        A% = MsgBox(m$, 0, "Ingreso de Procesos de Fabricacion")
        Exit Sub
    End If
    
    If Trim(ControlCambio.Text) = "" Then
        m$ = "Se debe informar el campo Control de Cambios"
        A% = MsgBox(m$, 0, "Ingreso de Procesos de Fabricacion")
        Exit Sub
    End If

           
    Terminado.Text = UCase(Terminado.Text)
    
    Erase ZTraspaso
    LugarTraspaso = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM CargaIV"
    Sql3 = " Where CargaIV.Terminado = " + "'" + Terminado.Text + "'"
    Sql4 = " Order by CargaIV.Clave"

    rsCargaIV = Sql1 + Sql2 + Sql3 + Sql4
    Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaIV.RecordCount > 0 Then
        With rstCargaIV
            .MoveFirst
            Do
                If .EOF = False Then
                
                    LugarTraspaso = LugarTraspaso + 1
                    
                    ZTraspaso(LugarTraspaso, 1) = rstCargaIV!Etapa
                    ZTraspaso(LugarTraspaso, 2) = rstCargaIV!LetraInstrucciones
                    ZTraspaso(LugarTraspaso, 3) = rstCargaIV!Instrucciones
                    ZTraspaso(LugarTraspaso, 4) = rstCargaIV!Equipo
                    ZTraspaso(LugarTraspaso, 5) = rstCargaIV!LetraTemperatura
                    ZTraspaso(LugarTraspaso, 6) = rstCargaIV!Temperatura
                    ZTraspaso(LugarTraspaso, 7) = rstCargaIV!LetraTiempo
                    ZTraspaso(LugarTraspaso, 8) = rstCargaIV!Tiempo
                    ZTraspaso(LugarTraspaso, 9) = rstCargaIV!LetraControl
                    ZTraspaso(LugarTraspaso, 10) = rstCargaIV!Control
                    ZTraspaso(LugarTraspaso, 11) = rstCargaIV!Seguridad
                    
            
                    ZFecha = rstCargaIV!Fecha
                    ZVersion = rstCargaIV!Version
            
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaIV.Close
    End If
    
    For Ciclo = 1 To LugarTraspaso
    
        ZEtapa = ZTraspaso(Ciclo, 1)
        ZLetraInstrucciones = ZTraspaso(Ciclo, 2)
        ZInstrucciones = ZTraspaso(Ciclo, 3)
        ZEquipo = ZTraspaso(Ciclo, 4)
        ZLetraTemperatura = ZTraspaso(Ciclo, 5)
        ZTemperatura = ZTraspaso(Ciclo, 6)
        ZLetraTiempo = ZTraspaso(Ciclo, 7)
        ZTiempo = ZTraspaso(Ciclo, 8)
        ZLetraControl = ZTraspaso(Ciclo, 9)
        ZControl = ZTraspaso(Ciclo, 10)
        ZSeguridad = ZTraspaso(Ciclo, 11)
        
        ZFechaInicio = ZFecha
        ZFechaFinal = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        
        WVersion = Str$(ZVersion)
        ZControlCambio = ControlCambio.Text
        
        WRenglon = Str$(Ciclo)
        Call Ceros(WRenglon, 2)
                
        Call Ceros(WVersion, 4)
        ZClave = Terminado.Text + WVersion + WRenglon
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaIVVersion ("
        ZSql = ZSql + "Clave, "
        ZSql = ZSql + "Terminado, "
        ZSql = ZSql + "Version, "
        ZSql = ZSql + "Renglon, "
        ZSql = ZSql + "FechaInicio, "
        ZSql = ZSql + "FechaFinal, "
        ZSql = ZSql + "ControlCambio, "
        ZSql = ZSql + "Etapa, "
        ZSql = ZSql + "LetraInstrucciones, "
        ZSql = ZSql + "Instrucciones, "
        ZSql = ZSql + "Equipo, "
        ZSql = ZSql + "LetraTemperatura, "
        ZSql = ZSql + "Temperatura, "
        ZSql = ZSql + "LetraTiempo, "
        ZSql = ZSql + "Tiempo, "
        ZSql = ZSql + "LetraControl, "
        ZSql = ZSql + "Control, "
        ZSql = ZSql + "Seguridad) "
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZClave + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + WVersion + "',"
        ZSql = ZSql + "'" + WRenglon + "',"
        ZSql = ZSql + "'" + ZFechaInicio + "',"
        ZSql = ZSql + "'" + ZFechaFinal + "',"
        ZSql = ZSql + "'" + ZControlCambio + "',"
        ZSql = ZSql + "'" + ZEtapa + "',"
        ZSql = ZSql + "'" + ZLetraInstrucciones + "',"
        ZSql = ZSql + "'" + ZInstrucciones + "',"
        ZSql = ZSql + "'" + ZEquipo + "',"
        ZSql = ZSql + "'" + ZLetraTemperatura + "',"
        ZSql = ZSql + "'" + ZTemperatura + "',"
        ZSql = ZSql + "'" + ZLetraTiempo + "',"
        ZSql = ZSql + "'" + ZTiempo + "',"
        ZSql = ZSql + "'" + ZLetraControl + "',"
        ZSql = ZSql + "'" + ZControl + "',"
        ZSql = ZSql + "'" + ZSeguridad + "')"
        
        spCargaIVVersion = ZSql
        Set rstCargaIVVersion = db.OpenRecordset(spCargaIVVersion, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
    

    Sql1 = "DELETE CargaIV"
    Sql2 = " Where Terminado = " + "'" + Terminado.Text + "'"
    rsCargaIV = Sql1 + Sql2
    Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)

    Erase ZLugar
    Erase ZDescri
   
  
    Sql1 = "Select *"
    Sql2 = " FROM EquipoFabrica"
    Sql3 = " Order by Codigo"
    spEquipoFabrica = Sql1 + Sql2 + Sql3
    Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipoFabrica.RecordCount > 0 Then
        With rstEquipoFabrica
            .MoveFirst
            Do
                If .EOF = False Then
                    
                    ZDescripcion = IIf(IsNull(rstEquipoFabrica!Descripcion), "", rstEquipoFabrica!Descripcion)
                    ZDescripcionII = IIf(IsNull(rstEquipoFabrica!DescripcionII), "", rstEquipoFabrica!DescripcionII)
                    ZDescripcionIII = IIf(IsNull(rstEquipoFabrica!DescripcionIII), "", rstEquipoFabrica!DescripcionIII)
                
                    WDescripcion = Trim(ZDescripcion) + " " + Trim(ZDescripcionII) + " " + Trim(ZDescripcionIII)
                    ZCodigo = rstEquipoFabrica!codigo
                
                    ZHasta = Len(WDescripcion)
                    Desde = 1
                    
                    Do
                        Hasta = Desde + 15
                        If Hasta > ZHasta Then
                            Hasta = ZHasta
                        End If
                        
                        ZLugar(ZCodigo) = ZLugar(ZCodigo) + 1
                        aa = Mid(WDescripcion, Desde, Hasta)
                        ZDescri(ZCodigo, ZLugar(ZCodigo)) = Mid(WDescripcion, Desde, Hasta)
                        For Cicla = Hasta To Desde Step -1
                            aa = Mid(WDescripcion, Cicla, 1)
                            If Mid(WDescripcion, Cicla, 1) = Space(1) Then
                                aa = Mid(WDescripcion, Desde, Cicla - Desde)
                                ZDescri(ZCodigo, ZLugar(ZCodigo)) = Mid(WDescripcion, Desde, Cicla - Desde)
                                Desde = Cicla + 1
                                Exit For
                            End If
                        Next Cicla
                    
                        If Hasta >= ZHasta Then
                            Exit Do
                        End If
                    Loop
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEquipoFabrica.Close
    End If

    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
    
        Etapa = WVector1.TextMatrix(iRow, 1)
        LetraInstrucciones = WVector1.TextMatrix(iRow, 2)
        Instrucciones = WVector1.TextMatrix(iRow, 3)
        Equipo = WVector1.TextMatrix(iRow, 4)
        LetraTemperatura = WVector1.TextMatrix(iRow, 5)
        Temperatura = WVector1.TextMatrix(iRow, 6)
        LetraTiempo = WVector1.TextMatrix(iRow, 7)
        Tiempo = WVector1.TextMatrix(iRow, 8)
        LetraControl = WVector1.TextMatrix(iRow, 9)
        Control = WVector1.TextMatrix(iRow, 10)
        Seguridad = WVector1.TextMatrix(iRow, 11)
        
        If Etapa <> "" Or Instrucciones <> "" Or Equipo <> "" Or Temperatura <> "" Or Tiempo <> "" Or Control <> "" Or Seguridad <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
        
    Next iRow

    WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    Erase ZLugar
    XEquipo = ""
    XControl = ""
    XSeguridad = ""

    WRenglon = 0
    For iRow = 1 To HastaRenglon

        ZLote = ""
    
        Etapa = WVector1.TextMatrix(iRow, 1)
        LetraInstrucciones = WVector1.TextMatrix(iRow, 2)
        Instrucciones = WVector1.TextMatrix(iRow, 3)
        Equipo = WVector1.TextMatrix(iRow, 4)
        LetraTemperatura = WVector1.TextMatrix(iRow, 5)
        Temperatura = WVector1.TextMatrix(iRow, 6)
        LetraTiempo = WVector1.TextMatrix(iRow, 7)
        Tiempo = WVector1.TextMatrix(iRow, 8)
        LetraControl = WVector1.TextMatrix(iRow, 9)
        Control = WVector1.TextMatrix(iRow, 10)
        Seguridad = WVector1.TextMatrix(iRow, 11)
    
        IMPRE11 = ""
        Impre12 = ""
        Impre13 = ""
        Impre14 = ""
        Impre15 = ""
        Impre16 = ""
        Impre17 = ""
        Impre18 = ""
        Impre19 = ""
    
        Impre31 = ""
        Impre32 = ""
        Impre33 = ""
        Impre34 = ""
        Impre35 = ""
        Impre36 = ""
        Impre37 = ""
        Impre38 = ""
        Impre39 = ""
    
        Impre41 = ""
        Impre42 = ""
        Impre43 = ""
        Impre44 = ""
        Impre45 = ""
        Impre46 = ""
        Impre47 = ""
        Impre48 = ""
        Impre49 = ""
    
        Impre51 = ""
        Impre52 = ""
        Impre53 = ""
        Impre54 = ""
        Impre55 = ""
        Impre56 = ""
        Impre57 = ""
        Impre58 = ""
        Impre59 = ""
    
        Select Case LetraInstrucciones
            Case "8"
                Impre12 = Instrucciones
            Case "12"
                Impre13 = Instrucciones
            Case "N8"
                Impre14 = Instrucciones
            Case "N10"
                Impre15 = Instrucciones
            Case "N12"
                Impre16 = Instrucciones
            Case "FS"
                Impre17 = Instrucciones
            Case "FO"
                Impre18 = Instrucciones
            Case "R"
                Impre19 = Instrucciones
            Case Else
                IMPRE11 = Instrucciones
        End Select
        
        If Val(Equipo) <> Val(XEquipo) Then
            ZLugar(Val(Equipo)) = 0
        End If
        If Val(Control) <> Val(XControl) Then
            ZLugar(Val(Control)) = 0
        End If
        If Val(Seguridad) <> Val(XSeguridad) Then
            ZLugar(Val(Seguridad)) = 0
        End If
        
        If Val(Equipo) <> 0 Then
            ZLugar(Val(Equipo)) = ZLugar(Val(Equipo)) + 1
            If ZDescri(Val(Equipo), ZLugar(Val(Equipo))) <> "" Then
                Impre2 = ZDescri(Val(Equipo), ZLugar(Val(Equipo)))
                    Else
                Impre2 = "."
            End If
                Else
            Impre2 = Equipo
        End If
        
        Select Case LetraTemperatura
            Case "8"
                Impre32 = Temperatura
            Case "12"
                Impre33 = Temperatura
            Case "N8"
                Impre34 = Temperatura
            Case "N10"
                Impre35 = Temperatura
            Case "N12"
                Impre36 = Temperatura
            Case "FS"
                Impre37 = Temperatura
            Case "FO"
                Impre38 = Temperatura
            Case "R"
                Impre39 = Temperatura
            Case Else
                Impre31 = Temperatura
        End Select
    
        Select Case LetraTiempo
            Case "8"
                Impre42 = Tiempo
            Case "12"
                Impre43 = Tiempo
            Case "N8"
                Impre44 = Tiempo
            Case "N10"
                Impre45 = Tiempo
            Case "N12"
                Impre46 = Tiempo
            Case "FS"
                Impre47 = Tiempo
            Case "FO"
                Impre48 = Tiempo
            Case "R"
                Impre49 = Tiempo
            Case Else
                Impre41 = Tiempo
        End Select
    
        If Val(Control) <> 0 Then
            ZLugar(Val(Control)) = ZLugar(Val(Control)) + 1
            If ZDescri(Val(Control), ZLugar(Val(Control))) <> "" Then
                LetraControl = "FS"
                ZControl = ZDescri(Val(Control), ZLugar(Val(Control)))
                    Else
                LetraControl = "FS"
                ZControl = "."
            End If
                Else
            ZControl = Control
        End If
    
        Select Case LetraControl
            Case "8"
                Impre52 = ZControl
            Case "12"
                Impre53 = ZControl
            Case "N8"
                Impre54 = ZControl
            Case "N10"
                Impre55 = ZControl
            Case "N12"
                Impre56 = ZControl
            Case "FS"
                Impre57 = ZControl
            Case "FO"
                Impre58 = ZControl
            Case "R"
                Impre59 = ZControl
            Case Else
                Impre51 = ZControl
        End Select
    
        If Val(Seguridad) <> 0 Then
            ZLugar(Val(Seguridad)) = ZLugar(Val(Seguridad)) + 1
            If ZDescri(Val(Seguridad), ZLugar(Val(Seguridad))) <> "" Then
                Impre6 = ZDescri(Val(Seguridad), ZLugar(Val(Seguridad)))
                    Else
                Impre6 = "."
            End If
                Else
            Impre6 = Seguridad
        End If
        
            
        XEquipo = Equipo
        XControl = Control
        XSeguridad = Seguridad
        XControlCambio = ControlCambio.Text
        
        
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
    
        WClave = Terminado.Text + Auxi
    
        XXVersion = Str$(Val(Version.Text) + 1)
        XXFechaVersion = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        XXAutorizado = "S"
        XXOrdFecha = Right$(XXFechaVersion, 4) + Mid$(XXFechaVersion, 4, 2) + Left$(XXFechaVersion, 2)
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaIV ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "OrdFecha ,"
        ZSql = ZSql + "ControlCambio ,"
        ZSql = ZSql + "Lote ,"
        ZSql = ZSql + "Version ,"
        ZSql = ZSql + "Autorizado ,"
        ZSql = ZSql + "Etapa ,"
        ZSql = ZSql + "LetraInstrucciones ,"
        ZSql = ZSql + "Instrucciones ,"
        ZSql = ZSql + "Equipo ,"
        ZSql = ZSql + "LetraTemperatura ,"
        ZSql = ZSql + "Temperatura ,"
        ZSql = ZSql + "LetraTiempo ,"
        ZSql = ZSql + "Tiempo ,"
        ZSql = ZSql + "LetraControl ,"
        ZSql = ZSql + "Control ,"
        ZSql = ZSql + "Seguridad ,"
        ZSql = ZSql + "DesTerminado )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + XXFechaVersion + "',"
        ZSql = ZSql + "'" + XXOrdFecha + "',"
        ZSql = ZSql + "'" + XControlCambio + "',"
        ZSql = ZSql + "'" + ZLote + "',"
        ZSql = ZSql + "'" + XXVersion + "',"
        ZSql = ZSql + "'" + XXAutorizado + "',"
        ZSql = ZSql + "'" + Etapa + "',"
        ZSql = ZSql + "'" + LetraInstrucciones + "',"
        ZSql = ZSql + "'" + Instrucciones + "',"
        ZSql = ZSql + "'" + Equipo + "',"
        ZSql = ZSql + "'" + LetraTemperatura + "',"
        ZSql = ZSql + "'" + Temperatura + "',"
        ZSql = ZSql + "'" + LetraTiempo + "',"
        ZSql = ZSql + "'" + Tiempo + "',"
        ZSql = ZSql + "'" + LetraControl + "',"
        ZSql = ZSql + "'" + Control + "',"
        ZSql = ZSql + "'" + Seguridad + "',"
        ZSql = ZSql + "'" + DesTerminado.Caption + "')"
        
        rsCargaIV = ZSql
        Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE CargaIV SET "
        ZSql = ZSql + "Impre11 = " + "'" + IMPRE11 + "',"
        ZSql = ZSql + "Impre12 = " + "'" + Impre12 + "',"
        ZSql = ZSql + "Impre13 = " + "'" + Impre13 + "',"
        ZSql = ZSql + "Impre14 = " + "'" + Impre14 + "',"
        ZSql = ZSql + "Impre15 = " + "'" + Impre15 + "',"
        ZSql = ZSql + "Impre16 = " + "'" + Impre16 + "',"
        ZSql = ZSql + "Impre17 = " + "'" + Impre17 + "',"
        ZSql = ZSql + "Impre18 = " + "'" + Impre18 + "',"
        ZSql = ZSql + "Impre19 = " + "'" + Impre19 + "',"
        ZSql = ZSql + "Impre2 = " + "'" + Impre2 + "',"
        ZSql = ZSql + "Impre31 = " + "'" + Impre31 + "',"
        ZSql = ZSql + "Impre32 = " + "'" + Impre32 + "',"
        ZSql = ZSql + "Impre33 = " + "'" + Impre33 + "',"
        ZSql = ZSql + "Impre34 = " + "'" + Impre34 + "',"
        ZSql = ZSql + "Impre35 = " + "'" + Impre35 + "',"
        ZSql = ZSql + "Impre36 = " + "'" + Impre36 + "',"
        ZSql = ZSql + "Impre37 = " + "'" + Impre37 + "',"
        ZSql = ZSql + "Impre38 = " + "'" + Impre38 + "',"
        ZSql = ZSql + "Impre39 = " + "'" + Impre39 + "',"
        ZSql = ZSql + "Impre41 = " + "'" + Impre41 + "',"
        ZSql = ZSql + "Impre42 = " + "'" + Impre42 + "',"
        ZSql = ZSql + "Impre43 = " + "'" + Impre43 + "',"
        ZSql = ZSql + "Impre44 = " + "'" + Impre44 + "',"
        ZSql = ZSql + "Impre45 = " + "'" + Impre45 + "',"
        ZSql = ZSql + "Impre46 = " + "'" + Impre46 + "',"
        ZSql = ZSql + "Impre47 = " + "'" + Impre47 + "',"
        ZSql = ZSql + "Impre48 = " + "'" + Impre48 + "',"
        ZSql = ZSql + "Impre49 = " + "'" + Impre49 + "',"
        ZSql = ZSql + "Impre51 = " + "'" + Impre51 + "',"
        ZSql = ZSql + "Impre52 = " + "'" + Impre52 + "',"
        ZSql = ZSql + "Impre53 = " + "'" + Impre53 + "',"
        ZSql = ZSql + "Impre54 = " + "'" + Impre54 + "',"
        ZSql = ZSql + "Impre55 = " + "'" + Impre55 + "',"
        ZSql = ZSql + "Impre56 = " + "'" + Impre56 + "',"
        ZSql = ZSql + "Impre57 = " + "'" + Impre57 + "',"
        ZSql = ZSql + "Impre58 = " + "'" + Impre58 + "',"
        ZSql = ZSql + "Impre59 = " + "'" + Impre59 + "',"
        ZSql = ZSql + "Impre6 = " + "'" + Impre6 + "'"
        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"

        rsCargaIV = ZSql
        Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
        
    Next iRow
    
    ZOperador = "22"
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CargaIV SET "
    ZSql = ZSql + " Operador = " + "'" + ZOperador + "'"
    ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
                        
    spCargaIV = ZSql
    Set rstCargaIV = db.OpenRecordset(spCargaIV, dbOpenSnapshot, dbSQLPassThrough)
    

    XEmpresa = WEmpresa
    Erase CargaEmpresa
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7
            CargaEmpresa(1, 1) = "0001"
            CargaEmpresa(1, 2) = "Empresa01"
            CargaEmpresa(2, 1) = "0003"
            CargaEmpresa(2, 2) = "Empresa03"
            CargaEmpresa(3, 1) = "0005"
            CargaEmpresa(3, 2) = "Empresa05"
            CargaEmpresa(4, 1) = "0006"
            CargaEmpresa(4, 2) = "Empresa06"
            CargaEmpresa(5, 1) = "0007"
            CargaEmpresa(5, 2) = "Empresa07"
        Case 2, 4, 8, 9
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
        Case 10
            CargaEmpresa(1, 1) = "0010"
            CargaEmpresa(1, 2) = "Empresa10"
        Case Else
    End Select
            
    For Cicla = 1 To 5
        If CargaEmpresa(Cicla, 1) <> "" Then
        
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Terminado SET "
            ZSql = ZSql + " Metodo = " + "'" + Metodo.Text + "',"
            ZSql = ZSql + " VersionI = " + "'" + XXVersion + "',"
            ZSql = ZSql + " FechaVersionI = " + "'" + XXFechaVersion + "',"
            ZSql = ZSql + " EstadoI = " + "'" + "S" + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Terminado.Text + "'"
                
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    Next Cicla
    
    Select Case Val(XEmpresa)
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
        Case Else
    End Select

    Call Limpia_Click

    WVector1.Col = 1
    WVector1.Row = 1
    
    Terminado.SetFocus
        
End Sub



Private Sub Revalida_Click()

    If WGrabaII <> "S" Then
    
        Call Ingresa_ClaveII

               Else
               
        ZSql = ""
        ZSql = ZSql + "UPDATE CargaIV SET "
        ZSql = ZSql + " Autorizado = " + "'" + "S" + "'"
        ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
                    
        spCargaIV = ZSql
        Set rstCargaIV = db.OpenRecordset(spCargaIV, dbOpenSnapshot, dbSQLPassThrough)
    
        XEmpresa = WEmpresa
        Erase CargaEmpresa
        
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7
                CargaEmpresa(1, 1) = "0001"
                CargaEmpresa(1, 2) = "Empresa01"
                CargaEmpresa(2, 1) = "0003"
                CargaEmpresa(2, 2) = "Empresa03"
                CargaEmpresa(3, 1) = "0005"
                CargaEmpresa(3, 2) = "Empresa05"
                CargaEmpresa(4, 1) = "0006"
                CargaEmpresa(4, 2) = "Empresa06"
                CargaEmpresa(5, 1) = "0007"
                CargaEmpresa(5, 2) = "Empresa07"
            Case 2, 4, 8, 9
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
                CargaEmpresa(4, 1) = "0009"
                CargaEmpresa(4, 2) = "Empresa09"
            Case 10
                CargaEmpresa(1, 1) = "0010"
                CargaEmpresa(1, 2) = "Empresa10"
            Case Else
        End Select
                
        For Cicla = 1 To 5
            If CargaEmpresa(Cicla, 1) <> "" Then
            
                WEmpresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                ZSql = ""
                ZSql = ZSql + "UPDATE Terminado SET "
                ZSql = ZSql + " EstadoI = " + "'" + "S" + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + Terminado.Text + "'"
                    
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
        Next Cicla
        
        Select Case Val(XEmpresa)
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
            Case Else
        End Select
    
        Call Limpia_Click

        WVector1.Col = 1
        WVector1.Row = 1
        
        Terminado.SetFocus
        
    End If


End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Fecha.Text = "  /  /    "
    Version.Text = ""
    Autorizado.Text = ""
    DesOperador.Caption = ""
    Metodo.Text = ""
    ControlCambio.Text = ""
    
    Renglon = 0
    Graba.Enabled = True
    
    WGraba = ""
    WGrabaII = ""
    
    WVector1.Col = 1
    WVector1.Row = 1
    
    Terminado.SetFocus

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WTexto1.Text = WIndice.List(Indice)
            WVector1.Col = 4
            WVector1.Text = WIndice.List(Indice)
            
        Case 1
            Indice = Pantalla.ListIndex
            WTexto1.Text = WIndice.List(Indice)
            WVector1.Col = 10
            WVector1.Text = WIndice.List(Indice)
            
        Case 2
            Indice = Pantalla.ListIndex
            WTexto1.Text = WIndice.List(Indice)
            WVector1.Col = 11
            WVector1.Text = WIndice.List(Indice)
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    WVector1.Col = 1
    WVector1.Row = 1

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Fecha.Text = "  /  /    "
    Version.Text = ""
    Autorizado.Text = ""
    DesOperador.Caption = ""
    Metodo.Text = ""
    ControlCambio.Text = ""

    WGraba = ""
    WGrabaII = ""
    
    Renglon = 0
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    WRenglon = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM CargaIV"
    Sql3 = " Where CargaIV.Terminado = " + "'" + Terminado.Text + "'"
    Sql4 = " Order by CargaIV.Clave"
    
    rsCargaIV = Sql1 + Sql2 + Sql3 + Sql4
    Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaIV.RecordCount > 0 Then
        With rstCargaIV
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Fecha.Text = rstCargaIV!Fecha
                    Version.Text = rstCargaIV!Version
                    Autorizado.Text = rstCargaIV!Autorizado
                    ZOperador = IIf(IsNull(rstCargaIV!Operador), "O", rstCargaIV!Operador)
                    ControlCambio.Text = IIf(IsNull(rstCargaIV!ControlCambio), "", rstCargaIV!ControlCambio)
                
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector1.Col = 0
                    WVector1.Text = Trim(rstCargaIV!Etapa)
                    
                    WVector1.Col = 1
                    WVector1.Text = Trim(rstCargaIV!Etapa)
            
                    WVector1.Col = 2
                    WVector1.Text = Trim(rstCargaIV!LetraInstrucciones)
                    
                    WVector1.Col = 3
                    WVector1.Text = Trim(rstCargaIV!Instrucciones)
            
                    WVector1.Col = 4
                    WVector1.Text = Trim(rstCargaIV!Equipo)
            
                    WVector1.Col = 5
                    WVector1.Text = Trim(rstCargaIV!LetraTemperatura)
                    
                    WVector1.Col = 6
                    WVector1.Text = Trim(rstCargaIV!Temperatura)
            
                    WVector1.Col = 7
                    WVector1.Text = Trim(rstCargaIV!LetraTiempo)
                    
                    WVector1.Col = 8
                    WVector1.Text = Trim(rstCargaIV!Tiempo)
            
                    WVector1.Col = 9
                    WVector1.Text = Trim(rstCargaIV!LetraControl)
                    
                    WVector1.Col = 10
                    WVector1.Text = Trim(rstCargaIV!Control)
            
                    WVector1.Col = 11
                    WVector1.Text = Trim(rstCargaIV!Seguridad)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaIV.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM Terminado"
    Sql3 = " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
    spTerminado = Sql1 + Sql2 + Sql3
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        DesTerminado.Caption = Trim(rstTerminado!Descripcion)
        rstTerminado.Close
    End If
    
    If Val(ZOperador) <> 0 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Operador = " + "'" + ZOperador + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            DesOperador.Caption = IIf(IsNull(rstOperador!Descripcion), "", rstOperador!Descripcion)
            rstOperador.Close
        End If
    End If
    
    Graba.Enabled = True

End Sub

Private Sub Salva_Click()


    Hasta = WVector1.Row

    For iRow = 100 To Hasta Step -1
        WVector1.TextMatrix(iRow, 0) = WVector1.TextMatrix(iRow - 1, 0)
        WVector1.TextMatrix(iRow, 1) = WVector1.TextMatrix(iRow - 1, 1)
        WVector1.TextMatrix(iRow, 2) = WVector1.TextMatrix(iRow - 1, 2)
        WVector1.TextMatrix(iRow, 3) = WVector1.TextMatrix(iRow - 1, 3)
        WVector1.TextMatrix(iRow, 4) = WVector1.TextMatrix(iRow - 1, 4)
        WVector1.TextMatrix(iRow, 5) = WVector1.TextMatrix(iRow - 1, 5)
        WVector1.TextMatrix(iRow, 6) = WVector1.TextMatrix(iRow - 1, 6)
        WVector1.TextMatrix(iRow, 7) = WVector1.TextMatrix(iRow - 1, 7)
        WVector1.TextMatrix(iRow, 8) = WVector1.TextMatrix(iRow - 1, 8)
        WVector1.TextMatrix(iRow, 9) = WVector1.TextMatrix(iRow - 1, 9)
        WVector1.TextMatrix(iRow, 10) = WVector1.TextMatrix(iRow - 1, 10)
        WVector1.TextMatrix(iRow, 11) = WVector1.TextMatrix(iRow - 1, 11)
    Next iRow

    WVector1.TextMatrix(Hasta, 0) = ""
    WVector1.TextMatrix(Hasta, 1) = ""
    WVector1.TextMatrix(Hasta, 2) = ""
    WVector1.TextMatrix(Hasta, 3) = "En CASO de MEZCLA o CORTE, PROCEDER SEGUN INSTRUCTIVO I-AT-001-Vigente"
    WVector1.TextMatrix(Hasta, 4) = ""
    WVector1.TextMatrix(Hasta, 5) = ""
    WVector1.TextMatrix(Hasta, 6) = ""
    WVector1.TextMatrix(Hasta, 7) = ""
    WVector1.TextMatrix(Hasta, 8) = ""
    WVector1.TextMatrix(Hasta, 9) = ""
    WVector1.TextMatrix(Hasta, 10) = ""
    WVector1.TextMatrix(Hasta, 11) = ""
    
    WTexto1.Text = ""
    WTexto2.Text = ""

    ControlCambio.Text = "se agrega opcion de proceder por instructivo de Mezcla I AT 001"



End Sub

Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Terminado.Text = UCase(Terminado.Text)
        
        Sql1 = "Select *"
        Sql2 = " FROM Terminado"
        Sql3 = " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
        spTerminado = Sql1 + Sql2 + Sql3
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesTerminado.Caption = Trim(rstTerminado!Descripcion)
            Metodo.Text = IIf(IsNull(rstTerminado!Metodo), "", rstTerminado!Metodo)
            rstTerminado.Close
            
            Call Limpia_Vector

            Fecha.Text = "  /  /    "
            Version.Text = ""
            Autorizado.Text = ""
            ControlCambio.Text = ""
            
            Sql1 = "Select *"
            Sql2 = " FROM CargaIV"
            Sql3 = " Where CargaIV.Terminado = " + "'" + Terminado.Text + "'"
            rsCargaIV = Sql1 + Sql2 + Sql3
            Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaIV.RecordCount > 0 Then
                rstCargaIV.Close
                Call Proceso_Click
                WVector1.Col = 1
                WVector1.Row = 1
                Call StartEdit
                    Else
                Graba.Enabled = True
                WTerminado = Terminado.Text
                Terminado.Text = WTerminado
                WVector1.Col = 1
                WVector1.Row = 1
                Call StartEdit
            End If
                Else
            Terminado.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Terminado.Text = "  -     -   "
        DesTerminado.Caption = ""
    End If
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
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
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
            
        Case 123
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Col > 1 Then
                WVector1.Col = WVector1.Col - 1
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
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
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
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
        Case 11
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
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
            WVector1.TextMatrix(WVector1.Row, 0) = WVector1.TextMatrix(WVector1.Row, 1)
        Case 3, 6, 7
            Rem If Val(WVector1.Text) <> 0 Then
            Rem     ZCodigo = Val(WVector1.Text)
            Rem     Call Ceros(ZCodigo, 4)
            Rem
            Rem     Sql1 = "Select *"
            Rem     Sql2 = " FROM EquipoFabrica"
            Rem     Sql3 = " Where EquipoFabrica.Codigo = " + "'" + ZCodigo + "'"
            Rem     spEquipoFabrica = Sql1 + Sql2 + Sql3
            Rem     Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstEquipoFabrica.RecordCount > 0 Then
            Rem         rstEquipoFabrica.Close
            Rem     End If
            Rem End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
    
    RenglonAuxiliar = WVector1.Row

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
        
        Etapa = WVector1.TextMatrix(iRow, 1)
        Instrucciones = WVector1.TextMatrix(iRow, 3)
        Equipo = WVector1.TextMatrix(iRow, 4)
        Temperatura = WVector1.TextMatrix(iRow, 6)
        Tiempo = WVector1.TextMatrix(iRow, 8)
        Control = WVector1.TextMatrix(iRow, 10)
        Seguridad = WVector1.TextMatrix(iRow, 11)
            
        If Etapa <> "" Or Instrucciones <> "" Or Equipo <> "" Or Temperatura <> "" Or Tiempo <> "" Or Control <> "" Or Seguridad <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For da = 0 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
End Sub

Private Sub AgregaRenglon_Click()

    Hasta = WVector1.Row

    For iRow = 100 To Hasta Step -1
        WVector1.TextMatrix(iRow, 0) = WVector1.TextMatrix(iRow - 1, 0)
        WVector1.TextMatrix(iRow, 1) = WVector1.TextMatrix(iRow - 1, 1)
        WVector1.TextMatrix(iRow, 2) = WVector1.TextMatrix(iRow - 1, 2)
        WVector1.TextMatrix(iRow, 3) = WVector1.TextMatrix(iRow - 1, 3)
        WVector1.TextMatrix(iRow, 4) = WVector1.TextMatrix(iRow - 1, 4)
        WVector1.TextMatrix(iRow, 5) = WVector1.TextMatrix(iRow - 1, 5)
        WVector1.TextMatrix(iRow, 6) = WVector1.TextMatrix(iRow - 1, 6)
        WVector1.TextMatrix(iRow, 7) = WVector1.TextMatrix(iRow - 1, 7)
        WVector1.TextMatrix(iRow, 8) = WVector1.TextMatrix(iRow - 1, 8)
        WVector1.TextMatrix(iRow, 9) = WVector1.TextMatrix(iRow - 1, 9)
        WVector1.TextMatrix(iRow, 10) = WVector1.TextMatrix(iRow - 1, 10)
        WVector1.TextMatrix(iRow, 11) = WVector1.TextMatrix(iRow - 1, 11)
    Next iRow

    WVector1.TextMatrix(Hasta, 0) = ""
    WVector1.TextMatrix(Hasta, 1) = ""
    WVector1.TextMatrix(Hasta, 2) = ""
    WVector1.TextMatrix(Hasta, 3) = ""
    WVector1.TextMatrix(Hasta, 4) = ""
    WVector1.TextMatrix(Hasta, 5) = ""
    WVector1.TextMatrix(Hasta, 6) = ""
    WVector1.TextMatrix(Hasta, 7) = ""
    WVector1.TextMatrix(Hasta, 8) = ""
    WVector1.TextMatrix(Hasta, 9) = ""
    WVector1.TextMatrix(Hasta, 10) = ""
    WVector1.TextMatrix(Hasta, 11) = ""
    
    WTexto1.Text = ""
    WTexto2.Text = ""

End Sub


Private Sub WTexto2_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos Terminados a Utilizar"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click
    
    End If
    
    If WVector1.Col = 2 Then

    Opcion.Clear
    
     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Procesos (Equipo)"
     Opcion.AddItem "Procesos (Tiempo)"
     Opcion.AddItem "Procesos (Control)"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click
    
    End If
    
    
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
    WVector1.Cols = 12
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
    
    WVector1.ColWidth(0) = 400
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Etapa"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "L"
                WVector1.ColWidth(Ciclo) = 550
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Instrucciones"
                WVector1.ColWidth(Ciclo) = 8900
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 90
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Equipo"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "L"
                WVector1.ColWidth(Ciclo) = 550
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Temperatura"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 20
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "L"
                WVector1.ColWidth(Ciclo) = 550
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = "Tiempo"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = "L"
                WVector1.ColWidth(Ciclo) = 550
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 10
                WVector1.Text = "Control"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 20
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 11
                WVector1.Text = "Seguridad"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector1.Text
        Rem WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        Rem WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        Rem WTitulo(Ciclo).Width = WVector1.CellWidth
        Rem WTitulo(Ciclo).Height = WVector1.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 340
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector1.Width = WAncho

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

Sub Ingresa_clave()
    WClave.Text = ""
    XClave.Visible = True
    WClave.SetFocus
End Sub

Private Sub CancelaGraba_Click()
    XClave.Visible = False
End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WGraba = "N"
        ZGrabaIII = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClave.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZOperador = rstOperador!Operador
            ZGrabaIII = IIf(IsNull(rstOperador!GrabaIII), "", rstOperador!GrabaIII)
            rstOperador.Close
        End If
        
        If ZGrabaIII = "S" Then
            WGraba = "S"
            XClave.Visible = False
            Call Graba_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Ingreso de Procesos de Fabricacion")
            WClave.SetFocus
        End If
        
    End If
End Sub


Sub Ingresa_ClaveII()
    WClaveII.Text = ""
    XClaveII.Visible = True
    WClaveII.SetFocus
End Sub

Private Sub CancelaGrabaII_Click()
    XClaveII.Visible = False
End Sub

Private Sub WClaveII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WGrabaII = "N"
        ZGrabaIII = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClaveII.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZOperador = rstOperador!Operador
            ZGrabaIII = IIf(IsNull(rstOperador!GrabaIII), "", rstOperador!GrabaIII)
            rstOperador.Close
        End If
        
        If ZGrabaIII = "S" Then
            WGrabaII = "S"
            XClaveII.Visible = False
            Call Revalida_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Ingreso de Procesos de Fabricacion")
            WClaveII.SetFocus
        End If
        
    End If
End Sub



