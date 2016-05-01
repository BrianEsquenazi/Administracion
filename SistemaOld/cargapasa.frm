VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCargaPasa 
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
   Begin VB.Frame IngresaBase 
      Height          =   1215
      Left            =   3960
      TabIndex        =   29
      Top             =   2040
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
      Left            =   3480
      TabIndex        =   24
      Top             =   2160
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
      Left            =   6960
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
   Begin VB.TextBox WTexto12 
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
   Begin VB.ComboBox WCombo12 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto22 
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
      ItemData        =   "cargapasa.frx":0000
      Left            =   120
      List            =   "cargapasa.frx":0007
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox WTexto32 
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
   Begin MSFlexGridLib.MSFlexGrid WVector2 
      Height          =   4695
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8281
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
      Left            =   5640
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
      MouseIcon       =   "cargapasa.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "cargapasa.frx":031F
      ToolTipText     =   "Salida"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7320
      MouseIcon       =   "cargapasa.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "cargapasa.frx":0E6B
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   8160
      MouseIcon       =   "cargapasa.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "cargapasa.frx":19B7
      ToolTipText     =   "Consulta de Datos"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   9000
      MouseIcon       =   "cargapasa.frx":21F9
      MousePointer    =   99  'Custom
      Picture         =   "cargapasa.frx":2503
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
Attribute VB_Name = "PrgCargaPasa"
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

Dim ZLugar(100) As Integer
Dim ZDescri(1000, 100) As String

Rem para el vector

Dim WBorraII(1000, 20) As String
Dim WParametrosII(10, 20) As Double
Dim WFormatoII(20) As String
Dim WControlII As String

Private WGraba As String
Private WGrabaII As String

Dim CargaEmpresa(10, 2) As String

Private Sub Base_Click()

    IngresaBase.Visible = True
    
    ProductoBase.Text = "  -     -   "
    ProductoBase.SetFocus

End Sub

Private Sub ProductoBase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        Call Limpia_VectorII
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
                        WVector2.Row = WRenglon
                        Renglon = WRenglon
                
                        WVector2.Col = 0
                        WVector2.Text = Trim(rstCargaIV!Etapa)
                    
                        WVector2.Col = 1
                        WVector2.Text = Trim(rstCargaIV!Etapa)
            
                        WVector2.Col = 2
                        WVector2.Text = Trim(rstCargaIV!LetraInstrucciones)
                    
                        WVector2.Col = 3
                        WVector2.Text = Trim(rstCargaIV!Instrucciones)
            
                        WVector2.Col = 4
                        WVector2.Text = Trim(rstCargaIV!Equipo)
            
                        WVector2.Col = 5
                        WVector2.Text = Trim(rstCargaIV!LetraTemperatura)
                    
                        WVector2.Col = 6
                        WVector2.Text = Trim(rstCargaIV!Temperatura)
            
                        WVector2.Col = 7
                        WVector2.Text = Trim(rstCargaIV!LetraTiempo)
                    
                        WVector2.Col = 8
                        WVector2.Text = Trim(rstCargaIV!Tiempo)
            
                        WVector2.Col = 9
                        WVector2.Text = Trim(rstCargaIV!LetraControl)
                    
                        WVector2.Col = 10
                        WVector2.Text = Trim(rstCargaIV!Control)
            
                        WVector2.Col = 11
                        WVector2.Text = Trim(rstCargaIV!Seguridad)
                    
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
                        ZCodigo = rstEquipoFabrica!Codigo
                    
                        ZHAsta = Len(WDescripcion)
                        Desde = 1
                    
                        Do
                            Hasta = Desde + 15
                            If Hasta > ZHAsta Then
                                Hasta = ZHAsta
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
                        
                            If Hasta >= ZHAsta Then
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
        For IRow = 100 To 1 Step -1
        
            Etapa = WVector2.TextMatrix(IRow, 1)
            LetraInstrucciones = WVector2.TextMatrix(IRow, 2)
            Instrucciones = WVector2.TextMatrix(IRow, 3)
            Equipo = WVector2.TextMatrix(IRow, 4)
            LetraTemperatura = WVector2.TextMatrix(IRow, 5)
            Temperatura = WVector2.TextMatrix(IRow, 6)
            LetraTiempo = WVector2.TextMatrix(IRow, 7)
            Tiempo = WVector2.TextMatrix(IRow, 8)
            LetraControl = WVector2.TextMatrix(IRow, 9)
            Control = WVector2.TextMatrix(IRow, 10)
            Seguridad = WVector2.TextMatrix(IRow, 11)
            
            If Etapa <> "" Or Instrucciones <> "" Or Equipo <> "" Or Temperatura <> "" Or Tiempo <> "" Or Control <> "" Or Seguridad <> "" Then
                HastaRenglon = IRow
                Exit For
            End If
            
        Next IRow
    
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        Erase ZLugar

        WRenglon = 0
        For IRow = 1 To HastaRenglon
    
            ZLote = ""
        
            Etapa = WVector2.TextMatrix(IRow, 1)
            LetraInstrucciones = WVector2.TextMatrix(IRow, 2)
            Instrucciones = WVector2.TextMatrix(IRow, 3)
            Equipo = WVector2.TextMatrix(IRow, 4)
            LetraTemperatura = WVector2.TextMatrix(IRow, 5)
            Temperatura = WVector2.TextMatrix(IRow, 6)
            LetraTiempo = WVector2.TextMatrix(IRow, 7)
            Tiempo = WVector2.TextMatrix(IRow, 8)
            LetraControl = WVector2.TextMatrix(IRow, 9)
            Control = WVector2.TextMatrix(IRow, 10)
            Seguridad = WVector2.TextMatrix(IRow, 11)
        
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
            
        Next IRow



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
                            IngresaItem = Str$(rstEquipoFabrica!Codigo) + " " + rstEquipoFabrica!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = Str$(rstEquipoFabrica!Codigo)
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

    If WGraba <> "S" Then
    
        Call Ingresa_clave

               Else
               
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
                        ZCodigo = rstEquipoFabrica!Codigo
                    
                        ZHAsta = Len(WDescripcion)
                        Desde = 1
                    
                        Do
                            Hasta = Desde + 15
                            If Hasta > ZHAsta Then
                                Hasta = ZHAsta
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
                        
                            If Hasta >= ZHAsta Then
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
        For IRow = 100 To 1 Step -1
        
            Etapa = WVector2.TextMatrix(IRow, 1)
            LetraInstrucciones = WVector2.TextMatrix(IRow, 2)
            Instrucciones = WVector2.TextMatrix(IRow, 3)
            Equipo = WVector2.TextMatrix(IRow, 4)
            LetraTemperatura = WVector2.TextMatrix(IRow, 5)
            Temperatura = WVector2.TextMatrix(IRow, 6)
            LetraTiempo = WVector2.TextMatrix(IRow, 7)
            Tiempo = WVector2.TextMatrix(IRow, 8)
            LetraControl = WVector2.TextMatrix(IRow, 9)
            Control = WVector2.TextMatrix(IRow, 10)
            Seguridad = WVector2.TextMatrix(IRow, 11)
            
            If Etapa <> "" Or Instrucciones <> "" Or Equipo <> "" Or Temperatura <> "" Or Tiempo <> "" Or Control <> "" Or Seguridad <> "" Then
                HastaRenglon = IRow
                Exit For
            End If
            
        Next IRow
    
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        Erase ZLugar

        WRenglon = 0
        For IRow = 1 To HastaRenglon
    
            ZLote = ""
        
            Etapa = WVector2.TextMatrix(IRow, 1)
            LetraInstrucciones = WVector2.TextMatrix(IRow, 2)
            Instrucciones = WVector2.TextMatrix(IRow, 3)
            Equipo = WVector2.TextMatrix(IRow, 4)
            LetraTemperatura = WVector2.TextMatrix(IRow, 5)
            Temperatura = WVector2.TextMatrix(IRow, 6)
            LetraTiempo = WVector2.TextMatrix(IRow, 7)
            Tiempo = WVector2.TextMatrix(IRow, 8)
            LetraControl = WVector2.TextMatrix(IRow, 9)
            Control = WVector2.TextMatrix(IRow, 10)
            Seguridad = WVector2.TextMatrix(IRow, 11)
        
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
            
        Next IRow
    
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
            Case 2, 4, 8
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
            Case 9
                CargaEmpresa(1, 1) = "0009"
                CargaEmpresa(1, 2) = "Empresa09"
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

        WVector2.Col = 1
        WVector2.Row = 1
        
        Terminado.SetFocus
        
    End If
        
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
            Case 2, 4, 8
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
            Case 9
                CargaEmpresa(1, 1) = "0009"
                CargaEmpresa(1, 2) = "Empresa09"
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

        WVector2.Col = 1
        WVector2.Row = 1
        
        Terminado.SetFocus
        
    End If


End Sub

Private Sub Limpia_Click()
    
    Call Limpia_VectorII

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Fecha.Text = "  /  /    "
    Version.Text = ""
    Autorizado.Text = ""
    
    Renglon = 0
    Graba.Enabled = True
    
    WGraba = ""
    WGrabaII = ""
    
    WVector2.Col = 1
    WVector2.Row = 1
    
    Terminado.SetFocus

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WTexto12.Text = WIndice.List(Indice)
            WVector2.Col = 4
            WVector2.Text = WIndice.List(Indice)
            
        Case 1
            Indice = Pantalla.ListIndex
            WTexto12.Text = WIndice.List(Indice)
            WVector2.Col = 10
            WVector2.Text = WIndice.List(Indice)
            
        Case 2
            Indice = Pantalla.ListIndex
            WTexto12.Text = WIndice.List(Indice)
            WVector2.Col = 11
            WVector2.Text = WIndice.List(Indice)
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_VectorII
    
    WVector2.Col = 1
    WVector2.Row = 1

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Fecha.Text = "  /  /    "
    Version.Text = ""
    Autorizado.Text = ""
    
    WGraba = ""
    WGrabaII = ""
    
    Renglon = 0
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_VectorII
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
                
                    WRenglon = WRenglon + 1
                    WVector2.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector2.Col = 0
                    WVector2.Text = Trim(rstCargaIV!Etapa)
                    
                    WVector2.Col = 1
                    WVector2.Text = Trim(rstCargaIV!Etapa)
            
                    WVector2.Col = 2
                    WVector2.Text = Trim(rstCargaIV!LetraInstrucciones)
                    
                    WVector2.Col = 3
                    WVector2.Text = Trim(rstCargaIV!Instrucciones)
            
                    WVector2.Col = 4
                    WVector2.Text = Trim(rstCargaIV!Equipo)
            
                    WVector2.Col = 5
                    WVector2.Text = Trim(rstCargaIV!LetraTemperatura)
                    
                    WVector2.Col = 6
                    WVector2.Text = Trim(rstCargaIV!Temperatura)
            
                    WVector2.Col = 7
                    WVector2.Text = Trim(rstCargaIV!LetraTiempo)
                    
                    WVector2.Col = 8
                    WVector2.Text = Trim(rstCargaIV!Tiempo)
            
                    WVector2.Col = 9
                    WVector2.Text = Trim(rstCargaIV!LetraControl)
                    
                    WVector2.Col = 10
                    WVector2.Text = Trim(rstCargaIV!Control)
            
                    WVector2.Col = 11
                    WVector2.Text = Trim(rstCargaIV!Seguridad)
                    
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
    
    Graba.Enabled = True

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
            rstTerminado.Close
            
            Call Limpia_VectorII

            Fecha.Text = "  /  /    "
            Version.Text = ""
            Autorizado.Text = ""
            
            Sql1 = "Select *"
            Sql2 = " FROM CargaIV"
            Sql3 = " Where CargaIV.Terminado = " + "'" + Terminado.Text + "'"
            rsCargaIV = Sql1 + Sql2 + Sql3
            Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaIV.RecordCount > 0 Then
                rstCargaIV.Close
                Call Proceso_Click
                WVector2.Col = 1
                WVector2.Row = 1
                Call StartEditII
                    Else
                Graba.Enabled = True
                WTerminado = Terminado.Text
                Terminado.Text = WTerminado
                WVector2.Col = 1
                WVector2.Row = 1
                Call StartEditII
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
Rem Controles de la WVector2
Rem

Private Sub GridEditTextII(ByVal KeyAscii As Integer)

    XColumna = WVector2.Col
    XTipoDato = WParametrosII(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto12.Left = WVector2.CellLeft + WVector2.Left
            WTexto12.Top = WVector2.CellTop + WVector2.Top
            WTexto12.Width = WVector2.CellWidth
            WTexto12.Height = WVector2.CellHeight
            WTexto12.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto12.Text = WVector2.Text
                    WTexto12.SelStart = Len(WTexto12.Text)
                Case Else
                    WTexto12.Text = Chr$(KeyAscii)
                    WTexto12.SelStart = 1
            End Select
            WTexto12.Visible = True
            WTexto12.SetFocus
        Case 1
            WTexto22.Left = WVector2.CellLeft + WVector2.Left
            WTexto22.Top = WVector2.CellTop + WVector2.Top
            WTexto22.Width = WVector2.CellWidth
            WTexto22.Height = WVector2.CellHeight
            WTexto22.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto22.Text = WVector2.Text
                    Rem WTexto22.SelStart = Len(WTexto22.Text)
                    WTexto22.SelStart = 0
                Case Else
                    WTexto22.Text = Chr$(KeyAscii)
                    WTexto22.SelStart = 1
            End Select
            WTexto22.Visible = True
            WTexto22.SetFocus
        Case 2
            WTexto32.Left = WVector2.CellLeft + WVector2.Left
            WTexto32.Top = WVector2.CellTop + WVector2.Top
            WTexto32.Width = WVector2.CellWidth
            WTexto32.Height = WVector2.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector2.Text) = 10 Then
                        WTexto32.Text = WVector2.Text
                            Else
                        WTexto32.Text = "  /  /    "
                    End If
                    WTexto32.SelStart = 0
                Case Else
                    WTexto32.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto32.SelStart = 1
            End Select
            WTexto32.Visible = True
            WTexto32.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditII()
    Pasa = 0
    If WCombo12.Visible Then
        Pasa = 0
        WVector2.Text = WCombo12.Text
        WCombo12.Visible = False
            Else
        If WTexto12.Visible Then
            Pasa = 1
            WVector2.Text = WTexto12.Text
            WTexto12.Visible = False
                Else
            If WTexto22.Visible Then
                Pasa = 1
                WVector2.Text = WTexto22.Text
                WTexto22.Visible = False
                    Else
                If WTexto32.Visible Then
                    Pasa = 1
                    WVector2.Text = WTexto32.Text
                    WTexto32.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoII(WVector2.Col) <> "" Then
            WVector2.Text = Pusing(WFormatoII(WVector2.Col), WVector2.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboII()
    ' Position the ComboBox over the cell.
    WCombo12.Left = WVector2.CellLeft + WVector2.Left
    WCombo12.Top = WVector2.CellTop + WVector2.Top
    WCombo12.Width = WVector2.CellWidth
    WCombo12.Visible = True
    WCombo12.SetFocus
End Sub

Private Sub WTexto12_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto12.Text = ""
            
        Rem F1
        Case 113
            WTexto12.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            DoEvents
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 123
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Col > 1 Then
                WVector2.Col = WVector2.Col - 1
            End If
            Call StartEditII

    End Select
End Sub

Private Sub WTexto22_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto22.Text = ""
            
        Rem F1
        Case 113
            WTexto22.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            DoEvents
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

Private Sub Wtexto32_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto32.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto32.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto12_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto22_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub Wtexto32_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo12_Click()
    WVector2.SetFocus
End Sub


Private Sub WVector2_Click()
    StartEditII
End Sub

Private Sub WVector2_LeaveCell()
    EndEditII
End Sub

Private Sub WVector2_GotFocus()
    EndEditII
End Sub

Private Sub WVector2_KeyPress(KeyAscii As Integer)
    XColumna = WVector2.Col
    Select Case WParametrosII(4, WVector2.Col)
        Case 1
        Case Else
            If WParametrosII(2, XColumna) = 0 Then
                GridEditTextII KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditII()
    Select Case WParametrosII(4, WVector2.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo12.Clear
            WCombo12.AddItem "Campo1"
            WCombo12.AddItem "Campo2"
            On Error Resume Next
            WCombo12.Text = WVector2.Text
            On Error GoTo 0
            GridEditComboII
        Case Else
            If WParametrosII(2, WVector2.Col) = 0 Then
                GridEditTextII Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVector2()
    Select Case WVector2.Col
        Case 11
            If WVector2.Row < WVector2.Rows - 1 Then
                WVector2.Row = WVector2.Row + 1
            End If
            WVector2.Col = 1
        Case Else
            If WVector2.Col < WVector2.Cols - 1 Then
                WVector2.Col = WVector2.Col + 1
            End If
    End Select
    WVector2.SetFocus
    GridEditTextII KeyAscii
End Sub

Private Sub Control_CampoII()
    XColumna = WVector2.Col
    XFila = WVector2.Row
    WControlII = "S"
    Select Case XColumna
        Case 1
            WVector2.TextMatrix(WVector2.Row, 0) = WVector2.TextMatrix(WVector2.Row, 1)
        Case 3, 6, 7
            Rem If Val(WVector2.Text) <> 0 Then
            Rem     ZCodigo = Val(WVector2.Text)
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
            WVector2.Col = XColumna
    End Select
End Sub

Private Sub WVector2_DblClick()

    If WVector2.Col = 0 Or WVector2.Col = 1 Then
    
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
    
    RenglonAuxiliar = WVector2.Row

    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        WVector2.Text = ""
    Next Ciclo
    
    Erase WBorraII
    EntraVector = 0
    
    HastaRenglon = 0
    For IRow = 100 To 1 Step -1
        
        Etapa = WVector2.TextMatrix(IRow, 1)
        Instrucciones = WVector2.TextMatrix(IRow, 3)
        Equipo = WVector2.TextMatrix(IRow, 4)
        Temperatura = WVector2.TextMatrix(IRow, 6)
        Tiempo = WVector2.TextMatrix(IRow, 8)
        Control = WVector2.TextMatrix(IRow, 10)
        Seguridad = WVector2.TextMatrix(IRow, 11)
            
        If Etapa <> "" Or Instrucciones <> "" Or Equipo <> "" Or Temperatura <> "" Or Tiempo <> "" Or Control <> "" Or Seguridad <> "" Then
            HastaRenglon = IRow
            Exit For
        End If
            
    Next IRow
    
    For Ciclo = 1 To HastaRenglon
        WVector2.Row = Ciclo
        WVector2.Col = 1
        WAuxi1 = WVector2.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector2.Cols - 1
                WVector2.Col = Ciclo1
                WBorraII(EntraVector, Ciclo1) = WVector2.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_VectorII
    
    For Ciclo = 1 To EntraVector
        WVector2.Row = Ciclo
        For da = 0 To WVector2.Cols - 1
            WVector2.Col = da
            WVector2.Text = WBorraII(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
End Sub

Private Sub AgregaRenglon_Click()

    Hasta = WVector2.Row

    For IRow = 100 To Hasta Step -1
        WVector2.TextMatrix(IRow, 0) = WVector2.TextMatrix(IRow - 1, 0)
        WVector2.TextMatrix(IRow, 1) = WVector2.TextMatrix(IRow - 1, 1)
        WVector2.TextMatrix(IRow, 2) = WVector2.TextMatrix(IRow - 1, 2)
        WVector2.TextMatrix(IRow, 3) = WVector2.TextMatrix(IRow - 1, 3)
        WVector2.TextMatrix(IRow, 4) = WVector2.TextMatrix(IRow - 1, 4)
        WVector2.TextMatrix(IRow, 5) = WVector2.TextMatrix(IRow - 1, 5)
        WVector2.TextMatrix(IRow, 6) = WVector2.TextMatrix(IRow - 1, 6)
        WVector2.TextMatrix(IRow, 7) = WVector2.TextMatrix(IRow - 1, 7)
        WVector2.TextMatrix(IRow, 8) = WVector2.TextMatrix(IRow - 1, 8)
        WVector2.TextMatrix(IRow, 9) = WVector2.TextMatrix(IRow - 1, 9)
        WVector2.TextMatrix(IRow, 10) = WVector2.TextMatrix(IRow - 1, 10)
        WVector2.TextMatrix(IRow, 11) = WVector2.TextMatrix(IRow - 1, 11)
    Next IRow

    WVector2.TextMatrix(Hasta, 0) = ""
    WVector2.TextMatrix(Hasta, 1) = ""
    WVector2.TextMatrix(Hasta, 2) = ""
    WVector2.TextMatrix(Hasta, 3) = ""
    WVector2.TextMatrix(Hasta, 4) = ""
    WVector2.TextMatrix(Hasta, 5) = ""
    WVector2.TextMatrix(Hasta, 6) = ""
    WVector2.TextMatrix(Hasta, 7) = ""
    WVector2.TextMatrix(Hasta, 8) = ""
    WVector2.TextMatrix(Hasta, 9) = ""
    WVector2.TextMatrix(Hasta, 10) = ""
    WVector2.TextMatrix(Hasta, 11) = ""
    
    WTexto12.Text = ""
    WTexto22.Text = ""

End Sub


Private Sub WTexto22_DblClick()

    If WVector2.Col = 1 Then

    Opcion.Clear
    
     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos Terminados a Utilizar"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click
    
    End If
    
    If WVector2.Col = 2 Then

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

Private Sub Limpia_VectorII()

    WVector2.Clear

    Rem ponga la WVector2 en negritas
    WVector2.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto12.FontName = WVector2.FontName
    WTexto12.FontSize = WVector2.FontSize
    WTexto12.Visible = False
    WTexto22.FontName = WVector2.FontName
    WTexto22.FontSize = WVector2.FontSize
    WTexto22.Visible = False
    WTexto32.FontName = WVector2.FontName
    WTexto32.FontSize = WVector2.FontSize
    WTexto32.Visible = False
    WCombo12.FontName = WVector2.FontName
    WCombo12.FontSize = WVector2.FontSize
    WCombo12.Visible = False

    ' Establesco loa Valores de la WVector2
    
    WVector2.FixedCols = 1
    WVector2.Cols = 12
    WVector2.FixedRows = 1
    WVector2.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector2.Text = "Articulo"
    
    Rem Longitud
    Rem WVector2.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametrosII(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametrosII(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametrosII(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametrosII(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector2.ColWidth(0) = 400
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Etapa"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 10
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 2
                WVector2.Text = "L"
                WVector2.ColWidth(Ciclo) = 550
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 4
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 3
                WVector2.Text = "Instrucciones"
                WVector2.ColWidth(Ciclo) = 8900
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 90
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 4
                WVector2.Text = "Equipo"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 15
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 5
                WVector2.Text = "L"
                WVector2.ColWidth(Ciclo) = 550
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 4
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 6
                WVector2.Text = "Temperatura"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 20
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 7
                WVector2.Text = "L"
                WVector2.ColWidth(Ciclo) = 550
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 4
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 8
                WVector2.Text = "Tiempo"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 15
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 9
                WVector2.Text = "L"
                WVector2.ColWidth(Ciclo) = 550
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 4
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 10
                WVector2.Text = "Control"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 20
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 11
                WVector2.Text = "Seguridad"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 15
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector2.Text
        Rem WTitulo(Ciclo).Left = WVector2.CellLeft + WVector2.Left
        Rem WTitulo(Ciclo).Top = WVector2.CellTop + WVector2.Top
        Rem WTitulo(Ciclo).Width = WVector2.CellWidth
        Rem WTitulo(Ciclo).Height = WVector2.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector2
    
    WAncho = 340
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Col = 1
    WVector2.Row = 1
    
End Sub

Private Sub WVector2_Scroll()
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
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
        If UCase(WClave.Text) = "NEGRO" Then
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
        If UCase(WClaveII.Text) = "NEGRO" Then
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



