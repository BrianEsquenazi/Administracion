VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsFicMatDyTransito 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Ficha de Stock de Materias Primas"
   ClientHeight    =   8130
   ClientLeft      =   90
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form2"
   ScaleHeight     =   8130
   ScaleWidth      =   11775
   Begin VB.CommandButton AvisoError 
      Caption         =   "Sistema sin Conexion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      Picture         =   "consficmatdyTransito.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
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
      Index           =   9
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WStock4 
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   " "
      Top             =   1200
      Width           =   975
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
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
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
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox Tipo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7800
      TabIndex        =   23
      Top             =   2400
      Width           =   2415
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
      Index           =   8
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3000
      Visible         =   0   'False
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
      Index           =   7
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3000
      Visible         =   0   'False
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
      Index           =   6
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3000
      Visible         =   0   'False
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
      Index           =   5
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3000
      Visible         =   0   'False
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3000
      Visible         =   0   'False
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
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
      Index           =   3
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
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
      Index           =   4
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10920
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WLotematdy.rpt"
   End
   Begin VB.TextBox XStock 
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   " "
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox XSalidas 
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   " "
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox XEntradas 
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin MSMask.MaskEdBox Articulo 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Proceso"
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
      Left            =   3480
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   840
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
      Height          =   1620
      ItemData        =   "consficmatdyTransito.frx":0742
      Left            =   120
      List            =   "consficmatdyTransito.frx":0749
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   3255
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
      Height          =   300
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   5175
      Left            =   0
      TabIndex        =   18
      Top             =   2880
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9128
      _Version        =   393216
      BackColor       =   16777152
   End
   Begin VB.Label WLabel4 
      Caption         =   "Stock "
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
      Left            =   4800
      TabIndex        =   28
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "O/C Pend."
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
      Height          =   375
      Left            =   7680
      TabIndex        =   25
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Saldo Final"
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
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Salidas"
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
      Left            =   4800
      TabIndex        =   9
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Entradas"
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
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label DesArticulo 
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
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Articulo"
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
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "PrgConsFicMatDyTransito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WClave As String
Private Vector(1000, 13) As String
Private XLote(100, 7) As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim XParam As String
Private WWOrden As String
Private WXInicial As Double
Private WXSalidas As Double
Private WXEntradas As Double
Private WXStock As Double
Private WCanti As Double
Private WSaldo As Double
Dim WWVector(10000, 4) As String
Dim WOrden As String
Private NombreEmpresa As String
Private WGrilla(100, 10) As String

Private Sub AvisoError_Click()
    Rem AvisoError.Visible = False
End Sub

Private Sub cmdClose_Click()
    Articulo.SetFocus
    PrgConsFicMatDyTransito.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String
    
    Pantalla.Clear
    WIndice.Clear

    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    With rstArticulo
        .MoveFirst
        Do
            If .EOF = False Then
                If Left$(rstArticulo!Codigo, 2) = "DY" Or Left$(rstArticulo!Codigo, 2) = "DS" Or Left$(rstArticulo!Codigo, 2) = "DQ" Then
                    IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                    Pantalla.AddItem IngresaItem
                    IngresaItem = rstArticulo!Codigo
                    WIndice.AddItem IngresaItem
                End If
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstArticulo.Close
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub Tipo_Click()
    If Articulo.Text <> "  -   -   " Then
        Call Proceso_Click
    End If
End Sub

Private Sub WVector1_DblClick()
    
    spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        WFechaCierre = IIf(IsNull(rstArticulo!FechaCierre), "00/00/0000", rstArticulo!FechaCierre)
        WOrdFechaCierre = IIf(IsNull(rstArticulo!OrdFechaCierre), "00000000", rstArticulo!OrdFechaCierre)
        rstArticulo.Close
    End If

    If Left$(Articulo.Text, 2) = "DY" Or Left$(Articulo.Text, 2) = "DS" Or Left$(Articulo.Text, 2) = "DQ" Then
    
        WVector1.Col = 7
        WPartiOri = WVector1.Text
        nrolote = PartiOri
        WEntra = "N"
                
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + Articulo.Text + "'"
        ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WPartiOri + "'"
        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            With rstLaudo
                .MoveFirst
                nrolote = IIf(IsNull(rstLaudo!Laudo), "", Str$(rstLaudo!Laudo))
                WEntra = "S"
                rstLaudo.Close
            End With
        End If
                    
        If WEntra = "N" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Guia"
            ZSql = ZSql + " Where Guia.Articulo = " + "'" + Articulo.Text + "'"
            ZSql = ZSql + " and Guia.PartiOri = " + "'" + WPartiOri + "'"
            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
            spMovguia = ZSql
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                With rstMovguia
                    .MoveFirst
                    nrolote = IIf(IsNull(rstMovguia!Lote), "", Str$(rstMovguia!Lote))
                    rstMovguia.Close
                End With
            End If
        End If
        
            Else
            
        WPartiOri = ""
        WVector1.Col = 7
        nrolote = WVector1.Text
        
    End If
    
    Da = 0
    With rstFichaMat
        .Index = "Articulo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Erase WGrilla
    LugarGrilla = 0
    
    WArticulo = Articulo.Text
    WLote = nrolote

    If (Left$(WArticulo, 2) = "DY" Or Left$(WArticulo, 2) = "DS" Or Left$(WArticulo, 2) = "DQ") And Trim(WPartiOri) <> "" Then
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.PartiOri = " + "'" + WPartiOri + "'"
        ZSql = ZSql + " and Laudo.Articulo = " + "'" + WArticulo + "'"
        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            With rstLaudo
                .MoveFirst
                Do
                    If .EOF = False Then
                        WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                        If WLiberada <> 0 Then
                            LugarGrilla = LugarGrilla + 1
                            WGrilla(LugarGrilla, 1) = rstLaudo!Fecha
                            WGrilla(LugarGrilla, 2) = Right$(rstLaudo!Fecha, 4) + Mid$(rstLaudo!Fecha, 4, 2) + Left$(rstLaudo!Fecha, 2)
                            WGrilla(LugarGrilla, 3) = rstLaudo!Laudo
                            WGrilla(LugarGrilla, 4) = Str$(rstLaudo!Liberada)
                            WGrilla(LugarGrilla, 5) = rstLaudo!Orden
                            WGrilla(LugarGrilla, 6) = "Laudo"
                            WGrilla(LugarGrilla, 7) = rstLaudo!Laudo
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstLaudo.Close
        End If
        
            Else
        
        XParam = "'" + WLote + "','" _
                     + WArticulo + "'"
    
        spLaudo = "ListaLaudoArticulo" + XParam
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
        
            WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
            If WLiberada <> 0 Then
                LugarGrilla = LugarGrilla + 1
                WGrilla(LugarGrilla, 1) = rstLaudo!Fecha
                WGrilla(LugarGrilla, 2) = Right$(rstLaudo!Fecha, 4) + Mid$(rstLaudo!Fecha, 4, 2) + Left$(rstLaudo!Fecha, 2)
                WGrilla(LugarGrilla, 3) = rstLaudo!Laudo
                WGrilla(LugarGrilla, 4) = Str$(rstLaudo!Liberada)
                WGrilla(LugarGrilla, 5) = rstLaudo!Orden
                WGrilla(LugarGrilla, 6) = "Laudo"
                WGrilla(LugarGrilla, 7) = rstLaudo!Laudo
            End If
            rstLaudo.Close
           
        End If
            
    End If
    
    For Ciclo = 1 To LugarGrilla
    
        WFecha = WGrilla(Ciclo, 1)
        WFechaord = WGrilla(Ciclo, 2)
        WCodigo = WGrilla(Ciclo, 3)
        WCantidad = Val(WGrilla(Ciclo, 4))
        WComprobante = WGrilla(Ciclo, 5)
        WDescri = WGrilla(Ciclo, 6)
        WLote = WGrilla(Ciclo, 7)

        If WDescri = "Guia In" Then
        
            Select Case Val(WComprobante)
                Case 1
                    WObservaciones = "Recepcion de Surfactan"
                Case 2
                    WObservaciones = "Recepcion de Pellital"
                Case 3
                    WObservaciones = "Recepcion de Surfactan II"
                Case 4
                    WObservaciones = "Recepcion de Pellital II"
                Case 5
                    WObservaciones = "Recepcion de Surfactan III"
                Case 6
                    WObservaciones = "Recepcion de Surfactan IV"
                Case 7
                    WObservaciones = "Recepcion de Surfactan V"
                Case 8
                    WObservaciones = "Recepcion de Pellital V"
                Case 9
                    WObservaciones = "Recepcion de Pellital IV"
                Case 10
                    WObservaciones = "Recepcion de Surfactan VI"
                Case 11
                    WObservaciones = "Recepcion de Surfactan VII"
                Case Else
            End Select
            
                Else
                
            spOrden = "ListaOrden" + "'" + WComprobante + "'"
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                WProveedor = rstOrden!Proveedor
                rstOrden.Close
            End If
        
            WObservaciones = ""
                
            spProveedor = "ConsultaProveedores" + "'" + WProveedor + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                WObservaciones = RstProveedor!Nombre
                RstProveedor.Close
            End If
            
        End If
            
        WDesArticulo = ""
            
        spArticulo = "ConsultaArticulo " + " '" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WDesArticulo = rstArticulo!Descripcion
            rstArticulo.Close
        End If
                
        With rstFichaMat
            .AddNew
            !Articulo = WArticulo
            !Fecha = WFecha
            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
            !Tipo = 0
            !Numero = WCodigo
            !Inicial = 0
            !Entrada = WCantidad
            !Salida = 0
            !Descripcion = WDesArticulo
            !Observaciones = WObservaciones
            !Lista1 = WDescri
            !Lista2 = ""
            !Lote = WLote
            !Saldo = 0
            !Empresa = NombreEmpresa
            !PartiOri = WPartiOri
            .Update
        End With
    
    Next Ciclo
            
    
    Erase Vector
    Renglon = 0
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spHoja = "ListaHojaArticuloDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                If rstHoja!Marca = "X" Or XFec < WOrdFechaCierre Then
                
                        Else
                
                If !Tipo = "M" Then
                
                    XLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                    XLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "", rstHoja!Canti1)
                    XLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                    XLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "", rstHoja!Canti2)
                    XLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                    XLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "", rstHoja!Canti3)
                    
                    Rem If Val(XLote(1, 1)) = 0 And rstHoja!Lote <> 0 Then
                    Rem     XLote(1, 1) = rstHoja!Lote
                    Rem     XLote(1, 2) = rstHoja!Cantidad
                    Rem End If
                    
                    For Da = 1 To 3
                        If Val(XLote(Da, 1)) = Val(nrolote) Then
                
                            WArticulo = rstHoja!Articulo
                            WCantidad = XLote(Da, 2)
                            WFecha = rstHoja!Fecha
                            WHoja = rstHoja!Hoja
                            WSaldo = 0
                
                            With rstFichaMat
                
                                .AddNew
                                !Articulo = WArticulo
                                !Fecha = WFecha
                                !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                !Tipo = 0
                                !Numero = WHoja
                                !Inicial = 0
                                !Entrada = 0
                                !Salida = WCantidad
                                !Observaciones = ""
                                !Descripcion = WDesArticulo
                                !Lista1 = "Hoja"
                                !Lista2 = ""
                                !Lote = Val(nrolote)
                                !Saldo = WSaldo
                                !Empresa = NombreEmpresa
                                !PartiOri = WPartiOri
                                .Update
                            End With
                        End If
                    Next Da
                        
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstHoja.Close
    End If
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spMovvar = "ListaMovvarArticuloDesdeHasta" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
    
        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovvar!Marca = "X" Then
                
                        Else
                
                
                If !Tipo = "M" Then
                
                    WLote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                    
                    If Val(WLote) = Val(nrolote) Then
                
                        WArticulo = rstMovvar!Articulo
                        WCantidad = rstMovvar!Cantidad
                        WFecha = rstMovvar!Fecha
                        WCodigo = rstMovvar!Codigo
                        WMovi = rstMovvar!Movi
                        WTipomov = Val(rstMovvar!Tipomov)
                        WObservaciones = rstMovvar!Observaciones
                        WSaldo = 0
                    
                        With rstFichaMat
                    
                            .AddNew
                            !Articulo = WArticulo
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            If WMovi = "E" Then
                                !Entrada = WCantidad
                                !Salida = 0
                                    Else
                                !Entrada = 0
                                !Salida = WCantidad
                            End If
                            !Observaciones = WObservaciones
                            !Descripcion = WDesArticulo
                            If WTipomov = 0 Or WTipomov = 1 Then
                                !Lista1 = "Mov.Var"
                                    Else
                                !Lista1 = "Guia In"
                            End If
                            !Lista2 = ""
                            !Lote = WLote
                            !Saldo = WSaldo
                            !Empresa = NombreEmpresa
                            !PartiOri = WPartiOri
                            .Update
                        End With
                    End If
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstMovvar.Close
    End If
    
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spMovguia = "ListaMovguiaArticuloDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovguia!Marca = "X" Then
                
                        Else
                        
                If rstMovguia!Tipo = "M" Then
            
                    WArticulo = rstMovguia!Articulo
                    WCantidad = rstMovguia!Cantidad
                    WFecha = rstMovguia!Fecha
                    WCodigo = rstMovguia!Codigo
                    WMovi = rstMovguia!Movi
                    WDestino = rstMovguia!Destino
                    WTipomov = rstMovguia!Tipomov
                    Rem WObservaciones = rstMovvar!Observaciones
                        
                    If WMovi = "E" Then
                        WLote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(WSaldo)
                            Else
                        WLote = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                        WSaldo = 0
                    End If

                        
                    If WMovi = "S" Then
                        Select Case WDestino
                            Case 1
                                WObservaciones = "Envio a Surfactan"
                            Case 2
                                WObservaciones = "Envio a Pellital"
                            Case 3
                                WObservaciones = "Envio a Surfactan II"
                            Case 4
                                WObservaciones = "Envio a Pellital II"
                            Case 5
                                WObservaciones = "Envio a Surfactan III"
                            Case 6
                                WObservaciones = "Envio a Surfactan IV"
                            Case 7
                                WObservaciones = "Envio a Surfactan V"
                            Case 8
                                WObservaciones = "Envio a Pellital V"
                            Case 9
                                WObservaciones = "Envio a Pellital IV"
                            Case 10
                                WObservaciones = "Envio a Surfactan VI"
                            Case 11
                                WObservaciones = "Envio a Surfactan VII"
                            Case Else
                        End Select
                            
                                Else
                                
                        Select Case WTipomov
                            Case 1
                                WObservaciones = "Recepcion de Surfactan"
                            Case 2
                                WObservaciones = "Recepcion de Pellital"
                            Case 3
                                WObservaciones = "Recepcion de Surfactan II"
                            Case 4
                                WObservaciones = "Recepcion de Pellital II"
                            Case 5
                                WObservaciones = "Recepcion de Surfactan III"
                            Case 6
                                WObservaciones = "Recepcion de Surfactan IV"
                            Case 7
                                WObservaciones = "Recepcion de Surfactan V"
                            Case 8
                                WObservaciones = "Recepcion de Pellital V"
                            Case 9
                                WObservaciones = "Recepcion de Pellital IV"
                            Case 10
                                WObservaciones = "Recepcion de Surfactan VI"
                            Case 11
                                WObservaciones = "Recepcion de Surfactan VII"
                            Case Else
                        End Select
                            
                    End If
                    
                    If Val(WLote) = Val(nrolote) Then
                    
                        With rstFichaMat
                    
                            .AddNew
                            !Articulo = WArticulo
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            If WMovi = "E" Then
                                !Entrada = WCantidad
                                !Salida = 0
                                    Else
                                !Entrada = 0
                                !Salida = WCantidad
                            End If
                            !Observaciones = WObservaciones
                            !Descripcion = WDesArticulo
                            If !Numero > 900000 Then
                                !Lista1 = "Prestamo"
                                !Numero = !Numero - 900000
                                    Else
                                !Lista1 = "Guia In"
                            End If
                            !Lista2 = ""
                            !Lote = WLote
                            !Saldo = WSaldo
                            !Empresa = NombreEmpresa
                            !PartiOri = WPartiOri
                            .Update
                        End With
                    End If
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstMovguia.Close
    End If
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spMovlab = "ListaMovlabArticuloDesdeHasta" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovlab!Marca = "X" Then
                
                        Else
                
                If !Tipo = "M" Then
                
                    WArticulo = rstMovlab!Articulo
                    WCantidad = rstMovlab!Cantidad
                    WFecha = rstMovlab!Fecha
                    WCodigo = rstMovlab!Codigo
                    WMovi = rstMovlab!Movi
                    WTipomov = rstMovlab!Tipomov
                    WObservaciones = rstMovlab!Observaciones
                    WLote = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)
                    Rem WSaldo = IIf(IsNull(rstMovlab!Saldo), "0", rstMovlab!Saldo)
                    
                    If Val(WLote) = Val(nrolote) Then
                        
                        With rstFichaMat
                    
                            .AddNew
                            !Articulo = WArticulo
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            If WMovi = "E" Then
                                !Entrada = WCantidad
                                !Salida = 0
                                    Else
                                !Entrada = 0
                                !Salida = WCantidad
                            End If
                            !Observaciones = WObservaciones
                            !Descripcion = WDesArticulo
                            !Lista1 = "Mov.Lab"
                            !Lista2 = ""
                            !Lote = WLote
                            !Saldo = WSaldo
                            !Empresa = NombreEmpresa
                            !PartiOri = WPartiOri
                            .Update
                        End With
                    End If
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstMovlab.Close
    End If
    
    Rem PROCESA LAS VENTAS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spEstadistica = "ListaEstadisticaArticuloDesdeHasta" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstEstadistica!Marca = "X" Then
                
                        Else
                
                    If rstEstadistica!TipoproDy = "M" And rstEstadistica!ArticuloDy = Articulo.Text Then
                    
                        WArticulo = rstEstadistica!ArticuloDy
                        WFecha = rstEstadistica!Fecha
                        WCodigo = rstEstadistica!Numero
                        WObservaciones = rstEstadistica!Cliente
                        WTipo = rstEstadistica!Tipo
                        
                        XLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                        XLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                        XLote(2, 1) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                        XLote(2, 2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                        XLote(3, 1) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                        XLote(3, 2) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                        XLote(4, 1) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                        XLote(4, 2) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                        XLote(5, 1) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                        XLote(5, 2) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                        
                        WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                        
                        If Len(Trim(WLoteAdicional)) = 98 Then
                            XLote(6, 1) = Mid$(WLoteAdicional, 1, 8)
                            XLote(6, 2) = Mid$(WLoteAdicional, 9, 6)
                            XLote(7, 1) = Mid$(WLoteAdicional, 15, 8)
                            XLote(7, 2) = Mid$(WLoteAdicional, 23, 6)
                            XLote(8, 1) = Mid$(WLoteAdicional, 29, 8)
                            XLote(8, 2) = Mid$(WLoteAdicional, 37, 6)
                            XLote(9, 1) = Mid$(WLoteAdicional, 43, 8)
                            XLote(9, 2) = Mid$(WLoteAdicional, 51, 6)
                            XLote(10, 1) = Mid$(WLoteAdicional, 57, 8)
                            XLote(10, 2) = Mid$(WLoteAdicional, 65, 6)
                            XLote(11, 1) = Mid$(WLoteAdicional, 71, 8)
                            XLote(11, 2) = Mid$(WLoteAdicional, 79, 6)
                            XLote(12, 1) = Mid$(WLoteAdicional, 85, 8)
                            XLote(12, 2) = Mid$(WLoteAdicional, 93, 6)
                                Else
                            XLote(6, 1) = "0"
                            XLote(6, 2) = "0"
                            XLote(7, 1) = "0"
                            XLote(7, 2) = "0"
                            XLote(8, 1) = "0"
                            XLote(8, 2) = "0"
                            XLote(9, 1) = "0"
                            XLote(9, 2) = "0"
                            XLote(10, 1) = "0"
                            XLote(10, 2) = "0"
                            XLote(11, 1) = "0"
                            XLote(11, 2) = "0"
                            XLote(12, 1) = "0"
                            XLote(12, 2) = "0"
                        End If
                        
                        For Da = 1 To 12
                        
                            WLote = XLote(Da, 1)
                            WCantidad = XLote(Da, 2)
                    
                            If Val(WLote) = Val(nrolote) Then
                        
                                With rstFichaMat
                    
                                    .AddNew
                                    !Articulo = WArticulo
                                    !Fecha = WFecha
                                    !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                    !Tipo = 0
                                    !Numero = WCodigo
                                    !Inicial = 0
                                    If WTipo = 2 Then
                                        !Entrada = Abs(Val(WCantidad))
                                        !Salida = 0
                                        !Lista1 = "Devol."
                                            Else
                                        !Entrada = 0
                                        !Salida = WCantidad
                                        !Lista1 = "Factura"
                                    End If
                                    !Observaciones = WObservaciones
                                    !Descripcion = ""
                                    !Lista2 = ""
                                    !Lote = WLote
                                    !Saldo = 0
                                    !Empresa = NombreEmpresa
                                    !PartiOri = WPartiOri
                                    .Update
                                End With
                            End If
                            
                        Next Da
                        
                    End If
                
                End If
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
            
        End With
        rstEstadistica.Close
    End If
    
    Da = 0
    With rstFichaMat
        .Index = "Articulo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                WArticulo = !Articulo
                WObservaciones = !Observaciones
                WDescripcion = ""
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                If !Lista1 = "Devol." Or !Lista1 = "Factura" Then
                    spCliente = "ConsultaCliente" + "'" + WObservaciones + "'"
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        WObservaciones = rstCliente!Razon
                        rstCliente.Close
                    End If
                End If
                !Descripcion = WDescripcion
                !Observaciones = WObservaciones
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    If Left$(WArticulo, 2) = "DY" Or Left$(WArticulo, 2) = "DS" Or Left$(WArticulo, 2) = "DQ" Then
        Listado.ReportFileName = "WLotematdy.rpt"
            Else
        Listado.ReportFileName = "WLotemat.rpt"
    End If

    Listado.WindowTitle = "Listado de Ficha Lote de Materias Primas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.Destination = 0
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
    Listado.Action = 1
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_FichaMat
End Sub

Private Sub pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    
    Indice = Pantalla.ListIndex
    WArticulo = WIndice.List(Indice)
    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Articulo.Text = rstArticulo!Codigo
        DesArticulo.Caption = rstArticulo!Descripcion
        rstArticulo.Close
        Call Proceso_Click
        WVector1.SetFocus
            Else
        Articulo.Text = WArticulo
    End If
    Articulo.SetFocus
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector

    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgConsFicMatDyTransito.Caption = "Consulta de Ficha de Stock de Materias Primas :  " + !Nombre
            NombreEmpresa = !Nombre
        End If
    End With
    
    Tipo.Clear
    
    Tipo.AddItem "Con Saldo"
    Tipo.AddItem "Todos los movimientos"
    
    Tipo.ListIndex = 0
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    Articulo.Text = UCase(Articulo.Text)
    
    WXInicial = 0
    WXEntradas = 0
    WXSalidas = 0
    WXStock = 0

    Renglon = 0
    
    spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
                
        WOrden = Str$(rstArticulo!Pedido)
                
        WArticulo = rstArticulo!Codigo
        WInicial = rstArticulo!Inicial
        
        WFechaCierre = IIf(IsNull(rstArticulo!FechaCierre), "00/00/0000", rstArticulo!FechaCierre)
        WOrdFechaCierre = IIf(IsNull(rstArticulo!OrdFechaCierre), "00000000", rstArticulo!OrdFechaCierre)
                                        
        Renglon = Renglon + 1
        WVector1.Row = Renglon
                   
        WVector1.Col = 1
        WVector1.Text = WFechaCierre
                        
        WVector1.Col = 2
        WVector1.Text = ""
                        
        WVector1.Col = 3
        WVector1.Text = ""
                        
        WVector1.Col = 4
        WVector1.Text = "Saldo Inicial"
                        
        WVector1.Col = 5
        WVector1.Text = Pusing("###,###", Str$(rstArticulo!Inicial))
                
        WVector1.Col = 6
        WVector1.Text = ""
                
        WXInicial = rstArticulo!Inicial
        
        rstArticulo.Close
                
    End If
    
                
    Rem PROCESA LOS LAUDOS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                 + Articulo.Text + "'"
    spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
        With rstLaudo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstLaudo!Marca = "X" And rstLaudo!Saldo = 0 Then
                
                        Else
                    
                    If rstLaudo!Articulo = Articulo.Text Then
                
                        WArticulo = rstLaudo!Articulo
                        WCantidad = rstLaudo!Liberada
                        WFecha = rstLaudo!Fecha
                        WLaudo = rstLaudo!Laudo
                        WWOrden = rstLaudo!Orden
                        WDevuelta = IIf(IsNull(rstLaudo!devuelta), "0", rstLaudo!devuelta)
                        WRechazo = IIf(IsNull(rstLaudo!Rechazo), "0", rstLaudo!Rechazo)
                        WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                        WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                        Call Redondeo(WSaldo)
                        WPartiOri = rstLaudo!PartiOri
                        WEnvase = IIf(IsNull(rstLaudo!Envase), "", rstLaudo!Envase)
                        WTransito = IIf(IsNull(rstLaudo!Transito), "", rstLaudo!Transito)
                        WSaldoTransito = IIf(IsNull(rstLaudo!SaldoTransito), "0", rstLaudo!SaldoTransito)
                        
                        If WLiberada <> 0 Or WSaldoTransito <> 0 Then
                
                            Lugar = Lugar + 1
                        
                            Vector(Lugar, 1) = !Fecha
                            Vector(Lugar, 2) = "Laudo"
                            Vector(Lugar, 3) = WLaudo
                            Vector(Lugar, 4) = WDEsProveedor
                            Vector(Lugar, 5) = Pusing("###,###", Str$(WLiberada))
                            Vector(Lugar, 6) = ""
                            Vector(Lugar, 7) = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                            Vector(Lugar, 8) = WWOrden
                            If Left$(Articulo.Text, 2) = "DY" Or Left$(Articulo.Text, 2) = "DS" Or Left$(Articulo.Text, 2) = "DQ" Then
                                Vector(Lugar, 9) = Left$(WPartiOri, 10)
                                    Else
                                Vector(Lugar, 9) = WLaudo
                            End If
                            Vector(Lugar, 10) = WSaldo
                            Vector(Lugar, 11) = WEnvase
                            Vector(Lugar, 12) = WTransito
                            Vector(Lugar, 13) = WSaldoTransito
                    
                            WXEntradas = WXEntradas + WLiberada
                            
                        End If
                        
                        If WDevuelta <> 0 Then
                
                            Lugar = Lugar + 1
                        
                            Vector(Lugar, 1) = !Fecha
                            Vector(Lugar, 2) = "Rechazo"
                            Vector(Lugar, 3) = WRechazo
                            Vector(Lugar, 4) = WDEsProveedor
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = "(" + Pusing("###,###", Str$(rstLaudo!devuelta)) + ")"
                            Vector(Lugar, 7) = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                            Vector(Lugar, 8) = WWOrden
                            Vector(Lugar, 9) = WRechazo
                            Vector(Lugar, 10) = "0"
                            Vector(Lugar, 11) = ""
                            Vector(Lugar, 12) = ""
                            Vector(Lugar, 13) = ""
                            
                        End If
                
                    End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstLaudo.Close
    End If
    
    For Ciclo = 1 To Lugar
    
        WWOrden = Vector(Ciclo, 8)
        
        spOrden = "ListaOrden" + "'" + WWOrden + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WProveedor = rstOrden!Proveedor
            rstOrden.Close
        End If
        
        WDEsProveedor = ""
                
        spProveedor = "ConsultaProveedores" + "'" + WProveedor + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            WDEsProveedor = RstProveedor!Nombre
            RstProveedor.Close
        End If
    
        Vector(Ciclo, 4) = WDEsProveedor
        
        spEnvase = "ConsultaEnvases " + "'" + Vector(Ciclo, 11) + "'"
        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvase.RecordCount > 0 Then
            Vector(Ciclo, 11) = rstEnvase!Abreviatura
            rstEnvase.Close
        End If
        
    Next Ciclo
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(dada, 7) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                Auxi11 = Vector(Ciclo, 11)
                Auxi12 = Vector(Ciclo, 12)
                Auxi13 = Vector(Ciclo, 13)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                Vector(Ciclo, 11) = Vector(dada, 11)
                Vector(Ciclo, 12) = Vector(dada, 12)
                Vector(Ciclo, 13) = Vector(dada, 13)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10
                Vector(dada, 11) = Auxi11
                Vector(dada, 12) = Auxi12
                Vector(dada, 13) = Auxi13

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        WSaldo = Val(Vector(Cicla, 10))
        Call Redondeo(WSaldo)
        If Tipo.ListIndex = 1 Or WSaldo <> 0 Then
    
            Renglon = Renglon + 1
            WVector1.Row = Renglon
                
            WVector1.Col = 1
            WVector1.Text = Vector(Cicla, 1)
                        
            WVector1.Col = 2
            WVector1.Text = Vector(Cicla, 2)
                                               
            WVector1.Col = 3
            WVector1.Text = Vector(Cicla, 3)
                        
            WVector1.Col = 4
            WVector1.Text = Vector(Cicla, 4)
                        
            WVector1.Col = 5
            WVector1.Text = Vector(Cicla, 5)
                
            WVector1.Col = 6
            WVector1.Text = Vector(Cicla, 6)
        
            WVector1.Col = 7
            WVector1.Text = Vector(Cicla, 9)
        
            WSaldo = Val(Vector(Cicla, 10))
            Call Redondeo(WSaldo)
            WVector1.Col = 8
            WVector1.Text = Str$(WSaldo)
            
            WVector1.Col = 9
            WVector1.Text = Vector(Cicla, 11)
            
            WVector1.Col = 10
            WVector1.Text = Vector(Cicla, 12)
            
            WVector1.Col = 11
            If Val(Vector(Cicla, 13)) <> 0 Then
                WVector1.Text = Vector(Cicla, 13)
                    Else
                WVector1.Text = ""
            End If
            
        End If
    
    Next Cicla
    
    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                 + Articulo.Text + "'"
    spHoja = "ListaHojaArticuloDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                If rstHoja!Marca = "X" Or XFec < WOrdFechaCierre Then
                
                        Else
                        
                    fr = rstHoja!Clave
                        
                    If rstHoja!Tipo = "M" And rstHoja!Articulo = Articulo.Text Then
                    
                
                        XLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                        XLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "0", rstHoja!Canti1)
                        XLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                        XLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "0", rstHoja!Canti2)
                        XLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                        XLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "0", rstHoja!Canti3)
                        
                        If Val(XLote(1, 1)) = 0 Then
                            XLote(1, 1) = rstHoja!Lote
                            XLote(1, 2) = rstHoja!Cantidad
                        End If
                        
                        For Da = 1 To 3
                        
                            If XLote(Da, 2) = "" Then
                                XLote(Da, 2) = "0"
                            End If
                        
                            WCanti = XLote(Da, 2)
                            If WCanti <> 0 Then
                
                                WArticulo = rstHoja!Articulo
                                WCanti = XLote(Da, 2)
                                WFecha = rstHoja!Fecha
                                WHoja = rstHoja!Hoja
                                WLote = XLote(Da, 1)
                        
                                Lugar = Lugar + 1
                        
                                Vector(Lugar, 1) = !Fecha
                                Vector(Lugar, 2) = "Hoja"
                                Vector(Lugar, 3) = WHoja
                                Vector(Lugar, 4) = ""
                                Vector(Lugar, 5) = ""
                                Vector(Lugar, 6) = Pusing("###,###", Str$(WCanti * 1))
                                Vector(Lugar, 7) = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                                Vector(Lugar, 9) = WLote
                                Vector(Lugar, 10) = ""
                        
                                WXSalidas = WXSalidas + WCanti
                                
                            End If
                        Next Da

                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
                If !Articulo > Articulo.Text Then
                    Exit Do
                End If
                
            Loop
            End If
        
        End With
    End If
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(dada, 7) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        WSaldo = Val(Vector(Cicla, 10))
        Call Redondeo(WSaldo)
        If Tipo.ListIndex = 1 Or WSaldo <> 0 Then
    
            Renglon = Renglon + 1
            WVector1.Row = Renglon
                
            WVector1.Col = 1
            WVector1.Text = Vector(Cicla, 1)
                        
            WVector1.Col = 2
            WVector1.Text = Vector(Cicla, 2)
                                               
            WVector1.Col = 3
            WVector1.Text = Vector(Cicla, 3)
                        
            WVector1.Col = 4
            WVector1.Text = Vector(Cicla, 4)
                        
            WVector1.Col = 5
            WVector1.Text = Vector(Cicla, 5)
                
            WVector1.Col = 6
            WVector1.Text = Vector(Cicla, 6)
        
            WVector1.Col = 7
            WVector1.Text = Vector(Cicla, 9)
        
            WSaldo = Val(Vector(Cicla, 10))
            Call Redondeo(WSaldo)
            WVector1.Col = 8
            WVector1.Text = Str$(WSaldo)
            
        End If
    
    Next Cicla
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                + Articulo.Text + "'"
    spMovvar = "ListaMovvarArticuloDesdeHasta" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then

        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovvar!Marca = "X" Then
                
                        Else
                        
                    If rstMovvar!Tipo = "M" And rstMovvar!Articulo = Articulo.Text Then
                    
                        WArticulo = rstMovvar!Articulo
                        WCantidad = rstMovvar!Cantidad
                        WFecha = rstMovvar!Fecha
                        WCodigo = rstMovvar!Codigo
                        WMovi = rstMovvar!Movi
                        
                        Lugar = Lugar + 1
                        
                        Vector(Lugar, 1) = rstMovvar!Fecha
                        If rstMovvar!Tipomov = 0 Or rstMovvar!Tipomov = 1 Then
                            Vector(Lugar, 2) = "Mov.Var"
                                Else
                            Vector(Lugar, 2) = "Guia In"
                        End If
                        Vector(Lugar, 3) = WCodigo
                        Vector(Lugar, 4) = rstMovvar!Observaciones
                        If rstMovvar!Movi = "E" Then
                            Vector(Lugar, 5) = Pusing("###,###", Str$(rstMovvar!Cantidad))
                            Vector(Lugar, 6) = ""
                            WXEntradas = WXEntradas + rstMovvar!Cantidad
                                Else
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = Pusing("###,###", Str$(rstMovvar!Cantidad))
                            WXSalidas = WXSalidas + rstMovvar!Cantidad
                        End If
                        Vector(Lugar, 7) = Right$(rstMovvar!Fecha, 4) + Mid$(rstMovvar!Fecha, 4, 2) + Left$(rstMovvar!Fecha, 2)
                        Vector(Lugar, 9) = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                        Vector(Lugar, 10) = ""
                        
                    End If
                End If
                
                .MoveNext
            
                If .EOF = True Then
                    Exit Do
                End If
                                                                            
            Loop
            End If
            
        End With
    End If
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(dada, 7) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        WSaldo = Val(Vector(Cicla, 10))
        Call Redondeo(WSaldo)
        If Tipo.ListIndex = 1 Or WSaldo <> 0 Then
    
            Renglon = Renglon + 1
            WVector1.Row = Renglon
                
            WVector1.Col = 1
            WVector1.Text = Vector(Cicla, 1)
                        
            WVector1.Col = 2
            WVector1.Text = Vector(Cicla, 2)
                                               
            WVector1.Col = 3
            WVector1.Text = Vector(Cicla, 3)
                        
            WVector1.Col = 4
            WVector1.Text = Vector(Cicla, 4)
                        
            WVector1.Col = 5
            WVector1.Text = Vector(Cicla, 5)
                
            WVector1.Col = 6
            WVector1.Text = Vector(Cicla, 6)
        
            WVector1.Col = 7
            WVector1.Text = Vector(Cicla, 9)
        
            WSaldo = Val(Vector(Cicla, 10))
            Call Redondeo(WSaldo)
            WVector1.Col = 8
            WVector1.Text = Str$(WSaldo)
            
        End If
    
    Next Cicla
    
    Rem PROCESA LAS GUIAS DE TRASLADO INTERNOS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                + Articulo.Text + "'"
    spMovguia = "ListaMovguiaArticuloDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then

        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                        
                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = Articulo.Text Then
                    
                        WArticulo = rstMovguia!Articulo
                        WCantidad = rstMovguia!Cantidad
                        WFecha = rstMovguia!Fecha
                        WCodigo = rstMovguia!Codigo
                        WMovi = rstMovguia!Movi
                        WDestino = rstMovguia!Destino
                        WTipomov = rstMovguia!Tipomov
                        
                        Lugar = Lugar + 1
                        
                        Vector(Lugar, 1) = rstMovguia!Fecha
                        If Val(WCodigo) > 900000 Then
                            Vector(Lugar, 2) = "Prestamo"
                            Vector(Lugar, 3) = WCodigo - 900000
                                Else
                            Vector(Lugar, 2) = "Guia In"
                            Vector(Lugar, 3) = WCodigo
                        End If
                        Rem Vector(Lugar, 4) = rstMovguia!Observaciones
                                
                        If rstMovguia!Movi = "E" Then
                            Select Case WTipomov
                                Case 1
                                    Vector(Lugar, 4) = "Recepcion de Surfactan"
                                Case 2
                                    Vector(Lugar, 4) = "Recepcion de Pellital"
                                Case 3
                                    Vector(Lugar, 4) = "Recepcion de Surfactan II"
                                Case 4
                                    Vector(Lugar, 4) = "Recepcion de Pellital II"
                                Case 5
                                    Vector(Lugar, 4) = "Recepcion de Surfactan III"
                                Case 6
                                    Vector(Lugar, 4) = "Recepcion de Surfactan IV"
                                Case 7
                                    Vector(Lugar, 4) = "Recepcion de Surfactan V"
                                Case 8
                                    Vector(Lugar, 4) = "Recepcion de Pellital V"
                                Case 9
                                    Vector(Lugar, 4) = "Recepcion de Pellital IV"
                                Case 10
                                    Vector(Lugar, 4) = "Recepcion de Surfactan VI"
                                Case 11
                                    Vector(Lugar, 4) = "Recepcion de Surfactan VII"
                                Case Else
                            End Select
                            Vector(Lugar, 5) = Pusing("###,###", Str$(rstMovguia!Cantidad))
                            Vector(Lugar, 6) = ""
                            WXEntradas = WXEntradas + rstMovguia!Cantidad
                            WPartiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                            If Trim(WPartiOri) <> "" Then
                                Vector(Lugar, 9) = WPartiOri
                                    Else
                                Vector(Lugar, 9) = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                            End If
                            Vector(Lugar, 10) = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                Else
                            Select Case WDestino
                                Case 1
                                    Vector(Lugar, 4) = "Envio a Surfactan"
                                Case 2
                                    Vector(Lugar, 4) = "Envio a Pellital"
                                Case 3
                                    Vector(Lugar, 4) = "Envio a Surfactan II"
                                Case 4
                                    Vector(Lugar, 4) = "Envio a Pellital II"
                                Case 5
                                    Vector(Lugar, 4) = "Envio a Surfactan III"
                                Case 6
                                    Vector(Lugar, 4) = "Envio a Surfactan IV"
                                Case 7
                                    Vector(Lugar, 4) = "Envio a Surfactan V"
                                Case 8
                                    Vector(Lugar, 4) = "Envio a Pellital V"
                                Case 9
                                    Vector(Lugar, 4) = "Envio a Pellital IV"
                                Case 10
                                    Vector(Lugar, 4) = "Envio a Surfactan VI"
                                Case 11
                                    Vector(Lugar, 4) = "Envio a Surfactan VII"
                                Case Else
                            End Select
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = Pusing("###,###", Str$(rstMovguia!Cantidad))
                            WXSalidas = WXSalidas + rstMovguia!Cantidad
                            WPartiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                            If Trim(WPartiOri) <> "" Then
                                Vector(Lugar, 9) = WPartiOri
                                    Else
                                Vector(Lugar, 9) = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                            End If
                            Vector(Lugar, 10) = ""
                        End If
                        Vector(Lugar, 7) = Right$(rstMovguia!Fecha, 4) + Mid$(rstMovguia!Fecha, 4, 2) + Left$(rstMovguia!Fecha, 2)
                        
                        
                    End If
                End If
                
                .MoveNext
            
                If .EOF = True Then
                    Exit Do
                End If
                                                                            
            Loop
            End If
            
        End With
        rstMovguia.Close
    End If
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(dada, 7) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        WSaldo = Val(Vector(Cicla, 10))
        Call Redondeo(WSaldo)
        If Tipo.ListIndex = 1 Or WSaldo <> 0 Then
    
            Renglon = Renglon + 1
            WVector1.Row = Renglon
                
            WVector1.Col = 1
            WVector1.Text = Vector(Cicla, 1)
                        
            WVector1.Col = 2
            WVector1.Text = Vector(Cicla, 2)
                                               
            WVector1.Col = 3
            WVector1.Text = Vector(Cicla, 3)
                        
            WVector1.Col = 4
            WVector1.Text = Vector(Cicla, 4)
                        
            WVector1.Col = 5
            WVector1.Text = Vector(Cicla, 5)
                
            WVector1.Col = 6
            WVector1.Text = Vector(Cicla, 6)
    
            WVector1.Col = 7
            WVector1.Text = Vector(Cicla, 9)
                    
            WSaldo = Val(Vector(Cicla, 10))
            Call Redondeo(WSaldo)
            WVector1.Col = 8
            WVector1.Text = Str$(WSaldo)
            
        End If
    
    Next Cicla
    
    Rem PROCESA LAS HOJAS DE LABORATORIO
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                 + Articulo.Text + "'"
    
    spMovlab = "ListaMovlabArticuloDesdeHasta" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovlab!Marca = "X" Then
                
                        Else
                
                    If rstMovlab!Tipo = "M" And rstMovlab!Articulo = Articulo.Text Then
                
                        WArticulo = rstMovlab!Articulo
                        WCantidad = rstMovlab!Cantidad
                        WFecha = rstMovlab!Fecha
                        WCodigo = rstMovlab!Codigo
                        WMovi = rstMovlab!Movi
                        
                        Lugar = Lugar + 1
                        
                        Vector(Lugar, 1) = rstMovlab!Fecha
                        Vector(Lugar, 2) = "Mov.Lab"
                        Vector(Lugar, 3) = WCodigo
                        Vector(Lugar, 4) = rstMovlab!Observaciones
                        If rstMovlab!Movi = "E" Then
                            Vector(Lugar, 5) = Pusing("###,###", Str$(rstMovlab!Cantidad))
                            Vector(Lugar, 6) = ""
                            WXEntradas = WXEntradas + rstMovlab!Cantidad
                                Else
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = Pusing("###,###", Str$(rstMovlab!Cantidad))
                            WXSalidas = WXSalidas + rstMovlab!Cantidad
                        End If
                        Vector(Lugar, 7) = Right$(rstMovlab!Fecha, 4) + Mid$(rstMovlab!Fecha, 4, 2) + Left$(rstMovlab!Fecha, 2)
                        Vector(Lugar, 9) = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)
                        Vector(Lugar, 10) = ""
                        
                    End If
                
                End If
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
            
        End With
    End If
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(dada, 7) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        WSaldo = Val(Vector(Cicla, 10))
        Call Redondeo(WSaldo)
        If Tipo.ListIndex = 1 Or WSaldo <> 0 Then
        
            Renglon = Renglon + 1
            WVector1.Row = Renglon
                
            WVector1.Col = 1
            WVector1.Text = Vector(Cicla, 1)
                        
            WVector1.Col = 2
            WVector1.Text = Vector(Cicla, 2)
                                               
            WVector1.Col = 3
            WVector1.Text = Vector(Cicla, 3)
                        
            WVector1.Col = 4
            WVector1.Text = Vector(Cicla, 4)
                        
            WVector1.Col = 5
            WVector1.Text = Vector(Cicla, 5)
                
            WVector1.Col = 6
            WVector1.Text = Vector(Cicla, 6)
        
            WVector1.Col = 7
            WVector1.Text = Vector(Cicla, 9)
        
            WSaldo = Val(Vector(Cicla, 10))
            Call Redondeo(WSaldo)
            WVector1.Col = 8
            WVector1.Text = Str$(WSaldo)
            
        End If
    
    Next Cicla
    
    Rem PROCESA LAS VENTAS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                 + Articulo.Text + "'"
    
    spEstadistica = "ListaEstadisticaArticuloDesdeHasta" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstEstadistica!Marca = "X" Then
                
                        Else
                
                    If rstEstadistica!TipoproDy = "M" And rstEstadistica!ArticuloDy = Articulo.Text Then
                
                        WArticulo = rstEstadistica!ArticuloDy
                        WFecha = rstEstadistica!Fecha
                        WCodigo = rstEstadistica!Numero
                        
                        XLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                        XLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                        XLote(2, 1) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                        XLote(2, 2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                        XLote(3, 1) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                        XLote(3, 2) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                        XLote(4, 1) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                        XLote(4, 2) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                        XLote(5, 1) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                        XLote(5, 2) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                        
                        WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                        
                        If Len(Trim(WLoteAdicional)) = 98 Then
                            XLote(6, 1) = Mid$(WLoteAdicional, 1, 8)
                            XLote(6, 2) = Mid$(WLoteAdicional, 9, 6)
                            XLote(7, 1) = Mid$(WLoteAdicional, 15, 8)
                            XLote(7, 2) = Mid$(WLoteAdicional, 23, 6)
                            XLote(8, 1) = Mid$(WLoteAdicional, 29, 8)
                            XLote(8, 2) = Mid$(WLoteAdicional, 37, 6)
                            XLote(9, 1) = Mid$(WLoteAdicional, 43, 8)
                            XLote(9, 2) = Mid$(WLoteAdicional, 51, 6)
                            XLote(10, 1) = Mid$(WLoteAdicional, 57, 8)
                            XLote(10, 2) = Mid$(WLoteAdicional, 65, 6)
                            XLote(11, 1) = Mid$(WLoteAdicional, 71, 8)
                            XLote(11, 2) = Mid$(WLoteAdicional, 79, 6)
                            XLote(12, 1) = Mid$(WLoteAdicional, 85, 8)
                            XLote(12, 2) = Mid$(WLoteAdicional, 93, 6)
                                Else
                            XLote(6, 1) = "0"
                            XLote(6, 2) = "0"
                            XLote(7, 1) = "0"
                            XLote(7, 2) = "0"
                            XLote(8, 1) = "0"
                            XLote(8, 2) = "0"
                            XLote(9, 1) = "0"
                            XLote(9, 2) = "0"
                            XLote(10, 1) = "0"
                            XLote(10, 2) = "0"
                            XLote(11, 1) = "0"
                            XLote(11, 2) = "0"
                            XLote(12, 1) = "0"
                            XLote(12, 2) = "0"
                        End If
                        
                        For Da = 1 To 12
                        
                            WLote = XLote(Da, 1)
                            WCanti = Val(XLote(Da, 2))
                        
                            If WCanti <> 0 Then
                                WCantidad = WCanti
                                Lugar = Lugar + 1
                                Vector(Lugar, 1) = WFecha
                                If rstEstadistica!Tipo = 1 Then
                                    Vector(Lugar, 2) = "Factura"
                                        Else
                                    Vector(Lugar, 2) = "Devol"
                                End If
                                Vector(Lugar, 3) = WCodigo
                                Vector(Lugar, 4) = rstEstadistica!Cliente
                                If rstEstadistica!Tipo = 2 Then
                                    Vector(Lugar, 5) = Pusing("###,###", Str$(WCantidad))
                                    Vector(Lugar, 6) = ""
                                    WXEntradas = WXEntradas + WCantidad
                                        Else
                                    Vector(Lugar, 5) = ""
                                    Vector(Lugar, 6) = Pusing("###,###", Str$(WCantidad))
                                    WXSalidas = WXSalidas + WCantidad
                                End If
                                Vector(Lugar, 7) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                Vector(Lugar, 9) = WLote
                                Vector(Lugar, 10) = ""
                            End If
                            
                        Next Da
                        
                    End If
                
                End If
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
            
        End With
        rstEstadistica.Close
    End If
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(dada, 7) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        WSaldo = Val(Vector(Cicla, 10))
        Call Redondeo(WSaldo)
        If Tipo.ListIndex = 1 Or WSaldo <> 0 Then
        
            Renglon = Renglon + 1
            WVector1.Row = Renglon
                
            WVector1.Col = 1
            WVector1.Text = Vector(Cicla, 1)
                        
            WVector1.Col = 2
            WVector1.Text = Vector(Cicla, 2)
                                               
            WVector1.Col = 3
            WVector1.Text = Vector(Cicla, 3)
        
            WVector1.Col = 4
            spCliente = "ConsultaCliente" + "'" + Vector(Cicla, 4) + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                WVector1.Text = rstCliente!Razon
                    Else
                WVector1.Text = ""
            End If
                        
            WVector1.Col = 5
            WVector1.Text = Vector(Cicla, 5)
                
            WVector1.Col = 6
            WVector1.Text = Vector(Cicla, 6)
            
            If Left$(Articulo.Text, 2) = "DY" Or Left$(Articulo.Text, 2) = "DS" Or Left$(Articulo.Text, 2) = "DQ" Then
            
                WNroLaudo = Vector(Cicla, 9)
                WEntra = "N"
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Laudo"
                ZSql = ZSql + " Where Laudo.Articulo = " + "'" + Articulo.Text + "'"
                ZSql = ZSql + " and Laudo.Laudo = " + "'" + WNroLaudo + "'"
                spLaudo = ZSql
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WVector1.Col = 7
                    WVector1.Text = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                        
                If WEntra = "N" Then
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Guia"
                    ZSql = ZSql + " Where Guia.Articulo = " + "'" + Articulo.Text + "'"
                    ZSql = ZSql + " and Guia.Lote = " + "'" + WNroLaudo + "'"
                    spMovguia = ZSql
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WVector1.Col = 7
                        WVector1.Text = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WVector1.Col = 7
                WVector1.Text = Vector(Cicla, 9)
                
            End If
            
            Rem WVector1.Col = 7
            Rem WVector1.Text = Vector(Cicla, 9)
        
            WSaldo = Val(Vector(Cicla, 10))
            Call Redondeo(WSaldo)
            WVector1.Col = 8
            WVector1.Text = Str$(WSaldo)
            
        End If
        
    Next Cicla
    
    WXStock = WXInicial + WXEntradas - WXSalidas
    XEmpresa = WEmpresa
    
    Rem XInicial.Text = Pusing("###,###", Str$(WXInicial))
    XEntradas.Text = Pusing("###,###", Str$(WXEntradas))
    XSalidas.Text = Pusing("###,###", Str$(WXSalidas))
    XStock.Text = Pusing("###,###", Str$(WXStock))
    
    Orden.Text = Pusing("###,###", WOrden)
    
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector1.TopRow = 1
    
    Exit Sub
    
Control_error:
    Rem MsgBox Err.Description
    WSalidaError = "N"
    AvisoError.Visible = True
    WLabel4.Visible = False
    Label8.Visible = False
    WStock4.Visible = False
    Disponible.Visible = False
    Label2.Visible = False
    Orden.Visible = False
    Resume Next
    
End Sub

Private Sub Articulo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Articulo.Text = UCase(Articulo.Text)
        WArticulo = Articulo.Text
        Articulo.Text = WArticulo
        
        spArticulo = "ConsultaArticulo" + "'" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DesArticulo.Caption = rstArticulo!Descripcion
            rstArticulo.Close
            Call Proceso_Click
            WVector1.SetFocus
                Else
            Articulo.SetFocus
        End If
    End If
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear
    WVector1.Font.Bold = True
    
    WVector1.FixedCols = 1
    WVector1.Cols = 12
    WVector1.FixedRows = 1
    WVector1.Rows = 1001
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                WVector1.Text = "Entrada"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector1.Text = "Salida"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector1.Text = "Partida"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                WVector1.Text = "Saldo"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 9
                WVector1.Text = "Envase"
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 10
                WVector1.Text = "Transito"
                WVector1.ColWidth(Ciclo) = 2100
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 11
                WVector1.Text = "Saldo"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
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
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
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

Private Sub Labo_Click()

    Erase WWVector
    WWRenglon = 0
    
    spInforme = "ModificaInformeProcesoSaldo"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
    XParam = "'" + "20020101" + "'"
    spInforme = "ModificaInformeProceso0 " + XParam
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
    XParam = "'" + Articulo.Text + "'"
    spInforme = "ListaInformeArticulo " + XParam
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        With rstInforme
            .MoveFirst
            Do
                If .EOF = False Then
                    If !FechaOrd > "20020101" Then
                        If rstInforme!Articulo = Articulo.Text Then
                            WWRenglon = WWRenglon + 1
                            WWVector(WWRenglon, 1) = rstInforme!Clave
                            WWVector(WWRenglon, 2) = rstInforme!Informe
                            WWVector(WWRenglon, 3) = rstInforme!Articulo
                            WWVector(WWRenglon, 4) = rstInforme!Cantidad
                        End If
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstInforme.Close
    End If
    
    For Ciclo = 1 To WWRenglon
    
        WClave = WWVector(Ciclo, 1)
        WInforme = WWVector(Ciclo, 2)
        WArticulo = WWVector(Ciclo, 3)
        WCantidad = Val(WWVector(Ciclo, 4))
        WResta = 0
    
        XParam = "'" + WInforme + "','" _
                 + WArticulo + "'"
        spLaudo = "ListaLaudoInforme " + XParam
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
    
            With rstLaudo
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        WLiberada = rstLaudo!Liberada
                        WDevuelta = rstLaudo!devuelta
                        WSuma = WLiberada + WDevuelta
                        
                        WLiberadaAnt = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)
                        WDevueltaAnt = IIf(IsNull(rstLaudo!devueltaant), "0", rstLaudo!devueltaant)
                        WSumaAnt = WLiberadaAnt + WDevueltaAnt
                        
                        If WSumaAnt <> 0 Then
                            WResta = WResta + WSumaAnt
                                Else
                            WResta = WResta + WSuma
                        End If
                        
                        Rem WResta = WResta + rstLaudo!Liberada + rstLaudo!Devuelta
                
                        .MoveNext
                
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                End If
            End With
            rstLaudo.Close
        End If
        
        XParam = "'" + WClave + "','" _
                 + Str$(WResta) + "'"
        spInforme = "ModificaInformeProceso " + XParam
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    spInforme = "ModificaInformeProcesoDife"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)

    WDesde = "00000000"
    WHasta = "99999999"
    
    Listado.WindowTitle = "Listado de Informe de Recepcion Pendientes de Aprobacion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Informe.Articulo} in " + Chr$(34) + Articulo.Text + Chr$(34) + " to " + Chr$(34) + Articulo.Text + Chr$(34) + " and {Informe.fechaord} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Informe.Informe, Informe.Fecha, Informe.Remito, Informe.Proveedor, Informe.Orden, Informe.Articulo, Informe.Cantidad, Informe.Fechaord, Informe.CantidadLaudo, Informe.Dife, " _
                        + "Articulo.Descripcion, " _
                        + "Proveedor.Nombre " _
                        + "From " _
                        + DSQ + ".dbo.Informe Informe, " _
                        + DSQ + ".dbo.Articulo Articulo, " _
                        + DSQ + ".dbo.Proveedor Proveedor " _
                        + "Where " _
                        + "Informe.Articulo = Articulo.Codigo AND " _
                        + "Informe.Proveedor = Proveedor.Proveedor AND " _
                        + "Informe.Fechaord >= '" + WDesde + "' AND Informe.Fechaord <= '" + WHasta + "' AND " _
                        + "Informe.Dife <> 0."
                        
    Listado.DataFiles(0) = ""
    Listado.DataFiles(1) = ""
    Listado.DataFiles(2) = ""
    Listado.DataFiles(3) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    WListado = Listado.ReportFileName
    Listado.ReportFileName = "Wlistinfpend.rpt"
    Listado.Action = 1
    Listado.ReportFileName = WListado

End Sub

Private Sub Orden_dblclick()

    XParam = "'" + "'"

    spOrden = "ModificaOrdenSaldo " + XParam
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    Erase WWVector
    WWLugar = 0

    spOrden = "ListaOrdenTotal "
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
    
    With rstOrden
         .MoveFirst
         Do
             If .EOF = False Then
                 WWClave = rstOrden!Clave
                 WWOrden = rstOrden!Orden
                 WWFecha2 = rstOrden!fecha2
                 WWSaldo = Str$(rstOrden!Cantidad - rstOrden!Recibida)
                 If Val(WWSaldo) > 0 Then
                    Entra = "S"
                    For XX = 1 To WWLugar
                        If Val(WWVector(XX, 1)) = WWOrden Then
                            Entra = "N"
                            Exit For
                        End If
                    Next XX
                    
                    If Entra = "S" Then
                        WWLugar = WWLugar + 1
                        WWVector(WWLugar, 1) = WWOrden
                        WWVector(WWLugar, 2) = Right$(WWFecha2, 4) + Mid$(WWFecha2, 4, 2) + Left$(WWFecha2, 2)
                    End If
                    
                 End If
                 .MoveNext
                     Else
                 Exit Do
             End If
         Loop
    End With
    rstOrden.Close
    
    End If
    
    For XX = 1 To WWLugar
        WWOrden = WWVector(XX, 1)
        WWFecha2 = WWVector(XX, 2)
        XParam = "'" + WWOrden + "','" _
                     + WWFecha2 + "'"
    
        spOrden = "ModificaOrdenFecha2 " + XParam
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    Next XX
    
    Listado.WindowTitle = "Listado de Ordenes Pendientes por Articulo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Orden.Articulo} in " + Chr$(34) + Articulo.Text + Chr$(34) + " to " + Chr$(34) + Articulo.Text + Chr$(34)
   
    Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Orden.Orden, Orden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Fecha2, Orden.Saldo, Orden.OrdFecha2, Proveedor.Nombre, Articulo.Descripcion " _
                        + "From " + DSQ + ".dbo.Orden Orden, " _
                        + DSQ + ".dbo.Proveedor Proveedor, " _
                        + DSQ + ".dbo.Articulo Articulo " _
                        + "Where Orden.Proveedor = Proveedor.Proveedor AND Orden.Articulo = Articulo.Codigo AND Orden.Articulo >= '" + Articulo.Text + "' AND Orden.Articulo <= '" + Articulo.Text + "' AND Orden.Saldo > 0. AND Orden.OrdFecha2 >= '00000000' AND Orden.OrdFecha2 <= '9999999'"
    
    Listado.DataFiles(0) = ""
    Listado.DataFiles(1) = ""
    Listado.DataFiles(2) = ""
    Listado.DataFiles(3) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.ReportFileName = "WOrdPenArt.rpt"
    Listado.Action = 1

End Sub


Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spArticulo = "ListaArticuloConsulta"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
    
    With rstArticulo
        .MoveFirst
        Do
            If .EOF = False Then
            
                If Left$(rstArticulo!Codigo, 2) = "DY" Or Left$(rstArticulo!Codigo, 2) = "DS" Or Left$(rstArticulo!Codigo, 2) = "DQ" Then
                    Da = Len(rstArticulo!Descripcion) - WEspacios
                    For Aaa = 1 To Da
                        If Left$(Ayuda.Text, WEspacios) = Mid$(rstArticulo!Descripcion, Aaa, WEspacios) Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                            Exit For
                        End If
                    Next Aaa
                End If
                .MoveNext
                    
                        Else
                        
                Exit Do
                
            End If
        Loop
    End With
    
    rstArticulo.Close
    
    End If
    
    End If

End Sub


