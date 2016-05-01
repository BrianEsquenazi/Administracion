VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAvance 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Avance de Proyectos"
   ClientHeight    =   5940
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   11790
   LinkTopic       =   "Form2"
   ScaleHeight     =   5940
   ScaleWidth      =   11790
   Begin VB.OptionButton Tipo4 
      Caption         =   "Obra - Equipo"
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
      Left            =   9240
      TabIndex        =   31
      Top             =   1200
      Width           =   1575
   End
   Begin VB.OptionButton Tipo3 
      Caption         =   "Obra - Mano de Obra"
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
      Left            =   6840
      TabIndex        =   30
      Top             =   1200
      Width           =   2295
   End
   Begin VB.OptionButton Tipo2 
      Caption         =   "Obra - Material"
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
      Left            =   4800
      TabIndex        =   29
      Top             =   1200
      Width           =   2055
   End
   Begin VB.OptionButton Tipo1 
      Caption         =   "Equipo"
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
      Left            =   3360
      TabIndex        =   28
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Proveedor 
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
      Left            =   1680
      MaxLength       =   11
      TabIndex        =   23
      Text            =   " "
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Importe 
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
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   21
      Text            =   " "
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Proyecto 
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   16
      Text            =   " "
      Top             =   480
      Width           =   1095
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
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4440
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
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   2040
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox TipoLista 
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
         Left            =   2280
         TabIndex        =   33
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
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
         Left            =   2520
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
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
         Left            =   720
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   285
         Left            =   2280
         TabIndex        =   26
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   285
         Left            =   2280
         TabIndex        =   27
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label3 
         Caption         =   "Tipo"
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
         Left            =   720
         TabIndex        =   32
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Image Acepta 
         Height          =   480
         Left            =   4320
         MouseIcon       =   "Avance.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Avance.frx":030A
         ToolTipText     =   "Confirma la Impresion"
         Top             =   1200
         Width           =   480
      End
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   4320
         MouseIcon       =   "Avance.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "Avance.frx":0A56
         ToolTipText     =   "Cancela la Impresion"
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
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
         Left            =   720
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
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
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   5040
      TabIndex        =   12
      Top             =   1920
      Width           =   3015
      Begin VB.Image Anterior 
         Height          =   480
         Left            =   840
         MouseIcon       =   "Avance.frx":0E98
         MousePointer    =   99  'Custom
         Picture         =   "Avance.frx":11A2
         ToolTipText     =   "Registro Anterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Siguiente 
         Height          =   480
         Left            =   1560
         MouseIcon       =   "Avance.frx":15E4
         MousePointer    =   99  'Custom
         Picture         =   "Avance.frx":18EE
         ToolTipText     =   "Registro Posterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Ultimo 
         Height          =   480
         Left            =   2280
         MouseIcon       =   "Avance.frx":1D30
         MousePointer    =   99  'Custom
         Picture         =   "Avance.frx":203A
         ToolTipText     =   "Ultimo Registro"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Primer 
         Height          =   480
         Left            =   240
         MouseIcon       =   "Avance.frx":247C
         MousePointer    =   99  'Custom
         Picture         =   "Avance.frx":2786
         ToolTipText     =   "Primer Registro"
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   8175
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6960
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Avance.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Efluentes de Lavado"
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
      Left            =   7920
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Descripcion 
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
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1560
      Width           =   9015
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   2160
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   2415
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4260
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4560
      TabIndex        =   19
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label Label5 
      Caption         =   "Proveedor"
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
      TabIndex        =   25
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label DesProveedor 
      BackColor       =   &H00FFFF00&
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
      Height          =   255
      Left            =   3360
      TabIndex        =   24
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label Label4 
      Caption         =   "Importe "
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
      TabIndex        =   22
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Fecha "
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
      Left            =   3000
      TabIndex        =   20
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label DesProyecto 
      BackColor       =   &H00FFFF00&
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
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label6 
      Caption         =   "Proyecto"
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
      TabIndex        =   17
      Top             =   480
      Width           =   1095
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   3480
      MouseIcon       =   "Avance.frx":2BC8
      MousePointer    =   99  'Custom
      Picture         =   "Avance.frx":2ED2
      ToolTipText     =   "Impresion "
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   1800
      MouseIcon       =   "Avance.frx":3714
      MousePointer    =   99  'Custom
      Picture         =   "Avance.frx":3A1E
      ToolTipText     =   "Limpia la pantalla"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   120
      MouseIcon       =   "Avance.frx":4260
      MousePointer    =   99  'Custom
      Picture         =   "Avance.frx":456A
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   960
      MouseIcon       =   "Avance.frx":4DAC
      MousePointer    =   99  'Custom
      Picture         =   "Avance.frx":50B6
      ToolTipText     =   "Elimina el Registro"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   4320
      MouseIcon       =   "Avance.frx":58F8
      MousePointer    =   99  'Custom
      Picture         =   "Avance.frx":5C02
      ToolTipText     =   "Salida"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   2640
      MouseIcon       =   "Avance.frx":6444
      MousePointer    =   99  'Custom
      Picture         =   "Avance.frx":674E
      ToolTipText     =   "Consulta de Datos"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lblLabels 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo "
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
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   1095
   End
End
Attribute VB_Name = "PrgAvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstAvance As Recordset
Dim spAvance As String
Dim rstProyecto As Recordset
Dim spProyecto As String
Dim rstProveedor As Recordset
Dim spProveedor As String

Sub Imprime_Datos()

    sql1 = "Select *"
    Sql2 = " FROM Avance"
    Sql3 = " Where Avance.Codigo = " + "'" + codigo.Text + "'"
    spAvance = sql1 + Sql2 + Sql3
    Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
    If rstAvance.RecordCount > 0 Then
        Fecha.Text = rstAvance!Fecha
        proyecto.Text = Str$(rstAvance!proyecto)
        descripcion.Text = Trim(rstAvance!descripcion)
        proveedor.Text = rstAvance!proveedor
        Importe.Text = Str$(rstAvance!Importe)
        Select Case rstAvance!Tipo
            Case 1
                Tipo1.Value = True
                Tipo2.Value = False
                Tipo3.Value = False
                Tipo4.Value = False
            Case 2
                Tipo1.Value = False
                Tipo2.Value = True
                Tipo3.Value = False
                Tipo4.Value = False
            Case 3
                Tipo1.Value = False
                Tipo2.Value = False
                Tipo3.Value = True
                Tipo4.Value = False
            Case 4
                Tipo1.Value = False
                Tipo2.Value = False
                Tipo3.Value = False
                Tipo4.Value = True
            Case Else
                Tipo1.Value = False
                Tipo2.Value = False
                Tipo3.Value = False
                Tipo4.Value = False
        End Select
        rstAvance.Close
    End If
    
    sql1 = "Select *"
    Sql2 = " FROM Proyecto"
    Sql3 = " Where Proyecto.Codigo = " + "'" + proyecto.Text + "'"
    spProyecto = sql1 + Sql2 + Sql3
    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    If rstProyecto.RecordCount > 0 Then
        DesProyecto.Caption = Trim(rstProyecto!descripcion)
        rstProyecto.Close
            Else
        DesProyecto.Caption = ""
    End If
    
    sql1 = "Select *"
    Sql2 = " FROM Proveedor"
    Sql3 = " Where Proveedor.Proveedor = " + "'" + proveedor.Text + "'"
    spProveedor = sql1 + Sql2 + Sql3
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        DesProveedor.Caption = Trim(rstProveedor!nombre)
        rstProveedor.Close
    End If
    
    Importe.Text = Pusing("###,###.##", Importe.Text)
    
End Sub

Private Sub Acepta_Click()

    WWTitulo = "del " + DesdeFecha.Text + " al " + HastaFecha.Text
    If Val(WEmpresa) = 1 Then
        WDesEmpresa = "SURFACTAN"
            Else
        WDesEmpresa = "PELLITAL"
    End If

    ZSql = ""
    ZSql = ZSql + "UPDATE Avance SET "
    ZSql = ZSql + " Titulo = " + "'" + WWTitulo + "',"
    ZSql = ZSql + " Empresa = " + "'" + WDesEmpresa + "'"
    spAvance = ZSql
    Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)


    WDesde = Right$(DesdeFecha.Text, 4) + Mid$(DesdeFecha.Text, 4, 2) + Left$(DesdeFecha.Text, 2)
    WHasta = Right$(HastaFecha.Text, 4) + Mid$(HastaFecha.Text, 4, 2) + Left$(HastaFecha.Text, 2)
        
    Listado.WindowTitle = "Listado de Avances de Proyectos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Avance.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    If TipoLista.ListIndex = 0 Then
    
        Listado.SQLQuery = "SELECT Avance.Codigo, Avance.Fecha, Avance.OrdFecha, Avance.Proyecto, Avance.Importe, Avance.Descripcion, Avance.Proveedor, Avance.Titulo, Avance.Empresa, Avance.Tipo, " _
                + "Proveedor.Nombre, " _
                + "Proyecto.Descripcion, Proyecto.Planta " _
                + "From " _
                + DSQ + ".dbo.Avance Avance, " _
                + DSQ + ".dbo.Proveedor Proveedor, " _
                + DSQ + ".dbo.Proyecto Proyecto " _
                + "Where " _
                + "Avance.Proveedor = Proveedor.Proveedor AND " _
                + "Avance.Proyecto = Proyecto.Codigo AND " _
                + "Avance.OrdFecha >= '" + WDesde + "' AND " _
                + "Avance.OrdFecha <= '" + WHasta + "'"
        Listado.ReportFileName = "Avance.rpt"
        
            Else
            
        Listado.SQLQuery = "SELECT Avance.Codigo, Avance.Fecha, Avance.OrdFecha, Avance.Proyecto, Avance.Importe, Avance.Descripcion, Avance.Proveedor, Avance.Titulo, Avance.Empresa, Avance.Tipo, " _
                + "Proveedor.Nombre, " _
                + "Proyecto.Descripcion, Proyecto.Planta " _
                + "From " _
                + DSQ + ".dbo.Avance Avance, " _
                + DSQ + ".dbo.Proveedor Proveedor, " _
                + DSQ + ".dbo.Proyecto Proyecto " _
                + "Where " _
                + "Avance.Proveedor = Proveedor.Proveedor AND " _
                + "Avance.Proyecto = Proyecto.Codigo AND " _
                + "Avance.OrdFecha >= '" + WDesde + "' AND " _
                + "Avance.OrdFecha <= '" + WHasta + "'"
        Listado.ReportFileName = "AvanceRubro.rpt"
            
            
    End If
    
    Listado.Connect = Connect()
    Listado.Action = 1
    Frame2.Visible = False
    
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Val(codigo.Text) <> 0 Then
    
        If Tipo1.Value = False And Tipo2.Value = False And Tipo3.Value = False And Tipo4.Value = False Then
            m$ = "Se debe informar el tipo de gasto"
            A% = MsgBox(m$, 0, "Archivo de Avances de Proyecto")
            Exit Sub
        End If
        
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi <> "S" Then
            m$ = "Se debe informar la fecha de Gasto"
            A% = MsgBox(m$, 0, "Archivo de Avances de Proyecto")
            Exit Sub
        End If
        
        
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        
        sql1 = "Select *"
        Sql2 = " FROM Avance"
        Sql3 = " Where Avance.Codigo = " + "'" + codigo.Text + "'"
        spAvance = sql1 + Sql2 + Sql3
        Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
        If rstAvance.RecordCount > 0 Then
            rstAvance.Close
            
            If Tipo1.Value = True Then
                ZTipo = "1"
            End If
            If Tipo2.Value = True Then
                ZTipo = "2"
            End If
            If Tipo3.Value = True Then
                ZTipo = "3"
            End If
            If Tipo4.Value = True Then
                ZTipo = "4"
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Avance SET "
            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
            ZSql = ZSql + " OrdFecha = " + "'" + WOrdFecha + "',"
            ZSql = ZSql + " Proyecto = " + "'" + proyecto.Text + "',"
            ZSql = ZSql + " Tipo = " + "'" + ZTipo + "',"
            ZSql = ZSql + " Importe = " + "'" + Importe.Text + "',"
            ZSql = ZSql + " Descripcion = " + "'" + descripcion.Text + "',"
            ZSql = ZSql + " Proveedor = " + "'" + proveedor.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + codigo.Text + "'"
            spAvance = ZSql
            Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            If Tipo1.Value = True Then
                ZTipo = "1"
            End If
            If Tipo2.Value = True Then
                ZTipo = "2"
            End If
            If Tipo3.Value = True Then
                ZTipo = "3"
            End If
            If Tipo4.Value = True Then
                ZTipo = "4"
            End If
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Avance ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Proyecto ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Proveedor )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + codigo.Text + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + WOrdFecha + "',"
            ZSql = ZSql + "'" + proyecto.Text + "',"
            ZSql = ZSql + "'" + ZTipo + "',"
            ZSql = ZSql + "'" + Importe.Text + "',"
            ZSql = ZSql + "'" + descripcion.Text + "',"
            ZSql = ZSql + "'" + proveedor.Text + "')"
            spAvance = ZSql
            Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    
        Call CmdLimpiar_Click
        codigo.SetFocus
        
    End If
    
End Sub

Private Sub cmdDelete_Click()

    If Val(codigo.Text) <> 0 Then
        sql1 = "Select *"
        Sql2 = " FROM Avance"
        Sql3 = " Where Avance.Codigo = " + "'" + codigo.Text + "'"
        spAvance = sql1 + Sql2 + Sql3
        Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
        If rstAvance.RecordCount > 0 Then
            rstAvance.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                sql1 = "DELETE Avance"
                Sql2 = " Where Codigo = " + "'" + codigo.Text + "'"
                spAvance = sql1 + Sql2
                Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    End If
    
    codigo.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()

    codigo.Text = ""
    Fecha.Text = "  /  /    "
    proyecto.Text = ""
    Importe.Text = ""
    descripcion.Text = ""
    proveedor.Text = ""
    
    DesProyecto.Caption = ""
    DesProveedor.Caption = ""
    
    Tipo1.Value = False
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False

    sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Avance"
    spAvance = sql1 + Sql2
    Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
    If rstAvance.RecordCount > 0 Then
        rstAvance.MoveLast
        ZCodigo = IIf(IsNull(rstAvance!CodigoMayor), "0", rstAvance!CodigoMayor)
        codigo.Text = ZCodigo + 1
        rstAvance.Close
    End If
    If Val(codigo.Text) = 0 Then
        codigo.Text = "1"
    End If
    
    codigo.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgAvance.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Anterior_Click()
    sql1 = "Select *"
    Sql2 = " FROM Avance"
    Sql3 = " Where Avance.Codigo < " + "'" + codigo.Text + "'"
    Sql4 = " Order by Avance.Codigo"
    spAvance = sql1 + Sql2 + Sql3 + Sql4
    Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
    If rstAvance.RecordCount > 0 Then
        With rstAvance
            .MoveLast
            codigo.Text = rstAvance!codigo
        End With
        rstAvance.Close
        Call Imprime_Datos
        codigo.SetFocus
            Else
        m$ = "No existe registro Anterior"
        A% = MsgBox(m$, 0, "Archivo de Avances de Proyecto")
    End If
End Sub


Private Sub Lista_Click()
    TipoLista.ListIndex = 0
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    DesdeFecha.SetFocus
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFecha.Text, Auxi)
        If Auxi = "S" Then
            HastaFecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        DesdeFecha.Text = "  /  /    "
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFecha.Text, Auxi)
        If Auxi = "S" Then
            DesdeFecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFecha.Text = "  /  /    "
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            proyecto.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub proyecto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        sql1 = "Select *"
        Sql2 = " FROM Proyecto"
        Sql3 = " Where Proyecto.Codigo = " + "'" + proyecto.Text + "'"
        spProyecto = sql1 + Sql2 + Sql3
        Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
        If rstProyecto.RecordCount > 0 Then
            DesProyecto.Caption = rstProyecto!descripcion
            rstProyecto.Close
            proveedor.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        proyecto.Text = ""
        DesProyecto.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Proveedor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        sql1 = "Select *"
        Sql2 = " FROM Proveedor"
        Sql3 = " Where Proveedor.Proveedor = " + "'" + proveedor.Text + "'"
        spProveedor = sql1 + Sql2 + Sql3
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            DesProveedor.Caption = rstProveedor!nombre
            rstProveedor.Close
            Importe.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        proveedor.Text = ""
        DesProveedor.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Importe_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Importe.Text = Pusing("###,###.##", Importe.Text)
        descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Importe.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        proyecto.SetFocus
    End If
    If KeyAscii = 27 Then
        descripcion.Text = ""
    End If
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(codigo.Text) <> 0 Then
            sql1 = "Select *"
            Sql2 = " FROM Avance"
            Sql3 = " Where Avance.Codigo = " + "'" + codigo.Text + "'"
            spAvance = sql1 + Sql2 + Sql3
            Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
            If rstAvance.RecordCount > 0 Then
                rstAvance.Close
                Call Imprime_Datos
                    Else
                WCodigo = codigo.Text
                CmdLimpiar_Click
                codigo.Text = WCodigo
            End If
        End If
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     pantalla.Visible = False
     WTitulo(1).Visible = False
     WTitulo(2).Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Proyectos"
     Opcion.AddItem "Proveedores"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False
     
    Dim IngresaItem As String

    Call Limpia_Ayuda
    Lugarayuda = 0
    windice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            sql1 = "Select *"
            Sql2 = " FROM Proyecto"
            Sql3 = " Order by Proyecto.Codigo"
            spProyecto = sql1 + Sql2 + Sql3
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                With rstProyecto
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Lugarayuda = Lugarayuda + 1
                            pantalla.Row = Lugarayuda
                            pantalla.Col = 1
                            pantalla.Text = rstProyecto!codigo
                            pantalla.Col = 2
                            pantalla.Text = rstProyecto!descripcion
                            IngresaItem = rstProyecto!codigo
                            windice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProyecto.Close
            End If
            
        Case 1
            sql1 = "Select *"
            Sql2 = " FROM Proveedor"
            Sql3 = " Order by Proveedor.Nombre"
            spProveedor = sql1 + Sql2 + Sql3
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                With rstProveedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Lugarayuda = Lugarayuda + 1
                            pantalla.Row = Lugarayuda
                            pantalla.Col = 1
                            pantalla.Text = rstProveedor!proveedor
                            pantalla.Col = 2
                            pantalla.Text = rstProveedor!nombre
                            IngresaItem = rstProveedor!proveedor
                            windice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProveedor.Close
            End If
            
        Case Else
    End Select
            
    pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub pantalla_Click()

    pantalla.Visible = False
    Ayuda.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    Select Case XIndice
        Case 0
            Indice = pantalla.Row - 1
            proyecto.Text = windice.List(Indice)
            Call proyecto_KeyPress(13)
            
        Case 1
            Indice = pantalla.Row - 1
            proveedor.Text = windice.List(Indice)
            Call Proveedor_Keypress(13)
            
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()

    sql1 = "Select Min(Codigo) as [CodigoMenor]"
    Sql2 = " FROM Avance"
    spAvance = sql1 + Sql2
    Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
    If rstAvance.RecordCount > 0 Then
        rstAvance.MoveFirst
        codigo.Text = rstAvance!CodigoMenor
        rstAvance.Close
        Call Imprime_Datos
        codigo.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Avance"
    spAvance = sql1 + Sql2
    Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
    If rstAvance.RecordCount > 0 Then
        rstAvance.MoveLast
        codigo.Text = rstAvance!CodigoMayor
        rstAvance.Close
        Call Imprime_Datos
        codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    sql1 = "Select *"
    Sql2 = " FROM Avance"
    Sql3 = " Where Avance.Codigo > " + "'" + codigo.Text + "'"
    Sql4 = " Order by Avance.Codigo"
    spAvance = sql1 + Sql2 + Sql3 + Sql4
    Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
    If rstAvance.RecordCount > 0 Then
        With rstAvance
            .MoveFirst
            codigo.Text = rstAvance!codigo
        End With
        rstAvance.Close
        Call Imprime_Datos
        codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        A% = MsgBox(m$, 0, "Archivo de Avances de Produccion")
    End If

End Sub

Sub Form_Load()

    codigo.Text = ""
    Fecha.Text = "  /  /    "
    proyecto.Text = ""
    Importe.Text = ""
    descripcion.Text = ""
    proveedor.Text = ""
    
    DesProyecto.Caption = ""
    DesProveedor.Caption = ""
    
    TipoLista.Clear
    
    TipoLista.AddItem "Por Proyecto"
    TipoLista.AddItem "Por Rubro"
    
    TipoLista.ListIndex = 0
    
    Tipo1.Value = False
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False
    
    sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Avance"
    spAvance = sql1 + Sql2
    Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
    If rstAvance.RecordCount > 0 Then
        rstAvance.MoveLast
        ZCodigo = IIf(IsNull(rstAvance!CodigoMayor), "0", rstAvance!CodigoMayor)
        codigo.Text = ZCodigo + 1
        rstAvance.Close
    End If
    
    If Val(codigo.Text) = 0 Then
        codigo.Text = "1"
    End If
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    Call Limpia_Ayuda
    Lugarayuda = 0
    windice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            sql1 = "Select *"
            Sql2 = " FROM Proyecto"
            Sql3 = " Order by Proyecto.Codigo"
            spProyecto = sql1 + Sql2 + Sql3
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                With rstProyecto
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstProyecto!descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstProyecto!descripcion, aa, WEspacios) Then
                                    Lugarayuda = Lugarayuda + 1
                                    pantalla.Row = Lugarayuda
                                    pantalla.Col = 1
                                    pantalla.Text = rstProyecto!codigo
                                    pantalla.Col = 2
                                    pantalla.Text = rstProyecto!descripcion
                                    IngresaItem = rstProyecto!codigo
                                    windice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProyecto.Close
            End If
                
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Nombre LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            ZSql = ZSql + " Order by Nombre"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                With rstProveedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Lugarayuda = Lugarayuda + 1
                            pantalla.Row = Lugarayuda
                            pantalla.Col = 1
                            pantalla.Text = rstProveedor!proveedor
                            pantalla.Col = 2
                            pantalla.Text = rstProveedor!nombre
                            IngresaItem = rstProveedor!proveedor
                            windice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProveedor.Close
            End If
                
        Case Else
    End Select
    
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Proyecto_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proyecto"
    Opcion.AddItem "Proveedores"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

End Sub

Private Sub Proveedor_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proyecto"
    Opcion.AddItem "Proveedores"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Limpia_Ayuda()

    pantalla.Clear
    pantalla.Font.Bold = True
    
    ' Establesco loa Valores de la pantalla
    
    XIndice = Opcion.ListIndex
    Select Case XIndice
        Case 0
            pantalla.FixedCols = 1
            pantalla.Cols = 3
            pantalla.FixedRows = 1
            pantalla.Rows = 10001
        Case 1
            pantalla.FixedCols = 1
            pantalla.Cols = 3
            pantalla.FixedRows = 1
            pantalla.Rows = 10001
    End Select
    
    pantalla.ColWidth(0) = 200
    pantalla.Row = 0
    
    Select Case XIndice
        Case 0
            For Ciclo = 1 To pantalla.Cols - 1
                pantalla.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        pantalla.Text = "Proyecto"
                        pantalla.ColWidth(Ciclo) = 1500
                        pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
                    Case 2
                        pantalla.Text = "Nombre"
                        pantalla.ColWidth(Ciclo) = 5000
                        pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                End Select
            Next Ciclo
            
        Case 1
            For Ciclo = 1 To pantalla.Cols - 1
                pantalla.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        pantalla.Text = "Proveedor"
                        pantalla.ColWidth(Ciclo) = 1500
                        pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
                    Case 2
                        pantalla.Text = "Nombre"
                        pantalla.ColWidth(Ciclo) = 5000
                        pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                End Select
            Next Ciclo
            
        Case Else
            
    End Select
    
    Rem DESPILEGA LOS TITULOS
    
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    pantalla.Row = 0
    For Ciclo = 1 To pantalla.Cols - 1
        pantalla.Col = Ciclo
        WTitulo(Ciclo).Text = pantalla.Text
        WTitulo(Ciclo).Left = pantalla.CellLeft + pantalla.Left
        WTitulo(Ciclo).Top = pantalla.CellTop + pantalla.Top
        WTitulo(Ciclo).Width = pantalla.CellWidth
        WTitulo(Ciclo).Height = pantalla.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA pantalla
    
    WAncho = 400
    For Ciclo = 0 To pantalla.Cols - 1
        WAncho = WAncho + pantalla.ColWidth(Ciclo)
    Next Ciclo
    pantalla.Width = WAncho

    ' Size the columns.
    Font.Name = pantalla.Font.Name
    Font.Size = pantalla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    pantalla.AllowUserResizing = flexResizeBoth
    
    pantalla.Col = 1
    pantalla.Row = 1
    
End Sub





