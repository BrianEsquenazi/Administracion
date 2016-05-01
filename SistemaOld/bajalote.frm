VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgBajaLote 
   AutoRedraw      =   -1  'True
   Caption         =   "Verificacion de Lotes de Materias Primas y Productos Terminados Inactivos"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   11850
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11850
   Begin VB.Frame PantaMp 
      Height          =   2775
      Left            =   1920
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox DiasMp 
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
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   11
         Text            =   " "
         Top             =   1200
         Width           =   1095
      End
      Begin MSMask.MaskEdBox HastaMp 
         Height          =   300
         Left            =   2760
         TabIndex        =   7
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox DesdeMp 
         Height          =   300
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
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
      Begin VB.Image ConfirmaMp 
         Height          =   480
         Left            =   2040
         MouseIcon       =   "bajalote.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "bajalote.frx":030A
         ToolTipText     =   "Finaliza la Consulta de Datos"
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image CancelaMp 
         Height          =   480
         Left            =   3240
         MouseIcon       =   "bajalote.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "bajalote.frx":0A56
         ToolTipText     =   "Salida"
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Dias Inactividad"
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
         Left            =   840
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Articulo"
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
         Left            =   840
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Articulo"
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
         Left            =   840
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame PantaPt 
      Height          =   2535
      Left            =   2040
      TabIndex        =   27
      Top             =   1680
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox DiasPt 
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
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   28
         Text            =   " "
         Top             =   1200
         Width           =   1095
      End
      Begin MSMask.MaskEdBox HastaPt 
         Height          =   300
         Left            =   2760
         TabIndex        =   29
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox DesdePt 
         Height          =   300
         Left            =   2760
         TabIndex        =   30
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
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
      Begin VB.Label Label8 
         Caption         =   "Hasta Articulo"
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
         Left            =   840
         TabIndex        =   33
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Desde Articulo"
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
         Left            =   840
         TabIndex        =   32
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Dias Inactividad"
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
         Left            =   840
         TabIndex        =   31
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Image CancelaPt 
         Height          =   480
         Left            =   3240
         MouseIcon       =   "bajalote.frx":1298
         MousePointer    =   99  'Custom
         Picture         =   "bajalote.frx":15A2
         ToolTipText     =   "Salida"
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image ConfirmaPt 
         Height          =   480
         Left            =   2040
         MouseIcon       =   "bajalote.frx":1DE4
         MousePointer    =   99  'Custom
         Picture         =   "bajalote.frx":20EE
         ToolTipText     =   "Finaliza la Consulta de Datos"
         Top             =   1800
         Width           =   480
      End
   End
   Begin VB.Frame PantaAsigna 
      Height          =   3135
      Left            =   840
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   9495
      Begin VB.TextBox Responsable 
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
         Left            =   2040
         TabIndex        =   25
         Top             =   1320
         Width           =   6975
      End
      Begin VB.TextBox Observaciones 
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
         Left            =   2040
         TabIndex        =   19
         Top             =   840
         Width           =   6975
      End
      Begin VB.ComboBox TipoProceso 
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
         Left            =   2040
         TabIndex        =   18
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton FibnPantaAsigna 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4680
         MouseIcon       =   "bajalote.frx":2530
         MousePointer    =   99  'Custom
         Picture         =   "bajalote.frx":283A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Salida"
         Top             =   1920
         Width           =   855
      End
      Begin MSMask.MaskEdBox Producto 
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.Label Label5 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Destino"
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
         Height          =   285
         Left            =   480
         TabIndex        =   24
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Observaciones"
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
         Left            =   480
         TabIndex        =   23
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label22 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label DesProducto 
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
         Height          =   255
         Left            =   3960
         TabIndex        =   21
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame PantaConsulta 
      Height          =   6255
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton CancelaPantaConsulta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5160
         MouseIcon       =   "bajalote.frx":307C
         MousePointer    =   99  'Custom
         Picture         =   "bajalote.frx":3386
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Salida"
         Top             =   5160
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid ZConsulta 
         Height          =   4815
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8493
         _Version        =   327680
         Rows            =   4000
         Cols            =   9
         BackColor       =   16777152
      End
      Begin VB.Image Lista 
         Height          =   480
         Left            =   6600
         MouseIcon       =   "bajalote.frx":3BC8
         MousePointer    =   99  'Custom
         Picture         =   "bajalote.frx":3ED2
         ToolTipText     =   "Impresion "
         Top             =   5520
         Width           =   480
      End
   End
   Begin VB.CommandButton ConsultaPt 
      Caption         =   "Consulta PT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton ConsultaMp 
      Caption         =   "Consulta MP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Impre 
      Caption         =   "Impresion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10920
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Lee datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
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
      Height          =   540
      Left            =   7320
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   6375
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11245
      _Version        =   327680
      Rows            =   10000
      Cols            =   13
   End
End
Attribute VB_Name = "PrgBajaLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String

Dim rstBajaLote As Recordset
Dim spBajaLote As String
Dim rstListaInactivos As Recordset
Dim spListaInactivos As String

Dim XParam As String

Dim ZArticulo(10000) As String
Dim ZTerminado(10000) As String
Dim Vector(10000, 10) As String
Dim xLote(100, 7) As String
Dim WTipoProceso(10) As String
Dim Empe(12, 10) As String

Dim ZNumero As String
Dim ZProducto As String
Dim ZDescripcion As String
Dim ZPartida As String
Dim ZPlanta As String
Dim ZSaldo As String
Dim ZDias As String
Dim ZProceso As String
Dim ZDestino As String
Dim ZObservaciones As String
Dim ZResponsable As String
Dim ZTipoProceso As String

Dim ZStock As Double

Private Sub CancelaPantaConsulta_Click()
    PantaConsulta.Visible = False
End Sub

Private Sub cmdClose_Click()

    For Ciclo = 1 To 9999
    
        ZNumero = Muestra.TextMatrix(Ciclo, 1)
        ZProducto = Muestra.TextMatrix(Ciclo, 2)
        ZDescripcion = Muestra.TextMatrix(Ciclo, 3)
        ZPartida = Muestra.TextMatrix(Ciclo, 4)
        ZPlanta = Muestra.TextMatrix(Ciclo, 5)
        Auxi = Muestra.TextMatrix(Ciclo, 6)
        Auxi = Pusing("###,###.##", Auxi)
        ZSaldo = Auxi
        ZSaldoOriginal = Muestra.TextMatrix(Ciclo, 6)
        ZDias = Muestra.TextMatrix(Ciclo, 7)
        ZProceso = Muestra.TextMatrix(Ciclo, 8)
        ZDestino = UCase(Muestra.TextMatrix(Ciclo, 9))
        ZObservaciones = Muestra.TextMatrix(Ciclo, 10)
        ZResponsable = Muestra.TextMatrix(Ciclo, 11)
        ZTipoProceso = Muestra.TextMatrix(Ciclo, 12)
        
        If Val(ZTipoProceso) <> 0 Then
        
            If Len(Trim(ZProducto)) = 10 Then
                ZZArticulo = ZProducto
                ZZTerminado = "  -     -   "
                    Else
                ZZArticulo = "  -   -   "
                ZZTerminado = ZProducto
            End If
            
            If Val(ZSaldo) = 0 Then
                ZEstado = "1"
                    Else
                ZEstado = "0"
            End If
        
            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZOrdFecha = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
        
            If Val(ZNumero) <> 0 Then
            
                ZSql = ""
                ZSql = ZSql + "UPDATE BajaLote SET "
                ZSql = ZSql + " Saldo = " + "'" + ZSaldo + "',"
                ZSql = ZSql + " Destino = " + "'" + ZDestino + "',"
                ZSql = ZSql + " Observaciones = " + "'" + ZObservaciones + "',"
                ZSql = ZSql + " Responsable = " + "'" + ZResponsable + "',"
                ZSql = ZSql + " Estado = " + "'" + ZEstado + "',"
                ZSql = ZSql + " TipoProceso = " + "'" + ZTipoProceso + "'"
                ZSql = ZSql + " Where Numero = " + "'" + ZNumero + "'"
                spBajaLote = ZSql
                Set rstBajaLote = db.OpenRecordset(spBajaLote, dbOpenSnapshot, dbSQLPassThrough)
                
                    Else
                    
                ZSql = ""
                ZSql = ZSql + "Select Max(Numero) as [NumeroMayor]"
                ZSql = ZSql + " FROM BajaLote"
                spBajaLote = ZSql
                Set rstBajaLote = db.OpenRecordset(spBajaLote, dbOpenSnapshot, dbSQLPassThrough)
                If rstBajaLote.RecordCount > 0 Then
                    rstBajaLote.MoveLast
                    ZNumero = IIf(IsNull(rstBajaLote!NumeroMayor), "0", rstBajaLote!NumeroMayor)
                    ZNumero = Str$(Val(ZNumero) + 1)
                    rstBajaLote.Close
                        Else
                    ZNumero = "1"
                End If
                    
                ZSql = ""
                ZSql = ZSql + "INSERT INTO BajaLote ("
                ZSql = ZSql + "Numero ,"
                ZSql = ZSql + "Articulo ,"
                ZSql = ZSql + "Terminado ,"
                ZSql = ZSql + "Descripcion ,"
                ZSql = ZSql + "Partida ,"
                ZSql = ZSql + "Planta ,"
                ZSql = ZSql + "Dias ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "OrdFecha ,"
                ZSql = ZSql + "SaldoOriginal ,"
                ZSql = ZSql + "Saldo     ,"
                ZSql = ZSql + "Destino ,"
                ZSql = ZSql + "Observaciones ,"
                ZSql = ZSql + "Responsable ,"
                ZSql = ZSql + "Estado ,"
                ZSql = ZSql + "Tipoproceso) "
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + ZNumero + "',"
                ZSql = ZSql + "'" + ZZArticulo + "',"
                ZSql = ZSql + "'" + ZZTerminado + "',"
                ZSql = ZSql + "'" + ZDescripcion + "',"
                ZSql = ZSql + "'" + ZPartida + "',"
                ZSql = ZSql + "'" + ZPlanta + "',"
                ZSql = ZSql + "'" + ZDias + "',"
                ZSql = ZSql + "'" + ZFecha + "',"
                ZSql = ZSql + "'" + ZOrdFecha + "',"
                ZSql = ZSql + "'" + ZSaldoOriginal + "',"
                ZSql = ZSql + "'" + ZSaldo + "',"
                ZSql = ZSql + "'" + ZDestino + "',"
                ZSql = ZSql + "'" + ZObservaciones + "',"
                ZSql = ZSql + "'" + ZResponsable + "',"
                ZSql = ZSql + "'" + ZEstado + "',"
                ZSql = ZSql + "'" + ZTipoProceso + "')"
                
                spBajaLote = ZSql
                Set rstBajaLote = db.OpenRecordset(spBajaLote, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        End If
        
    Next Ciclo
    
    T$ = "Lotes Inactivos"
    m$ = "Desea imprimir los lotes inactivos"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
    
        Listado.WindowTitle = "Lotes Inactivos"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
    
        Listado.Destination = 1
        Listado.Destination = 0
                
        Listado.ReportFileName = "BajaLote.rpt"
                    
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
    
        Listado.SQLQuery = "SELECT BajaLote.Articulo, BajaLote.Terminado, BajaLote.Descripcion, BajaLote.Partida, BajaLote.Planta, BajaLote.Dias, BajaLote.Saldo, BajaLote.Destino, BajaLote.Observaciones, BajaLote.Responsable, BajaLote.Estado, BajaLote.TipoProceso " _
                + "From " _
                + DSQ + ".dbo.BajaLote BajaLote " _
                + "Where " _
                + "BajaLote.Estado = 0"
    
        Listado.Connect = Connect()
        
        Rem Listado.Destination = 0
        Listado.Destination = 1
        Listado.Action = 1
        
    End If

    PrgBajaLote.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub FibnPantaAsigna_Click()
    If TipoProceso.ListIndex > 0 Then
        Muestra.TextMatrix(Muestra.Row, 8) = TipoProceso.Text
        Muestra.TextMatrix(Muestra.Row, 9) = Producto.Text
        Muestra.TextMatrix(Muestra.Row, 10) = Observaciones.Text
        Muestra.TextMatrix(Muestra.Row, 11) = Responsable.Text
        Muestra.TextMatrix(Muestra.Row, 12) = TipoProceso.ListIndex
    End If
    PantaAsigna.Visible = False
End Sub

Private Sub Form_Load()

    TipoProceso.Clear
    
    TipoProceso.AddItem ""
    TipoProceso.AddItem "Fabricacion"
    TipoProceso.AddItem "Planta de Ttatamiento"
    TipoProceso.AddItem "Incineracion"
    
    TipoProceso.ListIndex = 0
    
    WTipoProceso(1) = "Fabricacion"
    WTipoProceso(2) = "Planta de Tratamiento"
    WTipoProceso(3) = "Incineracion"

    Call Limpia_Vector
    Call Limpia_VectorII
    
    WPosi1 = 1
    WPosi2 = 1
    
End Sub

Private Sub Lista_Click()

    ZSql = "DELETE ListaInactivos"
    spListaInactivos = ZSql
    Set rstListaInactivos = db.OpenRecordset(spListaInactivos, dbOpenSnapshot, dbSQLPassThrough)
    
    For Ciclo = 1 To 10000
    
        If Trim(ZConsulta.TextMatrix(Ciclo, 1)) = "" Then
        
            Exit For
            
                Else
            
            ZSql = ""
            ZSql = ZSql & "INSERT INTO ListaInactivos ("
            ZSql = ZSql & "Articulo ,"
            ZSql = ZSql & "Descripcion ,"
            ZSql = ZSql & "FechaLote ,"
            ZSql = ZSql & "Lote ,"
            ZSql = ZSql & "Saldo ,"
            ZSql = ZSql & "Dias ,"
            ZSql = ZSql & "Ultimo ,"
            ZSql = ZSql & "Planta )"
            ZSql = ZSql & "Values ("
            ZSql = ZSql & "'" + ZConsulta.TextMatrix(Ciclo, 1) + "',"
            ZSql = ZSql & "'" + ZConsulta.TextMatrix(Ciclo, 2) + "',"
            ZSql = ZSql & "'" + ZConsulta.TextMatrix(Ciclo, 3) + "',"
            ZSql = ZSql & "'" + ZConsulta.TextMatrix(Ciclo, 4) + "',"
            ZSql = ZSql & "'" + ZConsulta.TextMatrix(Ciclo, 5) + "',"
            ZSql = ZSql & "'" + ZConsulta.TextMatrix(Ciclo, 6) + "',"
            ZSql = ZSql & "'" + ZConsulta.TextMatrix(Ciclo, 7) + "',"
            ZSql = ZSql & "'" + ZConsulta.TextMatrix(Ciclo, 8) + "')"

            spListaInactivos = ZSql
            Set rstListaInactivos = db.OpenRecordset(spListaInactivos, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    
    Next Ciclo
            
            
    Listado.WindowTitle = "Lotes Inactivos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.Destination = 1
    Listado.Destination = 0
            
    Listado.ReportFileName = "ListaInactivos.rpt"
                
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)

    Listado.SQLQuery = "SELECT ListaInactivos.Articulo, ListaInactivos.Descripcion, ListaInactivos.FechaLote, ListaInactivos.Lote, ListaInactivos.Saldo, ListaInactivos.Dias, ListaInactivos.Ultimo, ListaInactivos.Planta " _
            + "From " _
            + DSQ + ".dbo.ListaInactivos ListaInactivos " _
            + "Where " _
            + "ListaInactivos.Articulo >= ' ' AND " _
            + "ListaInactivos.Articulo <= 'ZZZZZZZZZZ'"
    
    Listado.Connect = Connect()
    
    Listado.Destination = 1
    Listado.Destination = 0
    Listado.Action = 1

End Sub

Private Sub Proceso_Click()

    WSalida = "N"
    Call Limpia_Vector
    Renglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM BajaLote"
    ZSql = ZSql + " Where BajaLote.Estado = 0"
    ZSql = ZSql + " Order by Numero"
    spBajaLote = ZSql
    Set rstBajaLote = db.OpenRecordset(spBajaLote, dbOpenSnapshot, dbSQLPassThrough)
    If rstBajaLote.RecordCount > 0 Then
        With rstBajaLote
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                    Renglon = Renglon + 1
            
                    Muestra.Row = Renglon
                    
                    Muestra.Col = 1
                    Muestra.Text = Pusing("######", Str$(rstBajaLote!Numero))
                        
                    If rstBajaLote!Articulo = "  -   -   " Then
                        Muestra.Col = 2
                        Muestra.Text = rstBajaLote!Terminado
                            Else
                        Muestra.Col = 2
                        Muestra.Text = rstBajaLote!Articulo
                    End If
                
                    Muestra.Col = 3
                    Muestra.Text = rstBajaLote!Descripcion
                        
                    Muestra.Col = 4
                    Muestra.Text = rstBajaLote!Partida
                            
                    Muestra.Col = 5
                    Muestra.Text = rstBajaLote!Planta
                    
                    Muestra.Col = 6
                    Rem Muestra.Text = Str$(rstBajaLote!Saldo)
                    Muestra.Text = ""
                            
                    Muestra.Col = 7
                    Muestra.Text = rstBajaLote!Dias
                    
                    Muestra.Col = 8
                    Muestra.Text = WTipoProceso(rstBajaLote!TipoProceso)
                            
                    Muestra.Col = 9
                    Muestra.Text = rstBajaLote!Destino
                            
                    Muestra.Col = 10
                    Muestra.Text = rstBajaLote!Observaciones
                            
                    Muestra.Col = 11
                    Muestra.Text = rstBajaLote!Responsable
                            
                    Muestra.Col = 12
                    Muestra.Text = rstBajaLote!TipoProceso
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            End If
            
        End With
        rstBajaLote.Close
    End If
    
    XEmpresa = WEmpresa
    
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
    
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
            
            Else
            
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
            
    End If
    
    For Ciclo = 1 To Renglon
    
        ZEmpresa = Val(Muestra.TextMatrix(Ciclo, 5))

        WEmpresa = Empe(ZEmpresa, 1)
        txtOdbc = Empe(ZEmpresa, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        If Len(Muestra.TextMatrix(Ciclo, 2)) = 10 Then
    
            WEntra = "N"
            XParam = "'" + Muestra.TextMatrix(Ciclo, 4) + "','" _
                        + Muestra.TextMatrix(Ciclo, 2) + "'"
            spLaudo = "ListaLaudoArticulo " + XParam
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                WEntra = "S"
                Muestra.TextMatrix(Ciclo, 6) = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                rstLaudo.Close
            End If
            
            If WEntra = "N" Then
                XParam = "'" + Muestra.TextMatrix(Ciclo, 2) + "','" _
                        + Muestra.TextMatrix(Ciclo, 4) + "'"
                spMovguia = "ListaMovguiaLote " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    Muestra.TextMatrix(Ciclo, 6) = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    rstMovguia.Close
                End If
            End If
            
                Else
                
            WEntra = "N"
            
            XParam = "'" + Muestra.TextMatrix(Ciclo, 4) + "','" _
                    + Muestra.TextMatrix(Ciclo, 2) + "'"
            spHoja = "ListaHojaProducto " + XParam
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                Muestra.TextMatrix(Ciclo, 6) = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                WEntra = "S"
                rstHoja.Close
            End If
            
            If WEntra = "N" Then
                XParam = "'" + Muestra.TextMatrix(Ciclo, 2) + "','" _
                        + Muestra.TextMatrix(Ciclo, 4) + "'"
                spMovguia = "ListaMovguiaLote1 " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    Muestra.TextMatrix(Ciclo, 6) = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    WEntra = "S"
                    rstMovguia.Close
                End If
            End If
    
        End If
    
    Next Ciclo
    
    Call Conecta_Empresa
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    Muestra.TopRow = 1

End Sub

Private Sub Limpia_Vector()
    
    Muestra.Clear
    Muestra.Font.Bold = True
    
    Muestra.ColWidth(0) = 50
    Muestra.ColWidth(1) = 50
    Muestra.ColWidth(2) = 1350
    Muestra.ColWidth(3) = 1750
    Muestra.ColWidth(4) = 800
    Muestra.ColWidth(5) = 800
    Muestra.ColWidth(6) = 800
    Muestra.ColWidth(7) = 700
    Muestra.ColWidth(8) = 1200
    Muestra.ColWidth(9) = 1300
    Muestra.ColWidth(10) = 1900
    Muestra.ColWidth(11) = 650
    Muestra.ColWidth(12) = 50
    
    Muestra.ColAlignment(1) = flexAlignLeftCenter
    Muestra.ColAlignment(2) = flexAlignLeftCenter
    Muestra.ColAlignment(3) = flexAlignLeftCenter
    Muestra.ColAlignment(4) = flexAlignRightCenter
    Muestra.ColAlignment(5) = flexAlignRightCenter
    Muestra.ColAlignment(6) = flexAlignRightCenter
    Muestra.ColAlignment(7) = flexAlignRightCenter
    Muestra.ColAlignment(8) = flexAlignLeftCenter
    Muestra.ColAlignment(9) = flexAlignLeftCenter
    Muestra.ColAlignment(10) = flexAlignLeftCenter
    Muestra.ColAlignment(11) = flexAlignLeftCenter
    Muestra.ColAlignment(12) = flexAlignLeftCenter
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Numero"
    
    Muestra.Col = 2
    Muestra.Text = "Producto"
    
    Muestra.Col = 3
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 4
    Muestra.Text = "Partida"
    
    Muestra.Col = 5
    Muestra.Text = "Planta"
    
    Muestra.Col = 6
    Muestra.Text = "Saldo"
    
    Muestra.Col = 7
    Muestra.Text = "Dias"
    
    Muestra.Col = 8
    Muestra.Text = "Proceso"
    
    Muestra.Col = 9
    Muestra.Text = "Destino"
    
    Muestra.Col = 10
    Muestra.Text = "Observaciones"
    
    Muestra.Col = 11
    Muestra.Text = "Resp."
    
    Muestra.Col = 12
    Muestra.Text = ""
    
End Sub

Private Sub Limpia_VectorII()

    ZConsulta.Clear
    ZConsulta.Font.Bold = True
    
    ZConsulta.ColWidth(0) = 50
    ZConsulta.ColWidth(1) = 1400
    ZConsulta.ColWidth(2) = 2700
    ZConsulta.ColWidth(3) = 1200
    ZConsulta.ColWidth(4) = 1000
    ZConsulta.ColWidth(5) = 1000
    ZConsulta.ColWidth(6) = 1000
    ZConsulta.ColWidth(7) = 1200
    ZConsulta.ColWidth(8) = 900
    
    ZConsulta.ColAlignment(1) = flexAlignLeftCenter
    ZConsulta.ColAlignment(2) = flexAlignLeftCenter
    ZConsulta.ColAlignment(3) = flexAlignLeftCenter
    ZConsulta.ColAlignment(4) = flexAlignRightCenter
    ZConsulta.ColAlignment(5) = flexAlignRightCenter
    ZConsulta.ColAlignment(6) = flexAlignRightCenter
    ZConsulta.ColAlignment(7) = flexAlignLeftCenter
    ZConsulta.ColAlignment(8) = flexAlignLeftCenter
    
    ZConsulta.Row = 0
    
    ZConsulta.Col = 1
    ZConsulta.Text = "Articulo"
    
    ZConsulta.Col = 2
    ZConsulta.Text = "Descripcion"
    
    ZConsulta.Col = 3
    ZConsulta.Text = "F.Lote"
    
    ZConsulta.Col = 4
    ZConsulta.Text = "Lote"
    
    ZConsulta.Col = 5
    ZConsulta.Text = "Saldo"
    
    ZConsulta.Col = 6
    ZConsulta.Text = "Dias"
    
    ZConsulta.Col = 7
    ZConsulta.Text = "Ult.Mov."
    
    ZConsulta.Col = 8
    ZConsulta.Text = "Planta"
    
End Sub

Private Sub Muestra_DblClick()

    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    If Val(Muestra.TextMatrix(Muestra.Row, 12)) <> 0 Then
        TipoProceso.ListIndex = Val(Muestra.TextMatrix(Muestra.Row, 12))
        Producto.Text = Muestra.TextMatrix(Muestra.Row, 9)
        spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesProducto.Caption = rstTerminado!Descripcion
            rstTerminado.Close
                Else
            DesProducto.Caption = ""
        End If
        Observaciones.Text = Muestra.TextMatrix(Muestra.Row, 10)
        Responsable.Text = Muestra.TextMatrix(Muestra.Row, 11)
            Else
        TipoProceso.ListIndex = 0
        Producto.Text = "  -     -   "
        DesProducto.Caption = ""
        Observaciones.Text = ""
        Responsable.Text = ""
    End If

    PantaAsigna.Visible = True
    
End Sub

Private Sub Form_Activate()
    Call Proceso_Click
    Muestra.TopRow = WPosi1
    Muestra.Row = WPosi2
End Sub



Private Sub ZConsulta_Click()

    If Trim(ZConsulta.TextMatrix(ZConsulta.Row, 1)) <> "" Then
        
        For Ciclo = 1 To 10000
        
            If Trim(Muestra.TextMatrix(Ciclo, 2)) = "" Then
            
                ZLugar = Ciclo
                
                Muestra.TextMatrix(ZLugar, 2) = ZConsulta.TextMatrix(ZConsulta.Row, 1)
                Muestra.TextMatrix(ZLugar, 3) = ZConsulta.TextMatrix(ZConsulta.Row, 2)
                Muestra.TextMatrix(ZLugar, 4) = ZConsulta.TextMatrix(ZConsulta.Row, 4)
                Muestra.TextMatrix(ZLugar, 6) = ZConsulta.TextMatrix(ZConsulta.Row, 5)
                Muestra.TextMatrix(ZLugar, 7) = ZConsulta.TextMatrix(ZConsulta.Row, 6)
                Muestra.TextMatrix(ZLugar, 5) = ZConsulta.TextMatrix(ZConsulta.Row, 8)
                
                Exit For
            End If
        Next Ciclo
        
        ZConsulta.TextMatrix(ZConsulta.Row, 1) = ""
        ZConsulta.TextMatrix(ZConsulta.Row, 2) = ""
        ZConsulta.TextMatrix(ZConsulta.Row, 3) = ""
        ZConsulta.TextMatrix(ZConsulta.Row, 4) = ""
        ZConsulta.TextMatrix(ZConsulta.Row, 5) = ""
        ZConsulta.TextMatrix(ZConsulta.Row, 6) = ""
        ZConsulta.TextMatrix(ZConsulta.Row, 7) = ""
        
    End If
    
End Sub

Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "  -     -   " Then
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                DesProducto.Caption = rstTerminado!Descripcion
                rstTerminado.Close
                Observaciones.SetFocus
            End If
                Else
            Observaciones.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Producto.Text = "  -     -   "
        DesProducto.Caption = ""
    End If
End Sub

Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Sub Responsable_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TipoProceso.SetFocus
    End If
    If KeyAscii = 27 Then
        Responsable.Text = ""
    End If
End Sub

Sub TipoProceso_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Producto.SetFocus
    End If
End Sub
















Private Sub ConsultaMp_Click()

    DesdeMp.Text = "AA-000-000"
    HastaMp.Text = "ZZ-999-999"
    DiasMp.Text = "200"
    
    PantaMp.Visible = True
    
    DesdeMp.SetFocus

End Sub

Private Sub CancelaMp_Click()
    PantaMp.Visible = False
End Sub

Private Sub DesdeMp_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeMp.Text = UCase(DesdeMp.Text)
        HastaMp.SetFocus
    End If
    If KeyAscii = 27 Then
        DesdeMp.Text = "  -   -   "
    End If
End Sub

Private Sub HastaMp_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaMp.Text = UCase(HastaMp.Text)
        DiasMp.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaMp.Text = "  -   -   "
    End If
End Sub

Private Sub DiasMp_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeMp.SetFocus
    End If
    If KeyAscii = 27 Then
        DiasMp.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ConfirmaMp_Click()

    On Error GoTo WError
    
    ZZDesde = DesdeMp.Text
    ZZHasta = HastaMp.Text
    
    Call Limpia_VectorII
    ZLugarII = 0
    
    XEmpresa = WEmpresa
    
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
    
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        
        XHasta = 7
            
            Else
            
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        
        XHasta = 4
            
    End If
                
    For Ciclo2 = 1 To XHasta
    
        WEmpresa = Empe(Ciclo2, 1)
        txtOdbc = Empe(Ciclo2, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        Erase ZArticulo
        LugarArticulo = 0
        
        ZSql = ""
        ZSql = ZSql + "Select Articulo.Codigo, Articulo.Entradas, Articulo.Salidas"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Codigo >= " + "'" + ZZDesde + "'"
        ZSql = ZSql + " and Articulo.Codigo <= " + "'" + ZZHasta + "'"
        ZSql = ZSql + " Order by Articulo.Codigo"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            With rstArticulo
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        If .EOF = True Then
                            Exit Do
                        End If
                    
                        WArticulo = rstArticulo!Codigo
                        ZStock = rstArticulo!Entradas - rstArticulo!Salidas
                        Call Redondeo(ZStock)
                        If ZStock > 0 Then
                            LugarArticulo = LugarArticulo + 1
                            ZArticulo(LugarArticulo) = WArticulo
                        End If
                        
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
                End If
            End With
            rstArticulo.Close
        End If
        
        Erase Vector
        LugarVector = 0
        
        Rem dada
        
        For Ciclo = 1 To LugarArticulo
            
            WArticulo = ZArticulo(Ciclo)
            
            ZSql = ""
            ZSql = ZSql + "Select Laudo.Articulo, Laudo.Saldo, Laudo.Fecha, Laudo.Laudo, Laudo.Saldo"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArticulo + "'"
            ZSql = ZSql + " and Laudo.Saldo <> 0"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
        
                With rstLaudo
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            If .EOF = True Then
                                Exit Do
                            End If
                            
                            ZStock = rstLaudo!Saldo
                            Call Redondeo(ZStock)
                            If ZStock > 0 Then
                                LugarVector = LugarVector + 1
                                Vector(LugarVector, 1) = rstLaudo!Articulo
                                Vector(LugarVector, 2) = ""
                                Vector(LugarVector, 3) = rstLaudo!Fecha
                                Vector(LugarVector, 4) = Str$(rstLaudo!Laudo)
                                Vector(LugarVector, 5) = Str$(rstLaudo!Saldo)
                                Vector(LugarVector, 6) = ""
                                Vector(LugarVector, 7) = ""
                                Vector(LugarVector, 8) = rstLaudo!Fecha
                                Vector(LugarVector, 9) = ""
                                Vector(LugarVector, 10) = ""
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
            
            
            
        
            ZSql = ""
            ZSql = ZSql + "Select Guia.Articulo, Guia.Saldo, Guia.Fecha, Guia.Lote, Guia.Saldo"
            ZSql = ZSql + " FROM Guia"
            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArticulo + "'"
            ZSql = ZSql + " and Guia.Saldo <> 0"
            spGuia = ZSql
            Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
            If rstGuia.RecordCount > 0 Then
        
                With rstGuia
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            If .EOF = True Then
                                Exit Do
                            End If
                    
                            ZStock = rstGuia!Saldo
                            Call Redondeo(ZStock)
                            If ZStock > 0 Then
                                LugarVector = LugarVector + 1
                                Vector(LugarVector, 1) = rstGuia!Articulo
                                Vector(LugarVector, 2) = ""
                                Vector(LugarVector, 3) = rstGuia!Fecha
                                Vector(LugarVector, 4) = Str$(rstGuia!Lote)
                                Vector(LugarVector, 5) = Str$(rstGuia!Saldo)
                                Vector(LugarVector, 6) = ""
                                Vector(LugarVector, 7) = ""
                                Vector(LugarVector, 8) = rstGuia!Fecha
                                Vector(LugarVector, 9) = ""
                                Vector(LugarVector, 10) = ""
                            End If
                        
                            .MoveNext
                    
                            If .EOF = True Then
                                Exit Do
                            End If
                        Loop
                    End If
                End With
                rstGuia.Close
            End If
            
        Next Ciclo
        
        For Ciclo = 1 To LugarVector
       
            WArticulo = Vector(Ciclo, 1)
            WFecha = Vector(Ciclo, 3)
            WLote = Val(Vector(Ciclo, 4))
            WSaldo = Val(Vector(Ciclo, 5))
            
            WArticuloDy = Left$(WArticulo, 3) + "00" + Right$(WArticulo, 7)
            
            ZSql = ""
            ZSql = ZSql + "Select Estadistica.Articulo, Estadistica.OrdFecha, Estadistica.fecha, Estadistica.Numero, Estadistica.Lote1, Estadistica.Canti1, Estadistica.Lote2, Estadistica.Canti2, Estadistica.Lote3, Estadistica.Canti3, Estadistica.Lote4, Estadistica.Canti4, Estadistica.Lote5, Estadistica.Canti5, Estadistica.cantidad "
            ZSql = ZSql + " FROM Estadistica "
            ZSql = ZSql + " Where Estadistica.Articulo = " + "'" + WArticuloDy + "'"
            ZSql = ZSql + " order by Estadistica.OrdFecha desc"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
                With rstEstadistica
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            If .EOF = True Then
                                Exit Do
                            End If
                                    
                            WWFecha = rstEstadistica!Fecha
                            WWNumero = rstEstadistica!Numero
                            
                            Erase xLote
                    
                            ZLote1 = IIf(IsNull(rstEstadistica!lote1), "0", rstEstadistica!lote1)
                            ZCanti1 = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                            ZLote2 = IIf(IsNull(rstEstadistica!lote2), "0", rstEstadistica!lote2)
                            ZCanti2 = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                            ZLote3 = IIf(IsNull(rstEstadistica!lote3), "0", rstEstadistica!lote3)
                            ZCanti3 = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                            ZLote4 = IIf(IsNull(rstEstadistica!lote4), "0", rstEstadistica!lote4)
                            ZCanti4 = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                            ZLote5 = IIf(IsNull(rstEstadistica!lote5), "0", rstEstadistica!lote5)
                            ZCanti5 = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                            
                            xLote(1, 1) = Str$(ZLote1)
                            xLote(1, 2) = Str$(ZCanti1)
                            xLote(2, 1) = Str$(ZLote2)
                            xLote(2, 2) = Str$(ZCanti2)
                            xLote(3, 1) = Str$(ZLote3)
                            xLote(3, 2) = Str$(ZCanti3)
                            xLote(4, 1) = Str$(ZLote4)
                            xLote(4, 2) = Str$(ZCanti4)
                            xLote(5, 1) = Str$(ZLote5)
                            xLote(5, 2) = Str$(ZCanti5)
                        
                            If Val(xLote(1, 2)) = 0 Then
                                xLote(1, 2) = rstEstadistica!Cantidad
                            End If
                            For x = 1 To 5
                                If Val(xLote(x, 1)) = WLote Then
                                    WFecha1 = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                    WFecha2 = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
                                    If WFecha2 > WFecha1 Then
                                        WFecha = WWFecha
                                        Vector(Ciclo, 6) = "Factura"
                                        Vector(Ciclo, 7) = Str$(WWNumero)
                                        Vector(Ciclo, 8) = WWFecha
                                        Vector(Ciclo, 9) = xLote(x, 2)
                                    End If
                                    Exit Do
                                End If
                            Next x
                    
                            .MoveNext
                            If .EOF = True Then
                                Exit Do
                            End If
                    
                        Loop
                    End If
                End With
                rstEstadistica.Close
            End If
            
            
            
            
            ZSql = ""
            ZSql = ZSql + "Select Hoja.Articulo, Hoja.FechaOrd, Hoja.Fecha, Hoja.Hoja, Hoja.Lote1, Hoja.Canti1, Hoja.Lote2, Hoja.Canti2, Hoja.Lote3, Hoja.Canti3, Hoja.Lote, Hoja.Cantidad "
            ZSql = ZSql + " FROM Hoja "
            ZSql = ZSql + " Where Hoja.Articulo = " + "'" + WArticulo + "'"
            ZSql = ZSql + " Order by Hoja.FechaOrd desc"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                With rstHoja
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            If .EOF = True Then
                                Exit Do
                            End If
                                    
                            WWFecha = rstHoja!Fecha
                            WWNumero = rstHoja!Hoja
                            
                            Erase xLote
                            
                            ZLote1 = IIf(IsNull(rstHoja!lote1), "0", rstHoja!lote1)
                            ZCanti1 = IIf(IsNull(rstHoja!Canti1), "0", rstHoja!Canti1)
                            ZLote2 = IIf(IsNull(rstHoja!lote2), "0", rstHoja!lote2)
                            ZCanti2 = IIf(IsNull(rstHoja!Canti2), "0", rstHoja!Canti2)
                            ZLote3 = IIf(IsNull(rstHoja!lote3), "0", rstHoja!lote3)
                            ZCanti3 = IIf(IsNull(rstHoja!Canti3), "0", rstHoja!Canti3)
                            
                            xLote(1, 1) = Str$(ZLote1)
                            xLote(1, 2) = Str$(ZCanti1)
                            xLote(2, 1) = Str$(ZLote2)
                            xLote(2, 2) = Str$(ZCanti2)
                            xLote(3, 1) = Str$(ZLote3)
                            xLote(3, 2) = Str$(ZCanti3)
                            If Val(xLote(1, 1)) = 0 Then
                                xLote(1, 1) = Str$(rstHoja!Lote)
                                xLote(1, 2) = Str$(rstHoja!Cantidad)
                            End If
                            
                            For x = 1 To 3
                                If Val(xLote(x, 1)) = WLote Then
                                    WFecha1 = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                    WFecha2 = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
                                    If WFecha2 > WFecha1 Then
                                        WFecha = WWFecha
                                        Vector(Ciclo, 6) = "Hoja"
                                        Vector(Ciclo, 7) = Str$(WWNumero)
                                        Vector(Ciclo, 8) = WWFecha
                                        Vector(Ciclo, 9) = xLote(x, 2)
                                    End If
                                    Exit Do
                                End If
                            Next x
                    
                            .MoveNext
                            If .EOF = True Then
                                Exit Do
                            End If
                    
                        Loop
                    End If
                End With
                rstHoja.Close
            End If
            
            
            
            ZSql = ""
            ZSql = ZSql + "Select * "
            ZSql = ZSql + " FROM Movvar "
            ZSql = ZSql + " Where Movvar.Articulo = " + "'" + WArticulo + "'"
            ZSql = ZSql + " and Movvar.Lote = " + "'" + Str$(WLote) + " '"
            ZSql = ZSql + " and Movvar.Movi = " + "'" + "S" + " '"
            spMovvar = ZSql
            Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovvar.RecordCount > 0 Then
                With rstMovvar
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            If .EOF = True Then
                                Exit Do
                            End If
                            
                            WWCantidad = rstMovvar!Cantidad
                            WWFecha = rstMovvar!Fecha
                            WWNumero = rstMovvar!Codigo
                            WWLote = rstMovvar!Lote
                            
                            WFecha1 = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            WFecha2 = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
                            If WFecha2 > WFecha1 Then
                                WFecha = WWFecha
                                Vector(Ciclo, 6) = "Mov.Var."
                                Vector(Ciclo, 7) = Str$(WWNumero)
                                Vector(Ciclo, 8) = WWFecha
                                Vector(Ciclo, 9) = Str$(WWCantidad)
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
        
        
        
            Rem ZSql = ""
            Rem ZSql = ZSql + "Select * "
            Rem ZSql = ZSql + " FROM Guia "
            Rem ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArticulo + "'"
            Rem ZSql = ZSql + " and Guia.Lote = " + "'" + Str$(WLote) + " '"
            Rem ZSql = ZSql + " and Guia.Movi = " + "'" + "S" + " '"
            Rem spGuia = ZSql
            Rem Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstGuia.RecordCount > 0 Then
            Rem     With rstGuia
            Rem         .MoveFirst
            Rem         If .NoMatch = False Then
            Rem             Do
            Rem                 If .EOF = True Then
            Rem                     Exit Do
            Rem                 End If
            Rem
            Rem                 WWCantidad = rstGuia!Cantidad
            Rem                 WWFecha = rstGuia!Fecha
            Rem                 WWNumero = rstGuia!Codigo
            Rem                 WWLote = rstGuia!Lote
            Rem
            Rem                 WFecha1 = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
            Rem                 WFecha2 = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
            Rem                 If WFecha2 > WFecha1 Then
            Rem                     WFecha = WWFecha
            Rem                     Vector(Ciclo, 6) = "Guia"
            Rem                     Vector(Ciclo, 7) = Str$(WWNumero)
            Rem                     Vector(Ciclo, 8) = WWFecha
            Rem                     Vector(Ciclo, 9) = Str$(WWCantidad)
            Rem                 End If
            Rem
            Rem                 .MoveNext
            Rem
            Rem                 If .EOF = True Then
            Rem                     Exit Do
            Rem                 End If
            Rem
            Rem             Loop
            Rem         End If
            Rem     End With
            Rem     rstGuia.Close
            Rem End If
            
            ZComparaI = "01/01/1900"
            ZComparaII = "01/01/1900"
            
            ZComparaI = Vector(Ciclo, 8)
            ZComparaII = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            
            WDias = DateDiff("d", ZComparaI, ZComparaII)
            
            If WDias > Val(DiasMp.Text) Then
            
                ZEntra = "S"
                
                For ZCiclo = 1 To 9999
                    If Trim(Muestra.TextMatrix(ZCiclo, 2)) = "" Then
                        Exit For
                    End If
                    If Muestra.TextMatrix(ZCiclo, 2) = Vector(Ciclo, 1) Then
                        If Val(Muestra.TextMatrix(ZCiclo, 4)) = Val(Vector(Ciclo, 4)) Then
                            ZEntra = "N"
                            Exit For
                        End If
                    End If
                Next ZCiclo
                            
                If ZEntra = "S" Then
            
                    ZZDescripcion = ""
                    ZSql = ""
                    ZSql = ZSql + "Select Articulo.Codigo, Articulo.Descripcion"
                    ZSql = ZSql + " FROM Articulo"
                    ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Vector(Ciclo, 1) + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        ZZDescripcion = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                
                    ZLugarII = ZLugarII + 1
                
                    ZConsulta.TextMatrix(ZLugarII, 1) = Vector(Ciclo, 1)
                    ZConsulta.TextMatrix(ZLugarII, 2) = ZZDescripcion
                    ZConsulta.TextMatrix(ZLugarII, 3) = Vector(Ciclo, 3)
                    ZConsulta.TextMatrix(ZLugarII, 4) = Vector(Ciclo, 4)
                    ZConsulta.TextMatrix(ZLugarII, 5) = Vector(Ciclo, 5)
                    ZConsulta.TextMatrix(ZLugarII, 6) = Str$(WDias)
                    ZConsulta.TextMatrix(ZLugarII, 7) = Vector(Ciclo, 8)
                    ZConsulta.TextMatrix(ZLugarII, 8) = Str$(Ciclo2)
                    
                End If
        
            End If
        
        Next Ciclo
        
    Next Ciclo2
    
    Call Conecta_Empresa
    
    PantaMp.Visible = False
    PantaConsulta.Visible = True
    
    Exit Sub
    
WError:
     Resume Next
    
End Sub


















Private Sub ConsultaPT_Click()

    DesdePt.Text = "PT-00000-000"
    HastaPt.Text = "PT-99999-999"
    DiasPt.Text = "200"
    
    PantaPt.Visible = True
    
    DesdePt.SetFocus

End Sub

Private Sub CancelaPt_Click()
    PantaPt.Visible = False
End Sub

Private Sub DesdePt_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdePt.Text = UCase(DesdePt.Text)
        HastaPt.SetFocus
    End If
    If KeyAscii = 27 Then
        DesdePt.Text = "  -     -   "
    End If
End Sub

Private Sub HastaPt_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaPt.Text = UCase(HastaPt.Text)
        DiasPt.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaPt.Text = "  -     -   "
    End If
End Sub

Private Sub DiasPt_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdePt.SetFocus
    End If
    If KeyAscii = 27 Then
        DiasPt.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub ConfirmaPT_Click()


    On Error GoTo WError
    
    ZZDesde = DesdePt.Text
    ZZHasta = HastaPt.Text
    
    Call Limpia_VectorII
    ZLugarII = 0
    
    XEmpresa = WEmpresa
    
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
    
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        
        XHasta = 7
            
            Else
            
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        
        XHasta = 4
            
    End If
                
    For Ciclo2 = 1 To XHasta
    
        WEmpresa = Empe(Ciclo2, 1)
        txtOdbc = Empe(Ciclo2, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        Erase ZTerminado
        LugarTerminado = 0
    
        ZSql = ""
        ZSql = ZSql + "Select Terminado.Codigo, Terminado.Entradas, Terminado.Salidas"
        ZSql = ZSql + " FROM Terminado"
        ZSql = ZSql + " Where Terminado.Codigo >= " + "'" + ZZDesde + "'"
        ZSql = ZSql + " and Terminado.Codigo <= " + "'" + ZZHasta + "'"
        ZSql = ZSql + " Order by Terminado.Codigo"
        spTerminado = ZSql
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            With rstTerminado
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        If .EOF = True Then
                            Exit Do
                        End If
                    
                        WTerminado = rstTerminado!Codigo
                        ZStock = rstTerminado!Entradas - rstTerminado!Salidas
                        If ZStock > 0 Then
                            LugarTerminado = LugarTerminado + 1
                            ZTerminado(LugarTerminado) = WTerminado
                        End If
                        
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
                End If
            End With
            rstTerminado.Close
        End If
        
        Erase Vector
        LugarVector = 0
    
        For Ciclo = 1 To LugarTerminado
            
            WTerminado = ZTerminado(Ciclo)
            
            ZSql = ""
            ZSql = ZSql + "Select Hoja.Producto, Hoja.Saldo, Hoja.Renglon, Hoja.Real, Hoja.Fecha, Hoja.Hoja"
            ZSql = ZSql + " FROM Hoja"
            ZSql = ZSql + " Where Hoja.Producto = " + "'" + WTerminado + "'"
            ZSql = ZSql + " and Hoja.Saldo <> 0"
            ZSql = ZSql + " and Hoja.Renglon = 1"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
        
                With rstHoja
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            If .EOF = True Then
                                Exit Do
                            End If
                    
                            LugarVector = LugarVector + 1
                            Vector(LugarVector, 1) = rstHoja!Producto
                            Vector(LugarVector, 2) = Str$(rstHoja!Real)
                            Vector(LugarVector, 3) = rstHoja!Fecha
                            Vector(LugarVector, 4) = Str$(rstHoja!Hoja)
                            Vector(LugarVector, 5) = Str$(rstHoja!Saldo)
                            Vector(LugarVector, 6) = ""
                            Vector(LugarVector, 7) = ""
                            Vector(LugarVector, 8) = rstHoja!Fecha
                            Vector(LugarVector, 9) = ""
                            Vector(LugarVector, 10) = ""
                    
                            .MoveNext
                            If .EOF = True Then
                                Exit Do
                            End If
                        Loop
                    End If
                End With
                rstHoja.Close
            End If
            
            
            
        
            ZSql = ""
            ZSql = ZSql + "Select Guia.Terminado, Guia.Saldo, Guia.Fecha, Guia.Lote"
            ZSql = ZSql + " FROM Guia"
            ZSql = ZSql + " Where Guia.Terminado = " + "'" + WTerminado + "'"
            ZSql = ZSql + " and Guia.Saldo <> 0"
            spMovguia = ZSql
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
            
                 With rstMovguia
                     .MoveFirst
                     If .NoMatch = False Then
                         Do
                             If .EOF = True Then
                                 Exit Do
                             End If
            
                             LugarVector = LugarVector + 1
                             Vector(LugarVector, 1) = rstMovguia!Terminado
                             Vector(LugarVector, 2) = ""
                             Vector(LugarVector, 3) = rstMovguia!Fecha
                             Vector(LugarVector, 4) = Str$(rstMovguia!Lote)
                             Vector(LugarVector, 5) = Str$(rstMovguia!Saldo)
                             Vector(LugarVector, 6) = ""
                             Vector(LugarVector, 7) = ""
                             Vector(LugarVector, 8) = rstMovguia!Fecha
                             Vector(LugarVector, 9) = ""
                             Vector(LugarVector, 10) = ""
            
                             .MoveNext
            
                             If .EOF = True Then
                                 Exit Do
                             End If
                         Loop
                     End If
                 End With
                 rstMovguia.Close
            End If
        
        
            ZSql = ""
            ZSql = ZSql + "Select EntDev.Terminado, EntDev.Saldo, EntDev.Fecha, EntDev.Lote"
            ZSql = ZSql + " FROM Entdev"
            ZSql = ZSql + " Where Entdev.Terminado = " + "'" + WTerminado + "'"
            ZSql = ZSql + " and Entdev.Saldo <> 0"
            spEntdev = ZSql
            Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
            If rstEntdev.RecordCount > 0 Then
                 With rstEntdev
                     .MoveFirst
                     If .NoMatch = False Then
                         Do
                             If .EOF = True Then
                                 Exit Do
                             End If
            
                             LugarVector = LugarVector + 1
                             Vector(LugarVector, 1) = rstEntdev!Terminado
                             Vector(LugarVector, 2) = ""
                             Vector(LugarVector, 3) = rstEntdev!Fecha
                             Vector(LugarVector, 4) = Str$(rstEntdev!Lote)
                             Vector(LugarVector, 5) = Str$(rstEntdev!Saldo)
                             Vector(LugarVector, 6) = ""
                             Vector(LugarVector, 7) = ""
                             Vector(LugarVector, 8) = rstEntdev!Fecha
                             Vector(LugarVector, 9) = ""
                             Vector(LugarVector, 10) = ""
            
                             .MoveNext
                             If .EOF = True Then
                                 Exit Do
                             End If
                         Loop
                     End If
                 End With
                 rstEntdev.Close
            End If
            
        Next Ciclo
    
        For Ciclo = 1 To LugarVector
       
            WTerminado = Vector(Ciclo, 1)
            WFecha = Vector(Ciclo, 3)
            WLote = Val(Vector(Ciclo, 4))
            WSaldo = Val(Vector(Ciclo, 5))
            
            ZSql = ""
            ZSql = ZSql + "Select Estadistica.Articulo, Estadistica.Fecha, Estadistica.Numero, Estadistica.Lote1, Estadistica.Canti1, Estadistica.Lote2, Estadistica.Canti2, Estadistica.Lote3, Estadistica.Canti3, Estadistica.Lote4, Estadistica.Canti4, Estadistica.Lote5, Estadistica.Canti5, Estadistica.Cantidad"
            ZSql = ZSql + " FROM Estadistica "
            ZSql = ZSql + " Where Estadistica.Articulo = " + "'" + WTerminado + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
                With rstEstadistica
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            If .EOF = True Then
                                Exit Do
                            End If
                                    
                            WWFecha = rstEstadistica!Fecha
                            WWNumero = rstEstadistica!Numero
                            
                            Erase xLote
                    
                            ZLote1 = IIf(IsNull(rstEstadistica!lote1), "0", rstEstadistica!lote1)
                            ZCanti1 = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                            ZLote2 = IIf(IsNull(rstEstadistica!lote2), "0", rstEstadistica!lote2)
                            ZCanti2 = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                            ZLote3 = IIf(IsNull(rstEstadistica!lote3), "0", rstEstadistica!lote3)
                            ZCanti3 = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                            ZLote4 = IIf(IsNull(rstEstadistica!lote4), "0", rstEstadistica!lote4)
                            ZCanti4 = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                            ZLote5 = IIf(IsNull(rstEstadistica!lote5), "0", rstEstadistica!lote5)
                            ZCanti5 = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                            
                            xLote(1, 1) = Str$(ZLote1)
                            xLote(1, 2) = Str$(ZCanti1)
                            xLote(2, 1) = Str$(ZLote2)
                            xLote(2, 2) = Str$(ZCanti2)
                            xLote(3, 1) = Str$(ZLote3)
                            xLote(3, 2) = Str$(ZCanti3)
                            xLote(4, 1) = Str$(ZLote4)
                            xLote(4, 2) = Str$(ZCanti4)
                            xLote(5, 1) = Str$(ZLote5)
                            xLote(5, 2) = Str$(ZCanti5)
                        
                            If xLote(1, 2) = 0 Then
                                xLote(1, 2) = Str$(rstEstadistica!Cantidad)
                            End If
                            For x = 1 To 5
                                If Val(xLote(x, 1)) = WLote Then
                                    WFecha1 = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                    WFecha2 = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
                                    If WFecha2 > WFecha1 Then
                                        WFecha = WWFecha
                                        Vector(Ciclo, 6) = "Factura"
                                        Vector(Ciclo, 7) = Str$(WWNumero)
                                        Vector(Ciclo, 8) = WWFecha
                                        Vector(Ciclo, 9) = xLote(x, 2)
                                    End If
                                End If
                            Next x
                    
                            .MoveNext
                            If .EOF = True Then
                                Exit Do
                            End If
                    
                        Loop
                    End If
                End With
                rstEstadistica.Close
            End If
            
            
            
            ZSql = ""
            ZSql = ZSql + "Select Hoja.Terminado, Hoja.Fecha, Hoja.Hoja, Hoja.Lote1, Hoja.Canti1, Hoja.Lote2, Hoja.Canti2, Hoja.Lote3, Hoja.Canti3, Hoja.Lote, Hoja.Cantidad"
            ZSql = ZSql + " FROM Hoja "
            ZSql = ZSql + " Where Hoja.Terminado = " + "'" + WTerminado + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                With rstHoja
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            If .EOF = True Then
                                Exit Do
                            End If
                                    
                            WWFecha = rstHoja!Fecha
                            WWNumero = rstHoja!Hoja
                            
                            Erase xLote
                    
                            ZLote1 = IIf(IsNull(rstHoja!lote1), "0", rstHoja!lote1)
                            ZCanti1 = IIf(IsNull(rstHoja!Canti1), "0", rstHoja!Canti1)
                            ZLote2 = IIf(IsNull(rstHoja!lote2), "0", rstHoja!lote2)
                            ZCanti2 = IIf(IsNull(rstHoja!Canti2), "0", rstHoja!Canti2)
                            ZLote3 = IIf(IsNull(rstHoja!lote3), "0", rstHoja!lote3)
                            ZCanti3 = IIf(IsNull(rstHoja!Canti3), "0", rstHoja!Canti3)
                            
                            xLote(1, 1) = Str$(ZLote1)
                            xLote(1, 2) = Str$(ZCanti1)
                            xLote(2, 1) = Str$(ZLote2)
                            xLote(2, 2) = Str$(ZCanti2)
                            xLote(3, 1) = Str$(ZLote3)
                            xLote(3, 2) = Str$(ZCanti3)
                            
                            If Val(xLote(1, 1)) = 0 Then
                                xLote(1, 1) = Str$(rstHoja!Lote)
                                xLote(1, 2) = Str$(rstHoja!Cantidad)
                            End If
                            
                            For x = 1 To 3
                                If Val(xLote(x, 1)) = WLote Then
                                    WFecha1 = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                    WFecha2 = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
                                    If WFecha2 > WFecha1 Then
                                        WFecha = WWFecha
                                        Vector(Ciclo, 6) = "Hoja"
                                        Vector(Ciclo, 7) = Str$(WWNumero)
                                        Vector(Ciclo, 8) = WWFecha
                                        Vector(Ciclo, 9) = xLote(x, 2)
                                    End If
                                End If
                            Next x
                    
                            .MoveNext
                            If .EOF = True Then
                                Exit Do
                            End If
                    
                        Loop
                    End If
                End With
                rstHoja.Close
            End If
            
            
            
            ZSql = ""
            ZSql = ZSql + "Select Movvar.Terminado, Movvar.Lote, Movvar.Movi, Movvar.Cantidad, Movvar.Fecha, Movvar.Codigo"
            ZSql = ZSql + " FROM Movvar "
            ZSql = ZSql + " Where Movvar.Terminado = " + "'" + WTerminado + "'"
            ZSql = ZSql + " and Movvar.Lote = " + "'" + Str$(WLote) + " '"
            ZSql = ZSql + " and Movvar.Movi = " + "'" + "S" + " '"
            spMovvar = ZSql
            Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovvar.RecordCount > 0 Then
                With rstMovvar
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            If .EOF = True Then
                                Exit Do
                            End If
                            
                            WWCantidad = rstMovvar!Cantidad
                            WWFecha = rstMovvar!Fecha
                            WWNumero = rstMovvar!Codigo
                            WWLote = rstMovvar!Lote
                            
                            WFecha1 = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            WFecha2 = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
                            If WFecha2 > WFecha1 Then
                                WFecha = WWFecha
                                Vector(Ciclo, 6) = "Mov.Var."
                                Vector(Ciclo, 7) = Str$(WWNumero)
                                Vector(Ciclo, 8) = WWFecha
                                Vector(Ciclo, 9) = Str$(WWCantidad)
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
            
            
        
            Rem XParam = "'" + WTerminado + "','" _
            rem          + WTerminado + "'"
            Rem spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
            Rem Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstMovguia.RecordCount > 0 Then
            Rem     With rstMovguia
            Rem         .MoveFirst
            Rem         If .NoMatch = False Then
            Rem             Do
            Rem                 If .EOF = True Then
            Rem                     Exit Do
            Rem                 End If
            Rem
            Rem                 If rstMovguia!Partida = WLote And rstMovguia!Movi = "S" Then
            Rem
            Rem                     WWCantidad = rstMovguia!Cantidad
            Rem                     WWFecha = rstMovguia!Fecha
            Rem                     WWNumero = rstMovguia!Codigo
            Rem
            Rem                     WFecha1 = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
            Rem                     WFecha2 = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
            Rem                     If WFecha2 > WFecha1 Then
            Rem                         WFecha = WWFecha
            Rem                         Vector(Ciclo, 6) = "Guia"
            Rem                         Vector(Ciclo, 7) = Str$(WWNumero)
            Rem                         Vector(Ciclo, 8) = WWFecha
            Rem                         Vector(Ciclo, 9) = Str$(WWCantidad)
            Rem                     End If
            Rem
            Rem                 End If
            Rem
            Rem
            Rem                 .MoveNext
            Rem                 If .EOF = True Then
            Rem                     Exit Do
            Rem                 End If
            Rem             Loop
            Rem         End If
            Rem     End With
            Rem     rstMovguia.Close
            Rem End If
            
            ZComparaI = "01/01/1900"
            ZComparaII = "01/01/1900"
            
            ZComparaI = Vector(Ciclo, 8)
            ZComparaII = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            
            WDias = DateDiff("d", ZComparaI, ZComparaII)
            
            If WDias > Val(DiasPt.Text) Then
                
                ZEntra = "S"
                
                For ZCiclo = 1 To 9999
                    If Trim(Muestra.TextMatrix(ZCiclo, 2)) = "" Then
                        Exit For
                    End If
                    If Muestra.TextMatrix(ZCiclo, 2) = Vector(Ciclo, 1) Then
                        If Val(Muestra.TextMatrix(ZCiclo, 4)) = Val(Vector(Ciclo, 4)) Then
                            ZEntra = "N"
                            Exit For
                        End If
                    End If
                Next ZCiclo
                            
                If ZEntra = "S" Then
            
                    ZZDescripcion = ""
                    ZSql = ""
                    ZSql = ZSql + "Select Terminado.Codigo, Terminado.Descripcion"
                    ZSql = ZSql + " FROM Terminado"
                    ZSql = ZSql + " Where Terminado.Codigo = " + "'" + Vector(Ciclo, 1) + "'"
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        ZZDescripcion = rstTerminado!Descripcion
                        rstTerminado.Close
                    End If
                
                    ZLugarII = ZLugarII + 1
                
                    ZConsulta.TextMatrix(ZLugarII, 1) = Vector(Ciclo, 1)
                    ZConsulta.TextMatrix(ZLugarII, 2) = ZZDescripcion
                    ZConsulta.TextMatrix(ZLugarII, 3) = Vector(Ciclo, 3)
                    ZConsulta.TextMatrix(ZLugarII, 4) = Vector(Ciclo, 4)
                    ZConsulta.TextMatrix(ZLugarII, 5) = Vector(Ciclo, 5)
                    ZConsulta.TextMatrix(ZLugarII, 6) = Str$(WDias)
                    ZConsulta.TextMatrix(ZLugarII, 7) = Vector(Ciclo, 8)
                    ZConsulta.TextMatrix(ZLugarII, 8) = Str$(Ciclo2)
                    
                End If
            
            End If
            
        Next Ciclo
            
    Next Ciclo2
    
    Call Conecta_Empresa
    
    PantaPt.Visible = False
    PantaConsulta.Visible = True
    
    Exit Sub
    
WError:
     Resume Next

End Sub



