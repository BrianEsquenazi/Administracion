VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgModpedNuevoII 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Discriminacion de Partidas"
   ClientHeight    =   6495
   ClientLeft      =   1125
   ClientTop       =   780
   ClientWidth     =   9510
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   6495
   ScaleWidth      =   9510
   Visible         =   0   'False
   Begin VB.Frame CargaLote 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   9255
      Begin VB.TextBox Diferencia 
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
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   22
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Cantidad 
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   0
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Producto 
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   16
         Top             =   240
         Width           =   1455
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
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2400
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
         TabIndex        =   14
         Top             =   2280
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2160
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2040
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2160
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
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
         Left            =   2640
         TabIndex        =   9
         Top             =   2880
         Width           =   375
      End
      Begin VB.ComboBox WCombo1 
         Height          =   315
         Left            =   2640
         TabIndex        =   8
         Top             =   3480
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
         Left            =   3240
         TabIndex        =   7
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox Asignada 
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
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid WVector1 
         Height          =   3375
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5953
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   4080
         TabIndex        =   18
         Top             =   3360
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diferencia"
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
         Left            =   5760
         TabIndex        =   23
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
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
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Asignada"
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
         Left            =   3120
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid WVector2 
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      _Version        =   327680
      BackColor       =   12648384
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Cancela 
      Caption         =   "Cancela"
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
      Left            =   4800
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Confirma 
      Caption         =   "Confirma"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   4800
      Width           =   1335
   End
End
Attribute VB_Name = "PrgModpedNuevoII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstLaudo As Recordset
Dim spLaudo As String

Dim XParam As String

Dim WSaldo As Double

Dim WEstado As String
Dim XTerminado As String
Dim XCantidad  As Double
Dim WRow As Integer

Dim XCantidad1 As String
Dim xCantidad2 As String

Dim XMes As String
Dim XAno As String

Dim ControlLote(12, 2) As String

Dim WCanti As Double
Dim WLote As String
Dim WLugar As Integer

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String
Dim WControlII As String

Dim ZSaldo As Double

Private Sub Cancela_click()
    PrgModpedNuevoII.Hide
    Unload Me
    PrgModpedNuevo.Show
End Sub

Private Sub Confirma_Click()
    Call Verifica_Lote
    If WEstado = "S" Then

        ZZTrabajaLote(1) = WVector1.TextMatrix(1, 1)
        ZZTrabajaLote(3) = WVector1.TextMatrix(2, 1)
        ZZTrabajaLote(5) = WVector1.TextMatrix(3, 1)
        ZZTrabajaLote(7) = WVector1.TextMatrix(4, 1)
        ZZTrabajaLote(9) = WVector1.TextMatrix(5, 1)
        ZZTrabajaLote(11) = WVector1.TextMatrix(6, 1)
        ZZTrabajaLote(13) = WVector1.TextMatrix(7, 1)
        ZZTrabajaLote(15) = WVector1.TextMatrix(8, 1)
        ZZTrabajaLote(17) = WVector1.TextMatrix(9, 1)
        ZZTrabajaLote(19) = WVector1.TextMatrix(10, 1)
        ZZTrabajaLote(21) = WVector1.TextMatrix(11, 1)
        ZZTrabajaLote(23) = WVector1.TextMatrix(12, 1)
        
        ZZTrabajaLote(2) = WVector1.TextMatrix(1, 2)
        ZZTrabajaLote(4) = WVector1.TextMatrix(2, 2)
        ZZTrabajaLote(6) = WVector1.TextMatrix(3, 2)
        ZZTrabajaLote(8) = WVector1.TextMatrix(4, 2)
        ZZTrabajaLote(10) = WVector1.TextMatrix(5, 2)
        ZZTrabajaLote(12) = WVector1.TextMatrix(6, 2)
        ZZTrabajaLote(14) = WVector1.TextMatrix(7, 2)
        ZZTrabajaLote(16) = WVector1.TextMatrix(8, 2)
        ZZTrabajaLote(18) = WVector1.TextMatrix(9, 2)
        ZZTrabajaLote(20) = WVector1.TextMatrix(10, 2)
        ZZTrabajaLote(22) = WVector1.TextMatrix(11, 2)
        ZZTrabajaLote(24) = WVector1.TextMatrix(12, 2)
        
        ZZTrabajaLote(31) = WVector1.TextMatrix(1, 3)
        ZZTrabajaLote(32) = WVector1.TextMatrix(1, 5)
        ZZTrabajaLote(33) = WVector1.TextMatrix(2, 3)
        ZZTrabajaLote(34) = WVector1.TextMatrix(2, 5)
        ZZTrabajaLote(35) = WVector1.TextMatrix(3, 3)
        ZZTrabajaLote(36) = WVector1.TextMatrix(3, 5)
        ZZTrabajaLote(37) = WVector1.TextMatrix(4, 3)
        ZZTrabajaLote(38) = WVector1.TextMatrix(4, 5)
        ZZTrabajaLote(39) = WVector1.TextMatrix(5, 3)
        ZZTrabajaLote(40) = WVector1.TextMatrix(5, 5)
        ZZTrabajaLote(41) = WVector1.TextMatrix(6, 3)
        ZZTrabajaLote(42) = WVector1.TextMatrix(6, 5)
        ZZTrabajaLote(43) = WVector1.TextMatrix(7, 3)
        ZZTrabajaLote(44) = WVector1.TextMatrix(7, 5)
        ZZTrabajaLote(45) = WVector1.TextMatrix(8, 3)
        ZZTrabajaLote(46) = WVector1.TextMatrix(8, 5)
        ZZTrabajaLote(47) = WVector1.TextMatrix(9, 3)
        ZZTrabajaLote(48) = WVector1.TextMatrix(9, 5)
        ZZTrabajaLote(49) = WVector1.TextMatrix(10, 3)
        ZZTrabajaLote(50) = WVector1.TextMatrix(10, 5)
        ZZTrabajaLote(51) = WVector1.TextMatrix(11, 3)
        ZZTrabajaLote(52) = WVector1.TextMatrix(11, 5)
        ZZTrabajaLote(53) = WVector1.TextMatrix(12, 3)
        ZZTrabajaLote(54) = WVector1.TextMatrix(12, 5)
        
        ZZTrabajaLote(61) = WVector1.TextMatrix(1, 6)
        ZZTrabajaLote(62) = WVector1.TextMatrix(2, 6)
        ZZTrabajaLote(63) = WVector1.TextMatrix(3, 6)
        ZZTrabajaLote(64) = WVector1.TextMatrix(4, 6)
        ZZTrabajaLote(65) = WVector1.TextMatrix(5, 6)
        ZZTrabajaLote(66) = WVector1.TextMatrix(6, 6)
        ZZTrabajaLote(67) = WVector1.TextMatrix(7, 6)
        ZZTrabajaLote(68) = WVector1.TextMatrix(8, 6)
        ZZTrabajaLote(69) = WVector1.TextMatrix(9, 6)
        ZZTrabajaLote(70) = WVector1.TextMatrix(10, 6)
        ZZTrabajaLote(71) = WVector1.TextMatrix(11, 6)
        ZZTrabajaLote(72) = WVector1.TextMatrix(12, 6)
    
        PrgModpedNuevoII.Hide
        Unload Me
        PrgModpedNuevo.Show
    End If
End Sub

Private Sub Form_Load()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    Call Limpia_Vector
    
    Producto.Text = WPasaTerminado
    Cantidad.Text = Pusing("###,###.##", Str$(WPasaCantidad))
    WControlII = ""
    
    WVector1.TextMatrix(1, 1) = ZZTrabajaLote(1)
    WVector1.TextMatrix(2, 1) = ZZTrabajaLote(3)
    WVector1.TextMatrix(3, 1) = ZZTrabajaLote(5)
    WVector1.TextMatrix(4, 1) = ZZTrabajaLote(7)
    WVector1.TextMatrix(5, 1) = ZZTrabajaLote(9)
    WVector1.TextMatrix(6, 1) = ZZTrabajaLote(11)
    WVector1.TextMatrix(7, 1) = ZZTrabajaLote(13)
    WVector1.TextMatrix(8, 1) = ZZTrabajaLote(15)
    WVector1.TextMatrix(9, 1) = ZZTrabajaLote(17)
    WVector1.TextMatrix(10, 1) = ZZTrabajaLote(19)
    WVector1.TextMatrix(11, 1) = ZZTrabajaLote(21)
    WVector1.TextMatrix(12, 1) = ZZTrabajaLote(23)
    
    If Val(ZZTrabajaLote(2)) <> 0 Then
        WVector1.TextMatrix(1, 2) = ZZTrabajaLote(2)
    End If
    If Val(ZZTrabajaLote(4)) <> 0 Then
        WVector1.TextMatrix(2, 2) = ZZTrabajaLote(4)
    End If
    If Val(ZZTrabajaLote(6)) <> 0 Then
        WVector1.TextMatrix(3, 2) = ZZTrabajaLote(6)
    End If
    If Val(ZZTrabajaLote(8)) <> 0 Then
        WVector1.TextMatrix(4, 2) = ZZTrabajaLote(8)
    End If
    If Val(ZZTrabajaLote(10)) <> 0 Then
        WVector1.TextMatrix(5, 2) = ZZTrabajaLote(10)
    End If
    If Val(ZZTrabajaLote(12)) <> 0 Then
        WVector1.TextMatrix(6, 2) = ZZTrabajaLote(12)
    End If
    If Val(ZZTrabajaLote(14)) <> 0 Then
        WVector1.TextMatrix(7, 2) = ZZTrabajaLote(14)
    End If
    If Val(ZZTrabajaLote(16)) <> 0 Then
        WVector1.TextMatrix(8, 2) = ZZTrabajaLote(16)
    End If
    If Val(ZZTrabajaLote(18)) <> 0 Then
        WVector1.TextMatrix(9, 2) = ZZTrabajaLote(18)
    End If
    If Val(ZZTrabajaLote(20)) <> 0 Then
        WVector1.TextMatrix(10, 2) = ZZTrabajaLote(20)
    End If
    If Val(ZZTrabajaLote(22)) <> 0 Then
        WVector1.TextMatrix(11, 2) = ZZTrabajaLote(22)
    End If
    If Val(ZZTrabajaLote(24)) <> 0 Then
        WVector1.TextMatrix(12, 2) = ZZTrabajaLote(24)
    End If
    
    If Val(ZZTrabajaLote(31)) <> 0 Then
        WVector1.TextMatrix(1, 3) = ZZTrabajaLote(31)
    End If
    If Val(ZZTrabajaLote(32)) <> 0 Then
        WVector1.TextMatrix(1, 5) = ZZTrabajaLote(32)
    End If
    If Val(ZZTrabajaLote(33)) <> 0 Then
        WVector1.TextMatrix(2, 3) = ZZTrabajaLote(33)
    End If
    If Val(ZZTrabajaLote(34)) <> 0 Then
        WVector1.TextMatrix(2, 5) = ZZTrabajaLote(34)
    End If
    If Val(ZZTrabajaLote(35)) <> 0 Then
        WVector1.TextMatrix(3, 3) = ZZTrabajaLote(35)
    End If
    If Val(ZZTrabajaLote(36)) <> 0 Then
        WVector1.TextMatrix(3, 5) = ZZTrabajaLote(36)
    End If
    If Val(ZZTrabajaLote(37)) <> 0 Then
        WVector1.TextMatrix(4, 3) = ZZTrabajaLote(37)
    End If
    If Val(ZZTrabajaLote(38)) <> 0 Then
        WVector1.TextMatrix(4, 5) = ZZTrabajaLote(38)
    End If
    If Val(ZZTrabajaLote(39)) <> 0 Then
        WVector1.TextMatrix(5, 3) = ZZTrabajaLote(39)
    End If
    If Val(ZZTrabajaLote(40)) <> 0 Then
        WVector1.TextMatrix(5, 5) = ZZTrabajaLote(40)
    End If
    If Val(ZZTrabajaLote(41)) <> 0 Then
        WVector1.TextMatrix(6, 3) = ZZTrabajaLote(41)
    End If
    If Val(ZZTrabajaLote(42)) <> 0 Then
        WVector1.TextMatrix(6, 5) = ZZTrabajaLote(42)
    End If
    If Val(ZZTrabajaLote(43)) <> 0 Then
        WVector1.TextMatrix(7, 3) = ZZTrabajaLote(43)
    End If
    If Val(ZZTrabajaLote(44)) <> 0 Then
        WVector1.TextMatrix(7, 5) = ZZTrabajaLote(44)
    End If
    If Val(ZZTrabajaLote(45)) <> 0 Then
        WVector1.TextMatrix(8, 3) = ZZTrabajaLote(45)
    End If
    If Val(ZZTrabajaLote(46)) <> 0 Then
        WVector1.TextMatrix(8, 5) = ZZTrabajaLote(46)
    End If
    If Val(ZZTrabajaLote(47)) <> 0 Then
        WVector1.TextMatrix(9, 3) = ZZTrabajaLote(47)
    End If
    If Val(ZZTrabajaLote(48)) <> 0 Then
        WVector1.TextMatrix(9, 5) = ZZTrabajaLote(48)
    End If
    If Val(ZZTrabajaLote(49)) <> 0 Then
        WVector1.TextMatrix(10, 3) = ZZTrabajaLote(49)
    End If
    If Val(ZZTrabajaLote(50)) <> 0 Then
        WVector1.TextMatrix(10, 5) = ZZTrabajaLote(50)
    End If
    If Val(ZZTrabajaLote(51)) <> 0 Then
        WVector1.TextMatrix(11, 3) = ZZTrabajaLote(51)
    End If
    If Val(ZZTrabajaLote(52)) <> 0 Then
        WVector1.TextMatrix(11, 5) = ZZTrabajaLote(52)
    End If
    If Val(ZZTrabajaLote(53)) <> 0 Then
        WVector1.TextMatrix(12, 3) = ZZTrabajaLote(53)
    End If
    If Val(ZZTrabajaLote(54)) <> 0 Then
        WVector1.TextMatrix(12, 5) = ZZTrabajaLote(54)
    End If
    
    If Val(ZZTrabajaLote(61)) <> 0 Then
        WVector1.TextMatrix(1, 6) = ZZTrabajaLote(61)
    End If
    If Val(ZZTrabajaLote(62)) <> 0 Then
        WVector1.TextMatrix(2, 6) = ZZTrabajaLote(62)
    End If
    If Val(ZZTrabajaLote(63)) <> 0 Then
        WVector1.TextMatrix(3, 6) = ZZTrabajaLote(63)
    End If
    If Val(ZZTrabajaLote(64)) <> 0 Then
        WVector1.TextMatrix(4, 6) = ZZTrabajaLote(64)
    End If
    If Val(ZZTrabajaLote(65)) <> 0 Then
        WVector1.TextMatrix(5, 6) = ZZTrabajaLote(65)
    End If
    If Val(ZZTrabajaLote(66)) <> 0 Then
        WVector1.TextMatrix(6, 6) = ZZTrabajaLote(66)
    End If
    If Val(ZZTrabajaLote(67)) <> 0 Then
        WVector1.TextMatrix(7, 6) = ZZTrabajaLote(67)
    End If
    If Val(ZZTrabajaLote(68)) <> 0 Then
        WVector1.TextMatrix(8, 6) = ZZTrabajaLote(68)
    End If
    If Val(ZZTrabajaLote(69)) <> 0 Then
        WVector1.TextMatrix(9, 6) = ZZTrabajaLote(69)
    End If
    If Val(ZZTrabajaLote(70)) <> 0 Then
        WVector1.TextMatrix(10, 6) = ZZTrabajaLote(70)
    End If
    If Val(ZZTrabajaLote(71)) <> 0 Then
        WVector1.TextMatrix(11, 6) = ZZTrabajaLote(71)
    End If
    If Val(ZZTrabajaLote(72)) <> 0 Then
        WVector1.TextMatrix(12, 6) = ZZTrabajaLote(72)
    End If
    
    ZZLugar = 0
    For Cicla = 31 To 53 Step 2
        ZZLugar = ZZLugar + 1
        spEnvase = "ConsultaEnvases " + "'" + ZZTrabajaLote(Cicla) + "'"
        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvase.RecordCount > 0 Then
            WVector1.TextMatrix(ZZLugar, 4) = Left$(Trim(rstEnvase!Abreviatura), 7)
            rstEnvase.Close
                Else
            WVector1.TextMatrix(ZZLugar, 4) = ""
        End If
    Next Cicla
    
    Call Conecta_Empresa
    
    Call Suma_Lote
    WVector1.Row = 1
    WVector1.Col = 1
    Rem Call Cantidad_KeyPress(13)
    
    Rem Call StartEdit
     
End Sub

Private Sub Suma_Lote()
    Suma = 0
    For Ciclo = 1 To 12
        If Trim(WVector1.TextMatrix(Ciclo, 1)) <> "" Then
            Suma = Suma + Val(WVector1.TextMatrix(Ciclo, 2))
        End If
    Next Ciclo
    Asignada.Text = Str$(Suma)
    Diferencia.Text = Str$(Val(Cantidad.Text) - Suma)
    Asignada.Text = Pusing("###,###.##", Asignada.Text)
    Diferencia.Text = Pusing("###,###.##", Diferencia.Text)
End Sub

Private Sub Verifica_Lote()

    WEstado = "N"
    Suma = 0
    XTerminado = Producto.Text
    
    For Ciclo = 1 To 12
        If Trim(WVector1.TextMatrix(Ciclo, 1)) <> "" Then
            Suma = Suma + Val(WVector1.TextMatrix(Ciclo, 2))
        End If
    Next Ciclo
        
    If Suma = Val(Cantidad.Text) Then
        WEstado = "S"
            Else
        Rem m$ = "Las cantidades asignadas no concuerdan con las cantidades a facturar"
        Rem A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
    End If
    
    If WEstado = "S" Then
    
        Erase ControlLote
        For Ciclo = 1 To 12
            ControlLote(Ciclo, 1) = WVector1.TextMatrix(Ciclo, 1)
            ControlLote(Ciclo, 2) = WVector1.TextMatrix(Ciclo, 2)
        Next Ciclo
    
        For Ciclo1 = 1 To 12
            If Val(ControlLote(Ciclo1, 1)) <> 0 Then
                For Ciclo2 = 1 To 12
                    If Ciclo1 <> Ciclo2 Then
                        If Val(ControlLote(Ciclo1, 1)) = Val(ControlLote(Ciclo2, 1)) <> 0 Then
                            Rem dada
                            Rem dada
                            Rem dada
                            Rem dada
                            Rem dada
                            Rem dada
                            Rem m$ = "A asignado una misma partida 2 veces"
                            Rem a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                            Rem WEstado = "N"
                            Rem Exit For
                        End If
                    End If
                Next Ciclo2
            End If
            If WEstado = "N" Then
                Exit For
            End If
        Next Ciclo1
        
    End If

    If WEstado = "S" Then
    
        Erase ControlLote
        For Ciclo = 1 To 12
            ControlLote(Ciclo, 1) = WVector1.TextMatrix(Ciclo, 1)
            ControlLote(Ciclo, 2) = WVector1.TextMatrix(Ciclo, 2)
        Next Ciclo
    
        For Ciclo1 = 1 To 12
    
            WLote = ControlLote(Ciclo1, 1)
            WCanti = Val(ControlLote(Ciclo1, 2))
            
            If WLote <> "" Or Val(WCanti) <> 0 Then
            
                If Left$(XTerminado, 2) <> "PT" And Left$(XTerminado, 2) <> "YQ" And Left$(XTerminado, 2) <> "YF" And Left$(XTerminado, 2) <> "YP" And Left$(XTerminado, 2) <> "YH" Then
                    WTipopro = "M"
                        Else
                    WTipopro = "T"
                End If
                
                Select Case WTipopro
                    Case "M"
                        WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                        WEntra = "N"
                        
                        XEmpresa = WEmpresa
                        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                            Select Case WTipoPedido
                                Case "PG", "CO"
                                    WEmpresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "FA"
                                    WEmpresa = "0005"
                                    txtOdbc = "Empresa05"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "TA"
                                    WEmpresa = "0003"
                                    txtOdbc = "Empresa03"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    WEmpresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                        End If
                        
                        ZSql = ""
                        If Val(WLote) = 0 Then
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Laudo"
                            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote + "'"
                            ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                                Else
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Laudo"
                            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Laudo.Laudo = " + "'" + WLote + "'"
                            ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                        End If
                        spLaudo = ZSql
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            With rstLaudo
                                .MoveFirst
                                WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                                Call Redondeo(WSaldo)
                                WEntra = "S"
                                If WSaldo < WCanti Then
                                    m$ = "La cantidad informada supera al saldo disponible"
                                    a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                    WEstado = "N"
                                End If
                                ZEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                                ZEstadoII = IIf(IsNull(rstLaudo!EstadoII), "", rstLaudo!EstadoII)
                                If ZEstado = "N" Then
                                    If ZEstadoII = "V" Then
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                            Else
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    End If
                                    WEstado = "N"
                                End If
                                rstLaudo.Close
                            End With
                        End If
                            
                        If WEntra = "N" Then
                            ZSql = ""
                            If Val(WLote) = 0 Then
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Guia"
                                ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                                ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote + "'"
                                ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                                    Else
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Guia"
                                ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                                ZSql = ZSql + " and Guia.Lote = " + "'" + WLote + "'"
                                ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                            End If
                            spMovguia = ZSql
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                With rstMovguia
                                    .MoveFirst
                                    WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                    Call Redondeo(WSaldo)
                                    WEntra = "S"
                                    If WSaldo < WCanti Then
                                        m$ = "La cantidad informada supera al saldo disponible"
                                        a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                        WEstado = "N"
                                    End If
                                    ZEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                                    ZEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                                    If ZEstado = "N" Then
                                        If ZEstadoII = "V" Then
                                            m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                                Else
                                            m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        End If
                                        WEstado = "N"
                                    End If
                                    rstMovguia.Close
                                End With
                            End If
                        End If
                        
                        Call Conecta_Empresa
                        
                        If WEntra = "N" Then
                            m$ = "Partida Inexistente"
                            a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                            WEstado = "N"
                        End If
                    
                    Case Else
                        WEntra = "N"
                        WControla = 0
                        
                        XEmpresa = WEmpresa
                        Select Case Val(WEmpresa)
                            Case 1, 3, 5, 6, 7, 10
                                WEmpresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                WEmpresa = "0008"
                                txtOdbc = "Empresa08"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                        
                        spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                            rstTerminado.Close
                        End If
                        
                        Call Conecta_Empresa
                
                        If WControla = 0 Then
                        
                            XEmpresa = WEmpresa
                            If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                                Select Case WTipoPedido
                                    Case "PG", "CO"
                                        WEmpresa = "0001"
                                        txtOdbc = "Empresa01"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case "FA"
                                        WEmpresa = "0005"
                                        txtOdbc = "Empresa05"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case "TA"
                                        WEmpresa = "0003"
                                        txtOdbc = "Empresa03"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case Else
                                        WEmpresa = "0007"
                                        txtOdbc = "Empresa07"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                End Select
                            End If
                        
                            XParam = "'" + WLote + "','" _
                                    + XTerminado + "'"
                            spHoja = "ListaHojaProducto " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                                WSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                                Call Redondeo(WSaldo)
                                WEntra = "S"
                                If WSaldo < WCanti Then
                                    m$ = "La cantidad informada supera al saldo disponible"
                                    a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                    WEstado = "N"
                                End If
                                ZEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
                                ZEstadoII = IIf(IsNull(rstHoja!EstadoII), "", rstHoja!EstadoII)
                                If ZEstado = "N" Then
                                    If ZEstadoII = "V" Then
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                            Else
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    End If
                                    WEstado = "N"
                                End If
                                WFechaHoja = rstHoja!Fecha
                                rstHoja.Close
                                Rem WVida = 0
                                Rem
                                Rem spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                                Rem Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                Rem If rstTerminado.RecordCount > 0 Then
                                Rem     WVida = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
                                Rem     rstTerminado.Close
                                Rem End If
                                Rem
                                Rem If WVida <> 0 Then
                                Rem
                                Rem     WMes = Val(Mid$(WFechaHoja, 4, 2))
                                Rem     WAno = Val(Right$(WFechaHoja, 4))
                                Rem     For Ciclo = 1 To WVida
                                Rem         WMes = WMes + 1
                                Rem         If WMes > 12 Then
                                Rem             WAno = WAno + 1
                                Rem             WMes = 1
                                Rem         End If
                                Rem     Next Ciclo
                                Rem     XMes = Str$(WMes)
                                Rem     XAno = Str$(WAno)
                                Rem     Call Ceros(XMes, 2)
                                Rem     Call Ceros(XAno, 4)
                                Rem     WVencimiento = "01/" + XMes + "/" + XAno
                                Rem
                                Rem     WFechaActual = "01" + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                                Rem     WFechaActualOrd = Right$(WFechaActual, 4) + Mid$(WFechaActual, 4, 2) + Left$(WFechaActual, 2)
                                Rem
                                Rem     WFechaVencimiento = "01" + Mid$(WVencimiento, 3, 10)
                                Rem     WFechaVencimientoOrd = Right$(WFechaVencimiento, 4) + Mid$(WFechaVencimiento, 4, 2) + Left$(WFechaVencimiento, 2)
                                Rem
                                Rem     Pasa = "S"
                                Rem     If WFechaActualOrd >= WFechaVencimientoOrd Then
                                Rem         Pasa = "N"
                                Rem             Else
                                Rem         Meses = 0
                                Rem         WMes = Val(Mid$(WFechaActual, 4, 2))
                                Rem         WAno = Val(Right$(WFechaActual, 4))
                                Rem         Do
                                Rem             Meses = Meses + 1
                                Rem             WMes = WMes + 1
                                Rem             If WMes > 12 Then
                                Rem                 WAno = WAno + 1
                                Rem                 WMes = 1
                                Rem             End If
                                Rem             XMes = Str$(WMes)
                                Rem             XAno = Str$(WAno)
                                Rem             Call Ceros(XMes, 2)
                                Rem             Call Ceros(XAno, 4)
                                Rem             WCompara = "01/" + XMes + "/" + XAno
                                Rem             If WCompara = WFechaVencimiento Then
                                Rem                 Exit Do
                                Rem             End If
                                Rem         Loop
                                Rem         If Meses <= 12 Then
                                Rem             Pasa = "N"
                                Rem         End If
                                Rem     End If
                                Rem
                                Rem     If Pasa = "N" Then
                                Rem         m$ = "EL Producto tiene menos de un a�o de vida util"
                                Rem         G% = MsgBox(m$, 0, "Actualizacion de Pedido")
                                Rem         WEstado = "N"
                                Rem     End If
                                Rem
                                Rem End If
                            End If
                    
                            If WEntra = "N" Then
                                XParam = "'" + XTerminado + "','" _
                                            + WLote + "'"
                                spMovguia = "ListaMovguiaLote1 " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                If rstMovguia.RecordCount > 0 Then
                                    WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                    Call Redondeo(WSaldo)
                                    WEntra = "S"
                                    If WSaldo < WCanti Then
                                        m$ = "La cantidad informada supera al saldo disponible"
                                        a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                        WEstado = "N"
                                    End If
                                    ZEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                                    ZEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                                    If ZEstado = "N" Then
                                        If ZEstadoII = "V" Then
                                            m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                                Else
                                            m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        End If
                                        WEstado = "N"
                                    End If
                                    rstMovguia.Close
                                End If
                            End If
                                    
                            Call Conecta_Empresa
                            
                    
                                Else
                                
                            WEntra = "S"
                            
                        End If
                        
                        If WEntra = "N" Then
                            m$ = "Partida Inexistente"
                            a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                            WEstado = "N"
                        End If
                    
                End Select
            
            End If
            
        Next Ciclo1

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
        Call Suma_Lote
        If WVector1.Col = 1 Then
            If WVector1.Text = "" Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WControlII = "N"
                    Call Confirma_Click
                End If
            End If
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
            If WControlII = "" Then
                Call Control_Campo
                If WControl = "S" Then
                    Call Control_wvector1
                End If
                Call StartEdit
            End If

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
            If WControlII = "" Then
                If WControl = "S" Then
                    Call Control_wvector1
                End If
                Call StartEdit
            End If
    
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
            If WControlII = "" Then
                If WControl = "S" Then
                    Call Control_wvector1
                End If
                Call StartEdit
            End If

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
        Case 6
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
            XTerminado = Producto.Text
            WLote = WVector1.Text
            WSaldo = 0
        
            If Left$(XTerminado, 2) <> "PT" And Left$(XTerminado, 2) <> "YQ" And Left$(XTerminado, 2) <> "YF" And Left$(XTerminado, 2) <> "YP" And Left$(XTerminado, 2) <> "YH" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            WEstado = ""
            WBuscaEmpresa = ""
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
                    XEmpresa = WEmpresa
                    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                        Select Case WTipoPedido
                            Case "PG", "CO"
                                WEmpresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case "FA"
                                WEmpresa = "0005"
                                txtOdbc = "Empresa05"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case "TA"
                                WEmpresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                WEmpresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                    End If
                    
                    If Val(WLote) = 0 Then
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                            Else
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.Laudo = " + "'" + WLote + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    End If
                    spLaudo = ZSql
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            
                            WEntra = "S"
                            
                            WEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                            If WEstado <> "N" Then
                                WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                                    Else
                                WEstadoII = IIf(IsNull(rstLaudo!EstadoII), "", rstLaudo!EstadoII)
                                If WEstadoII = "V" Then
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        Else
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                End If
                                WSaldo = 0
                            End If
                            
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                        If Val(WLote) = 0 Then
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                                Else
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.Lote = " + "'" + WLote + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        End If
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
                                
                                WEntra = "S"
                                
                                WEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                                If WEstado <> "N" Then
                                    WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                        Else
                                    WEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                                    If WEstadoII = "V" Then
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                            Else
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    End If
                                    WSaldo = 0
                                End If
                                
                                rstMovguia.Close
                            End With
                        End If
                        
                    End If
                    
                    WBuscaEmpresa = WEmpresa
                    Call Conecta_Empresa
                    
                Case Else
                    WEntra = "N"
                    WControla = 0
                    
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10
                            WEmpresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                            WEmpresa = "0008"
                            txtOdbc = "Empresa08"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End Select
                    
                    spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        rstTerminado.Close
                    End If
                    
                    Call Conecta_Empresa
                    
                    If WControla = 0 Then
                    
                        XEmpresa = WEmpresa
                        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                            Select Case WTipoPedido
                                Case "PG", "CO"
                                    WEmpresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "FA"
                                    WEmpresa = "0005"
                                    txtOdbc = "Empresa05"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "TA"
                                    WEmpresa = "0003"
                                    txtOdbc = "Empresa03"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    WEmpresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                        End If
                    
                        XParam = "'" + WLote + "','" _
                                + XTerminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                        
                            WEntra = "S"
                            
                            WEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
                            If WEstado <> "N" Then
                                WSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                                    Else
                                WEstadoII = IIf(IsNull(rstHoja!EstadoII), "", rstHoja!EstadoII)
                                If WEstadoII = "V" Then
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        Else
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                End If
                                WSaldo = 0
                            End If
                            
                            WMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
                            If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                                m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                     "Por favor comuniquese con el laboratorio para su revalida"
                                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                WSaldo = 0
                            End If
                            
                            rstHoja.Close
                            
                        End If
                
                        If WEntra = "N" Then
                            XParam = "'" + XTerminado + "','" _
                                    + WLote + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                            
                                WEntra = "S"
                                
                                WEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                                If WEstado <> "N" Then
                                    WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                        Else
                                    WEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                                    If WEstadoII = "V" Then
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                            Else
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    End If
                                    WSaldo = 0
                                End If
                                
                                WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                                If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                                    m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                        "Por favor comuniquese con el laboratorio para su revalida"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    WSaldo = 0
                                End If
                                
                                aa = rstMovguia!Clave
                                rstMovguia.Close
                            End If
                        End If
                        
                        WBuscaEmpresa = WEmpresa
                        Call Conecta_Empresa
                
                            Else
                    
                        WEntra = "S"
                
                    End If
            End Select
                
            If WEntra = "S" Then
                WVector1.TextMatrix(WVector1.Row, 7) = Str$(WSaldo)
                    Else
                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                    Select Case Val(WBuscaEmpresa)
                        Case 1
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta I"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case 5
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta III (FARMA)"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case 7
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta V"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case Else
                    End Select
                End If
                WControl = "N"
            End If
        
        Case 2
            WSaldo = Val(WVector1.TextMatrix(WVector1.Row, 7))
            If WSaldo >= Val(WVector1.Text) Then
                    Else
                m$ = XTerminado + " Cantidad Insuficiente Stock : " + Str$(WSaldo)
                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                WControl = "N"
            End If
            
        Case 3
            If Val(WVector1.Text) <> 0 Then
        
                XEmpresa = WEmpresa
                Select Case Val(WEmpresa)
                    Case 1, 3, 5, 6, 7, 10
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        WEmpresa = "0008"
                        txtOdbc = "Empresa08"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
        
                spEnvases = "ConsultaEnvases " + "'" + WVector1.Text + "'"
                Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvases.RecordCount > 0 Then
                    WVector1.Col = 4
                    WVector1.Text = rstEnvases!Abreviatura
                    rstEnvases.Close
                        Else
                    WControl = "N"
                End If
            
                Call Conecta_Empresa
                
                    Else
                    
                WVector1.TextMatrix(WVector1.Row, 4) = ""
                WVector1.TextMatrix(WVector1.Row, 5) = ""
                WVector1.TextMatrix(WVector1.Row, 6) = ""
                
                If WVector1.Row < 12 Then
                    WVector1.Row = WVector1.Row + 1
                End If
                WVector1.Col = 1
                
                WControl = "N"
                    
            End If
            
        Case 5
            If Val(WEmpresa) = 1 Then
                Select Case Val(WVector1.TextMatrix(WVector1.Row, 3))
                    Case 20, 21, 22, 23, 24, 25, 26, 28, 30
                        WVector1.TextMatrix(WVector1.Row, 6) = WVector1.Text
                    Case Else
                End Select
            End If
        
            
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
    WVector1.Cols = 8
    WVector1.FixedRows = 1
    WVector1.Rows = 13
    
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
                WVector1.Text = "Lote"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Envase"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 1800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Cant.Envases"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Bultos"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
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
    Rem modificar el tama�o de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub



































Private Sub Conecta_Empresa()

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

End Sub

Private Sub Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WVector1.Row = 1
        WVector1.Col = 1
        Call StartEdit
    End If
End Sub

Private Sub WTexto1_DblClick()
    Call ficha_Pt
End Sub

Private Sub WTexto2_DblClick()
    Call ficha_Pt
End Sub

Private Sub ficha_Pt()

    XEmpresa = WEmpresa
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
        Select Case WTipoPedido
            Case "PG", "CO"
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case "FA"
                WEmpresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case "TA"
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    End If


    Call Limpia_Vector2
    WTerminado = Producto.Text
    XRenglon = 0
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spHoja = "ListaHojaProductoDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstHoja!Marca = "X" And rstHoja!Saldo = 0 Then
                
                    Else
                
                If Val(rstHoja!Renglon) = 1 Then
                Rem And rstHoja!Real <> 0 Then
                 
                    ZProducto = rstHoja!Producto
                    ZCantidad = rstHoja!Real
                    ZFecha = rstHoja!Fecha
                    ZHoja = rstHoja!Hoja
                    ZSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Call Redondeo(ZSaldo)
                    
                    If ZSaldo <> 0 Then
                    
                        XRenglon = XRenglon + 1
                        WVector2.Row = XRenglon
                
                        WVector2.Col = 1
                        WVector2.Text = "Hoja"
                        
                        WVector2.Col = 2
                        WVector2.Text = ZHoja
                                               
                        WVector2.Col = 3
                        WVector2.Text = ZFecha
                        
                        WVector2.Col = 4
                        WVector2.Text = ""
                        
                        WVector2.Col = 5
                        WVector2.Text = ZCantidad
                
                        WVector2.Col = 6
                        WVector2.Text = ZSaldo
                
                        WVector2.Col = 7
                        WVector2.Text = ZHoja
                        
                        WVector2.Col = 8
                        WVector2.Text = ""
                    
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
    End If
    
    
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
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
                
                If rstMovguia!Tipo = "T" Then
                
                    ZTerminado = rstMovguia!Terminado
                    ZCantidad = rstMovguia!Cantidad
                    ZFecha = rstMovguia!Fecha
                    ZCodigo = rstMovguia!Codigo
                    zMovi = rstMovguia!Movi
                    ZDestino = rstMovguia!Destino
                    ZTipomov = rstMovguia!Tipomov
                    WWLote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                    ZPartida = IIf(IsNull(rstMovguia!Partida), "", rstMovguia!Partida)
                    ZSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(ZSaldo)
                    If Val(ZCodigo) > 900000 Then
                        WWTipo = "Prestamo"
                        ZCodigo = WCodigo - 900000
                            Else
                        WWTipo = "Guia In"
                    End If
                    
                    If zMovi = "E" And ZSaldo <> 0 Then
                    
                        XRenglon = XRenglon + 1
                        WVector2.Row = XRenglon
                
                        WVector2.Col = 1
                        WVector2.Text = WWTipo
                        
                        WVector2.Col = 2
                        WVector2.Text = ZCodigo
                                               
                        WVector2.Col = 3
                        WVector2.Text = ZFecha
                        
                        WVector2.Col = 4
                        WVector2.Text = ""
                        
                        WVector2.Col = 5
                        WVector2.Text = ZCantidad
                
                        WVector2.Col = 6
                        WVector2.Text = ZSaldo
                
                        WVector2.Col = 7
                        WVector2.Text = WWLote
                        
                        WVector2.Col = 8
                        WVector2.Text = ""
                        
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
    End If
    
    
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spEntdev = "ListaEntdevTerminadoDesdeHasta" + XParam
    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    If rstEntdev.RecordCount > 0 Then
    
        With rstEntdev
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstEntdev!Marca = "X" Then
                
                        Else
                
                ZTerminado = rstEntdev!Terminado
                ZCantidad = rstEntdev!Cantidad
                ZFecha = rstEntdev!Fecha
                ZCodigo = rstEntdev!Codigo
                WWLote = IIf(IsNull(rstEntdev!Lote), "0", rstEntdev!Lote)
                ZSaldo = rstEntdev!Saldo
                Call Redondeo(ZSaldo)
                
                If ZSaldo <> 0 Then
                    
                    XRenglon = XRenglon + 1
                    WVector2.Row = XRenglon
                
                    WVector2.Col = 1
                    WVector2.Text = "Dev"
                        
                    WVector2.Col = 2
                    WVector2.Text = ZCodigo
                                               
                    WVector2.Col = 3
                    WVector2.Text = ZFecha
                        
                    WVector2.Col = 4
                    WVector2.Text = ""
                        
                    WVector2.Col = 5
                    WVector2.Text = ZCantidad
                
                    WVector2.Col = 6
                    WVector2.Text = ZSaldo
                
                    WVector2.Col = 7
                    WVector2.Text = WWLote
                    
                    WVector2.Col = 8
                    WVector2.Text = ""

                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
        
        rstEntdev.Close
        
    End If
                    
    WBuscaEmpresa = WEmpresa
    Call Conecta_Empresa
    
    WVector2.Col = 1
    WVector2.Row = 1
    
    WVector2.TopRow = 1
    
End Sub


Private Sub Limpia_Vector2()

    WVector2.Height = 6095
    WVector2.Left = 120
    WVector2.Top = 1050
    WVector2.Width = 12000

    WVector2.Clear
    WVector2.Font.Bold = True
    
    WVector2.FixedCols = 1
    WVector2.Cols = 12
    WVector2.FixedRows = 1
    WVector2.Rows = 5001
    
    WVector2.ColWidth(0) = 200
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Tipo"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector2.Text = "Numero"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                WVector2.Text = "Fecha"
                WVector2.ColWidth(Ciclo) = 1500
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                WVector2.Text = "Orden"
                WVector2.ColWidth(Ciclo) = 10
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WVector2.Text = "Cantidad"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector2.Text = "Saldo"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector2.Text = "Partida"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                WVector2.Text = "Partida"
                WVector2.ColWidth(Ciclo) = 10
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 9
                WVector2.Text = "Envase"
                WVector2.ColWidth(Ciclo) = 10
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 10
                WVector2.Text = "Cant.Ped."
                WVector2.ColWidth(Ciclo) = 10
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 11
                WVector2.Text = "Disponible"
                WVector2.ColWidth(Ciclo) = 10
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
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
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tama�o de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Visible = True
    
    WVector2.Col = 1
    WVector2.Row = 1
    
End Sub

Private Sub WVector2_Click()
    If Trim(WVector2.TextMatrix(WVector2.Row, 7)) <> "" Then
        WVector1.TextMatrix(WVector1.Row, 1) = WVector2.TextMatrix(WVector2.Row, 7)
    End If
    WVector2.Visible = False
    Call StartEdit
End Sub
