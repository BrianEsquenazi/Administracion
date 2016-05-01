VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgHojaRuta 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "consulta de pedido a despachar"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   495
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin VB.Frame EntraPedido 
      Height          =   6495
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid Muestra 
         Height          =   5295
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9340
         _Version        =   327680
         Rows            =   4000
         Cols            =   6
         BackColor       =   16777088
      End
      Begin VB.Image FinPedido 
         Height          =   480
         Left            =   5520
         MouseIcon       =   "hojaruta.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "hojaruta.frx":030A
         ToolTipText     =   "Salida"
         Top             =   5760
         Width           =   480
      End
   End
   Begin VB.CommandButton AltaCheckList 
      Caption         =   "CheckList"
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
      Left            =   8280
      TabIndex        =   31
      Top             =   6480
      Width           =   1215
   End
   Begin VB.ComboBox TipoEstado 
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
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox RetiraProv 
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
      Left            =   4320
      MaxLength       =   50
      TabIndex        =   29
      Top             =   840
      Width           =   7335
   End
   Begin VB.TextBox NroViaje 
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
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   27
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox TipoPedido 
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
      Left            =   4680
      TabIndex        =   25
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox TotalKilos 
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
      Left            =   10680
      MaxLength       =   6
      TabIndex        =   24
      Top             =   120
      Width           =   975
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
      Height          =   4215
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7435
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.CommandButton CargaPedido 
      Height          =   735
      Left            =   7320
      MouseIcon       =   "hojaruta.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "hojaruta.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Ingresa POE"
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox Camion 
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
      Left            =   7440
      MaxLength       =   6
      TabIndex        =   19
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Chofer 
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
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   16
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Hoja 
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
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   1095
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
      TabIndex        =   13
      Top             =   5880
      Width           =   1095
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
      ItemData        =   "hojaruta.frx":1720
      Left            =   120
      List            =   "hojaruta.frx":1727
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   3120
      TabIndex        =   12
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Retira Prov."
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
      Left            =   2760
      TabIndex        =   28
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro de Viaje"
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
      TabIndex        =   26
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Kilos"
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
      TabIndex        =   23
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label DesCamion 
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
      Left            =   8400
      TabIndex        =   18
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Camion"
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
      Left            =   5880
      TabIndex        =   17
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chofer"
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
      TabIndex        =   15
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label DesChofer 
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
      Left            =   2760
      TabIndex        =   14
      Top             =   480
      Width           =   2895
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
      Left            =   2400
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   9720
      MouseIcon       =   "hojaruta.frx":1735
      MousePointer    =   99  'Custom
      Picture         =   "hojaruta.frx":1A3F
      ToolTipText     =   "Salida"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7320
      MouseIcon       =   "hojaruta.frx":2281
      MousePointer    =   99  'Custom
      Picture         =   "hojaruta.frx":258B
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   8160
      MouseIcon       =   "hojaruta.frx":2DCD
      MousePointer    =   99  'Custom
      Picture         =   "hojaruta.frx":30D7
      ToolTipText     =   "Consulta de Datos"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   9000
      MouseIcon       =   "hojaruta.frx":3919
      MousePointer    =   99  'Custom
      Picture         =   "hojaruta.frx":3C23
      ToolTipText     =   "Limpia la pantalla"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Nro. Hoja"
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
Attribute VB_Name = "PrgHojaRuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstChofer As Recordset
Dim spChofer As String
Dim rstHojaRuta As Recordset
Dim rsHojaRuta As String
Dim rstCamion As Recordset
Dim spCamion As String
Dim rstCheckListExpo As Recordset
Dim spCheckListExpo As String
Dim rstMuestra As Recordset
Dim spMuestra As String

Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Cantidad As Double
Dim Renglon As Integer
Dim ZCodigo As String
Dim ZOperador As String
Dim WVersion As String
Dim WRenglon As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
Dim ZArti(100) As String
Dim ZBajaHoja(100) As String
Dim ZProceso As String
Dim ZZVector(100) As String

Dim ZZAyuda(5000, 10) As String


Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Private Sub AltaCheckList_Click()
    Rem dada
    Rem dada
    Rem dada
    
    ZExterior = "N"
    For Ciclo = 1 To 100
        ZCliente = WVector1.TextMatrix(Ciclo, 2)
        If Trim(ZCliente) <> "" Then
            spCliente = "ConsultaCliente " + ZCliente
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                WProvincia = rstCliente!Provincia
                If WProvincia = 24 Then
                    ZExterior = "S"
                End If
                rstCliente.Close
            End If
        End If
    Next Ciclo
    
    If ZExterior = "S" Then
        
        If Val(Hoja.Text) <> 0 Then
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM HojaRuta"
            ZSql = ZSql + " Where HojaRuta.Hoja = " + "'" + Hoja.Text + "'"
            ZSql = ZSql + " Order by HojaRuta.Clave"
            rsHojaRuta = ZSql
            Set rstHojaRuta = db.OpenRecordset(rsHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
            If rstHojaRuta.RecordCount > 0 Then
                rstHojaRuta.Close
                    Else
                Hoja.Text = "0"
            End If
        
        End If
        
        If Val(Hoja.Text) = 0 Then
            
            ZSql = ""
            ZSql = ZSql + "Select Max(Hoja) as [HojaMayor]"
            ZSql = ZSql + " FROM HojaRuta"
            spHojaRuta = ZSql
            Set rstHojaRuta = db.OpenRecordset(spHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
            If rstHojaRuta.RecordCount > 0 Then
                rstHojaRuta.MoveLast
                WHojaMayor = IIf(IsNull(rstHojaRuta!HojaMayor), "0", rstHojaRuta!HojaMayor)
                Hoja.Text = Str$(WHojaMayor + 1)
                rstHojaRuta.Close
                    Else
                Hoja.Text = "1"
            End If
            
        End If
    
        ZZPeligrosa = "N"
        ZZPasaHoja = Hoja.Text
        ZZPasaFecha = Fecha.Text
        
        For Ciclo = 1 To 100
        
            ZPedido = WVector1.TextMatrix(Ciclo, 1)
            ZCliente = WVector1.TextMatrix(Ciclo, 2)
            
            If Val(ZPedido) <> 0 Then
            
                Erase ZZVector
                ZLugar = 0
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Pedido"
                ZSql = ZSql + " Where Pedido.Pedido = " + "'" + ZPedido + "'"
                spPedido = ZSql
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                If rstPedido.RecordCount > 0 Then
                    With rstPedido
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                        
                                ZLugar = ZLugar + 1
                                ZZVector(ZLugar) = rstPedido!terminado
                                            
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    
                    rstPedido.Close
                    
                End If
                
                For CicloII = 1 To ZLugar
                
                    ZZTerminado = ZZVector(CicloII)
                    ZZClase = ""
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Terminado"
                    ZSql = ZSql + " Where Terminado.Codigo = " + "'" + ZZTerminado + "'"
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        ZZClase = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
                        rstTerminado.Close
                    End If
                    
                    If Trim(ZZClase) <> "" Then
                        ZZPeligrosa = "S"
                    End If
                    
                Next CicloII
            
            End If
            
        Next Ciclo
        
        PrgCheckListExpo.Show
        
    End If

End Sub

Private Sub CargaPedido_Click()

    If EntraPedido.Visible = True Then
        EntraPedido.Visible = False
        Exit Sub
    End If
    
    Opcion.Clear
    
     Opcion.AddItem ""
     Opcion.AddItem ""
     Opcion.AddItem "Pedidos"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 2
    
    Rem Call Opcion_Click
End Sub

Private Sub Command1_Click()
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Muestra SET "
            ZSql = ZSql + " HojaRuta = " + "'" + "" + "'"
            spMuestra = ZSql
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)

End Sub

Private Sub FinPedido_Click()
    EntraPedido.Visible = False
End Sub

Private Sub Form_Activate()
    If ZLLave = "N" Then
        If ZLugarAtraso < ZHastaAtraso Then
            ZLugarAtraso = ZLugarAtraso + 1
            PrgAtrasoExpedicion.Show
        End If
    End If
    ZLLave = ""
End Sub

Private Sub Hoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        Call Limpia_Vector
        WRenglon = 0
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM HojaRuta"
        ZSql = ZSql + " Where HojaRuta.Hoja = " + "'" + Hoja.Text + "'"
        ZSql = ZSql + " Order by HojaRuta.Clave"
    
        rsHojaRuta = ZSql
        Set rstHojaRuta = db.OpenRecordset(rsHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
        If rstHojaRuta.RecordCount > 0 Then
        
            rstHojaRuta.Close
            
            Call Proceso_Click
            
            WVector1.TopRow = 1
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
            
                Else
                
            Fecha.SetFocus
            
        End If
        
    End If
    If KeyAscii = 27 Then
        Hoja.Text = ""
    End If
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Chofer"
     Opcion.AddItem "Camion"
     Opcion.Visible = True
     
End Sub

Private Sub Muestra_Click()

    For Ciclo = 1 To 100
        If Trim(WVector1.TextMatrix(Ciclo, 1)) = "" Then
            WVector1.Row = Ciclo
            WVector1.Col = 1
            WVector1.Text = Muestra.TextMatrix(Muestra.Row, 1)
            WTexto2.Text = Muestra.TextMatrix(Muestra.Row, 1)
            Call WTexto2_KeyDown(13, 0)
            WVector1.Row = WVector1.Row + 1
            WVector1.Col = 1
            Call StartEdit
            Exit For
        End If
    Next Ciclo

    Muestra.Col = 1
    Muestra.Text = ""
    Muestra.Col = 2
    Muestra.Text = ""
    Muestra.Col = 3
    Muestra.Text = ""
    Muestra.Col = 4
    Muestra.Text = ""
    Muestra.Col = 5
    Muestra.Text = ""
    
End Sub

Private Sub Opcion_Click()

    Rem On Error GoTo WError
    
    Dim IngresaItem As String
    
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Ayuda.Visible = True
            Ayuda.Text = ""
            Sql1 = "Select *"
            Sql2 = " FROM Chofer"
            Sql3 = " Order by Codigo"
            spChofer = Sql1 + Sql2 + Sql3
            Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
            If rstChofer.RecordCount > 0 Then
                With rstChofer
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstChofer!Codigo) + " " + rstChofer!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = Str$(rstChofer!Codigo)
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstChofer.Close
            End If
            Ayuda.SetFocus
            Pantalla.Visible = True
            
        Case 1
            Ayuda.Visible = True
            Ayuda.Text = ""
            Sql1 = "Select *"
            Sql2 = " FROM Camion"
            Sql3 = " Order by Codigo"
            spCamion = Sql1 + Sql2 + Sql3
            Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
            If rstCamion.RecordCount > 0 Then
                With rstCamion
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstCamion!Codigo) + " " + rstCamion!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = Str$(rstCamion!Codigo)
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCamion.Close
            End If
            Ayuda.SetFocus
            Pantalla.Visible = True
            
        Case 2
            WTexto1.Visible = False
            WTexto2.Visible = False
            WTexto3.Visible = False
            Muestra.Clear
            Muestra.Row = 0
            
            Muestra.Font.Bold = True
    
            Muestra.ColWidth(0) = 50
            Muestra.ColWidth(1) = 900
            Muestra.ColWidth(2) = 3500
            Muestra.ColWidth(3) = 900
            Muestra.ColWidth(4) = 900
            Muestra.ColWidth(5) = 6000
    
            Muestra.Col = 1
            Muestra.Text = "Pedido"
    
            Muestra.Col = 2
            Muestra.Text = "Razon Social"
            
            Muestra.Col = 3
            Muestra.Text = "Tipo"
    
            Muestra.Col = 4
            Muestra.Text = "Kilos"
    
            Muestra.Col = 5
            Muestra.Text = "Direccion"
    
            ZZRenglon = 0
            Erase ZZAyuda
            ZZLugar = 0
            
            ZZCorte = 0
            ZZPasa = 0
            ZZZKilos = 0
            ZZTipoPed = 0
    
            ZSql = ""
            ZSql = ZSql + "Select Pedido.Pedido, Pedido.TipoPedido, Pedido.Autorizo, Pedido.Fecha, Pedido.Cliente, Pedido.TipoPed, Pedido.FecEntrega, Pedido.DirEntrega, Pedido.CantidadFac, Pedido.CantiLote1, Pedido.CantiLote2, Pedido.CantiLote3, Pedido.CantiLote4, Pedido.CantiLote5, Pedido.Cantidad, Pedido.Facturado, Cliente.Razon as [WRazon], Cliente.DirEntrega as [WDirEntrega], Cliente.DirEntregaII as [WDirEntregaII], Cliente.DirEntregaIII as [WDirEntregaIII], Cliente.DirEntregaIV as [WDirEntregaIV], Cliente.DirEntregaV as [WDirEntregaV]"
            ZSql = ZSql + " FROM Pedido, Cliente"
            ZSql = ZSql + " Where Pedido.TipoPed <> 4"
            ZSql = ZSql + " and Pedido.HojaRuta = 0"
            ZSql = ZSql + " and (Pedido.CantidadFac <> 0 or Pedido.Cantidad > Pedido.Facturado)"
            ZSql = ZSql + " and Pedido.Cliente = Cliente.Cliente"
            ZSql = ZSql + " Order by Pedido.Pedido"
    
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                With rstPedido
                    .MoveFirst
                    Do
                        
                        If .EOF = False Then
                        
                            If rstPedido!autorizo = "X" Then
        
                                Pasa = "N"
                                If TipoPedido.ListIndex = 0 Then
                                    Pasa = "S"
                                End If
                                If TipoPedido.ListIndex = 1 Then
                                    If rstPedido!TipoPedido = 2 Or rstPedido!TipoPedido = 3 Then
                                        Pasa = "S"
                                    End If
                                End If
                                If TipoPedido.ListIndex = 2 Then
                                    If rstPedido!TipoPedido = 1 Then
                                        Pasa = "S"
                                    End If
                                End If
                                If TipoPedido.ListIndex = 3 Then
                                    If rstPedido!TipoPedido = 4 Then
                                        Pasa = "S"
                                    End If
                                End If
                                If TipoPedido.ListIndex = 4 Then
                                    If rstPedido!TipoPedido = 5 Then
                                        Pasa = "S"
                                    End If
                                End If
                            
                                If Pasa = "S" And !Cliente <> "P00005" Then
                                
                                    If ZZPasa = 0 Then
                                        ZZPasa = 1
                                        ZZCorte = rstPedido!Pedido
                                        ZZZKilos = 0
                                        ZZTipoPed = rstPedido!tipoped
                                        
                                        ZZDirEntrega = ""
                                        ZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
                                        ZDirEntrega(1) = rstPedido!WDirentrega
                                        ZDirEntrega(2) = Trim(IIf(IsNull(rstPedido!WDirEntregaII), "", rstPedido!WDirEntregaII))
                                        ZDirEntrega(3) = Trim(IIf(IsNull(rstPedido!WDirEntregaIII), "", rstPedido!WDirEntregaIII))
                                        ZDirEntrega(4) = Trim(IIf(IsNull(rstPedido!WDirEntregaIV), "", rstPedido!WDirEntregaIV))
                                        ZDirEntrega(5) = Trim(IIf(IsNull(rstPedido!WDirEntregaV), "", rstPedido!WDirEntregaV))
                                        
                                        ZZDirEntrega = ZDirEntrega(ZLugarDirEntrega)
                                        ZZRazon = rstPedido!WRazon
                                        ZZTipoPedido = rstPedido!tipoped
                                    End If
                                
                                    If ZZCorte <> rstPedido!Pedido Then
                                
                                        ZZRenglon = ZZRenglon + 1
                            
                                        Muestra.TextMatrix(ZZRenglon, 0) = Pusing("###", Str$(ZZRenglon))
                                        Muestra.TextMatrix(ZZRenglon, 1) = Pusing("######", Str$(ZZCorte))
                                        Muestra.TextMatrix(ZZRenglon, 2) = ZZRazon
                                        If ZZTipoPed = 5 Then
                                            Muestra.TextMatrix(ZZRenglon, 3) = "Muestra"
                                                Else
                                            Muestra.TextMatrix(ZZRenglon, 3) = "Pedido"
                                        End If
                                        Muestra.TextMatrix(ZZRenglon, 4) = Pusing("######", Str$(ZZZKilos))
                                        Muestra.TextMatrix(ZZRenglon, 5) = ZZDirEntrega
                            
                                        ZZAyuda(ZZRenglon, 1) = Pusing("###", Str$(ZZRenglon))
                                        ZZAyuda(ZZRenglon, 2) = Pusing("######", Str$(ZZCorte))
                                        ZZAyuda(ZZRenglon, 3) = ZZRazon
                                        If ZZTipoPed = 5 Then
                                            ZZAyuda(ZZRenglon, 4) = "Muestra"
                                                Else
                                            ZZAyuda(ZZRenglon, 4) = "Pedido"
                                        End If
                                        ZZAyuda(ZZRenglon, 5) = ""
                                        ZZAyuda(ZZRenglon, 6) = ZZDirEntrega
                                        
                                        ZZCorte = rstPedido!Pedido
                                        ZZZKilos = 0
                                        
                                        ZZDirEntrega = ""
                                        ZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
                                        ZDirEntrega(1) = rstPedido!WDirentrega
                                        ZDirEntrega(2) = Trim(IIf(IsNull(rstPedido!WDirEntregaII), "", rstPedido!WDirEntregaII))
                                        ZDirEntrega(3) = Trim(IIf(IsNull(rstPedido!WDirEntregaIII), "", rstPedido!WDirEntregaIII))
                                        ZDirEntrega(4) = Trim(IIf(IsNull(rstPedido!WDirEntregaIV), "", rstPedido!WDirEntregaIV))
                                        ZDirEntrega(5) = Trim(IIf(IsNull(rstPedido!WDirEntregaV), "", rstPedido!WDirEntregaV))
                                        
                                        ZZDirEntrega = ZDirEntrega(ZLugarDirEntrega)
                                        ZZRazon = rstPedido!WRazon
                                        ZZTipoPedido = rstPedido!tipoped
                                        ZZTipoPed = rstPedido!tipoped
                                        
                                    End If
                                    
                                    ZCantidad1 = IIf(IsNull(rstPedido!CantiLote1), "0", rstPedido!CantiLote1)
                                    ZCantidad2 = IIf(IsNull(rstPedido!CantiLote2), "0", rstPedido!CantiLote2)
                                    ZCantidad3 = IIf(IsNull(rstPedido!CantiLote3), "0", rstPedido!CantiLote3)
                                    ZCantidad4 = IIf(IsNull(rstPedido!CantiLote4), "0", rstPedido!CantiLote4)
                                    ZCantidad5 = IIf(IsNull(rstPedido!CantiLote5), "0", rstPedido!CantiLote5)
                                    ZCantidadFac = IIf(IsNull(rstPedido!CantidadFac), "0", rstPedido!CantidadFac)
                                    ZSumaCantidad = ZCantidad1 + ZCantidad2 + ZCantidad3 + ZCantidad4 + ZCantidad5
                                        
                                    If ZSumaCantidad = 0 Then
                                        ZSumaCantidad = ZCantidadFac
                                    End If
                                    
                                    If ZSumaCantidad <> 0 Then
                                        ZZZKilos = ZZZKilos + ZSumaCantidad
                                            Else
                                        ZZZKilos = ZZZKilos + rstPedido!Cantidad
                                    End If
                                
                                End If
                            
                            End If
                
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPedido.Close
            End If
            
            If ZZPasa <> 0 Then
                        
                ZZRenglon = ZZRenglon + 1
                            
                Muestra.TextMatrix(ZZRenglon, 0) = Pusing("###", Str$(ZZRenglon))
                Muestra.TextMatrix(ZZRenglon, 1) = Pusing("######", Str$(ZZCorte))
                Muestra.TextMatrix(ZZRenglon, 2) = ZZRazon
                If ZZTipoPed = 5 Then
                    Muestra.TextMatrix(ZZRenglon, 3) = "Muestra"
                        Else
                    Muestra.TextMatrix(ZZRenglon, 3) = "Pedido"
                End If
                Muestra.TextMatrix(ZZRenglon, 4) = Pusing("######", Str$(ZZZKilos))
                Muestra.TextMatrix(ZZRenglon, 5) = ZZDirEntrega
                            
                ZZAyuda(ZZRenglon, 1) = Pusing("###", Str$(ZZRenglon))
                ZZAyuda(ZZRenglon, 2) = Pusing("######", Str$(ZZCorte))
                ZZAyuda(ZZRenglon, 3) = ZZRazon
                If ZZTipoPed = 5 Then
                    ZZAyuda(ZZRenglon, 4) = "Muestra"
                        Else
                    ZZAyuda(ZZRenglon, 4) = "Pedido"
                End If
                ZZAyuda(ZZRenglon, 5) = ""
                ZZAyuda(ZZRenglon, 6) = ZZDirEntrega
                
            End If
            
            EntraPedido.Visible = True

        Case Else
    End Select
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub cmdClose1_Click()

    Call Limpia_Click
    PrgHojaRuta.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()
    
    If Fecha.Text = "  /  /    " Then
        m$ = "Fecha no informada"
        a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
        Exit Sub
    End If
    ZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    
    ZFechaInicio = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    ZZDife = DateDiff("d", ZFechaInicio, Fecha.Text)
    If ZZDife > 4 Or ZZDife < 0 Then
        m$ = "La fecha de la hoja de ruta exece los 4 dias desde la fecha actual o no es anterior"
        a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
        Exit Sub
    End If
    
    ZNuevaFecha = DateAdd("D", 15, Fecha.Text)
    ZOrdNuevaFecha = Right$(ZNuevaFecha, 4) + Mid$(ZNuevaFecha, 4, 2) + Left$(ZNuevaFecha, 2)
    
    If Val(Camion.Text) = 0 Then
        m$ = "Camion no informada"
        a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
        Exit Sub
    End If
    
    If Val(NroViaje.Text) = 0 Then
        m$ = "Nro de Viaje no informado no informada"
        a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
        Exit Sub
    End If
    
    If Val(Chofer.Text) = 0 Then
        m$ = "Chofer no informada"
        a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
        Exit Sub
    End If
    
    Call Valida_fecha(Fecha.Text, Auxi)
    If Auxi <> "S" Then
        m$ = "Fecha invalida"
        a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
        Exit Sub
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Camion"
    ZSql = ZSql + " Where Camion.Codigo = " + "'" + Camion.Text + "'"
    spCamion = ZSql
    Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
    If rstCamion.RecordCount > 0 Then
    
        ZCamionFechaVtoI = IIf(IsNull(rstCamion!OrdFechaVtoI), "", rstCamion!OrdFechaVtoI)
        ZCamionFechaVtoII = IIf(IsNull(rstCamion!OrdFechaVtoII), "", rstCamion!OrdFechaVtoII)
        ZCamionFechaVtoIII = IIf(IsNull(rstCamion!OrdFechaVtoIII), "", rstCamion!OrdFechaVtoIII)
        ZCamionFechaVtoIV = IIf(IsNull(rstCamion!OrdFechaVtoIV), "", rstCamion!OrdFechaVtoIV)
        ZCamionFechaVtoV = IIf(IsNull(rstCamion!OrdFechaVtoV), "", rstCamion!OrdFechaVtoV)
        
        ZZCamionFechaVtoI = IIf(IsNull(rstCamion!FechaVtoI), "", rstCamion!FechaVtoI)
        ZZCamionFechaVtoII = IIf(IsNull(rstCamion!FechaVtoII), "", rstCamion!FechaVtoII)
        ZZCamionFechaVtoIII = IIf(IsNull(rstCamion!FechaVtoIII), "", rstCamion!FechaVtoIII)
        ZZCamionFechaVtoIV = IIf(IsNull(rstCamion!FechaVtoIV), "", rstCamion!FechaVtoIV)
        ZZCamionFechaVtoV = IIf(IsNull(rstCamion!FechaVtoV), "", rstCamion!FechaVtoV)
        
        ZCamionAplicaV = IIf(IsNull(rstCamion!AplicaV), "0", rstCamion!AplicaV)
        
        rstCamion.Close
        
        Rem
        Rem  Veririca que no esten vencidos
        Rem
        
        If Val(ZCamionFechaVtoI) > 0 Then
            If Val(ZCamionFechaVtoI) < Val(ZOrdFecha) Then
                m$ = "La fecha de vigencia de Ruta esta vencida " + ZZCamionFechaVtoI
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
                Exit Sub
            End If
        End If
        
        If Val(ZCamionFechaVtoII) > 0 Then
            If Val(ZCamionFechaVtoII) < Val(ZOrdFecha) Then
                m$ = "La fecha de vigencia de Revision Tecnica Obligatoria esta vencida " + ZZCamionFechaVtoII
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
                Exit Sub
            End If
        End If
        
        If Val(ZCamionFechaVtoIII) > 0 Then
            If Val(ZCamionFechaVtoIII) < Val(ZOrdFecha) Then
                m$ = "La fecha de vigencia de Habilitacion de Dominio esta vencida " + ZZCamionFechaVtoIII
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
                Exit Sub
            End If
        End If
        
        If Val(ZCamionFechaVtoIV) > 0 Then
            If Val(ZCamionFechaVtoIV) < Val(ZOrdFecha) Then
                m$ = "La fecha de vigencia de Seguro esta vencida " + ZZCamionFechaVtoIV
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
                Exit Sub
            End If
        End If
        
        If ZCamionAplicaV = 1 Then
            If Val(ZCamionFechaVtoV) < Val(ZOrdFecha) Then
                m$ = "La fecha de vigencia de Cert. para transporte de cargas peligrosas esta vencida " + ZZCamionFechaVtoV
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
                Exit Sub
            End If
        End If
        
        
        Rem
        Rem  Veririca que no esten proximos a vencer
        Rem
        
        If Val(ZCamionFechaVtoI) > 0 Then
            If Val(ZCamionFechaVtoI) < Val(ZOrdNuevaFecha) Then
                m$ = "La fecha de vigencia de Ruta esta proxima a " + ZZCamionFechaVtoI
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
            End If
        End If
        
        If Val(ZCamionFechaVtoII) > 0 Then
            If Val(ZCamionFechaVtoII) < Val(ZOrdNuevaFecha) Then
                m$ = "La fecha de vigencia de Revision Tecnica Obligatoria esta proxima a " + ZZCamionFechaVtoII
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
            End If
        End If
        
        If Val(ZCamionFechaVtoIII) > 0 Then
            If Val(ZCamionFechaVtoIII) < Val(ZOrdNuevaFecha) Then
                m$ = "La fecha de vigencia de Habilitacion de Dominio esta proxima a " + ZZCamionFechaVtoIII
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
            End If
        End If
        
        If Val(ZCamionFechaVtoIV) > 0 Then
            If Val(ZCamionFechaVtoIV) < Val(ZOrdNuevaFecha) Then
                m$ = "La fecha de vigencia de Seguro esta proxima a " + ZZCamionFechaVtoIV
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
            End If
        End If
        
        If ZCamionAplicaV = 1 Then
            If Val(ZCamionFechaVtoV) < Val(ZOrdNuevaFecha) Then
                m$ = "La fecha de vigencia de Cert. para transporte de cargas peligrosas esta proxima a " + ZZCamionFechaVtoV
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
            End If
        End If
                
            Else
            
        m$ = "Nro de Camion Inexistente"
        a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
        Exit Sub
        
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Chofer"
    ZSql = ZSql + " Where Chofer.Codigo = " + "'" + Chofer.Text + "'"
    spChofer = ZSql
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    If rstChofer.RecordCount > 0 Then
    
        ZChoferFechaVtoI = IIf(IsNull(rstChofer!OrdFechaVtoI), "", rstChofer!OrdFechaVtoI)
        ZChoferFechaVtoII = IIf(IsNull(rstChofer!OrdFechaVtoII), "", rstChofer!OrdFechaVtoII)
        ZChoferFechaVtoIII = IIf(IsNull(rstChofer!OrdFechaVtoIII), "", rstChofer!OrdFechaVtoIII)
        
        ZZChoferFechaVtoI = IIf(IsNull(rstChofer!FechaVtoI), "", rstChofer!FechaVtoI)
        ZZChoferFechaVtoII = IIf(IsNull(rstChofer!FechaVtoII), "", rstChofer!FechaVtoII)
        ZZChoferFechaVtoIII = IIf(IsNull(rstChofer!FechaVtoIII), "", rstChofer!FechaVtoIII)
        
        ZChoferAplicaIII = IIf(IsNull(rstChofer!AplicaIII), "0", rstChofer!AplicaIII)
        
        rstChofer.Close
        
        Rem
        Rem  Veririca que no esten vencidos
        Rem
        
        If Val(ZChoferFechaVtoI) > 0 Then
            If Val(ZChoferFechaVtoI) < Val(ZOrdFecha) Then
                m$ = "La fecha de vigencia de Licencia de Conducir esta vencida " + ZZChoferFechaVtoI
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
                Exit Sub
            End If
        End If
        
        If Val(ZChoferFechaVtoII) > 0 Then
            If Val(ZChoferFechaVtoII) < Val(ZOrdFecha) Then
                m$ = "La fecha de vigencia de ART esta vencida " + ZZChoferFechaVtoII
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
                Exit Sub
            End If
        End If
        
        If ZChoferAplicaIII = 1 Then
            If Val(ZChoferFechaVtoIII) < Val(ZOrdFecha) Then
                m$ = "La fecha de vigencia de Cert. para transporte de cargas peligrosas esta vencida " + ZZChoferFechaVtoIII
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
                Exit Sub
            End If
        End If
        
        Rem
        Rem  Veririca que no esten proximos a vencer
        Rem
        
        If Val(ZChoferFechaVtoI) > 0 Then
            If Val(ZChoferFechaVtoI) < Val(ZOrdNuevaFecha) Then
                m$ = "La fecha de vigencia de Licencia de Conducir esta proxima a " + ZZChoferFechaVtoI
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
            End If
        End If
        
        If Val(ZChoferFechaVtoII) > 0 Then
            If Val(ZChoferFechaVtoII) < Val(ZOrdNuevaFecha) Then
                m$ = "La fecha de vigencia de ART esta proxima a " + ZZChoferFechaVtoII
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
            End If
        End If
        
        If ZChoferAplicaIII = 1 Then
            If Val(ZChoferFechaVtoIII) < Val(ZOrdNuevaFecha) Then
                m$ = "La fecha de vigencia de Cert. para transporte de cargas peligrosas esta proxima a " + ZZChoferFechaVtoIII
                a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
            End If
        End If
        
            Else
            
        m$ = "Nro de Chofer Inexistente"
        a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
        Exit Sub
        
    End If
    
    
    
    
    If Val(Hoja.Text) <> 0 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM HojaRuta"
        ZSql = ZSql + " Where HojaRuta.Hoja = " + "'" + Hoja.Text + "'"
        ZSql = ZSql + " Order by HojaRuta.Clave"
        rsHojaRuta = ZSql
        Set rstHojaRuta = db.OpenRecordset(rsHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
        If rstHojaRuta.RecordCount > 0 Then
            rstHojaRuta.Close
                Else
            Hoja.Text = "0"
        End If
    End If
    
    If Val(Hoja.Text) = 0 Then
        ZSql = ""
        ZSql = ZSql + "Select Max(Hoja) as [HojaMayor]"
        ZSql = ZSql + " FROM HojaRuta"
        spHojaRuta = ZSql
        Set rstHojaRuta = db.OpenRecordset(spHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
        If rstHojaRuta.RecordCount > 0 Then
            rstHojaRuta.MoveLast
            WHojaMayor = IIf(IsNull(rstHojaRuta!HojaMayor), "0", rstHojaRuta!HojaMayor)
            Hoja.Text = Str$(WHojaMayor + 1)
            rstHojaRuta.Close
                Else
            Hoja.Text = "1"
        End If
    End If
    
    
    
    
    
    ZExterior = "N"
    Rem by nan
    Rem revisar el ciclo no se realiza para cada cliente
    Rem el checklistexpo hasta que no termina el ciclo
    
    
    
    For Ciclo = 1 To 100
        ZCliente = WVector1.TextMatrix(Ciclo, 2)
        If Trim(ZCliente) <> "" Then
            spCliente = "ConsultaCliente " + ZCliente
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                WProvincia = rstCliente!Provincia
                If WProvincia = 24 Then
                    ZExterior = "S"
                End If
                rstCliente.Close
            End If
        End If
    Next Ciclo
    
    If ZExterior = "S" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CheckListExpo"
        ZSql = ZSql + " Where CheckListExpo.Hoja = " + "'" + Hoja.Text + "'"
        rsCheckListExpo = ZSql
        Set rstCheckListExpo = db.OpenRecordset(rsCheckListExpo, dbOpenSnapshot, dbSQLPassThrough)
        If rstCheckListExpo.RecordCount > 0 Then
            rstCheckListExpo.Close
                Else
            m$ = "Se debe ingresar el Check List"
            a% = MsgBox(m$, 0, "Ingreso de Hoja de Ruta")
            Exit Sub
        End If
    End If

    WRenglon = 0

    Sql1 = "Select *"
    Sql2 = " FROM HojaRuta"
    Sql3 = " Where HojaRuta.Hoja = " + "'" + Hoja.Text + "'"
    Sql4 = " Order by HojaRuta.Clave"
    
    rsHojaRuta = Sql1 + Sql2 + Sql3 + Sql4
    Set rstHojaRuta = db.OpenRecordset(rsHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
    If rstHojaRuta.RecordCount > 0 Then
        With rstHojaRuta
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    ZBajaHoja(WRenglon) = Str$(rstHojaRuta!Pedido)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHojaRuta.Close
    End If
    
    For Ciclo = 1 To WRenglon
    
        ZPedido = ZBajaHoja(Ciclo)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Pedido SET "
        ZSql = ZSql + " HojaRuta = 0"
        ZSql = ZSql + " Where Pedido = " + "'" + ZPedido + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo



    ZSql = ""
    ZSql = ZSql + "DELETE HojaRuta"
    ZSql = ZSql + " Where Hoja = " + "'" + Hoja.Text + "'"
    rsHojaRuta = ZSql
    Set rstHojaRuta = db.OpenRecordset(rsHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
    
    ZLugarAtraso = 0
    Erase ZAtraso
    
    ZCantiPeligroso = 0
    ZCantiPeligrosoII = 0
    ZCodigoPeligroso = ""
    ZCodigoPeligrosoII = ""
    
    For Ciclo = 1 To 100
        
        ZHoja = Hoja.Text
        ZRenglon = Str$(Ciclo)
        ZFecha = Fecha.Text
        ZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        ZChofer = Chofer.Text
        ZCamion = Camion.Text
        ZPedido = WVector1.TextMatrix(Ciclo, 1)
        ZCliente = WVector1.TextMatrix(Ciclo, 2)
        ZRazon = WVector1.TextMatrix(Ciclo, 3)
        ZRemito = WVector1.TextMatrix(Ciclo, 4)
        ZSeguridad = WVector1.TextMatrix(Ciclo, 5)
        ZKilos = WVector1.TextMatrix(Ciclo, 6)
        ZPesos = "0"
        ZBultos = WVector1.TextMatrix(Ciclo, 7)
        ZObservaI = WVector1.TextMatrix(Ciclo, 8)
        ZObservaII = WVector1.TextMatrix(Ciclo, 9)
        ZNroViaje = NroViaje.Text
        ZRetiraProv = RetiraProv.Text
        ZTipoEstado = Str$(TipoEstado.ListIndex)
        
        Auxi1 = Hoja.Text
        Call Ceros(Auxi1, 6)
        Auxi2 = Str$(Ciclo)
        Call Ceros(Auxi2, 2)
        
        ZClave = Auxi1 + Auxi2
    
        If Val(ZPedido) <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO HojaRuta ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Hoja ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Chofer ,"
            ZSql = ZSql + "Camion ,"
            ZSql = ZSql + "Pedido ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Remito ,"
            ZSql = ZSql + "Seguridad ,"
            ZSql = ZSql + "Kilos ,"
            ZSql = ZSql + "Pesos ,"
            ZSql = ZSql + "Bultos ,"
            ZSql = ZSql + "Razon ,"
            ZSql = ZSql + "ObservaI ,"
            ZSql = ZSql + "ObservaII ,"
            ZSql = ZSql + "NroViaje ,"
            ZSql = ZSql + "TipoEstado ,"
            ZSql = ZSql + "RetiraProv )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZClave + "',"
            ZSql = ZSql + "'" + ZHoja + "',"
            ZSql = ZSql + "'" + ZRenglon + "',"
            ZSql = ZSql + "'" + ZFecha + "',"
            ZSql = ZSql + "'" + ZOrdFecha + "',"
            ZSql = ZSql + "'" + ZChofer + "',"
            ZSql = ZSql + "'" + ZCamion + "',"
            ZSql = ZSql + "'" + ZPedido + "',"
            ZSql = ZSql + "'" + ZCliente + "',"
            ZSql = ZSql + "'" + ZRemito + "',"
            ZSql = ZSql + "'" + ZSeguridad + "',"
            ZSql = ZSql + "'" + ZKilos + "',"
            ZSql = ZSql + "'" + ZPesos + "',"
            ZSql = ZSql + "'" + ZBultos + "',"
            ZSql = ZSql + "'" + ZRazon + "',"
            ZSql = ZSql + "'" + ZObservaI + "',"
            ZSql = ZSql + "'" + ZObservaII + "',"
            ZSql = ZSql + "'" + ZNroViaje + "',"
            ZSql = ZSql + "'" + ZTipoEstado + "',"
            ZSql = ZSql + "'" + ZRetiraProv + "')"
                
            rsHojaRuta = ZSql
            Set rstHojaRuta = db.OpenRecordset(rsHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Pedido SET "
            ZSql = ZSql + " HojaRuta = " + "'" + ZHoja + "'"
            ZSql = ZSql + " Where Pedido = " + "'" + ZPedido + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Muestra SET "
            ZSql = ZSql + " HojaRuta = " + "'" + ZHoja + "'"
            ZSql = ZSql + " WHERE Pedido = " + "'" + ZPedido + "'"
            spMuestra = ZSql
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            
            Erase ZArti
            ZLugarArti = 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pedido"
            ZSql = ZSql + " Where Pedido.Pedido = " + "'" + ZPedido + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                With rstPedido
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            ZBultos1 = IIf(IsNull(rstPedido!Bultos1), "0", rstPedido!Bultos1)
                            ZBultos2 = IIf(IsNull(rstPedido!Bultos2), "0", rstPedido!Bultos2)
                            ZBultos3 = IIf(IsNull(rstPedido!Bultos3), "0", rstPedido!Bultos3)
                            ZBultos4 = IIf(IsNull(rstPedido!Bultos4), "0", rstPedido!Bultos4)
                            ZBultos5 = IIf(IsNull(rstPedido!Bultos5), "0", rstPedido!Bultos5)
                            
                            ZBultos = ZBultos + ZBultos1 + ZBultos2 + ZBultos3 + ZBultos4 + ZBultos5
                            
                            ZCantidad1 = IIf(IsNull(rstPedido!CantiLote1), "0", rstPedido!CantiLote1)
                            ZCantidad2 = IIf(IsNull(rstPedido!CantiLote2), "0", rstPedido!CantiLote2)
                            ZCantidad3 = IIf(IsNull(rstPedido!CantiLote3), "0", rstPedido!CantiLote3)
                            ZCantidad4 = IIf(IsNull(rstPedido!CantiLote4), "0", rstPedido!CantiLote4)
                            ZCantidad5 = IIf(IsNull(rstPedido!CantiLote5), "0", rstPedido!CantiLote5)
                            ZCantidadFac = IIf(IsNull(rstPedido!CantidadFac), "0", rstPedido!CantidadFac)
                            ZSumaCantidad = ZCantidad1 + ZCantidad2 + ZCantidad3 + ZCantidad4 + ZCantidad5
                            
                            If ZSumaCantidad = 0 Then
                                ZSumaCantidad = ZCantidadFac
                            End If
                            
                            If ZSumaCantidad <> 0 Then
                                ZZZKilos = ZZZKilos + ZSumaCantidad
                                    Else
                                ZZZKilos = ZZZKilos + (rstPedido!Cantidad - rstPedido!Facturado)
                            End If
                            
                            ZLugarArti = ZLugarArti + 1
                            ZArti(ZLugarArti) = rstPedido!terminado
            
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPedido.Close
            End If
            
            For CicloArti = 1 To ZLugarArti
                ZTerminado = ZArti(CicloArti)
                
                spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZMarca = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
                    ZNaciones = IIf(IsNull(rstTerminado!naciones), "", rstTerminado!naciones)
                    rstTerminado.Close
                End If
                
                ZMarca = Trim(ZMarca)
                ZNaciones = Trim(ZNaciones)
                If ZMarca <> "" Then
                    If ZMarca <> ZCodigoPeligroso Then
                        ZCodigoPeligroso = ZMarca
                        ZCantiPeligroso = ZCantiPeligroso + 1
                    End If
                    If ZNaciones <> ZCodigoPeligrosoII Then
                        ZCodigoPeligrosoII = ZNaciones
                        ZCantiPeligrosoII = ZCantiPeligrosoII + 1
                    End If
                End If
                
            Next CicloArti
    
            ZZFecEntrega = "00/00/0000"
            ZZOrdFecEntrega = "00000000"
            ZFecha = Fecha.Text
            ZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pedido"
            ZSql = ZSql + " Where Pedido.Pedido = " + "'" + ZPedido + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                ZZFecEntrega = rstPedido!FecEntrega
                ZZOrdFecEntrega = rstPedido!OrdFecEntrega
                rstPedido.Close
            End If
            
            If ZOrdFecha > ZZOrdFecEntrega Then
            
                WAtraso = "S"
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Atraso"
                ZSql = ZSql + " Where Atraso.Pedido = " + "'" + ZPedido + "'"
                ZSql = ZSql + " Order by Atraso.Numero"
                spAtraso = ZSql
                Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
                If rstAtraso.RecordCount > 0 Then
                    With rstAtraso
                        .MoveFirst
                        If .NoMatch = False Then
                            Do
                                ZOrigen = IIf(IsNull(!Origen), "0", !Origen)
                                If ZOrigen = 3 Then
                                    WAtraso = "N"
                                End If
                                .MoveNext
                                If .EOF = True Then
                                    Exit Do
                                End If
                            Loop
                        End If
                    End With
                    rstAtraso.Close
                End If
            
                If WAtraso = "S" Then
                    ZLugarAtraso = ZLugarAtraso + 1
                    ZAtraso(ZLugarAtraso, 1) = ZPedido
                    ZAtraso(ZLugarAtraso, 2) = ZCliente
                    ZAtraso(ZLugarAtraso, 3) = ZRazon
                    ZAtraso(ZLugarAtraso, 4) = ZZFecEntrega
                    Rem ZSql = ""
                    Rem ZSql = ZSql + "Select *"
                    Rem ZSql = ZSql + " FROM Atraso"
                    Rem ZSql = ZSql + " Where Atraso.Pedido = " + "'" + ZPedido + "'"
                    Rem ZSql = ZSql + " and Atraso.Concepto > 6"
                    Rem spAtraso = ZSql
                    Rem Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
                    Rem If rstAtraso.RecordCount > 0 Then
                    Rem     rstAtraso.Close
                    Rem         Else
                    Rem     ZLugarAtraso = ZLugarAtraso + 1
                    Rem     ZAtraso(ZLugarAtraso, 1) = ZPedido
                    Rem     ZAtraso(ZLugarAtraso, 2) = ZCliente
                    Rem     ZAtraso(ZLugarAtraso, 3) = ZRazon
                    Rem     ZAtraso(ZLugarAtraso, 4) = ZZFecEntrega
                    Rem End If
                End If
                
            End If
            
        End If
        
    Next Ciclo
    
    Rem ZCantiPeligrosoII = 2
    
    ZSql = ""
    ZSql = ZSql + "UPDATE HojaRuta SET "
    ZSql = ZSql + " CodigoPeligroso = " + "'" + ZCodigoPeligroso + "',"
    ZSql = ZSql + " CodigoPeligrosoII = " + "'" + ZCodigoPeligrosoII + "',"
    ZSql = ZSql + " CantiPeligroso = " + "'" + Str$(ZCantiPeligroso) + "',"
    ZSql = ZSql + " CantiPeligrosoII = " + "'" + Str$(ZCantiPeligrosoII) + "'"
    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
    spHojaRuta = ZSql
    Set rstHojaRuta = db.OpenRecordset(spHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem DADA
    Rem DADA
    Rem DADA
    
    T$ = "Hoja de Ruta"
    m$ = "Desea Imprimir la hoja de ruta"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Call Impresion
    End If
    
    Call Limpia_Click
    
    If ZLugarAtraso > 0 Then
        ZLLave = "N"
        ZHastaAtraso = ZLugarAtraso
        ZLugarAtraso = 1
        PrgHojaRuta.Hide
        Unload Me
        PrgAtrasoExpedicion.Show
    End If
        
End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector

    Hoja.Text = ""
    Fecha.Text = "  /  /    "
    Camion.Text = ""
    Chofer.Text = ""
    DesCamion.Caption = ""
    DesChofer.Caption = ""
    TotalKilos.Text = ""
    NroViaje.Text = ""
    RetiraProv.Text = ""
    ZLLave = ""
    TipoEstado.ListIndex = 0
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Hoja) as [HojaMayor]"
    Rem ZSql = ZSql + " FROM HojaRuta"
    Rem spHojaRuta = ZSql
    Rem Set rstHojaRuta = db.OpenRecordset(spHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstHojaRuta.RecordCount > 0 Then
    Rem     rstHojaRuta.MoveLast
    Rem     WHojaMayor = IIf(IsNull(rstHojaRuta!HojaMayor), "0", rstHojaRuta!HojaMayor)
    Rem     Hoja.Text = Str$(WHojaMayor + 1)
    Rem     rstHojaRuta.Close
    Rem         Else
    Rem     Hoja.Text = "0"
    Rem End If
    
    Renglon = 0
    
    WVector1.Col = 1
    WVector1.Row = 1
    
    Hoja.SetFocus
    
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Chofer.Text = WIndice.List(Indice)
            Call Chofer_KeyPress(13)
            
        Case 1
            Indice = Pantalla.ListIndex
            Chofer.Text = WIndice.List(Indice)
            Call Camion_KeyPress(13)
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    WVector1.Col = 1
    WVector1.Row = 1

    Hoja.Text = ""
    Fecha.Text = "  /  /    "
    Camion.Text = ""
    Chofer.Text = ""
    DesCamion.Caption = ""
    DesChofer.Caption = ""
    TotalKilos.Text = ""
    NroViaje.Text = ""
    RetiraProv.Text = ""
    
    TipoPedido.Clear
    
    TipoPedido.AddItem "Total"
    TipoPedido.AddItem "Quimicos"
    TipoPedido.AddItem "Colorantes"
    TipoPedido.AddItem "Farma"
    TipoPedido.AddItem "Pigmentos"
    
    TipoEstado.Clear
    
    TipoEstado.AddItem "Pendiente"
    TipoEstado.AddItem "Confirmada"
    
    
    ZProceso = "N"
    TipoPedido.ListIndex = 0
    TipoEstado.ListIndex = 0
    ZProceso = "S"
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Hoja) as [HojaMayor]"
    Rem ZSql = ZSql + " FROM HojaRuta"
    Rem spHojaRuta = ZSql
    Rem Set rstHojaRuta = db.OpenRecordset(spHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstHojaRuta.RecordCount > 0 Then
    Rem     rstHojaRuta.MoveLast
    Rem     WHojaMayor = IIf(IsNull(rstHojaRuta!HojaMayor), "0", rstHojaRuta!HojaMayor)
    Rem     Hoja.Text = Str$(WHojaMayor + 1)
    Rem     rstHojaRuta.Close
    Rem         Else
    Rem     Hoja.Text = "0"
    Rem End If
    
    Renglon = 0
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    WRenglon = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM HojaRuta"
    Sql3 = " Where HojaRuta.Hoja = " + "'" + Hoja.Text + "'"
    Sql4 = " Order by HojaRuta.Clave"
    
    rsHojaRuta = Sql1 + Sql2 + Sql3 + Sql4
    Set rstHojaRuta = db.OpenRecordset(rsHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
    If rstHojaRuta.RecordCount > 0 Then
        With rstHojaRuta
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Fecha.Text = rstHojaRuta!Fecha
                    Camion.Text = rstHojaRuta!Camion
                    Chofer.Text = rstHojaRuta!Chofer
                    NroViaje.Text = IIf(IsNull(rstHojaRuta!NroViaje), "", rstHojaRuta!NroViaje)
                    RetiraProv.Text = IIf(IsNull(rstHojaRuta!RetiraProv), "", rstHojaRuta!RetiraProv)
                    TipoEstado.ListIndex = IIf(IsNull(rstHojaRuta!TipoEstado), "0", rstHojaRuta!TipoEstado)
                
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = Str$(rstHojaRuta!Pedido)
            
                    WVector1.Col = 2
                    WVector1.Text = Trim(rstHojaRuta!Cliente)
                    
                    WVector1.Col = 3
                    WVector1.Text = Trim(rstHojaRuta!razon)
            
                    WVector1.Col = 4
                    WVector1.Text = Str$(rstHojaRuta!remito)
            
                    WVector1.Col = 5
                    WVector1.Text = Trim(rstHojaRuta!seguridad)
                    
                    WVector1.Col = 6
                    WVector1.Text = ""
            
                    WVector1.Col = 7
                    WVector1.Text = Str$(rstHojaRuta!Bultos)
                    
                    WVector1.Col = 8
                    WVector1.Text = Trim(rstHojaRuta!ObservaI)
            
                    WVector1.Col = 9
                    WVector1.Text = Trim(rstHojaRuta!ObservaII)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHojaRuta.Close
    End If
    
    
    
    For CicloII = 1 To WRenglon
    
        ZZZKilos = 0
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Pedido = " + "'" + WVector1.TextMatrix(CicloII, 1) + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            With rstPedido
                .MoveFirst
                Do
                    If .EOF = False Then
                                
                        ZCantidad1 = IIf(IsNull(rstPedido!CantiLote1), "0", rstPedido!CantiLote1)
                        ZCantidad2 = IIf(IsNull(rstPedido!CantiLote2), "0", rstPedido!CantiLote2)
                        ZCantidad3 = IIf(IsNull(rstPedido!CantiLote3), "0", rstPedido!CantiLote3)
                        ZCantidad4 = IIf(IsNull(rstPedido!CantiLote4), "0", rstPedido!CantiLote4)
                        ZCantidad5 = IIf(IsNull(rstPedido!CantiLote5), "0", rstPedido!CantiLote5)
                        ZCantidadFac = IIf(IsNull(rstPedido!CantidadFac), "0", rstPedido!CantidadFac)
                        ZSumaCantidad = ZCantidad1 + ZCantidad2 + ZCantidad3 + ZCantidad4 + ZCantidad5
                                    
                        If ZSumaCantidad = 0 Then
                            ZSumaCantidad = ZCantidadFac
                        End If
                                    
                        If ZSumaCantidad <> 0 Then
                            ZZZKilos = ZZZKilos + ZSumaCantidad
                                Else
                            ZZZKilos = ZZZKilos + rstPedido!Cantidad
                        End If
                                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedido.Close
        End If
        
        WVector1.TextMatrix(CicloII, 6) = Pusing("######", Str$(ZZZKilos))
        
    Next CicloII
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM Camion"
    Sql3 = " Where Camion.Codigo = " + "'" + Camion.Text + "'"
    spCamion = Sql1 + Sql2 + Sql3
    Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
    If rstCamion.RecordCount > 0 Then
        DesCamion.Caption = Trim(rstCamion!Descripcion)
        rstCamion.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM Chofer"
    Sql3 = " Where Chofer.Codigo = " + "'" + Chofer.Text + "'"
    spChofer = Sql1 + Sql2 + Sql3
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    If rstChofer.RecordCount > 0 Then
        DesChofer.Caption = Trim(rstChofer!Descripcion)
        rstChofer.Close
    End If
    
    Call Calcula_Click

End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Chofer.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Chofer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM Chofer"
        Sql3 = " Where Chofer.Codigo = " + "'" + Chofer.Text + "'"
        spChofer = Sql1 + Sql2 + Sql3
        Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
        If rstChofer.RecordCount > 0 Then
            DesChofer.Caption = Trim(rstChofer!Descripcion)
            rstChofer.Close
            Camion.SetFocus
                Else
            Chofer.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Chofer.Text = "  -     -   "
        DesChofer.Caption = ""
    End If
End Sub

Private Sub Camion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM Camion"
        Sql3 = " Where Camion.Codigo = " + "'" + Camion.Text + "'"
        spCamion = Sql1 + Sql2 + Sql3
        Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
        If rstCamion.RecordCount > 0 Then
            DesCamion.Caption = Trim(rstCamion!Descripcion)
            rstCamion.Close
            NroViaje.SetFocus
                Else
            Camion.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Camion.Text = "  -     -   "
        DesCamion.Caption = ""
    End If
End Sub

Private Sub NroViaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetiraProv.SetFocus
    End If
    If KeyAscii = 27 Then
        NroViaje.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub RetiraProv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WVector1.TopRow = 1
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        RetiraProv.Text = ""
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

Private Sub TipoPedido_Click()
    If ZProceso <> "N" Then
        
        Opcion.Clear
    
        Opcion.AddItem ""
        Opcion.AddItem ""
        Opcion.AddItem "Pedidos"

        Rem Opcion.Visible = True
    
        Opcion.ListIndex = 2
    
        Rem Call Opcion_Click
        
    End If
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
        Case 9
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    Call Calcula_Click
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            If WVector1.Text <> "" Then
                ZZMuestra = "N"
                ZPedido = WVector1.Text
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Pedido"
                ZSql = ZSql + " Where Pedido.Pedido = " + "'" + WVector1.Text + "'"
                ZSql = ZSql + " Order by Pedido.Remito desc"
                spPedido = ZSql
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                If rstPedido.RecordCount > 0 Then
                
                    WVector1.Col = 2
                    WVector1.Text = rstPedido!Cliente
                    ZCliente = rstPedido!Cliente
                    
                    WVector1.Col = 4
                    WVector1.Text = IIf(IsNull(rstPedido!remito), "0", rstPedido!remito)
                    ZZRemito = IIf(IsNull(rstPedido!remito), "0", rstPedido!remito)
                    
                    rstPedido.Close
                    
                    Rem If Val(ZZRemito) = 0 Then
                    Sql1 = "Select *"
                    Sql2 = " FROM Muestra"
                    Sql3 = " WHERE Pedido = " + "'" + ZPedido + "'"
                    spMuestra = Sql1 + Sql2 + Sql3
                    Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMuestra.RecordCount > 0 Then
                        ZZMuestra = "S"
                        WRemito = IIf(IsNull(rstMuestra!remito), "0", rstMuestra!remito)
                        WVector1.Col = 4
                        WVector1.Text = WRemito
                        rstMuestra.Close
                    End If
                    Rem End If
                    
                    spCliente = "ConsultaCliente " + "'" + ZCliente + "'"
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        WVector1.Col = 3
                        WVector1.Text = rstCliente!razon
                        rstCliente.Close
                    End If
                    
                    ZBultos = 0
                    ZZZKilos = 0
                    Erase ZArti
                    LugarArti = 0
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Pedido"
                    ZSql = ZSql + " Where Pedido.Pedido = " + "'" + ZPedido + "'"
                    spPedido = ZSql
                    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPedido.RecordCount > 0 Then
                        With rstPedido
                            .MoveFirst
                            Do
                                If .EOF = False Then
                                
                                    ZBultos1 = IIf(IsNull(rstPedido!Bultos1), "0", rstPedido!Bultos1)
                                    ZBultos2 = IIf(IsNull(rstPedido!Bultos2), "0", rstPedido!Bultos2)
                                    ZBultos3 = IIf(IsNull(rstPedido!Bultos3), "0", rstPedido!Bultos3)
                                    ZBultos4 = IIf(IsNull(rstPedido!Bultos4), "0", rstPedido!Bultos4)
                                    ZBultos5 = IIf(IsNull(rstPedido!Bultos5), "0", rstPedido!Bultos5)
                                    
                                    ZBultos = ZBultos + ZBultos1 + ZBultos2 + ZBultos3 + ZBultos4 + ZBultos5
                                    
                                    ZCantidad1 = IIf(IsNull(rstPedido!CantiLote1), "0", rstPedido!CantiLote1)
                                    ZCantidad2 = IIf(IsNull(rstPedido!CantiLote2), "0", rstPedido!CantiLote2)
                                    ZCantidad3 = IIf(IsNull(rstPedido!CantiLote3), "0", rstPedido!CantiLote3)
                                    ZCantidad4 = IIf(IsNull(rstPedido!CantiLote4), "0", rstPedido!CantiLote4)
                                    ZCantidad5 = IIf(IsNull(rstPedido!CantiLote5), "0", rstPedido!CantiLote5)
                                    ZCantidadFac = IIf(IsNull(rstPedido!CantidadFac), "0", rstPedido!CantidadFac)
                                    ZSumaCantidad = ZCantidad1 + ZCantidad2 + ZCantidad3 + ZCantidad4 + ZCantidad5
                                    
                                    If ZSumaCantidad = 0 Then
                                        ZSumaCantidad = ZCantidadFac
                                    End If
                                    
                                    If ZSumaCantidad <> 0 Then
                                        ZZZKilos = ZZZKilos + ZSumaCantidad
                                            Else
                                        ZZZKilos = ZZZKilos + (rstPedido!Cantidad - rstPedido!Facturado)
                                    End If
                                    
                                    If rstPedido!tipoped = 5 Then
                                        ZZZKilos = ZZZKilos + rstPedido!Cantidad
                                        ZBultos = 1
                                    End If
                                    
                                    ZLugarArti = ZLugarArti + 1
                                    ZArti(ZLugarArti) = rstPedido!terminado
                    
                                    .MoveNext
                                        Else
                                    Exit Do
                                End If
                            Loop
                        End With
                        rstPedido.Close
                    End If
                    
                    ZMarca = ""
                    WVector1.TextMatrix(WVector1.Row, 5) = ""
                    
                    For CicloArti = 1 To ZLugarArti
                        ZTerminado = ZArti(CicloArti)
                        
                        spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            ZMarca = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
                            rstTerminado.Close
                        End If
                        ZMarca = Trim(ZMarca)
                        If ZMarca <> "" Then
                            If Trim(WVector1.TextMatrix(WVector1.Row, 5)) <> "" Then
                                WVector1.TextMatrix(WVector1.Row, 5) = WVector1.TextMatrix(WVector1.Row, 5) + ";" + ZMarca
                                    Else
                                WVector1.TextMatrix(WVector1.Row, 5) = ZMarca
                            End If
                        End If
                    Next CicloArti
                    
                    If ZZMuestra = "S" Then
                        WVector1.Col = 9
                        WVector1.Text = "Muestras"
                    End If
                    
                    WVector1.Col = 6
                    WVector1.Text = Str$(ZZZKilos)
                    
                    WVector1.Col = 7
                    WVector1.Text = Str$(ZBultos)
                    
                        Else
                        
                    WControl = "N"
                    rstPedido.Close
                    
                End If
                
                    Else
                    
                WControl = "N"
                
            End If
           
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector1.Rows - 1
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        If WAuxi1 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For da = 1 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    Call Calcula_Click
    
    End If
    
End Sub


Private Sub AgregaRenglon_Click()

    Hasta = WVector1.Row

    For iRow = 100 To Hasta Step -1
        WVector1.TextMatrix(iRow, 1) = WVector1.TextMatrix(iRow - 1, 1)
        WVector1.TextMatrix(iRow, 2) = WVector1.TextMatrix(iRow - 1, 2)
        WVector1.TextMatrix(iRow, 3) = WVector1.TextMatrix(iRow - 1, 3)
        WVector1.TextMatrix(iRow, 4) = WVector1.TextMatrix(iRow - 1, 4)
        WVector1.TextMatrix(iRow, 5) = WVector1.TextMatrix(iRow - 1, 5)
        WVector1.TextMatrix(iRow, 6) = WVector1.TextMatrix(iRow - 1, 6)
        WVector1.TextMatrix(iRow, 7) = WVector1.TextMatrix(iRow - 1, 7)
        WVector1.TextMatrix(iRow, 8) = WVector1.TextMatrix(iRow - 1, 8)
    Next iRow

    WVector1.TextMatrix(Hasta, 1) = ""
    WVector1.TextMatrix(Hasta, 2) = ""
    WVector1.TextMatrix(Hasta, 3) = ""
    WVector1.TextMatrix(Hasta, 4) = ""
    WVector1.TextMatrix(Hasta, 5) = ""
    WVector1.TextMatrix(Hasta, 6) = ""
    WVector1.TextMatrix(Hasta, 7) = ""
    WVector1.TextMatrix(Hasta, 8) = ""
    WVector1.TextMatrix(Hasta, 9) = ""
    
    WTexto1.Text = ""
    WTexto2.Text = ""

End Sub


Private Sub WTexto2_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
     Opcion.AddItem ""
     Opcion.AddItem ""
     Opcion.AddItem "Pedidos"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 2
    
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
    WVector1.Cols = 10
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
    
    WVector1.ColWidth(0) = 300
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Pedido"
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Cliente"
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Razon"
                WVector1.ColWidth(Ciclo) = 2300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Remito"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Segur."
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Kilos"
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "Bultos"
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = "Envases a Retirar"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 3500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
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
    
    WAncho = 400
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


Private Sub Calcula_Click()

    ZTotalKilos = 0

    For a = 1 To 100
        ZTotalKilos = ZTotalKilos + Val(WVector1.TextMatrix(a, 6))
    Next a
    
    TotalKilos.Text = Str$(ZTotalKilos)
    
End Sub

Private Sub Impresion()

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    If Val(WEmpresa) = 1 Then
    
        Listado.SQLQuery = "SELECT HojaRuta.Hoja, HojaRuta.Renglon, HojaRuta.Fecha, HojaRuta.Chofer, HojaRuta.Camion, HojaRuta.Pedido, HojaRuta.Cliente, HojaRuta.Remito, HojaRuta.Seguridad, HojaRuta.Kilos, HojaRuta.ObservaI, HojaRuta.ObservaII, HojaRuta.Bultos, HojaRuta.Razon, HojaRuta.NroViaje, HojaRuta.RetiraProv, " _
            + "Chofer.Descripcion, " _
            + "Camion.Descripcion, Camion.Patente " _
            + "From " _
            + DSQ + ".dbo.HojaRuta HojaRuta, " _
            + DSQ + ".dbo.Chofer Chofer, " _
            + DSQ + ".dbo.Camion Camion " _
            + "Where " _
            + "HojaRuta.Chofer = Chofer.Codigo AND " _
            + "HojaRuta.Camion = Camion.Codigo AND " _
            + "HojaRuta.Hoja >= " + Hoja.Text + " AND " _
            + "HojaRuta.Hoja <= " + Hoja.Text
            
                Else
                
        Listado.SQLQuery = "SELECT HojaRuta.Hoja, HojaRuta.Renglon, HojaRuta.Fecha, HojaRuta.Chofer, HojaRuta.Camion, HojaRuta.Pedido, HojaRuta.Cliente, HojaRuta.Remito, HojaRuta.Seguridad, HojaRuta.Kilos, HojaRuta.ObservaI, HojaRuta.ObservaII, HojaRuta.Bultos, HojaRuta.Razon, " _
            + "Chofer.Descripcion, " _
            + "Camion.Descripcion, Camion.Patente " _
            + "From " _
            + DSQ + ".dbo.HojaRuta HojaRuta, " _
            + DSQ + ".dbo.Chofer Chofer, " _
            + DSQ + ".dbo.Camion Camion " _
            + "Where " _
            + "HojaRuta.Chofer = Chofer.Codigo AND " _
            + "HojaRuta.Camion = Camion.Codigo AND " _
            + "HojaRuta.Hoja >= " + Hoja.Text + " AND " _
            + "HojaRuta.Hoja <= " + Hoja.Text
            
    End If
            
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{HojaRuta.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.SelectionFormula = "{HojaRuta.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    Listado.CopiesToPrinter = 2
    
    If Val(WEmpresa) = 1 Then
        Listado.ReportFileName = "HojaRuta.rpt"
            Else
        Listado.ReportFileName = "HojaRutapelli.rpt"
    End If
    
    Listado.Action = 1
    Listado.CopiesToPrinter = 1
    
End Sub


