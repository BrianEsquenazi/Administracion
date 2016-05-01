VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCargaControl 
   AutoRedraw      =   -1  'True
   Caption         =   "Carga de Etapas de Fabricacion"
   ClientHeight    =   6660
   ClientLeft      =   1710
   ClientTop       =   570
   ClientWidth     =   8475
   LinkTopic       =   "Form2"
   ScaleHeight     =   6660
   ScaleWidth      =   8475
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
      Left            =   1320
      TabIndex        =   24
      Top             =   3840
      Visible         =   0   'False
      Width           =   3975
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
      TabIndex        =   23
      Top             =   3480
      Visible         =   0   'False
      Width           =   8055
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
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   4800
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Hora 
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
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   20
      Text            =   " "
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Operario 
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
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   19
      Text            =   " "
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Hasta 
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
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   16
      Text            =   " "
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Desde 
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
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   14
      Text            =   " "
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Etapa 
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
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   12
      Text            =   " "
      Top             =   1680
      Width           =   1095
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
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   6
      Text            =   " "
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Teorico 
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
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   4
      Text            =   " "
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Equipo 
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
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   240
      Width           =   855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7440
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WCampaña.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Bancos"
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
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   960
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
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
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
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   2655
      Left            =   120
      TabIndex        =   25
      Top             =   3840
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4683
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   1800
      MouseIcon       =   "CargaControl.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "CargaControl.frx":030A
      ToolTipText     =   "Consulta de Datos"
      Top             =   2880
      Width           =   480
   End
   Begin VB.Label Label8 
      Caption         =   "Operario"
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
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label DesOperario 
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
      Left            =   3720
      TabIndex        =   17
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Rango Temperatura"
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
      TabIndex        =   15
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Etapa"
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
      TabIndex        =   13
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Hoja de Produccion"
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
      Top             =   600
      Width           =   1695
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
      Left            =   3600
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
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
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1575
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
      Left            =   3720
      TabIndex        =   8
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Rendimiento teorico"
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
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   1080
      MouseIcon       =   "CargaControl.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "CargaControl.frx":0E56
      ToolTipText     =   "Limpia la pantalla"
      Top             =   2880
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   240
      MouseIcon       =   "CargaControl.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "CargaControl.frx":19A2
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   2880
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   2520
      MouseIcon       =   "CargaControl.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "CargaControl.frx":24EE
      ToolTipText     =   "Salida"
      Top             =   2880
      Width           =   480
   End
   Begin VB.Label lblLabels 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   1935
   End
End
Attribute VB_Name = "PrgCargaControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOperario As Recordset
Dim spOperario As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstHoja As Recordset
Dim spHoja As String

Dim ZTimer As String

Private Sub cmdAdd_Click()


    ZSql = ""
    ZSql = ZSql + "UPDATE Hoja SET "
    ZSql = ZSql + " EstadoHoja = 0"
    ZSql = ZSql + " Where EquipoII = " + "'" + Equipo.Text + "'"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZEtapa = 0
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Where Hoja.Hoja = " + "'" + Hoja.Text + "'"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        ZEtapa = rstHoja!Etapa
        ZFecha = IIf(IsNull(rstHoja!fechainicioetapa), "", rstHoja!fechainicioetapa)
        ZHora = IIf(IsNull(rstHoja!HoraInicioEtapa), "", rstHoja!HoraInicioEtapa)
        ZTimer = IIf(IsNull(rstHoja!timerinicioetapa), "", rstHoja!timerinicioetapa)
        rstHoja.Close
    End If
    
    If Val(Etapa.Text) > ZEtapa Then
        ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        ZHora = Left$(Time$, 5)
        ZTimer = Int(Timer)
    End If
    
    If Val(Hoja.Text) <> 0 Then
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + " EquipoII = " + "'" + Equipo.Text + "',"
        ZSql = ZSql + " Etapa = " + "'" + Etapa.Text + "',"
        ZSql = ZSql + " FechaInicioEtapa = " + "'" + ZFecha + "',"
        ZSql = ZSql + " HoraInicioEtapa = " + "'" + ZHora + "',"
        ZSql = ZSql + " TimerInicioEtapa = " + "'" + ZTimer + "',"
        ZSql = ZSql + " Operario = " + "'" + Operario.Text + "',"
        ZSql = ZSql + " Desde = " + "'" + Desde.Text + "',"
        ZSql = ZSql + " Hasta = " + "'" + Hasta.Text + "',"
        ZSql = ZSql + " EstadoHoja = 1"
        ZSql = ZSql + " Where Hoja = " + "'" + Hoja.Text + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    End If
    
    Call CmdLimpiar_Click
    Equipo.SetFocus
        
End Sub

Private Sub CmdLimpiar_Click()

    Equipo.Text = ""
    Hoja.Text = ""
    Fecha.Text = "  /  /    "
    Producto.Text = "  -     -   "
    Teorico.Text = ""
    Etapa.Text = ""
    Hora.Text = ""
    Desde.Text = ""
    Hasta.Text = ""
    Operario.Text = ""
    DesOperario.Caption = ""
    DesProducto.Caption = ""

    Equipo.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    PrgCargaControl.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Equipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Equipo.Text) <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Hoja"
            ZSql = ZSql + " Where Hoja.EquipoII = " + "'" + Equipo.Text + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
            
                Hoja.Text = rstHoja!Hoja
                Producto.Text = rstHoja!Producto
                Fecha.Text = rstHoja!Fecha
                Teorico.Text = rstHoja!Teorico
                
                Etapa.Text = IIf(IsNull(rstHoja!Etapa), "", rstHoja!Etapa)
                Hora.Text = IIf(IsNull(rstHoja!HoraInicioEtapa), "", rstHoja!HoraInicioEtapa)
                Desde.Text = IIf(IsNull(rstHoja!Desde), "", rstHoja!Desde)
                Hasta.Text = IIf(IsNull(rstHoja!Hasta), "", rstHoja!Hasta)
                Operario.Text = IIf(IsNull(rstHoja!Operario), "", rstHoja!Operario)
                
                rstHoja.Close
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Operarios"
                ZSql = ZSql + " Where Operarios.Codigo = " + "'" + Operario.Text + "'"
                spOperarios = ZSql
                Set rstOperarios = db.OpenRecordset(spOperarios, dbOpenSnapshot, dbSQLPassThrough)
                If rstOperarios.RecordCount > 0 Then
                    DesOperario.Caption = Trim(rstOperarios!Descripcion)
                    rstOperarios.Close
                End If
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Terminado"
                ZSql = ZSql + " Where Terminado.Codigo = " + "'" + Producto.Text + "'"
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    DesProducto.Caption = Trim(rstTerminado!Descripcion)
                    rstTerminado.Close
                End If

                    Else
                    
                WEquipo = Equipo.Text
                CmdLimpiar_Click
                Equipo.Text = WEquipo
                
            End If
        End If
        
        Hoja.SetFocus
        
    End If
    
    If KeyAscii = 27 Then
        Equipo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Hoja"
        ZSql = ZSql + " Where Hoja.Hoja = " + "'" + Hoja.Text + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            Hoja.Text = rstHoja!Hoja
            
            Hoja.Text = rstHoja!Hoja
            Producto.Text = rstHoja!Producto
            Fecha.Text = rstHoja!Fecha
            Teorico.Text = rstHoja!Teorico
                
            Etapa.Text = IIf(IsNull(rstHoja!Etapa), "", rstHoja!Etapa)
            Hora.Text = IIf(IsNull(rstHoja!HoraInicioEtapa), "", rstHoja!HoraInicioEtapa)
            Desde.Text = IIf(IsNull(rstHoja!Desde), "", rstHoja!Desde)
            Hasta.Text = IIf(IsNull(rstHoja!Hasta), "", rstHoja!Hasta)
            Operario.Text = IIf(IsNull(rstHoja!Operario), "", rstHoja!Operario)
                
            rstHoja.Close
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Operarios"
            ZSql = ZSql + " Where Operarios.Codigo = " + "'" + Operario.Text + "'"
            spOperarios = ZSql
            Set rstOperarios = db.OpenRecordset(spOperarios, dbOpenSnapshot, dbSQLPassThrough)
            If rstOperarios.RecordCount > 0 Then
                DesOperario.Caption = Trim(rstOperarios!Descripcion)
                rstOperarios.Close
            End If
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Terminado"
            ZSql = ZSql + " Where Terminado.Codigo = " + "'" + Producto.Text + "'"
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                DesProducto.Caption = Trim(rstTerminado!Descripcion)
                rstTerminado.Close
            End If
                
            Etapa.SetFocus
            
                Else
                
            WEquipo = Equipo.Text
            CmdLimpiar_Click
            Equipo.Text = WEquipo
        End If
    End If
    If KeyAscii = 27 Then
        Hoja.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Etapa_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Etapa.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Operario.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Operario_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operarios"
        ZSql = ZSql + " Where Operarios.Codigo = " + "'" + Operario.Text + "'"
        spOperarios = ZSql
        Set rstOperarios = db.OpenRecordset(spOperarios, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperarios.RecordCount > 0 Then
            DesOperario.Caption = Trim(rstOperarios!Descripcion)
            rstOperarios.Close
            Etapa.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Operario.Text = ""
        DesOperario.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()

    Equipo.Text = ""
    Hoja.Text = ""
    Fecha.Text = "  /  /    "
    Producto.Text = "  -     -   "
    Teorico.Text = ""
    Etapa.Text = ""
    Hora.Text = ""
    Desde.Text = ""
    Hasta.Text = ""
    Operario.Text = ""
    DesOperario.Caption = ""
    DesProducto.Caption = ""
    
End Sub

