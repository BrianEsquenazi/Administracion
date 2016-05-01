VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgConsultaHojanuevo 
   AutoRedraw      =   -1  'True
   Caption         =   " Consultas de Hojas de Produccion"
   ClientHeight    =   7320
   ClientLeft      =   90
   ClientTop       =   690
   ClientWidth     =   11850
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11850
   Begin VB.TextBox Operario 
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Estado 
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
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
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
      Height          =   300
      Left            =   10320
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10186
      _Version        =   327680
      Rows            =   4000
      Cols            =   11
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   5520
      MouseIcon       =   "consultahojanuevo.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "consultahojanuevo.frx":030A
      ToolTipText     =   "Salida"
      Top             =   6600
      Width           =   480
   End
End
Attribute VB_Name = "PrgConsultaHojanuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Dim XParam As String
Dim WGraba As String
Dim ZVector(100, 8) As String
Dim ZOpera(1000) As String
Dim XEmpresa As String

Private Sub cmdClose_Click()
    PrgConsultaHoja.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Estado_click()
    Rem Call Proceso_Click
End Sub

Private Sub Form_Activate()
    Call Proceso_Click
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    Muestra.Font.Bold = True
    
    Muestra.ColWidth(0) = 50
    Muestra.ColWidth(1) = 800
    Muestra.ColWidth(2) = 1200
    Muestra.ColWidth(3) = 1500
    Muestra.ColWidth(4) = 1000
    Muestra.ColWidth(5) = 2000
    Muestra.ColWidth(6) = 800
    Muestra.ColWidth(7) = 1000
    Muestra.ColWidth(8) = 1000
    Muestra.ColWidth(9) = 1000
    Muestra.ColWidth(10) = 1000
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Hoja"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Pt"
    
    Muestra.Col = 4
    Muestra.Text = "Cantidad"
    
    Muestra.Col = 5
    Muestra.Text = "Operador"
    
    Muestra.Col = 6
    Muestra.Text = "Equipo"
    
    Muestra.Col = 7
    Muestra.Text = "Hora"
    
    Muestra.Col = 8
    Muestra.Text = "Etapa"
    
    Muestra.Col = 9
    Muestra.Text = "Hora"
    
    Muestra.Col = 10
    Muestra.Text = "Alerta"
    
    Estado.Clear
    
    Estado.AddItem "Todos"
    Estado.AddItem "Pendientes"
    Estado.AddItem "En Proceso"
    
    Estado.ListIndex = 0
    
    Rem Call Proceso_Click
    
End Sub




Private Sub Operario_KeyPress(KeyAscii As Integer)
    Call Proceso_Click
End Sub

Private Sub Proceso_Click()

    XEmpresa = WEmpresa
    
    WSalida = "N"
        
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Hoja"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Pt"
    
    Muestra.Col = 4
    Muestra.Text = "Cantidad"
    
    Muestra.Col = 5
    Muestra.Text = "Operador"
    
    
    Muestra.Col = 6
    Muestra.Text = "Equipo"
    
    Muestra.Col = 7
    Muestra.Text = "Hora"
    
    Muestra.Col = 8
    Muestra.Text = "Etapa"
    
    Muestra.Col = 9
    Muestra.Text = "Hora"
    
    Muestra.Col = 10
    Muestra.Text = "Alerta"
    
    Renglon = 0
    
    Select Case Estado.ListIndex
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Hoja"
            ZSql = ZSql + " Where (Hoja.EstadoHoja = 0 or Hoja.EstadoHoja = 1) and Hoja.Renglon = 1"
            ZSql = ZSql + " Order by Hoja.Hoja"

        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Hoja"
            ZSql = ZSql + " Where Hoja.EstadoHoja = 0 and Hoja.Renglon = 1"
            ZSql = ZSql + " Order by Hoja.Hoja"
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Hoja"
            ZSql = ZSql + " Where Hoja.EstadoHoja = 1 and Hoja.Renglon = 1"
            ZSql = ZSql + " Order by Hoja.Hoja"

        Case Else
    End Select
            
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If Val(Operario.Text) = 0 Or Val(Operario.Text) = rstHoja!Operario Then
                
                    Renglon = Renglon + 1
            
                    Muestra.TextMatrix(Renglon, 1) = Pusing("######", Str$(rstHoja!Hoja))
                    Muestra.TextMatrix(Renglon, 2) = rstHoja!Fecha
                    Muestra.TextMatrix(Renglon, 3) = rstHoja!Producto
                    Muestra.TextMatrix(Renglon, 4) = rstHoja!Teorico
                    Muestra.TextMatrix(Renglon, 5) = ""
                    Muestra.TextMatrix(Renglon, 6) = rstHoja!Equipo
                    ZHoraInicio = IIf(IsNull(rstHoja!HoraInicio), "", rstHoja!HoraInicio)
                    Muestra.TextMatrix(Renglon, 7) = ZHoraInicio
                    If rstHoja!estadohoja = 0 Then
                        Muestra.TextMatrix(Renglon, 9) = ""
                        Muestra.TextMatrix(Renglon, 10) = ""
                            Else
                        ZHoraInicioEtapa = IIf(IsNull(rstHoja!HoraInicioEtapa), "", rstHoja!HoraInicioEtapa)
                        Muestra.TextMatrix(Renglon, 9) = ZHoraInicioEtapa
                        ZAlarma = IIf(IsNull(rstHoja!Alarma), " ", rstHoja!Alarma)
                        ZAlarmaI = IIf(IsNull(rstHoja!AlarmaI), " ", rstHoja!AlarmaI)
                        ZAlarmaII = IIf(IsNull(rstHoja!AlarmaII), " ", rstHoja!AlarmaII)
                        Muestra.TextMatrix(Renglon, 10) = ZAlarma + "-" + ZAlarmaI + "-" + ZAlarmaII
                    End If
                    Muestra.TextMatrix(Renglon, 8) = IIf(IsNull(rstHoja!etapa), "", rstHoja!etapa)
                    
                    End If
                    
                    ZOpera(Renglon) = rstHoja!Operario
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If
    
    For Ciclo = 1 To Renglon
    
        Sql1 = "Select *"
        Sql2 = " FROM Operarios"
        Sql3 = " Where Operarios.Codigo = " + "'" + ZOpera(Ciclo) + "'"
        spOperarios = Sql1 + Sql2 + Sql3
        Set rstOperarios = db.OpenRecordset(spOperarios, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperarios.RecordCount > 0 Then
            Muestra.TextMatrix(Ciclo, 5) = rstOperarios!Descripcion
            rstOperarios.Close
        End If
        
    Next Ciclo
    
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    
    Muestra.TopRow = 1
    
End Sub

Private Sub Limpia_Vector()
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Hoja"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Pt"
    
    Muestra.Col = 4
    Muestra.Text = "Cantidad"
    
    
    Muestra.Col = 6
    Muestra.Text = "Equipo"
    
    Muestra.Col = 7
    Muestra.Text = "Hora"
    
    Muestra.Col = 8
    Muestra.Text = "Etapa"
    
    Muestra.Col = 9
    Muestra.Text = "Hora"
    
    Muestra.Col = 10
    Muestra.Text = "Alerta"
    
End Sub

Private Sub Muestra_DblClick()

    Muestra.Col = 1
    ZHojaProceso = Muestra.Text
    Muestra.Col = 3
    ZTerminadoProceso = Muestra.Text
    Muestra.Col = 4
    ZCantidadProceso = Muestra.Text
    Muestra.Col = 8
    ZEtapaProceso = Muestra.Text
    If Val(ZEtapaProceso) = 0 Then
    
        ZEtapaProceso = "1"
        ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        ZHora = Left$(Time$, 5)
        ZTimer = Int(Timer)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + " Etapa = " + "'" + ZEtapaProceso + "',"
        ZSql = ZSql + " EstadoHoja = " + "'" + "1" + "',"
        ZSql = ZSql + " FechaInicioEtapa = " + "'" + ZFecha + "',"
        ZSql = ZSql + " HoraInicioEtapa = " + "'" + ZHora + "',"
        ZSql = ZSql + " TimerInicioEtapa = " + "'" + Str$(ZTimer) + "',"
        ZSql = ZSql + " FechaInicio = " + "'" + ZFecha + "',"
        ZSql = ZSql + " HoraInicio = " + "'" + ZHora + "',"
        ZSql = ZSql + " Alarma = " + "'" + "" + "',"
        ZSql = ZSql + " AlarmaI = " + "'" + "" + "',"
        ZSql = ZSql + " AlarmaII = " + "'" + "" + "'"
        ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    End If
    
    Rem PrgConsultaHoja.Hide
    Rem Unload Me
    PrgProcesoNueva.Show
    
End Sub

