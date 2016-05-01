VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgLeePlanilla 
   Caption         =   "Proceso de Lectura de Planillas de Excel"
   ClientHeight    =   2265
   ClientLeft      =   3465
   ClientTop       =   1020
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   6765
   Begin MSMask.MaskEdBox Terminado 
      Height          =   285
      Left            =   2640
      TabIndex        =   0
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
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   3840
      MouseIcon       =   "leeplanilla.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "leeplanilla.frx":030A
      ToolTipText     =   "Salida"
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Proceso 
      Height          =   480
      Left            =   2160
      MouseIcon       =   "leeplanilla.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "leeplanilla.frx":0E56
      ToolTipText     =   "Inicia el proceso de Grabacion"
      Top             =   1200
      Width           =   480
   End
End
Attribute VB_Name = "PrgLeePlanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objLibro As Object
Dim ruta As String
Dim ZPlanilla(1000, 7) As String
Dim ZProceso(1000, 3) As String
Dim ZPro As Integer

Dim ZFecha As String
Dim ZVersion As String

Private Sub Proceso_Click()

    Screen.MousePointer = vbHourglass
    
    Set appExcel = CreateObject("Excel.application")
    ruta = "c:\Planillas\" + Terminado + ".xls"

    If Len(Dir(ruta)) > 0 Then
    
        Set objLibro = appExcel.workbooks.Open(ruta)
        
        Lugar1 = 0
        Lugar2 = 0
        Erase ZPlanilla
        Erase ZProceso
        
        LugarPlanilla = 6
        WProceso = 0
        
        ZFecha = appExcel.cells(3, 2).Value
        ZVersion = appExcel.cells(4, 2).Value
        
        Do
        
            LugarPlanilla = LugarPlanilla + 1
            
            Campo1 = appExcel.cells(LugarPlanilla, 1).Value
            If UCase(Campo1) = "XXXXXX" Then
                WProceso = 1
                    Else
                If UCase(Campo1) = "ZZZZZZ" Then
                    Exit Do
                        Else
                    If WProceso = 0 Then
                        Lugar1 = Lugar1 + 1
                        ZPlanilla(Lugar1, 1) = appExcel.cells(LugarPlanilla, 1).Value
                        ZPlanilla(Lugar1, 2) = appExcel.cells(LugarPlanilla, 2).Value
                        ZPlanilla(Lugar1, 3) = appExcel.cells(LugarPlanilla, 3).Value
                        ZPlanilla(Lugar1, 4) = appExcel.cells(LugarPlanilla, 4).Value
                        ZPlanilla(Lugar1, 5) = appExcel.cells(LugarPlanilla, 5).Value
                        ZPlanilla(Lugar1, 6) = appExcel.cells(LugarPlanilla, 6).Value
                        ZPlanilla(Lugar1, 7) = appExcel.cells(LugarPlanilla, 7).Value
                            Else
                        Lugar2 = Lugar2 + 1
                        ZPro = Val(appExcel.cells(LugarPlanilla, 1).Value)
                        If ZProceso(ZPro, 1) = "" Then
                            ZProceso(ZPro, 1) = appExcel.cells(LugarPlanilla, 2).Value
                                Else
                            If ZProceso(ZPro, 2) = "" Then
                                ZProceso(ZPro, 2) = appExcel.cells(LugarPlanilla, 2).Value
                                    Else
                                ZProceso(ZPro, 3) = appExcel.cells(LugarPlanilla, 2).Value
                            End If
                        End If
                    End If
                End If
            End If
        Loop
            
        appExcel.Quit
        Set appExcel = Nothing
        
        ZOrdFecha = ""
        ZLote = ""
        ZAutorizado = "S"
        ZDesTerminado = ""
        
        Sql1 = "DELETE CargaIV"
        Sql2 = " Where Terminado = " + "'" + Terminado.Text + "'"
        spCargaIV = Sql1 + Sql2
        Set rstCargaIV = db.OpenRecordset(spCargaIV, dbOpenSnapshot, dbSQLPassThrough)
        
        For Ciclo = 1 To Lugar1
        
            Etapa = ZPlanilla(Ciclo, 1)
            Instrucciones = ZPlanilla(Ciclo, 2)
            Equipo = ZPlanilla(Ciclo, 3)
            Temperatura = ZPlanilla(Ciclo, 4)
            Tiempo = ZPlanilla(Ciclo, 5)
            Control = ZPlanilla(Ciclo, 6)
            Seguridad = ZPlanilla(Ciclo, 7)
            
            LetraInstrucciones = ""
            LetraTemperatura = ""
            LetraTiempo = ""
            LetraControl = ""
        
            Auxi = Str$(Ciclo)
            Call Ceros(Auxi, 2)
        
            WClave = Terminado.Text + Auxi
        
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
            Sql25 = "'" + Str$(Ciclo) + "',"
            Sql26 = "'" + ZFecha + "',"
            Sql27 = "'" + ZOrdFecha + "',"
            Sql28 = "'" + ZLote + "',"
            Sql29 = "'" + ZVersion + "',"
            Sql30 = "'" + ZAutorizado + "',"
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
            Sql42 = "'" + ZDesTerminado + "')"
                
            rsCargaIV = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 _
                 + Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 _
                 + Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 _
                 + Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 _
                 + Sql41 + Sql42
            Set rstCargaIV = db.OpenRecordset(rsCargaIV, dbOpenSnapshot, dbSQLPassThrough)
            
        Next Ciclo
    
    End If
        
        
        
    For Ciclo = 1 To 100
    
        If ZProceso(Ciclo, 1) <> "" Then
        
            WPRo = Str$(Ciclo)
            WDescri1 = ZProceso(Ciclo, 1)
            WDescri2 = ZProceso(Ciclo, 2)
            WDescri3 = ZProceso(Ciclo, 3)
        
            Sql1 = "Select *"
            Sql2 = " FROM EquipoFabrica"
            Sql3 = " Where EquipoFabrica.Codigo = " + "'" + WPRo + "'"
            spEquipoFabrica = Sql1 + Sql2 + Sql3
            Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEquipoFabrica.RecordCount > 0 Then
                rstEquipoFabrica.Close
                Sql1 = "UPDATE EquipoFabrica SET "
                Sql2 = " Descripcion = " + "'" + WDescri1 + "',"
                Sql3 = " DescripcionII = " + "'" + WDescri2 + "',"
                Sql4 = " DescripcionIII = " + "'" + WDescri3 + "'"
                Sql5 = " Where Codigo = " + "'" + WPRo + "'"
                spEquipoFabrica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                Sql1 = "INSERT INTO EquipoFabrica ("
                Sql2 = "Codigo ,"
                Sql3 = "Descripcion ,"
                Sql4 = "DescripcionII ,"
                Sql5 = "DescripcionIII )"
                Sql6 = "Values ("
                Sql7 = "'" + WPRo + "',"
                Sql8 = "'" + WDescri1 + "',"
                Sql9 = "'" + WDescri2 + "',"
                Sql10 = "'" + WDescri3 + "')"
                spEquipoFabrica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10
                Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
            End If
        
        End If
        
    Next Ciclo
        
    
    Screen.MousePointer = vbDefault
    
    m$ = "El Procesa a finalizado"
    A% = MsgBox(m$, 0, "Proceso de Lectura de Grabacion de Planillas de Excel")
    
    Call cmdClose_Click

End Sub

Private Sub cmdClose_Click()
    PrgLeePlanilla.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()
    Terminado.Text = "  -     -   "
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Sucursal.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Sucursal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


