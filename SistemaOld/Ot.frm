VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgOt 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Muestras a Clientes"
   ClientHeight    =   8325
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11790
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   8325
   ScaleWidth      =   11790
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
      Left            =   1800
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11040
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton ListadoII 
      Caption         =   "Listado (F5)"
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
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin Crystal.CrystalReport ListaGrilla 
      Left            =   10080
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ImpreOt.rpt"
   End
   Begin VB.ListBox Lista 
      Height          =   645
      Left            =   2640
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox Pantalla 
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
      Height          =   5715
      ItemData        =   "Ot.frx":0000
      Left            =   2040
      List            =   "Ot.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   4815
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   7455
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13150
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   4210752
      FocusRect       =   2
      GridLines       =   0
   End
   Begin VB.CommandButton Modifica 
      Caption         =   "Modifica (F3)"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Baja 
      Caption         =   "Elimina (F2)"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Alta 
      Caption         =   "Alta (F1)"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Impresion 
      Caption         =   "Impresion (F9)"
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
      Left            =   6360
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Fin (F10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7920
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgOt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Dim rstCompo As Recordset
Dim spCompo As String
Dim rstSolidez As Recordset
Dim spSolidez As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstOt As Recordset
Dim spOt As String
Dim rstOtImpre As Recordset
Dim spOtImpre As String
Dim XParam As String
Dim Auxiliar(10000, 20)
Dim XEmpresa As String
Dim WFecha As String
Dim WFecha2 As String
Dim SeparaFecha As Integer
Dim SumaDia As Integer
Dim SumaMes As Integer
Dim WDia As String
Dim WMes As String
Dim WCod As String
Dim ColumnaOpcion As Integer
Dim Seleccion As String
Dim WPasa(10000) As String

Private Sub Alta_Click()
    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    Muestra.Visible = False
    WOt = 0
    PrgAltaOt.Show
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Composicion_Click()
    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    Muestra.Visible = False
    WOt = 0
    PrgComposicion.Show
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Ayuda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 ColumnaOpcion = Muestra.Col
    Pantalla.Clear
Rem    WIndice.Clear
    Ayuda.Text = UCase(Ayuda.Text)
    WEspacios = Len(Ayuda.Text)
   ingresaItem = ""
    Pantalla.AddItem ingresaItem
          spOt = "ListaOtCliente"
            Set rstOt = db.OpenRecordset(spOt, dbOpenSnapshot, dbSQLPassThrough)
            With rstOt
                .MoveFirst
              
            Do
                If .EOF = False Then
            
                    da = Len(rstOt!Razon) - WEspacios
                
                    For aa = 1 To da
                        If Left$(Ayuda.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                            Auxi = rstOt!Razon
                            ingresaItem = Auxi + "    " + rstOt!Razon
                            Pantalla.AddItem ingresaItem
                            ingresaItem = rstOt!Razon
                        Rem    WIndice.AddItem IngresaItem
                            Exit For
                        End If
                    Next aa
                    .MoveNext
                    
                        Else
                        
                    Exit Do
                
                End If
            Loop
        End With
        rstOt.Close
    Pantalla.Visible = True
    
    End If
   

End Sub

Private Sub ListadoII_Click()

    Rem Sql1 = "UPDATE Insumos SET "
    Rem Sql2 = " TipoSolicitud = 0"
    Rem spInsumo = Sql1 + Sql2
    Rem Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
     Ayuda.Visible = False
    RowIni = Muestra.Row
    RowFin = Muestra.RowSel
    
    ZDesdeNumero = Muestra.TextMatrix(RowIni, 1)
    ZHastaNumero = Muestra.TextMatrix(RowFin, 1)
    
    For Ciclo = RowIni To RowFin
    
        ZNumero = Muestra.TextMatrix(Ciclo, 1)
        ZFechaInicio = Muestra.TextMatrix(Ciclo, 4)
        zfechaFinal = Muestra.TextMatrix(Ciclo, 5)
        
        WDias = "0"
        
        If zfechaFinal <> "  /  /    " Then
            ZDias = DateDiff("d", ZFechaInicio, zfechaFinal)
            WDias = Str$(ZDias)
        End If
        
        Sql1 = "UPDATE Ot SET "
        Sql2 = " Dias = " + "'" + WDias + "'"
        Sql3 = " Where Codigo = " + "'" + ZNumero + "'"
        spOt = Sql1 + Sql2 + Sql3
        Set rstOt = db.OpenRecordset(spOt, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    Listado.WindowTitle = "Listado de Ordenes de Trabajo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Ot.Codigo} in " + ZDesdeNumero + " to " + ZHastaNumero
    Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Ot.Codigo, Ot.Fecha, Ot.Cliente, Ot.Razon, Ot.Preparacion, Ot.Observaciones1, Ot.Solicitante, Ot.FechaCompro, Ot.FechaSalida, Ot.Dias " _
                    + "From " _
                    + DSQ + ".dbo.Ot Ot " _
                    + "Where " _
                    + "Ot.Codigo >= " + ZDesdeNumero + " AND " _
                    + "Ot.Codigo <= " + ZHastaNumero
    
    Listado.Connect = Connect()
    Listado.ReportFileName = "ListaOt.rpt"
    Listado.Action = 1

End Sub

Private Sub Modifica_Click()
    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    Muestra.Visible = False
    Fila = Muestra.Row
    WOt = Auxiliar(Fila, 1)
    If Val(WOt) <> 0 Then
        PrgAltaOt.Show
    End If
End Sub

Private Sub Baja_Click()
    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WPosi3 = Muestra.Col
    Fila = Muestra.Row
    WOt = Auxiliar(Fila, 1)
    If Val(WOt) <> 0 Then
        T$ = "Borrar Registro"
        m$ = "Desea Borrar la Orden de Trabajo"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            spOt = "BorrarOt " + "'" + WOt + "'"
            Set rstOt = db.OpenRecordset(spOt, dbOpenDynaset, dbSQLPassThrough)
            Call Proceso_Click
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    PrgOt.Hide
    Unload Me
    Close
    End
End Sub

Private Sub Impresion_Click()

    RowIni = Muestra.Row
    RowFin = Muestra.RowSel
    
    For Ciclo = RowIni To RowFin
    
        ZNumero = Muestra.TextMatrix(Ciclo, 1)
        
        ListaGrilla.GroupSelectionFormula = "{Ot.Codigo} in " + ZNumero + " to " + ZNumero
        ListaGrilla.Destination = 1
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        ListaGrilla.SQLQuery = "SELECT  Ot.Codigo, Ot.Fecha, Ot.Razon, Ot.Preparacion, Ot.Solidez, Ot.Observaciones1, Ot.Observaciones2, " _
                        + "Ot.Observaciones3, Ot.Solicitante, Ot.FechaCompro, " _
                        + "Ot.Compo, Ot.Compo1, Ot.Compo2, Ot.Compo3, Ot.Compo4, Ot.Compo5, Ot.Compo6, Ot.Compo7, Ot.Compo8, Ot.Compo9, Ot.Compo10, Ot.Compo11, Ot.Compo12, Ot.Compo13, Ot.Compo14, " _
                        + "Ot.Traba, Ot.Trabajo1, Ot.Trabajo2, Ot.Trabajo3, Ot.Trabajo4, Ot.Trabajo5, Ot.Trabajo6, Ot.Trabajo7, Ot.Trabajo8, Ot.Trabajo9, Ot.Trabajo10, Ot.Trabajo11, Ot.Trabajo12, Ot.Trabajo13, Ot.Trabajo14, " _
                        + "Ot.Color, Ot.Color1, Ot.Color2, Ot.Color3, Ot.Color4, Ot.Color5, Ot.Color6, Ot.Color7, Ot.Color8, Ot.Color9, Ot.Color10, Ot.Color11, Ot.Color12, Ot.Color13, Ot.Color14, Ot.Color15, Ot.Color16, Ot.Color17, Ot.Color18, Ot.Color19, Ot.Color20, Ot.Color21, " _
                        + "Ot.Maqui, Ot.Maquina1, Ot.Maquina2, Ot.Maquina3, Ot.Maquina4, Ot.Maquina5, Ot.Maquina6, Ot.Maquina7, Ot.Maquina8, Ot.Maquina9, Ot.Maquina10, Ot.Maquina11, Ot.Maquina12, Ot.Maquina13, Ot.Maquina14, " _
                        + "OtConfig.Compo1, OtConfig.Compo2, OtConfig.Compo3, OtConfig.Compo4, OtConfig.Compo5, OtConfig.Compo6, OtConfig.Compo7, OtConfig.Compo8, OtConfig.Compo9, OtConfig.Compo10, OtConfig.Compo11, OtConfig.Compo12, OtConfig.Compo13, OtConfig.Compo14, " _
                        + "OtConfig.Trabajo1, OtConfig.Trabajo2, OtConfig.Trabajo3, OtConfig.Trabajo4, OtConfig.Trabajo5, OtConfig.Trabajo6, OtConfig.Trabajo7, OtConfig.Trabajo8, OtConfig.Trabajo9, OtConfig.Trabajo10, OtConfig.Trabajo11, OtConfig.Trabajo12, OtConfig.Trabajo13, OtConfig.Trabajo14, " _
                        + "OtConfig.Color1, OtConfig.Color2, OtConfig.Color3, OtConfig.Color4, OtConfig.Color5, OtConfig.Color6, OtConfig.Color7, OtConfig.Color8, OtConfig.Color9, OtConfig.Color10, OtConfig.Color11, OtConfig.Color12, OtConfig.Color13, OtConfig.Color14, " _
                        + "OtConfig.Color15, OtConfifig.Color19, OtConfig.Color20, OtConfig.Color21, " _
                        + "OtConfig.Maquina1, OtConfig.Maquina2, OtConfig.Maquina3, OtConfig.Maquina4, OtConfig.Maquina5, OtConfig.Maquina6, OtConfig.Maquina7, OtConfig.Maquina8, OtConfig.Maquina9, OtConfig.Maquina10, OtConfig.Maquina11, OtConfig.Maquina12, OtConfig.Maquina13, OtConfig.Maquina14 " _
                        + "From " _
                        + DSQ + ".dbo.Ot Ot, " _
                        + DSQ + ".dbo.OtConfig OtConfig " _
                        + "Where " _
                        + "Ot.Clave = OtConfig.Clave AND " _
                        + "Ot.Codigo >= " + ZNumero + " AND Ot.Codigo <= " + ZNumero
        ListaGrilla.ReportFileName = "ImpreOt.rpt"
        ListaGrilla.Connect = Connect()
        ListaGrilla.Action = 1
        
    Next Ciclo
    
End Sub

Private Sub Proceso_Click()

    Muestra.Visible = True
    Call Limpia_Vector
        
    Select Case ColumnaOpcion
        Case 0, 1
            spOt = "ListaOtTotal "
        Case 2
            spOt = "ListaOtFechaSolo " + "'" + Seleccion + "'"
        Case 3
            spOt = "ListaOtClienteSolo " + "'" + Seleccion + "'"
        Case 4
            spOt = "ListaOtFComproSolo " + "'" + Seleccion + "'"
        Case 5
            spOt = "ListaOtFSalidaSolo " + "'" + Seleccion + "'"
        Case 6
            spOt = "ListaOtSolicitanteSolo " + "'" + Seleccion + "'"
        Case 7
            spOt = "ListaOtObservaciones1Solo " + "'" + Seleccion + "'"
        Case Else
    End Select
            
    Set rstOt = db.OpenRecordset(spOt, dbOpenSnapshot, dbSQLPassThrough)
    If rstOt.RecordCount > 0 Then
        With rstOt
    
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    WLugar = WLugar + 1
                    Auxiliar(WLugar, 1) = Str$(rstOt!Codigo)
                    Auxiliar(WLugar, 2) = rstOt!Fecha
                    Auxiliar(WLugar, 3) = rstOt!Cliente
                    Auxiliar(WLugar, 4) = rstOt!Razon
                    Auxiliar(WLugar, 5) = rstOt!FechaCompro
                    Auxiliar(WLugar, 6) = rstOt!FechaSalida
                    Auxiliar(WLugar, 7) = rstOt!Solicitante
                    Auxiliar(WLugar, 8) = rstOt!Observaciones1
                    Auxiliar(WLugar, 9) = rstOt!Preparacion
                    
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
        
        End With
        rstOt.Close
    End If
    
    For Cicla = 1 To WLugar
    
        Muestra.Row = Cicla
        
        Muestra.Col = 1
        Muestra.Text = Auxiliar(Cicla, 1)
        
        Muestra.Col = 2
        Muestra.Text = Auxiliar(Cicla, 2)
        
        Muestra.Col = 3
        Muestra.Text = Auxiliar(Cicla, 4)
        
        Muestra.Col = 4
        Muestra.Text = Auxiliar(Cicla, 5)
        
        Muestra.Col = 5
        Muestra.Text = Auxiliar(Cicla, 6)
        
        Muestra.Col = 6
        Muestra.Text = Auxiliar(Cicla, 7)
        
        Muestra.Col = 7
        Muestra.Text = Auxiliar(Cicla, 8)
        
        Muestra.Col = 8
        Muestra.Text = Auxiliar(Cicla, 9)
        
    Next Cicla
    
    Muestra.Visible = True
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    If WPosi1 <> 0 And WPosi2 <> 0 And WPosi3 <> 0 Then
        Muestra.TopRow = WPosi1
        Muestra.Col = WPosi3
        Muestra.Row = WPosi2
            Else
        If WLugar > 20 Then
            Muestra.TopRow = WLugar - 20
                Else
            Muestra.TopRow = 1
        End If
        Muestra.Col = 1
        Muestra.Row = WLugar
    End If
    
    Muestra.SetFocus
    
End Sub

Private Sub Muestra_Click()
 Ayuda.Visible = False
Rem ColumnaOpcion = Muestra.Col
End Sub

Private Sub Muestra_DblClick()
    Ayuda.Visible = False
    ColumnaOpcion = Muestra.Col
    WPosi1 = 1
    WPosi2 = 1
    WPosi3 = 1
    
    Pantalla.Clear
    Select Case ColumnaOpcion
        Case 1
            Call Modifica_Click
            
        Case 2
            Pasa = 0
            corte = ""
            spOt = "ListaOtFecha"
            Set rstOt = db.OpenRecordset(spOt, dbOpenSnapshot, dbSQLPassThrough)
            With rstOt
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            Pantalla.AddItem ""
                            Pasa = 1
                            corte = rstOt!Fecha
                        End If
                        If corte <> rstOt!Fecha Then
                            Pantalla.AddItem corte
                            corte = rstOt!Fecha
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            Pantalla.AddItem corte
            rstOt.Close
            Pantalla.Visible = True
            
        Case 3
            Pasa = 0
            corte = ""
            spOt = "ListaOtCliente"
            Set rstOt = db.OpenRecordset(spOt, dbOpenSnapshot, dbSQLPassThrough)
            With rstOt
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            Pantalla.AddItem ""
                            Pasa = 1
                            corte = rstOt!Razon
                        End If
                        If corte <> rstOt!Razon Then
                            Pantalla.AddItem corte
                            corte = rstOt!Razon
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            Pantalla.AddItem corte
            rstOt.Close
            Pantalla.Visible = True
            
        Case 4
            Pasa = 0
            corte = ""
            spOt = "ListaOtFCompro"
            Set rstOt = db.OpenRecordset(spOt, dbOpenSnapshot, dbSQLPassThrough)
            With rstOt
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            Pantalla.AddItem ""
                            Pasa = 1
                            corte = rstOt!FechaCompro
                        End If
                        If corte <> rstOt!FechaCompro Then
                            Pantalla.AddItem corte
                            corte = rstOt!FechaCompro
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            Pantalla.AddItem corte
            rstOt.Close
            Pantalla.Visible = True
            
        Case 5
            Pasa = 0
            corte = ""
            spOt = "ListaOtFechaSalida"
            Set rstOt = db.OpenRecordset(spOt, dbOpenSnapshot, dbSQLPassThrough)
            With rstOt
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            Pantalla.AddItem ""
                            Pasa = 1
                            corte = rstOt!FechaSalida
                        End If
                        If corte <> rstOt!FechaSalida Then
                            Pantalla.AddItem corte
                            corte = rstOt!FechaSalida
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            Pantalla.AddItem corte
            rstOt.Close
            Pantalla.Visible = True
            
        Case 6
            Pasa = 0
            corte = ""
            spOt = "ListaOtSolicitante"
            Set rstOt = db.OpenRecordset(spOt, dbOpenSnapshot, dbSQLPassThrough)
            With rstOt
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            Pantalla.AddItem ""
                            Pasa = 1
                            corte = rstOt!Solicitante
                        End If
                        If corte <> rstOt!Solicitante Then
                            Pantalla.AddItem corte
                            corte = rstOt!Solicitante
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            Pantalla.AddItem corte
            rstOt.Close
            Pantalla.Visible = True
            
        Case 7
            Pasa = 0
            corte = ""
            spOt = "ListaOtObservaciones1"
            Set rstOt = db.OpenRecordset(spOt, dbOpenSnapshot, dbSQLPassThrough)
            With rstOt
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Pasa = 0 Then
                            Pantalla.AddItem ""
                            Pasa = 1
                            corte = rstOt!Observaciones1
                        End If
                        If corte <> rstOt!Observaciones1 Then
                            Pantalla.AddItem corte
                            corte = rstOt!Observaciones1
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            Pantalla.AddItem corte
            rstOt.Close
            Pantalla.Visible = True
            
            
        Case Else
        
    End Select
    
    Rem Muestra.Col = 10
    Rem Muestra.Col = 1
    Rem WXSol = Muestra.Text
    Rem PrgSol.Show
End Sub

Private Sub Form_Activate()
    Muestra.Visible = False
    Call Proceso_Click
End Sub

Private Sub Muestra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Alta_Click
        Case 113
            Call Baja_Click
        Case 114
            Call Modifica_Click
        Case 120
            Call Impresion_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector()

    Muestra.Clear

    Rem ponga la muestra en negritas
    Rem Muestra.Font.Bold = True

    ' Establesco loa Valores de la muestra
    
    Muestra.FixedCols = 1
    Muestra.Cols = 9
    Muestra.FixedRows = 1
    Muestra.Rows = 5000
    
    Muestra.ColWidth(0) = 200
    Muestra.Row = 0
    
    For Ciclo = 1 To Muestra.Cols - 1
        Muestra.Col = Ciclo
        Select Case Ciclo
            Case 1
                Muestra.Text = "Codigo"
                Muestra.ColWidth(Ciclo) = 600
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                Muestra.Text = "Fecha"
                Muestra.ColWidth(Ciclo) = 1000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                Muestra.Text = "Cliente"
                Muestra.ColWidth(Ciclo) = 2400
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                Muestra.Text = "F.Comprom."
                Muestra.ColWidth(Ciclo) = 1000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                Muestra.Text = "F.Salida"
                Muestra.ColWidth(Ciclo) = 1000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 6
                Muestra.Text = "Solicitante"
                Muestra.ColWidth(Ciclo) = 900
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 7
                Muestra.Text = "Descripcion"
                Muestra.ColWidth(Ciclo) = 2400
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                Muestra.Text = "Observaciones"
                Muestra.ColWidth(Ciclo) = 2400
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Muestra.AllowUserResizing = flexResizeBoth
    
    Muestra.Col = 1
    Muestra.Row = 1
    
End Sub

Private Sub pantalla_Click()
    If Pantalla.ListIndex <> 0 Then
        Seleccion = Pantalla.Text
            Else
        Seleccion = ""
        ColumnaOpcion = 0
    End If
    Pantalla.Visible = False
    Call Proceso_Click
Ayuda.Text = ""
Ayuda.Visible = False
End Sub

