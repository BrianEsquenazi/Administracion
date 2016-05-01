VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaSacEstadistica 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Estadistica"
   ClientHeight    =   6945
   ClientLeft      =   2070
   ClientTop       =   615
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   6945
   ScaleWidth      =   8085
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   975
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
      Left            =   600
      TabIndex        =   12
      Top             =   3240
      Width           =   7095
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
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   6120
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   6000
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   7095
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
         Left            =   1680
         TabIndex        =   15
         Top             =   1560
         Width           =   3855
      End
      Begin VB.TextBox HastaTipo 
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
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   9
         Text            =   " "
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Ano 
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
         TabIndex        =   8
         Text            =   " "
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox DesdeTipo 
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
         Top             =   600
         Width           =   855
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   2160
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
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
         Left            =   5040
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
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
         Left            =   5040
         TabIndex        =   2
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   8040
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wficter.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   3015
      Left            =   600
      TabIndex        =   13
      Top             =   3600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5318
      _Version        =   327680
      BackColor       =   16777152
   End
End
Attribute VB_Name = "PrgListaSacEstadistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCargaSac As Recordset
Dim spCargaSac As String
Dim rstTipoSac As Recordset
Dim spTipoSac As String
Dim XParam As String

Private Sub DesdeTipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaTipo.SetFocus
    End If
    If KeyAscii = 27 Then
        DesdeTipo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub HastaTipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ano.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaTipo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Ano.Text) <> 0 Then
            DesdeTipo.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Acepta_Click()

    WDesde = Ano.Text + "0101"
    WHasta = Ano.Text + "1231"
    
    Rem ZCantidad = 0
    Rem ZCantidadAcciones = 0
    Rem ZCantidadImple = 0
    Rem ZCantidadCerradas = 0
    
    Rem Sql1 = "Select Clave, Ordfecha, Plazo1, Plazo2, Plazo3, Plazo4, Plazo5, Plazo6, Fecha1, Fecha2, Fecha3, Fecha4, Fecha5, Fecha6, Fecha21, Fecha22, Fecha23, Fecha24, Fecha25, Fecha26, Responsable1, Responsable2, Responsable3, Responsable4, Responsable5, Responsable6"
    Rem Sql2 = " FROM CargaSac"
    Rem Sql3 = " Order by CargaSac.Clave"
    Rem spCargaSac = Sql1 + Sql2 + Sql3
    Rem Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstCargaSac.RecordCount > 0 Then
    Rem     With rstCargaSac
    Rem         .MoveFirst
    Rem         Do
    Rem             If .EOF = False Then
    Rem                 If WDesde <= rstCargaSac!ordfecha And WHasta >= rstCargaSac!ordfecha Then
    Rem
    Rem                     ZCantidad = ZCantidad + 1
    Rem
    Rem                     If rstCargaSac!Plazo1 <> "  /  /    " Then
    Rem
    Rem                         ZCantidadAcciones = ZCantidadAcciones + 1
    Rem
    Rem                         If rstCargaSac!Responsable1 <> 0 Then
    Rem
    Rem                             ZCantidadImple = ZCantidadImple + 1
    Rem
    Rem                             If rstCargaSac!Fecha21 <> "  /  /    " Then
    Rem                                 ZCantidadCerradas = ZCantidadCerradas + 1
    Rem                             End If
    Rem
    Rem                         End If
    Rem
    Rem                     End If
    Rem
    Rem                 End If
     Rem                .MoveNext
    Rem                     Else
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem     End With
    Rem     rstCargaSac.Close
    Rem End If
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "UPDATE CargaSac SET "
    Rem ZSql = ZSql + " Cantidad = " + "'" + Str$(ZCantidad) + "',"
    Rem ZSql = ZSql + " CantidadAcciones = " + "'" + Str$(ZCantidadAcciones) + "',"
    Rem ZSql = ZSql + " CantidadImplementadas = " + "'" + Str$(ZCantidadImple) + "',"
    Rem ZSql = ZSql + " CantidadCerradas = " + "'" + Str$(ZCantidadCerradas) + "',"
    Rem ZSql = ZSql + " Porce1 = " + "'" + "0" + "',"
    Rem ZSql = ZSql + " Porce2 = " + "'" + "0" + "'"
    Rem ZSql = ZSql + " Where Centro = " + "'" + Centro.Text + "'"
    Rem spCargaSac = ZSql
    Rem Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    
        
    Listado.WindowTitle = "Listado de Estadistica"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    
    Uno = "{CargaSAc.Tipo} in " + DesdeTipo.Text + " to " + HastaTipo.Text
    Dos = " and {CargaSAC.Ano} in " + Ano.Text + " to " + Ano.Text
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CargaSAC.Tipo, CargaSAC.Ano, CargaSAC.Numero, CargaSAC.Centro, CargaSAC.Fecha, CargaSAC.OrdFecha, CargaSAC.Estado, CargaSAC.IngresoNoCon, CargaSAC.IngresoCausa, CargaSAC.Titulo, CargaSAC.Referencia, " _
                + "CentroSac.Descripcion, " _
                + "TipoSac.Descripcion " _
                + "From " _
                + DSQ + ".dbo.CargaSAC CargaSAC, " _
                + DSQ + ".dbo.CentroSac CentroSac, " _
                + DSQ + ".dbo.TipoSac TipoSac " _
                + "Where " _
                + "CargaSAC.Centro = CentroSac.Codigo AND " _
                + "CargaSAC.Tipo = TipoSac.Codigo AND " _
                + "CargaSAC.Tipo >= " + DesdeTipo.Text + " AND " _
                + "CargaSAC.Tipo <= " + HastaTipo.Text + " AND " _
                + "CargaSAC.Ano >= " + Ano.Text + " AND " _
                + "CargaSAC.Ano <= " + Ano.Text
    
    Listado.Connect = Connect()
    
    If Tipo.ListIndex = 0 Then
        Listado.ReportFileName = "ListadoSacEstadistica.rpt"
            Else
        Listado.ReportFileName = "ListadoSacEstadisticaREsu.rpt"
    End If
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    PrgListaSacEstadistica.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()

    Tipo.Clear

    Tipo.AddItem "Completo"
    Tipo.AddItem "Resumido"
    
    Tipo.ListIndex = 0
    
    DesdeTipo.Text = ""
    HastaTipo.Text = ""
    Ano.Text = ""
    
    Call Opcion
    
    Panta.Value = False
    Impresora.Value = True
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

        Call Limpia_Ayuda
        LugarAyuda = 0
        WIndice.Clear
    
        WEspacios = Len(Ayuda.Text)
    
    
        Sql1 = "Select *"
        Sql2 = " FROM TipoSac"
        Sql3 = " Order by TipoSac.Codigo"
        spTipoSac = Sql1 + Sql2 + Sql3
        Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoSac.RecordCount > 0 Then
            With rstTipoSac
                .MoveFirst
                Do
                    If .EOF = False Then
                        da = Len(rstTipoSac!Descripcion) - WEspacios
                        For aa = 1 To da + 1
                            If Left$(Ayuda.Text, WEspacios) = Mid$(rstTipoSac!Descripcion, aa, WEspacios) Then
                                LugarAyuda = LugarAyuda + 1
                                Pantalla.Row = LugarAyuda
                                Pantalla.Col = 1
                                Pantalla.Text = rstTipoSac!Codigo
                                Pantalla.Col = 2
                                Pantalla.Text = rstTipoSac!Descripcion
                                IngresaItem = rstTipoSac!Codigo
                                WIndice.AddItem IngresaItem
                                Exit For
                            End If
                        Next aa
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstTipoSac.Close
        End If
    
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Limpia_Ayuda()

    Pantalla.Clear
    Pantalla.Font.Bold = True
    
    ' Establesco loa Valores de la pantalla
    
    Pantalla.FixedCols = 1
    Pantalla.Cols = 3
    Pantalla.FixedRows = 1
    Pantalla.Rows = 10001
    
    Pantalla.ColWidth(0) = 200
    Pantalla.Row = 0
    
    For Ciclo = 1 To Pantalla.Cols - 1
        Pantalla.Col = Ciclo
        Select Case Ciclo
            Case 1
                Pantalla.Text = "Codigo"
                Pantalla.ColWidth(Ciclo) = 1000
                Pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                Pantalla.Text = "Nombre"
                Pantalla.ColWidth(Ciclo) = 6000
                Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    Pantalla.Row = 0
    For Ciclo = 1 To Pantalla.Cols - 1
        Pantalla.Col = Ciclo
        WTitulo(Ciclo).Text = Pantalla.Text
        WTitulo(Ciclo).Left = Pantalla.CellLeft + Pantalla.Left
        WTitulo(Ciclo).Top = Pantalla.CellTop + Pantalla.Top
        WTitulo(Ciclo).Width = Pantalla.CellWidth
        WTitulo(Ciclo).Height = Pantalla.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA pantalla
    
    WAncho = 400
    For Ciclo = 0 To Pantalla.Cols - 1
        WAncho = WAncho + Pantalla.ColWidth(Ciclo)
    Next Ciclo
    Pantalla.Width = WAncho

    ' Size the columns.
    Font.Name = Pantalla.Font.Name
    Font.Size = Pantalla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Pantalla.AllowUserResizing = flexResizeBoth
    
    Pantalla.Col = 1
    Pantalla.Row = 1
    
End Sub

Private Sub pantalla_Click()

    Indice = Pantalla.Row - 1
    DesdeTipo.Text = WIndice.List(Indice)
    HastaTipo.Text = WIndice.List(Indice)
    
End Sub

Private Sub Opcion()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Call Limpia_Ayuda
    LugarAyuda = 0
    WIndice.Clear

            Sql1 = "Select *"
            Sql2 = " FROM TipoSac"
            Sql3 = " Order by TipoSac.Codigo"
            spTipoSac = Sql1 + Sql2 + Sql3
            Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoSac.RecordCount > 0 Then
                With rstTipoSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            LugarAyuda = LugarAyuda + 1
                            Pantalla.Row = LugarAyuda
                            Pantalla.Col = 1
                            Pantalla.Text = rstTipoSac!Codigo
                            Pantalla.Col = 2
                            Pantalla.Text = rstTipoSac!Descripcion
                            IngresaItem = rstTipoSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTipoSac.Close
            End If
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

