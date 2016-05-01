VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEspecifi 
   Caption         =   "Ingreso de Especificaciones de Materia Prima (Historico)"
   ClientHeight    =   7350
   ClientLeft      =   450
   ClientTop       =   615
   ClientWidth     =   11160
   LinkTopic       =   "Form2"
   ScaleHeight     =   7350
   ScaleWidth      =   11160
   Begin MSMask.MaskEdBox Codigo 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin Crystal.CrystalReport lista 
      Left            =   4560
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wespec1.rpt"
      GroupSelectionFormula=   " "
      DiscardSavedData=   -1  'True
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control de Listado"
      Height          =   1575
      Left            =   4200
      TabIndex        =   52
      Top             =   4800
      Visible         =   0   'False
      Width           =   3135
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1320
         TabIndex        =   63
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1320
         TabIndex        =   62
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton ImpreListado 
         Caption         =   "Option2"
         Height          =   195
         Left            =   1920
         TabIndex        =   58
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton ImprePantalla 
         Caption         =   "Option1"
         Height          =   195
         Left            =   1920
         TabIndex        =   57
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   960
         TabIndex        =   56
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Impresora"
         Height          =   255
         Left            =   2160
         TabIndex        =   60
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Pantalla"
         Height          =   255
         Left            =   2160
         TabIndex        =   59
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta  Codigo"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   1215
      End
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
      Left            =   960
      TabIndex        =   51
      Top             =   4920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   5160
      TabIndex        =   50
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox valor10 
      Height          =   285
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   48
      Text            =   " "
      Top             =   4320
      Width           =   5055
   End
   Begin VB.TextBox valor9 
      Height          =   285
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   47
      Text            =   " "
      Top             =   3960
      Width           =   5055
   End
   Begin VB.TextBox valor8 
      Height          =   285
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   46
      Text            =   " "
      Top             =   3600
      Width           =   5055
   End
   Begin VB.TextBox valor7 
      Height          =   285
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   45
      Text            =   " "
      Top             =   3240
      Width           =   5055
   End
   Begin VB.TextBox valor6 
      Height          =   285
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   44
      Text            =   " "
      Top             =   2880
      Width           =   5055
   End
   Begin VB.TextBox valor5 
      Height          =   285
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   43
      Text            =   " "
      Top             =   2520
      Width           =   5055
   End
   Begin VB.TextBox valor4 
      Height          =   285
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   42
      Text            =   " "
      Top             =   2160
      Width           =   5055
   End
   Begin VB.TextBox Valor3 
      Height          =   285
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   41
      Text            =   " "
      Top             =   1800
      Width           =   5055
   End
   Begin VB.TextBox valor2 
      Height          =   285
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   40
      Text            =   " "
      Top             =   1440
      Width           =   5055
   End
   Begin VB.TextBox Valor1 
      Height          =   285
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   39
      Text            =   " "
      Top             =   1080
      Width           =   5055
   End
   Begin VB.TextBox Ensayo10 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   38
      Text            =   " "
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Ensayo9 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   37
      Text            =   " "
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox Ensayo8 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   36
      Text            =   " "
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Ensayo7 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   35
      Text            =   " "
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox Ensayo6 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      MaxLength       =   4
      TabIndex        =   34
      Text            =   " "
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Ensayo5 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   33
      Text            =   " "
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Ensayo4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   32
      Text            =   " "
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Ensayo3 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   31
      Text            =   " "
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Ensayo2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   30
      Text            =   " "
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Ensayo1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   29
      Text            =   " "
      Top             =   1080
      Width           =   735
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      ItemData        =   "espefici.frx":0000
      Left            =   0
      List            =   "espefici.frx":0007
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.CommandButton Listado 
      Caption         =   "Listado"
      Height          =   255
      Left            =   7920
      TabIndex        =   12
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   255
      Left            =   7920
      TabIndex        =   11
      Top             =   5880
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control"
      Height          =   1335
      Left            =   9120
      TabIndex        =   6
      Top             =   4800
      Width           =   1575
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   7920
      TabIndex        =   5
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   7920
      TabIndex        =   4
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   7920
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   7920
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   120
      TabIndex        =   61
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Descriprod 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   1080
      TabIndex        =   49
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Descri10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   1080
      TabIndex        =   28
      Top             =   4320
      Width           =   4860
   End
   Begin VB.Label Descri9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   1080
      TabIndex        =   27
      Top             =   3960
      Width           =   4860
   End
   Begin VB.Label Descri8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   1080
      TabIndex        =   26
      Top             =   3600
      Width           =   4860
   End
   Begin VB.Label Descri7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   1080
      TabIndex        =   25
      Top             =   3240
      Width           =   4860
   End
   Begin VB.Label Descri6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   1080
      TabIndex        =   24
      Top             =   2880
      Width           =   4860
   End
   Begin VB.Label Descri5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   1080
      TabIndex        =   23
      Top             =   2520
      Width           =   4860
   End
   Begin VB.Label Descri4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   1080
      TabIndex        =   22
      Top             =   2160
      Width           =   4860
   End
   Begin VB.Label Descri3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   1080
      TabIndex        =   21
      Top             =   1800
      Width           =   4860
   End
   Begin VB.Label descri2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   1080
      TabIndex        =   20
      Top             =   1440
      Width           =   4860
   End
   Begin VB.Label Descri1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   1080
      TabIndex        =   19
      Top             =   1080
      Width           =   4860
   End
   Begin VB.Label lblresultado 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor Standard"
      Height          =   255
      Left            =   6000
      TabIndex        =   18
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label lblDescri 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label lblensayo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ensayo"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   15
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "PrgEspecifi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstEspecificaciones As Recordset
Dim spEspecificaciones As String
Dim XParam As String
Dim EmpresaActual As String

Private Sub Imprime_Datos()

    spEspecificaciones = "ConsultaEspecificaciones " + "'" + Codigo.Text + "'"
    Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificaciones.RecordCount > 0 Then
        Codigo.Text = rstEspecificaciones!Producto
        Ensayo1.Text = rstEspecificaciones!Ensayo1
        Ensayo2.Text = rstEspecificaciones!Ensayo2
        Ensayo3.Text = rstEspecificaciones!Ensayo3
        Ensayo4.Text = rstEspecificaciones!Ensayo4
        Ensayo5.Text = rstEspecificaciones!Ensayo5
        Ensayo6.Text = rstEspecificaciones!Ensayo6
        Ensayo7.Text = rstEspecificaciones!Ensayo7
        Ensayo8.Text = rstEspecificaciones!Ensayo8
        Ensayo9.Text = rstEspecificaciones!Ensayo9
        Ensayo10.Text = rstEspecificaciones!Ensayo10
        Valor1.Text = rstEspecificaciones!Valor1
        valor2.Text = rstEspecificaciones!valor2
        Valor3.Text = rstEspecificaciones!Valor3
        valor4.Text = rstEspecificaciones!valor4
        valor5.Text = rstEspecificaciones!valor5
        valor6.Text = rstEspecificaciones!valor6
        valor7.Text = rstEspecificaciones!valor7
        valor8.Text = rstEspecificaciones!valor8
        valor9.Text = rstEspecificaciones!valor9
        valor10.Text = rstEspecificaciones!valor10
        
        rstEspecificaciones.Close
                        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri1.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri1.Caption = ""
        End If
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            descri2.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            descri2.Caption = ""
        End If
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri3.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri3.Caption = ""
        End If
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri4.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri4.Caption = ""
        End If
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri5.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri5.Caption = ""
        End If
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri6.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri6.Caption = ""
        End If
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri7.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri7.Caption = ""
        End If
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri8.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri8.Caption = ""
        End If
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri9.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri9.Caption = ""
        End If
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri10.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri10.Caption = ""
        End If
        
        spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Descriprod.Caption = rstArticulo!Descripcion
            rstArticulo.Close
        End If
    End If

End Sub

Private Sub Acepta_Click()
    
    lista.WindowTitle = "Listado de Ensayos"
    lista.WindowTop = 0
    lista.WindowLeft = 0
    lista.WindowWidth = Screen.Width
    lista.WindowHeight = Screen.Height

    lista.GroupSelectionFormula = "{Especificaciones.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If ImpreListado.Value = True Then
        lista.Destination = 1
            Else
        lista.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    lista.SQLQuery = "SELECT Especificaciones.Producto, Especificaciones.Ensayo1, Especificaciones.Valor1, Especificaciones.Ensayo2, Especificaciones.Valor2, Especificaciones.Ensayo3, Especificaciones.Valor3, Especificaciones.Ensayo4, Especificaciones.Valor4, Especificaciones.Ensayo5, Especificaciones.Valor5, Especificaciones.Ensayo6, Especificaciones.Valor6, Articulo.Descripcion " _
                     + "From " + DSQ + ".dbo.Especificaciones Especificaciones, " _
                     + DSQ + ".dbo.Articulo Articulo " _
                     + "Where Especificaciones.Producto = Articulo.Codigo AND Especificaciones.Producto >= ' ' AND Especificaciones.Producto <= 'ZZ-ZZZ-ZZZ'"
    
    lista.DataFiles(2) = WEmpresa + "auxi.mdb"
    lista.Connect = Connect()
    
    lista.Action = 1
    Frame2.Visible = False
    
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Codigo.Text <> "" Then
    
        spEspecificaciones = "ConsultaEspecificaciones " + "'" + Codigo.Text + "'"
        Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecificaciones.RecordCount > 0 Then
            rstEspecificaciones.Close
            WPasa = "S"
                Else
            WPasa = "N"
        End If
        
        If WPasa = "N" Then
            WProducto = Codigo.Text
            WEnsayo1 = Ensayo1.Text
            WEnsayo2 = Ensayo2.Text
            WEnsayo3 = Ensayo3.Text
            WEnsayo4 = Ensayo4.Text
            WEnsayo5 = Ensayo5.Text
            WEnsayo6 = Ensayo6.Text
            WEnsayo7 = Ensayo7.Text
            WEnsayo8 = Ensayo8.Text
            WEnsayo9 = Ensayo9.Text
            WEnsayo10 = Ensayo10.Text
            WValor1 = Valor1.Text
            WValor2 = valor2.Text
            WValor3 = Valor3.Text
            WValor4 = valor4.Text
            WValor5 = valor5.Text
            WValor6 = valor6.Text
            WValor7 = valor7.Text
            WValor8 = valor8.Text
            WValor9 = valor9.Text
            WValor10 = valor10.Text
            WDate = Date$
            XParam = "'" + WProducto + "','" _
                        + WEnsayo1 + "','" _
                        + WValor1 + "','" _
                        + WEnsayo2 + "','" _
                        + WValor2 + "','" _
                        + WEnsayo3 + "','" _
                        + WValor3 + "','" _
                        + WEnsayo4 + "','" _
                        + WValor4 + "','" _
                        + WEnsayo5 + "','" _
                        + WValor5 + "','" _
                        + WEnsayo6 + "','" _
                        + WValor6 + "','" _
                        + WEnsayo7 + "','" _
                        + WValor7 + "','" _
                        + WEnsayo8 + "','" _
                        + WValor8 + "','" _
                        + WEnsayo9 + "','" _
                        + WValor9 + "','" _
                        + WEnsayo10 + "','" _
                        + WValor10 + "','" _
                        + WDate + "'"
            Set rstEspecificaciones = db.OpenRecordset("AltaEspecificaciones " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                Else
            WProducto = Codigo.Text
            WEnsayo1 = Ensayo1.Text
            WEnsayo2 = Ensayo2.Text
            WEnsayo3 = Ensayo3.Text
            WEnsayo4 = Ensayo4.Text
            WEnsayo5 = Ensayo5.Text
            WEnsayo6 = Ensayo6.Text
            WEnsayo7 = Ensayo7.Text
            WEnsayo8 = Ensayo8.Text
            WEnsayo9 = Ensayo9.Text
            WEnsayo10 = Ensayo10.Text
            WValor1 = Valor1.Text
            WValor2 = valor2.Text
            WValor3 = Valor3.Text
            WValor4 = valor4.Text
            WValor5 = valor5.Text
            WValor6 = valor6.Text
            WValor7 = valor7.Text
            WValor8 = valor8.Text
            WValor9 = valor9.Text
            WValor10 = valor10.Text
            WDate = Date$
            XParam = "'" + WProducto + "','" _
                        + WEnsayo1 + "','" _
                        + WValor1 + "','" _
                        + WEnsayo2 + "','" _
                        + WValor2 + "','" _
                        + WEnsayo3 + "','" _
                        + WValor3 + "','" _
                        + WEnsayo4 + "','" _
                        + WValor4 + "','" _
                        + WEnsayo5 + "','" _
                        + WValor5 + "','" _
                        + WEnsayo6 + "','" _
                        + WValor6 + "','" _
                        + WEnsayo7 + "','" _
                        + WValor7 + "','" _
                        + WEnsayo8 + "','" _
                        + WValor8 + "','" _
                        + WEnsayo9 + "','" _
                        + WValor9 + "','" _
                        + WEnsayo10 + "','" _
                        + WValor10 + "','" _
                        + WDate + "'"
            Set rstEspecificaciones = db.OpenRecordset("ModificaEspecificaciones " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        End If
        Call CmdLimpiar_Click
        Codigo.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Codigo.Text <> "" Then
        spEspecificaciones = "ConsultaEspecificaciones " + "'" + Codigo.Text + "'"
        Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecificaciones.RecordCount > 0 Then
            rstEspecificaciones.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spEspecificaciones = "BorrarEspecificaciones " + "'" + Codigo.Text + "'"
                Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Codigo.Text = "  -   -   "
    Ensayo1.Text = ""
    Valor1.Text = ""
    Ensayo2.Text = ""
    valor2.Text = ""
    Ensayo3.Text = ""
    Valor3.Text = ""
    Ensayo4.Text = ""
    valor4.Text = ""
    Ensayo5.Text = ""
    valor5.Text = ""
    Ensayo6.Text = ""
    valor6.Text = ""
    Ensayo7.Text = ""
    valor7.Text = ""
    Ensayo8.Text = ""
    valor8.Text = ""
    Ensayo9.Text = ""
    valor9.Text = ""
    Ensayo10.Text = ""
    valor10.Text = ""
    Descriprod.Caption = ""
    Descri1.Caption = ""
    descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    Descri6.Caption = ""
    Descri7.Caption = ""
    Descri8.Caption = ""
    Descri9.Caption = ""
    Descri10.Caption = ""
    Codigo.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgEspecifi.Hide
    Unload Me
    Menu.SetFocus
End Sub

Private Sub Anterior_Click()
    spEspecificaciones = "AnteriorEspecificaciones " + "'" + Codigo.Text + "'"
    Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificaciones.RecordCount > 0 Then
        With rstEspecificaciones
            .MoveLast
            Codigo.Text = rstEspecificaciones!Producto
        End With
        rstEspecificaciones.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
End Sub

Private Sub Ensayo1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri1.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            Valor1.SetFocus
                    Else
            Descri1.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            descri2.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor2.SetFocus
                    Else
            descri2.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri3.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            Valor3.SetFocus
                    Else
            Descri3.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri4.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor4.SetFocus
                    Else
            Descri4.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri5.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor5.SetFocus
                    Else
            Descri5.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri6.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor6.SetFocus
                    Else
            Descri6.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo7_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri7.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor7.SetFocus
                    Else
            Descri7.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo8_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri8.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor8.SetFocus
                    Else
            Descri8.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo9_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri9.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor9.SetFocus
                    Else
            Descri9.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo10_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri10.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor10.SetFocus
                    Else
            Descri10.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Form_Activate()
    Select Case Val(EmpresaActual)
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
        Case 11
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgEspecifi.Caption = "Ingreso de Especificaciones de Materia Prima (Historico) :  " + !Nombre
        End If
    End With
    EmpresaActual = WEmpresa
End Sub

Private Sub Listado_Click()
    Desde.Text = "AA-000-000"
    Hasta.Text = "ZZ-999-999"
    ImprePantalla.Value = False
    ImpreListado.Value = True
    Frame2.Visible = True
End Sub



Private Sub Valor1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo2.SetFocus
    End If
End Sub
Private Sub Valor2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo3.SetFocus
    End If
End Sub
Private Sub Valor3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo4.SetFocus
    End If
End Sub
Private Sub Valor4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo5.SetFocus
    End If
End Sub
Private Sub Valor5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo6.SetFocus
    End If
End Sub
Private Sub Valor6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo7.SetFocus
    End If
End Sub
Private Sub Valor7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo8.SetFocus
    End If
End Sub
Private Sub Valor8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo9.SetFocus
    End If
End Sub
Private Sub Valor9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo10.SetFocus
    End If
End Sub
Private Sub Valor10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo1.SetFocus
    End If
End Sub

Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
            Codigo.Text = UCase(Codigo.Text)
            spEspecificaciones = "ConsultaEspecificaciones " + "'" + Codigo.Text + "'"
            Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecificaciones.RecordCount > 0 Then
                rstEspecificaciones.Close
                Call Imprime_Datos
                    Else
                WCodigo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WCodigo
            End If
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Descriprod.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                    Else
                Codigo.SetFocus
                Exit Sub
            End If
        End If
        Ensayo1.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()
    Opcion.Clear
    
    Opcion.AddItem "Codigos"
    Opcion.AddItem "Ensayos"
    
    Opcion.Visible = True
End Sub

Private Sub Opcion_Click()
    Opcion.Visible = False
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            spArticulo = "ListaArticuloConsulta"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstArticulo!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstArticulo.Close
            
            End If
            
        Case 1
            spEnsayo = "ListaEnsayos"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
            
            With rstEnsayo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstEnsayo!Codigo) + " " + rstEnsayo!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstEnsayo!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstEnsayo.Close
            
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            ClaveProd$ = WIndice.List(Indice)
            Codigo.Text = ClaveProd$
            spEspecificaciones = "ConsultaEspecificaciones " + "'" + Codigo.Text + "'"
            Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecificaciones.RecordCount > 0 Then
                rstEspecificaciones.Close
                Call Imprime_Datos
                    Else
                CmdLimpiar_Click
                Codigo.Text = ClaveProd$
                spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Descriprod.Caption = rstArticulo!Descripcion
                    rstArticulo.Close
                        Else
                    Codigo.SetFocus
                End If
            End If
            Codigo.SetFocus
            
        Case 1
            Entra$ = "S"
            If Val(Ensayo1.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo1.Text = Val(WIndice.List(Indice))
                    Valor1.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri1.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            
            If Val(Ensayo2.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo2.Text = Val(WIndice.List(Indice))
                    valor2.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        descri2.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            
            If Val(Ensayo3.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo3.Text = Val(WIndice.List(Indice))
                    Valor3.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri3.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            If Val(Ensayo4.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo4.Text = Val(WIndice.List(Indice))
                    valor4.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri4.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            If Val(Ensayo5.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo5.Text = Val(WIndice.List(Indice))
                    valor5.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri5.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            If Val(Ensayo6.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo6.Text = Val(WIndice.List(Indice))
                    valor6.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri6.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            If Val(Ensayo7.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo7.Text = Val(WIndice.List(Indice))
                    valor7.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri7.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            If Val(Ensayo8.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo8.Text = Val(WIndice.List(Indice))
                    valor8.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri8.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            If Val(Ensayo9.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo9.Text = Val(WIndice.List(Indice))
                    valor9.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri9.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            If Val(Ensayo10.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo10.Text = Val(WIndice.List(Indice))
                    valor10.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri10.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()
    spEspecificaciones = "ListaEspecificaciones"
    Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificaciones.RecordCount > 0 Then
        With rstEspecificaciones
            .MoveFirst
            Codigo.Text = rstEspecificaciones!Producto
        End With
        rstEspecificaciones.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
 End Sub

Private Sub Ultimo_Click()
    spEspecificaciones = "ListaEspecificaciones"
    Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificaciones.RecordCount > 0 Then
        With rstEspecificaciones
            .MoveLast
            Codigo.Text = rstEspecificaciones!Producto
        End With
        rstEspecificaciones.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If

 End Sub

Private Sub Siguiente_Click()

    spEspecificaciones = "PosteriorEspecificaciones " + "'" + Codigo.Text + "'"
    Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificaciones.RecordCount > 0 Then
        With rstEspecificaciones
            .MoveFirst
            Codigo.Text = rstEspecificaciones!Producto
        End With
        rstEspecificaciones.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
End Sub


