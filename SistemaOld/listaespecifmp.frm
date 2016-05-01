VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaEspecifMp 
   Caption         =   "Listado de Especificaciones de Materia Prima"
   ClientHeight    =   6555
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6555
   ScaleWidth      =   8145
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
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
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   6015
      Begin VB.TextBox DesdeMes 
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
         Left            =   2280
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox HastaMes 
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
         Left            =   2280
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Consulta 
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4560
         MaskColor       =   &H00000000&
         TabIndex        =   12
         Top             =   1680
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2280
         TabIndex        =   11
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   2280
         TabIndex        =   1
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   2760
         TabIndex        =   10
         Top             =   2400
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
         Left            =   1200
         TabIndex        =   9
         Top             =   2400
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
         Left            =   4560
         MaskColor       =   &H00000000&
         TabIndex        =   8
         Top             =   720
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
         Left            =   4560
         MaskColor       =   &H00000000&
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Desde Meses"
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
         Left            =   600
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta Meses"
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
         Left            =   600
         TabIndex        =   15
         Top             =   720
         Width           =   1215
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
         Left            =   600
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
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
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6840
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
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
      ItemData        =   "listaespecifmp.frx":0000
      Left            =   120
      List            =   "listaespecifmp.frx":0007
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "PrgListaEspecifMp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZVector(10000) As String
Dim XMes As String
Dim XAno As String

Private Sub Acepta_Click()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select


    ZTitulo = "de " + DesdeMes.Text + " a " + HastaMes.Text + " meses"
    If Val(WEmpresa) = 3 Then
        ZDesEmpresa = "Surfactan S.A."
            Else
        ZDesEmpresa = "Pellital S.A."
    End If

    Rem ZOrdDesdeFecha = Right$(DesdeFecha.Text, 4) + Mid$(DesdeFecha.Text, 4, 2) + Left$(DesdeFecha.Text, 2)
    Rem ZOrdHastaFecha = Right$(HastaFecha.Text, 4) + Mid$(HastaFecha.Text, 4, 2) + Left$(HastaFecha.Text, 2)
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    Erase ZVector
    ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "UPDATE EspecificacionesUnifica SET "
    ZSql = ZSql + " Marca = " + "'" + "" + "',"
    ZSql = ZSql + " DesEmpresa = " + "'" + ZDesEmpresa + "',"
    ZSql = ZSql + " Titulo = " + "'" + ZTitulo + "'"
    spEspecificacionesUnifica = ZSql
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EspecificacionesUnifica"
    spEspecificacionesUnifica = ZSql
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnifica.RecordCount > 0 Then
    
        With rstEspecificacionesUnifica
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                ZFecha = IIf(IsNull(!Fecha), "00/00/0000", !Fecha)
                ZOrdFecha = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
                WFechaActual = Left$(Date$, 2) + "/" + Right$(Date$, 4)
                
                Meses = 0
                WMes = Val(Mid$(ZFecha, 4, 2))
                WAno = Val(Right$(ZFecha, 4))
                Do
                    Meses = Meses + 1
                    WMes = WMes + 1
                    If WMes > 12 Then
                        WAno = WAno + 1
                        WMes = 1
                    End If
                    XMes = Str$(WMes)
                    XAno = Str$(WAno)
                    Call Ceros(XMes, 2)
                    Call Ceros(XAno, 4)
                    WCompara = XMes + "/" + XAno
                    If WCompara = WFechaActual Then
                        Exit Do
                    End If
                    If Meses > 1000 Then Exit Do
                Loop
            
                If Val(DesdeMes.Text) <= Meses And Meses <= Val(HastaMes.Text) Then
                    If Desde.Text <= !Producto And !Producto <= Hasta.Text Then
                        ZLugar = ZLugar + 1
                        ZVector(ZLugar) = !Producto
                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstEspecificacionesUnifica.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        WProducto = ZVector(Ciclo)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE EspecificacionesUnifica SET "
        ZSql = ZSql + " Marca = " + "'" + "S" + "'"
        ZSql = ZSql + " Where Producto = " + "'" + WProducto + "'"
        
        spEspecificacionesUnifica = ZSql
        Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
        
            
    
    
    
    
    
    

    Listado.WindowTitle = "Verificacion de Vencimientos de Materia Prima"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT EspecificacionesUnifica.Producto, EspecificacionesUnifica.Version, EspecificacionesUnifica.Fecha, EspecificacionesUnifica.Marca, EspecificacionesUnifica.Titulo, EspecificacionesUnifica.DesEmpresa, " _
                + "Articulo.Descripcion " _
                + "From " _
                + DSQ + ".dbo.EspecificacionesUnifica EspecificacionesUnifica, " _
                + DSQ + ".dbo.Articulo Articulo " _
                + "Where " _
                + "EspecificacionesUnifica.Producto = Articulo.Codigo AND " _
                + "EspecificacionesUnifica.Marca = 'S'"
    
    Listado.Connect = Connect()
    
    Rem Listado.GroupSelectionFormula = "{Articulo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.ReportFileName = "ListaEspecifMp.rpt"
    
    Listado.Action = 1
    
    Call Conecta_Empresa
    
End Sub

Private Sub Cancela_click()
    PrgListaEspecifMp.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub DesdeMes_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaMes.SetFocus
    End If
End Sub

Private Sub HastaMes_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    DesdeMes.Text = ""
    HastaMes.Text = ""
    
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    XIndice = 0
    
    Select Case XIndice
        Case 0
            spArticulo = "ListaArticulo"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Left$(rstArticulo!Codigo, 2) = "DY" Or Left$(rstArticulo!Codigo, 2) = "DW" Or Left$(rstArticulo!Codigo, 2) = "DS" Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstArticulo.Close
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus

End Sub


Private Sub pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    
    Select Case XIndice
        Case 0
            With rstArticulo
            
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                    Desde.Text = rstArticulo!Codigo
                    Hasta.Text = rstArticulo!Codigo
                End If
            End With
            Desde.SetFocus
        Case Else
    End Select
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spArticulo = "ListaArticuloConsulta"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
    
    With rstArticulo
        .MoveFirst
        Do
            If .EOF = False Then
            
                If Left$(rstArticulo!Codigo, 2) = "DY" Or Left$(rstArticulo!Codigo, 2) = "DW" Or Left$(rstArticulo!Codigo, 2) = "DS" Then
                    Da = Len(rstArticulo!Descripcion) - WEspacios
                    For Aaa = 1 To Da
                        If Left$(Ayuda.Text, WEspacios) = Mid$(rstArticulo!Descripcion, Aaa, WEspacios) Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                            Exit For
                        End If
                    Next Aaa
                End If
                .MoveNext
                    
                        Else
                        
                Exit Do
                
            End If
        Loop
    End With
    
    rstArticulo.Close
    
    End If
    
    End If

End Sub










