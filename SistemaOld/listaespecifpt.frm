VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaEspecifPt 
   Caption         =   "Listado de Especificaciones de Producto Terminado"
   ClientHeight    =   3600
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3600
   ScaleWidth      =   8145
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   6015
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
         TabIndex        =   13
         Top             =   840
         Width           =   735
      End
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
         Top             =   360
         Width           =   735
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2280
         TabIndex        =   10
         Top             =   1800
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   2280
         TabIndex        =   1
         Top             =   1320
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
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
         TabIndex        =   12
         Top             =   840
         Width           =   1215
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
         TabIndex        =   11
         Top             =   360
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaEspecifPt"
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
    ZSql = ZSql + "UPDATE EspecifUnifica SET "
    ZSql = ZSql + " Marca = " + "'" + "" + "',"
    ZSql = ZSql + " DesEmpresa = " + "'" + ZDesEmpresa + "',"
    ZSql = ZSql + " Titulo = " + "'" + ZTitulo + "'"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
    ZSql = ""
    ZSql = ZSql + "Select EspecifUnifica.Fecha, EspecifUnifica.Producto"
    ZSql = ZSql + " FROM EspecifUnifica"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
    
        With rstEspecifUnifica
    
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
            
                Rem If !Producto = "PT-06230-100" Then Stop
                
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
        rstEspecifUnifica.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        WProducto = ZVector(Ciclo)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE EspecifUnifica SET "
        ZSql = ZSql + " Marca = " + "'" + "S" + "'"
        ZSql = ZSql + " Where Producto = " + "'" + WProducto + "'"
        
        spEspecifUnifica = ZSql
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
        
            
    
    
    
    
    
    

    Listado.WindowTitle = "Verificacion de Vencimientos de Producto Terminado"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT EspecifUnifica.Producto, EspecifUnifica.Version, EspecifUnifica.Fecha, EspecifUnifica.Marca, EspecifUnifica.Titulo, EspecifUnifica.DesEmpresa, " _
                + "Terminado.Descripcion " _
                + "From " _
                + DSQ + ".dbo.EspecifUnifica EspecifUnifica, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "EspecifUnifica.Producto = Terminado.Codigo AND " _
                + "EspecifUnifica.Marca = 'S'"
    
    Listado.Connect = Connect()
    
    Rem Listado.GroupSelectionFormula = "{Articulo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.ReportFileName = "ListaEspecifPt.rpt"
    
    Listado.Action = 1
    
    Call Conecta_Empresa
    
    
End Sub

Private Sub Cancela_click()
    PrgListaEspecifPt.Hide
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
        DesdeMes.SetFocus
    End If
End Sub

Sub Form_Load()
    
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    DesdeMes.Text = ""
    HastaMes.Text = ""
    
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub









