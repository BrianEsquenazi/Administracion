VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAnalisisPtOtro 
   Caption         =   "Listado de Analisis de Productos Terminado"
   ClientHeight    =   3810
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3810
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   5535
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2040
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   2040
         TabIndex        =   2
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2640
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   960
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
         Height          =   495
         Left            =   4080
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
         Height          =   495
         Left            =   4080
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   300
         Left            =   2040
         TabIndex        =   1
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   300
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label Label4 
         Caption         =   "Desde Fecha"
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
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta Fecha"
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
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Terminado"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Terminado"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   360
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WAnalisisPt.rpt"
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
End
Attribute VB_Name = "PrgAnalisisPtOtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstMinimo As Recordset
Dim spMinimo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim XParam As String
Dim Empe(10, 10) As String
Dim Vector(50000, 4) As String
Dim WCodigo As String
Dim Articulo As String
Dim Terminado As String
Dim Descripcion As String
Dim Stock1 As String
Dim Stock2 As String
Dim Stock3 As String
Dim Stock4 As String
Dim Stock5 As String
Dim Stock As String
Dim Minimo As String

Private Sub Acepta_Click()

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    spMinimo = "BorrarMinimo "
    Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)

    XEmpresa = WEmpresa
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        XHasta = 5
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        XHasta = 3
    End If
    
    For A = 1 To XHasta
    
        Erase Vector
        Suma = 0
    
        WEmpresa = Empe(A, 1)
        txtOdbc = Empe(A, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        XParam = "'" + Desde.Text + "','" _
                     + Hasta.Text + "'"
    
        spTerminado = "ListaTerminadoDesdeHastaMinimo " + XParam
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
    
            With rstTerminado
                .MoveFirst
                Do
                    If .EOF = False Then
    
                        If Desde.Text <= UCase(rstTerminado!Codigo) And Hasta.Text >= UCase(rstTerminado!Codigo) Then
                        
                            XSaldo = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                            
                            Rem If XSaldo <> 0 Then
                            If XSaldo <> 0 Or rstTerminado!Minimo <> 0 Then
                        
                                Suma = Suma + 1
                            
                                Vector(Suma, 1) = UCase(rstTerminado!Codigo)
                                Vector(Suma, 2) = rstTerminado!Descripcion
                                Vector(Suma, 3) = Str$(rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas)
                                Vector(Suma, 4) = Str$(rstTerminado!Minimo)
                                
                            End If
                            
                        End If
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
            rstTerminado.Close
        
        End If
        
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
            Case Else
        End Select
        
        For Ciclo = 1 To Suma
        
            Terminado = Vector(Ciclo, 1)
            Articulo = ""
            WCodigo = Terminado
            Descri1 = Left$(Vector(Ciclo, 2), 50)
            Stock = Vector(Ciclo, 3)
            Minimo = Vector(Ciclo, 4)
            Descripcion = ""
            
            For Saca1 = 1 To 50
                cara = Mid$(Descri1, Saca1, 1)
                Ingre = "S"
                If Mid$(Descri1, Saca1, 1) <> "" Then
                    If Asc(Mid$(Descri1, Saca1, 1)) = 39 Then
                        Ingre = "N"
                    End If
                End If
                If Ingre = "S" Then
                    Descripcion = Descripcion + cara
                End If
            Next Saca1
        
            spMinimo = "ConsultaMinimo " + "'" + Terminado + "'"
            Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
            If rstMinimo.RecordCount > 0 Then
            
                    Stock1 = Str$(rstMinimo!Stock1)
                    Stock2 = Str$(rstMinimo!Stock2)
                    Stock3 = Str$(rstMinimo!Stock3)
                    Stock4 = Str$(rstMinimo!Stock4)
                    Stock5 = Str$(rstMinimo!Stock5)
            
                    Select Case A
                        Case 1
                            Stock1 = Str$(Val(Stock1) + Val(Stock))
                        Case 2
                            Stock2 = Str$(Val(Stock2) + Val(Stock))
                        Case 3
                            Stock3 = Str$(Val(Stock3) + Val(Stock))
                        Case 4
                            Stock4 = Str$(Val(Stock4) + Val(Stock))
                        Case 5
                            Stock5 = Str$(Val(Stock5) + Val(Stock))
                        Case Else
                    End Select
                    
                    rstMinimo.Close
                
                    XParam = "'" + WCodigo + "','" _
                                + Stock1 + "','" _
                                + Stock2 + "','" _
                                + Stock3 + "','" _
                                + Stock4 + "','" _
                                + Stock5 + "'"
                                            
                    spMinimo = "ModificaMinimo " + XParam
                    Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
                    
                            Else
                
                    Stock1 = "0"
                    Stock2 = "0"
                    Stock3 = "0"
                    Stock4 = "0"
                    Stock5 = "0"
            
                    Select Case A
                        Case 1
                            Stock1 = Stock
                        Case 2
                            Stock2 = Stock
                        Case 3
                            Stock3 = Stock
                        Case 4
                            Stock4 = Stock
                        Case 5
                            Stock5 = Stock
                        Case Else
                    End Select
                
                    XParam = "'" + WCodigo + "','" _
                                + Articulo + "','" _
                                + Terminado + "','" _
                                + Descripcion + "','" _
                                + Stock1 + "','" _
                                + Stock2 + "','" _
                                + Stock3 + "','" _
                                + Stock4 + "','" _
                                + Stock5 + "','" _
                                + Minimo + "'"
                                            
                    spMinimo = "AltaMinimo " + XParam
                    Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
        Next Ciclo
                
    Next A
    
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
            Case Else
    End Select

    spMinimo = "ModificaMinimoDife "
    Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
    
    
    

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    Rem Listado.GroupSelectionFormula = "{Terminado.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Minimo.Articulo, Minimo.Terminado, Minimo.Descripcion, Minimo.Stock1, Minimo.Stock2, Minimo.Stock3, Minimo.Stock4, Minimo.Stock5, Minimo.Minimo, Minimo.Dife " _
                    + "From " _
                    + DSQ + ".dbo.Minimo Minimo " _
                    + "Where " _
                    + "Minimo.Dife < 0."
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    PrgAnalisisPt.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFecha.Text, Auxi)
        If Auxi = "S" Then
            Hastafecha.SetFocus
                Else
            DesdeFecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        DesdeFecha.Text = "  /  /    "
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFecha.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Hastafecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Hastafecha.Text = "  /  /    "
    End If
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgAnalisisPt.Caption = "Listado de Analisis de Productos Terminado :  " + !Nombre
        End If
    End With
    
    DesdeFecha.Text = "  /  /    "
    Hastafecha.Text = "  /  /    "
    
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
End Sub


