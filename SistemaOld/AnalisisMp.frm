VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAnalisisMp 
   Caption         =   "Analisis de Materias Primas"
   ClientHeight    =   7560
   ClientLeft      =   2100
   ClientTop       =   585
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   7560
   ScaleWidth      =   8145
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6360
      TabIndex        =   20
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      ItemData        =   "AnalisisMp.frx":0000
      Left            =   360
      List            =   "AnalisisMp.frx":0007
      TabIndex        =   19
      Top             =   4920
      Visible         =   0   'False
      Width           =   7455
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
      Left            =   360
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   7455
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
         Height          =   500
         Left            =   6000
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Proveedor 
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
         MaxLength       =   11
         TabIndex        =   15
         Text            =   " "
         Top             =   2160
         Width           =   1455
      End
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
         Left            =   2280
         TabIndex        =   14
         Top             =   2760
         Width           =   2415
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2280
         TabIndex        =   8
         Top             =   840
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
         TabIndex        =   0
         Top             =   480
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   3480
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
         Left            =   1200
         TabIndex        =   6
         Top             =   3480
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
         Left            =   4560
         TabIndex        =   5
         Top             =   600
         Width           =   1215
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
         Left            =   4560
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   300
         Left            =   2280
         TabIndex        =   9
         Top             =   1680
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
         Left            =   2280
         TabIndex        =   10
         Top             =   1320
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
      Begin VB.Label Label6 
         Caption         =   "Proveedor"
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
         Height          =   285
         Left            =   360
         TabIndex        =   17
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label DesProveedor 
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
         Left            =   3840
         TabIndex        =   16
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo de Listado"
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
         TabIndex        =   13
         Top             =   2760
         Width           =   1575
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
         Left            =   360
         TabIndex        =   12
         Top             =   1680
         Width           =   1575
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
         Left            =   360
         TabIndex        =   11
         Top             =   1320
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   840
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   -120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WAnalisisMp.rpt"
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
Attribute VB_Name = "PrgAnalisisMp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstMinimo As Recordset
Dim spMinimo As String
Dim rstAnalisisMp As Recordset
Dim spAnalisisMp As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim XParam As String
Dim Empe(12, 10) As String
Dim Vector(5000, 4) As String
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
Dim Carga(1000, 10) As String
Dim ZDatosI(1000, 10) As String
Dim ZDatosII(1000, 10) As String
Dim LugarI As Integer
Dim LugarII As Integer
Dim FABRICAMAYOR As Double

Private VentaMes(100, 3) As String
Private LugarMes As Integer
Private WDesdeFecha As String
Private WHastaFecha As String
Dim WImpre(6, 3) As String
Dim WMes1 As String
Dim WAno1 As String
Dim WMes2 As String
Dim WAno2 As String


Private Sub Acepta_Click()

    WTitulo = "Del " + DesdeFecha.Text + " al " + HastaFecha.Text

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    WDesdeFecha = Right$(DesdeFecha.Text, 4) + Mid$(DesdeFecha.Text, 4, 2) + Left$(DesdeFecha.Text, 2)
    WHastaFecha = Right$(HastaFecha.Text, 4) + Mid$(HastaFecha.Text, 4, 2) + Left$(HastaFecha.Text, 2)
    
    If WDesdeFecha > WHastaFecha Then
        Exit Sub
    End If
    
    WMes1 = Mid$(WDesdeFecha, 5, 2)
    WAno1 = Mid$(WDesdeFecha, 1, 4)
    
    WMes2 = Mid$(WHastaFecha, 5, 2)
    WAno2 = Mid$(WHastaFecha, 1, 4)
    
    Call Ceros(WMes1, 2)
    Call Ceros(WMes2, 2)
    Call Ceros(WAno1, 4)
    Call Ceros(WAno2, 4)
    
    Erase VentaMes
    LugarMes = 0
   
    Do
    
        LugarMes = LugarMes + 1
        VentaMes(LugarMes, 1) = WAno1 + WMes1
        
        WMes1 = Str$(Val(WMes1) + 1)
        If Val(WMes1) > 12 Then
            WAno1 = Str$(Val(WAno1) + 1)
            WMes1 = "1"
        End If
        
        Call Ceros(WMes1, 2)
        Call Ceros(WAno1, 4)
        
        If WAno1 + WMes1 > WAno2 + WMes2 Then
            Exit Do
        End If
        
    Loop
    
    spMinimo = "BorrarMinimo "
    Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)

    XEmpresa = WEmpresa
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
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
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        XHasta = 7
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        XHasta = 4
    End If
    
    For a = 1 To XHasta
    
        Erase Vector
        Suma = 0
    
        WEmpresa = Empe(a, 1)
        txtOdbc = Empe(a, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        ZSql = ""
        ZSql = ZSql + "Select Articulo.Codigo, Articulo.Descripcion, Articulo.Inicial, Salidas,Entradas, Articulo.Minimo, Articulo.Minimo1, Articulo.Proveedor"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Codigo >= " + "'" + Desde.Text + "'"
        ZSql = ZSql + " and Articulo.Codigo <= " + "'" + Hasta.Text + "'"
        If Tipo.ListIndex = 0 Or Tipo.ListIndex = 1 Then
            ZSql = ZSql + " and ( (Articulo.Entradas-Articulo.Salidas) <> 0 OR Articulo.Minimo <> 0)"
        End If
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
        
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        WProveedor = IIf(IsNull(rstArticulo!Proveedor), "", rstArticulo!Proveedor)
                        If Proveedor.Text = "" Or Proveedor.Text = WProveedor Then
                    
                            XSaldo = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                            
                            Suma = Suma + 1
                            
                            Vector(Suma, 1) = rstArticulo!Codigo
                            Vector(Suma, 2) = rstArticulo!Descripcion
                            Vector(Suma, 3) = Str$(rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas)
                            Vector(Suma, 4) = Str$(rstArticulo!Minimo)
                            
                        End If
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
            rstArticulo.Close
        
        End If
        
        Call Conecta_Empresa
        
        For Ciclo = 1 To Suma
        
            Articulo = Vector(Ciclo, 1)
            WCodigo = Articulo
            Terminado = ""
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
        
            spMinimo = "ConsultaMinimo " + "'" + Articulo + "'"
            Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
            If rstMinimo.RecordCount > 0 Then
            
                    Stock1 = Str$(rstMinimo!Stock1)
                    Stock2 = Str$(rstMinimo!Stock2)
                    Stock3 = Str$(rstMinimo!Stock3)
                    Stock4 = Str$(rstMinimo!Stock4)
                    Stock5 = Str$(rstMinimo!Stock5)
            
                    Select Case a
                        Case 1
                            Stock1 = Str$(Val(Stock1) + Val(Stock))
                        Case 2
                            Stock2 = Str$(Val(Stock2) + Val(Stock))
                        Case 3
                            Stock3 = Str$(Val(Stock3) + Val(Stock))
                        Case 4, 6, 7
                            Stock4 = Str$(Val(Stock4) + Val(Stock))
                        Case Else
                            Stock5 = Str$(Val(Stock5) + Val(Stock))
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
            
                    Select Case a
                        Case 1
                            Stock1 = Stock
                        Case 2
                            Stock2 = Stock
                        Case 3
                            Stock3 = Stock
                        Case 4, 6, 7
                            Stock4 = Stock
                        Case Else
                            Stock5 = Stock
                    End Select
                    
                    ZTipo = "MP"
                    
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO Minimo ("
                    ZSql = ZSql + "Codigo ,"
                    ZSql = ZSql + "Articulo ,"
                    ZSql = ZSql + "Terminado ,"
                    ZSql = ZSql + "Descripcion ,"
                    ZSql = ZSql + "Stock1 ,"
                    ZSql = ZSql + "Stock2 ,"
                    ZSql = ZSql + "Stock3 ,"
                    ZSql = ZSql + "Stock4 ,"
                    ZSql = ZSql + "Stock5 ,"
                    ZSql = ZSql + "Minimo ,"
                    ZSql = ZSql + "Tipo )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + WCodigo + "',"
                    ZSql = ZSql + "'" + Articulo + "',"
                    ZSql = ZSql + "'" + Terminado + "',"
                    ZSql = ZSql + "'" + Descripcion + "',"
                    ZSql = ZSql + "'" + Stock1 + "',"
                    ZSql = ZSql + "'" + Stock2 + "',"
                    ZSql = ZSql + "'" + Stock3 + "',"
                    ZSql = ZSql + "'" + Stock4 + "',"
                    ZSql = ZSql + "'" + Stock5 + "',"
                    ZSql = ZSql + "'" + Minimo + "',"
                    ZSql = ZSql + "'" + ZTipo + "')"
           
                    spMinimo = ZSql
                    Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
        Next Ciclo
                
    Next a
    
    
    
    For a = 1 To XHasta
    
        Erase Vector
        Suma = 0
    
        WEmpresa = Empe(a, 1)
        txtOdbc = Empe(a, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        ZSql = ""
        ZSql = ZSql + "Select Terminado.Codigo, Terminado.Descripcion, Terminado.Inicial, Terminado.Salidas, Terminado.Entradas, Terminado.Minimo, Terminado.Minimo1, Terminado.Pedido"
        ZSql = ZSql + " FROM Terminado"
        ZSql = ZSql + " Where Terminado.Codigo >= " + "'" + "PT-00000-000" + "'"
        ZSql = ZSql + " and Terminado.Codigo <= " + "'" + "PT-99999-999" + "'"
        ZSql = ZSql + " and (Terminado.Entradas-Terminado.Salidas) <> 0 "
        spTerminado = ZSql
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
    
            With rstTerminado
                .MoveFirst
                Do
                    If .EOF = False Then
    
                        XSaldo = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        
                        Suma = Suma + 1
                            
                        Vector(Suma, 1) = UCase(rstTerminado!Codigo)
                        Vector(Suma, 2) = rstTerminado!Descripcion
                        Vector(Suma, 3) = Str$(rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas)
                        Vector(Suma, 4) = Str$(rstTerminado!Minimo)
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
            rstTerminado.Close
        
        End If
        
        Call Conecta_Empresa
        
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
            
                    Select Case a
                        Case 1
                            Stock1 = Str$(Val(Stock1) + Val(Stock))
                        Case 2
                            Stock2 = Str$(Val(Stock2) + Val(Stock))
                        Case 3
                            Stock3 = Str$(Val(Stock3) + Val(Stock))
                        Case 4, 6, 7
                            Stock4 = Str$(Val(Stock4) + Val(Stock))
                        Case Else
                            Stock5 = Str$(Val(Stock5) + Val(Stock))
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
            
                    Select Case a
                        Case 1
                            Stock1 = Stock
                        Case 2
                            Stock2 = Stock
                        Case 3
                            Stock3 = Stock
                        Case 4, 6, 7
                            Stock4 = Stock
                        Case Else
                            Stock5 = Stock
                    End Select
                
                    ZTipo = "PT"
                    
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO Minimo ("
                    ZSql = ZSql + "Codigo ,"
                    ZSql = ZSql + "Articulo ,"
                    ZSql = ZSql + "Terminado ,"
                    ZSql = ZSql + "Descripcion ,"
                    ZSql = ZSql + "Stock1 ,"
                    ZSql = ZSql + "Stock2 ,"
                    ZSql = ZSql + "Stock3 ,"
                    ZSql = ZSql + "Stock4 ,"
                    ZSql = ZSql + "Stock5 ,"
                    ZSql = ZSql + "Minimo ,"
                    ZSql = ZSql + "Tipo )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + WCodigo + "',"
                    ZSql = ZSql + "'" + Articulo + "',"
                    ZSql = ZSql + "'" + Terminado + "',"
                    ZSql = ZSql + "'" + Descripcion + "',"
                    ZSql = ZSql + "'" + Stock1 + "',"
                    ZSql = ZSql + "'" + Stock2 + "',"
                    ZSql = ZSql + "'" + Stock3 + "',"
                    ZSql = ZSql + "'" + Stock4 + "',"
                    ZSql = ZSql + "'" + Stock5 + "',"
                    ZSql = ZSql + "'" + Minimo + "',"
                    ZSql = ZSql + "'" + ZTipo + "')"
           
                    spMinimo = ZSql
                    Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
        Next Ciclo
                
    Next a
    
    Call Conecta_Empresa
    
    ZSql = "DELETE AnalisisMp"
    spAnalisisMp = ZSql
    Set rstAnalisisMp = db.OpenRecordset(spAnalisisMp, dbOpenSnapshot, dbSQLPassThrough)
    
    Erase Vector
    LugarVector = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Minimo"
    ZSql = ZSql + " Where Tipo = " + "'" + "MP" + "'"
            
    spMinimo = ZSql
    Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
    If rstMinimo.RecordCount > 0 Then
    
        With rstMinimo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZStock = rstMinimo!Stock1 + rstMinimo!Stock2 + rstMinimo!Stock3 + rstMinimo!Stock4 + rstMinimo!Stock5
                    ZMinimo = rstMinimo!Minimo
                    
                    Pasa = "N"
                    
                    Select Case Tipo.ListIndex
                        Case 0
                            If ZMinimo > ZStock Then
                                Pasa = "S"
                            End If
                        Case 1
                            If ZMinimo <> 0 Then
                                Pasa = "S"
                            End If
                        Case Else
                            Pasa = "S"
                    End Select
                            
                    If Pasa = "S" Then
                        LugarVector = LugarVector + 1
                        Vector(LugarVector, 1) = rstMinimo!Codigo
                        Vector(LugarVector, 2) = rstMinimo!Descripcion
                        Vector(LugarVector, 3) = Str$(rstMinimo!Stock1 + rstMinimo!Stock2 + rstMinimo!Stock3 + rstMinimo!Stock4 + rstMinimo!Stock5)
                        Vector(LugarVector, 4) = Str$(rstMinimo!Minimo)
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
            
        rstMinimo.Close
        
    End If
    
    For Ciclo = 1 To LugarVector
    
        WCodigo = Left$(Vector(Ciclo, 1), 10)
        WDescripcion = Vector(Ciclo, 2)
        WStock = Vector(Ciclo, 3)
        WMinimo = Vector(Ciclo, 4)
        
        LugarI = 0
        LugarII = 0
        
        Erase ZDatosI
        Erase ZDatosII
        
        ZSql = ""
        ZSql = ZSql + "Select Composicion.Articulo1, Composicion.Terminado, Composicion.Cantidad "
        ZSql = ZSql + " FROM Composicion"
        ZSql = ZSql + " Where Composicion.Articulo1 = " + "'" + WCodigo + "'"
        spComposicion = ZSql
        Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
        If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
            
                        LugarI = LugarI + 1
                        ZDatosI(LugarI, 1) = rstComposicion!Terminado
                        ZDatosI(LugarI, 2) = Str(rstComposicion!Cantidad)
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstComposicion.Close
        End If
        
        XEmpresa = WEmpresa
        
        If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
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
            Empe(6, 1) = "0010"
            Empe(6, 2) = "Empresa10"
            Empe(7, 1) = "0011"
            Empe(7, 2) = "Empresa11"
            XHasta = 7
                Else
            Empe(1, 1) = "0002"
            Empe(1, 2) = "Empresa02"
            Empe(2, 1) = "0004"
            Empe(2, 2) = "Empresa04"
            Empe(3, 1) = "0008"
            Empe(3, 2) = "Empresa08"
            Empe(4, 1) = "0009"
            Empe(4, 2) = "Empresa09"
            XHasta = 4
        End If
    
        For a = 1 To XHasta
        
            WEmpresa = Empe(a, 1)
            txtOdbc = Empe(a, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            Erase Carga
            LugarCarga = 0
        
            ZSql = ""
            ZSql = ZSql + "Select Orden.Articulo, Orden.FechaOrd, Orden.Clave, Orden.Orden, Orden.Articulo, Orden.Cantidad, Orden.Proveedor, Orden.Fecha1, Orden.Recibida "
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden.Articulo = " + "'" + WCodigo + "'"
            ZSql = ZSql + " and Orden.FechaOrd > " + "'" + "20050101" + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                With rstOrden
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            ZSaldo = rstOrden!Cantidad - rstOrden!Recibida
                            If (rstOrden!Orden < 900000) Then
                            Rem If (rstOrden!Orden < 900000) Or (rstOrden!Orden > 900000 And ZSaldo <> 0) Then
                                LugarCarga = LugarCarga + 1
                                Carga(LugarCarga, 1) = rstOrden!Clave
                                Carga(LugarCarga, 2) = rstOrden!Orden
                                Carga(LugarCarga, 3) = rstOrden!Articulo
                                Carga(LugarCarga, 4) = Str$(rstOrden!Cantidad)
                                Carga(LugarCarga, 5) = rstOrden!Proveedor
                                Carga(LugarCarga, 6) = rstOrden!Fecha1
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstOrden.Close
            End If
            
        
            For XCiclo = 1 To LugarCarga
    
                WClave = Carga(XCiclo, 1)
                WOrden = Carga(XCiclo, 2)
                WArticulo = Carga(XCiclo, 3)
                WCantidad = Val(Carga(XCiclo, 4))
                WProveedor = Carga(XCiclo, 5)
                WFechaEntrega = Carga(XCiclo, 6)
                WResta = 0
            
                XParam = "'" + WOrden + "','" _
                             + WArticulo + "'"
                spInforme = "ListaInformeOrdenArticulo" + XParam
                Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                If rstInforme.RecordCount > 0 Then
                    With rstInforme
                        .MoveFirst
                        If .NoMatch = False Then
                            Do
                                If .EOF = True Then
                                    Exit Do
                                End If
                                WResta = WResta + rstInforme!Resta
                                .MoveNext
                                If .EOF = True Then
                                    Exit Do
                                End If
                            Loop
                        End If
                    End With
                    rstInforme.Close
                End If
            
                WDife = WCantidad - WResta
                If WDife > 0 Then
                    LugarII = LugarII + 1
                    
                    spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    If RstProveedor.RecordCount > 0 Then
                        WProveedor = RstProveedor!Nombre
                        RstProveedor.Close
                    End If
                    
                    
                    ZDatosII(LugarII, 1) = WProveedor
                    ZDatosII(LugarII, 2) = Str$(WDife)
                    ZDatosII(LugarII, 3) = WFechaEntrega
                    ZDatosII(LugarII, 4) = WOrden
                    
                End If
        
            Next XCiclo
            
        Next a
        
        Call Conecta_Empresa
        
        CicloHasta = LugarI
        
        If LugarII > CicloHasta Then
            CicloHasta = LugarII
        End If
        
        If CicloHasta = 0 Then
            CicloHasta = 1
        End If
        
        For ZCiclo = 1 To CicloHasta
        
            WProducto = ZDatosI(ZCiclo, 1)
            WPorce = ZDatosI(ZCiclo, 2)
            WStockProducto = ""
            WVenta1 = ""
            WVenta2 = ""
            WVenta3 = ""
            WFabrica1 = ""
            WFabrica2 = ""
            
            WProveedor = ZDatosII(ZCiclo, 1)
            WCantidadOrden = ZDatosII(ZCiclo, 2)
            WFechaEntrega = ZDatosII(ZCiclo, 3)
            WOrden = ZDatosII(ZCiclo, 4)
            
            If WProducto <> "" Then
                        
                spMinimo = "ConsultaMinimo " + "'" + WProducto + "'"
                Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
                If rstMinimo.RecordCount > 0 Then
                    WStockProducto = Str$(rstMinimo!Stock1 + rstMinimo!Stock2 + rstMinimo!Stock3 + rstMinimo!Stock4 + rstMinimo!Stock5)
                    rstMinimo.Close
                End If
                
                For XCiclo = 1 To LugarMes
                    VentaMes(XCiclo, 2) = ""
                    VentaMes(XCiclo, 3) = ""
                Next XCiclo
                
                ZSql = ""
                ZSql = ZSql + "Select Estadistica.Articulo, Estadistica.OrdFecha, Estadistica.Cantidad, Estadistica.Tipo"
                ZSql = ZSql + " FROM Estadistica"
                ZSql = ZSql + " Where Estadistica.Articulo = " + "'" + WProducto + "'"
                ZSql = ZSql + " and Estadistica.OrdFecha >= " + "'" + WDesdeFecha + "'"
                ZSql = ZSql + " and Estadistica.OrdFecha <= " + "'" + WHastaFecha + "'"
                spEstadistica = ZSql
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEstadistica.RecordCount > 0 Then
                    With rstEstadistica
                        .MoveFirst
                        Do
                            If .EOF = False Then
                            
                                WCantidad = rstEstadistica!Cantidad
                                WTipo = rstEstadistica!Tipo
                    
                                WCompara = Left$(rstEstadistica!OrdFecha, 6)
                                For XCiclo = 1 To LugarMes
                                    If VentaMes(XCiclo, 1) = WCompara Then
                                        If WTipo = 1 Then
                                            VentaMes(XCiclo, 2) = Str$(Val(VentaMes(XCiclo, 2)) + WCantidad)
                                                Else
                                            VentaMes(XCiclo, 2) = Str$(Val(VentaMes(XCiclo, 2)) - WCantidad)
                                        End If
                                        Exit For
                                    End If
                                Next XCiclo
                                
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstEstadistica.Close
                End If
                
                SumaVenta = 0
                VentaMayor = 0
                For XCiclo = 1 To LugarMes
                    SumaVenta = SumaVenta + Val(VentaMes(XCiclo, 2))
                    If Val(VentaMes(XCiclo, 2)) > VentaMayor Then
                        VentaMayor = VentaMes(XCiclo, 2)
                    End If
                Next XCiclo
                
                WVenta1 = VentaMes(LugarMes, 2)
                WVenta2 = Str$(VentaMayor)
                WVenta3 = Str$(SumaVenta / LugarMes)
                
                
                XEmpresa = WEmpresa
                
                If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
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
                    Empe(6, 1) = "0010"
                    Empe(6, 2) = "Empresa10"
                    Empe(7, 1) = "0011"
                    Empe(7, 2) = "Empresa11"
                    XHasta = 7
                        Else
                    Empe(1, 1) = "0002"
                    Empe(1, 2) = "Empresa02"
                    Empe(2, 1) = "0004"
                    Empe(2, 2) = "Empresa04"
                    Empe(3, 1) = "0008"
                    Empe(3, 2) = "Empresa08"
                    Empe(4, 1) = "0009"
                    Empe(4, 2) = "Empresa09"
                    XHasta = 4
                End If
    
                For a = 1 To XHasta
        
                    WEmpresa = Empe(a, 1)
                    txtOdbc = Empe(a, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    ZSql = ""
                    ZSql = ZSql + "Select Hoja.Producto, Hoja.FechaOrd, Hoja.Real, Hoja.Renglon"
                    ZSql = ZSql + " FROM Hoja"
                    ZSql = ZSql + " Where Hoja.Producto = " + "'" + WProducto + "'"
                    ZSql = ZSql + " and Hoja.FechaOrd >= " + "'" + WDesdeFecha + "'"
                    ZSql = ZSql + " and Hoja.FechaOrd <= " + "'" + WHastaFecha + "'"
                    ZSql = ZSql + " and Hoja.Renglon = 1"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHoja.RecordCount > 0 Then
                        With rstHoja
                            .MoveFirst
                            Do
                                If .EOF = False Then
                            
                                    WCantidad = rstHoja!Real
                    
                                    WCompara = Left$(rstHoja!FechaOrd, 6)
                                    For XCiclo = 1 To LugarMes
                                        If VentaMes(XCiclo, 1) = WCompara Then
                                            VentaMes(XCiclo, 3) = Str$(Val(VentaMes(XCiclo, 3)) + WCantidad)
                                            Exit For
                                        End If
                                    Next XCiclo
                                
                                        .MoveNext
                                        Else
                                    Exit Do
                                End If
                            Loop
                        End With
                        rstHoja.Close
                    End If
                
                Next a
                
                Call Conecta_Empresa
                
                SumaFabrica = 0
                FABRICAMAYOR = 0
                For XCiclo = 1 To LugarMes
                    SumaFabrica = SumaFabrica + Val(VentaMes(XCiclo, 3))
                    If Val(VentaMes(XCiclo, 3)) > Val(FABRICAMAYOR) Then
                        FABRICAMAYOR = Val(VentaMes(XCiclo, 3))
                    End If
                Next XCiclo
                
                WFabrica1 = Str$(FABRICAMAYOR)
                WFabrica2 = Str$(SumaFabrica / LugarMes)
            
            End If
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO AnalisisMp ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Stock ,"
            ZSql = ZSql + "Minimo ,"
            ZSql = ZSql + "Producto ,"
            ZSql = ZSql + "StockProducto ,"
            ZSql = ZSql + "Venta1 ,"
            ZSql = ZSql + "Venta2 ,"
            ZSql = ZSql + "Venta3 ,"
            ZSql = ZSql + "Fabrica1 ,"
            ZSql = ZSql + "Fabrica2 ,"
            ZSql = ZSql + "CantidadOrden ,"
            ZSql = ZSql + "FechaEntrega ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Porcentaje ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "Orden )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "',"
            ZSql = ZSql + "'" + WStock + "',"
            ZSql = ZSql + "'" + WMinimo + "',"
            ZSql = ZSql + "'" + WProducto + "',"
            ZSql = ZSql + "'" + WStockProducto + "',"
            ZSql = ZSql + "'" + WVenta1 + "',"
            ZSql = ZSql + "'" + WVenta2 + "',"
            ZSql = ZSql + "'" + WVenta3 + "',"
            ZSql = ZSql + "'" + WFabrica1 + "',"
            ZSql = ZSql + "'" + WFabrica2 + "',"
            ZSql = ZSql + "'" + WCantidadOrden + "',"
            ZSql = ZSql + "'" + WFechaEntrega + "',"
            ZSql = ZSql + "'" + WProveedor + "',"
            ZSql = ZSql + "'" + WPorce + "',"
            ZSql = ZSql + "'" + WTitulo + "',"
            ZSql = ZSql + "'" + WOrden + "')"
           
            spAnalisisMp = ZSql
            Set rstAnalisisMp = db.OpenRecordset(spAnalisisMp, dbOpenSnapshot, dbSQLPassThrough)
        
        Next ZCiclo
    
    Next Ciclo
    
    Listado.ReportFileName = "WAnalisisMp.rpt"

    Listado.WindowTitle = "Listado de Stock de Materias Primas inferior al minimo (Consolidado)"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{Articulo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT AnalisisMp.Codigo, AnalisisMp.Descripcion, AnalisisMp.Stock, AnalisisMp.Minimo, AnalisisMp.Producto, AnalisisMp.StockProducto, AnalisisMp.Venta1, AnalisisMp.Venta2, AnalisisMp.Venta3, AnalisisMp.Fabrica1, AnalisisMp.Fabrica2, AnalisisMp.CantidadOrden, AnalisisMp.FechaEntrega, AnalisisMp.Proveedor, AnalisisMp.Porcentaje, AnalisisMp.Titulo, AnalisisMp.Orden " _
                    + "From " _
                    + DSQ + ".dbo.AnalisisMp AnalisisMp " _
                    + "Where " _
                    + "AnalisisMp.Codigo >= ' ' AND " _
                    + "AnalisisMp.Codigo <= 'ZZZZZZZZZZ'"
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    
    Desde.SetFocus
    PrgAnalisisMp.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()
    Call aYUDA_Keypress(13)
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFecha.Text, Auxi)
        If Auxi = "S" Then
            HastaFecha.SetFocus
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
            Proveedor.SetFocus
                Else
            HastaFecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFecha.Text = "  /  /    "
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
        DesdeFecha.SetFocus
    End If
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spProveedor = "Consultaproveedores " + "'" + Proveedor.Text + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            Proveedor.Text = RstProveedor!Proveedor
            DesProveedor.Caption = RstProveedor!Nombre
            RstProveedor.Close
            Desde.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Debajo Minimo"
    Tipo.AddItem "Minimo Informado"
    Tipo.AddItem "Todos"
    
    Tipo.ListIndex = 0

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgAnalisisMp.Caption = "Analisis de Materias Primas :  " + !Nombre
        End If
    End With
    
    Proveedor.Text = ""
    
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub


Private Sub aYUDA_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        Pantalla.Clear
        WIndice.Clear
        
        WEspacios = Len(Ayuda.Text)
    
        XEmpresa = WEmpresa
        
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        spProveedor = "ListaProveedoresOrdConsultaII " + "'" + Ayuda.Text + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
        If RstProveedor.RecordCount > 0 Then
            With RstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
                
                    Da = Len(RstProveedor!Nombre) - WEspacios
                        For aa = 1 To Da
                            If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                                Auxi = Str$(RstProveedor!Proveedor)
                                Call Ceros(Auxi, 11)
                                IngresaItem = Auxi + "    " + RstProveedor!Nombre
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Proveedor
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
            RstProveedor.Close
        End If
        
        Ayuda.Visible = True
        Pantalla.Visible = True
        
        Call Conecta_Empresa
        
    End If

End Sub

Private Sub pantalla_Click()

    Ayuda.Visible = False
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Proveedor.Text = WIndice.List(Indice)
    Call Proveedor_KeyPress(13)
    
End Sub

