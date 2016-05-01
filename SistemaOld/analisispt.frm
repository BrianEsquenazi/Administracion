VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAnalisisPt 
   Caption         =   "Listado de Analisis de Productos Terminado"
   ClientHeight    =   4500
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4500
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   5535
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
         TabIndex        =   13
         Top             =   2520
         Width           =   2415
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2280
         TabIndex        =   1
         Top             =   960
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
         Left            =   2280
         TabIndex        =   0
         Top             =   600
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
         Top             =   3120
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
         TabIndex        =   9
         Top             =   3120
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
         Left            =   2280
         TabIndex        =   3
         Top             =   1920
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
         TabIndex        =   2
         Top             =   1560
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
         TabIndex        =   14
         Top             =   2520
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
         TabIndex        =   12
         Top             =   1560
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
         Left            =   360
         TabIndex        =   11
         Top             =   1920
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
         Left            =   360
         TabIndex        =   6
         Top             =   960
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
         Left            =   360
         TabIndex        =   5
         Top             =   600
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
Attribute VB_Name = "PrgAnalisisPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstMinimo As Recordset
Dim spMinimo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim XParam As String
Dim Empe(12, 10) As String
Dim Vector(50000, 7) As String
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
                                Vector(Suma, 5) = ""
                                Vector(Suma, 6) = ""
                                Vector(Suma, 7) = ""
                                
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
        
        
        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 8 Then
        
            For ZCiclo = 1 To Suma
            
                ZProducto = Vector(ZCiclo, 1)
                WVenta1 = ""
                WVenta2 = ""
                WVenta3 = ""
            
                For XCiclo = 1 To LugarMes
                    VentaMes(XCiclo, 2) = ""
                    VentaMes(XCiclo, 3) = ""
                Next XCiclo
                
                ZSql = ""
                ZSql = ZSql + "Select Estadistica.Articulo, Estadistica.OrdFecha, Estadistica.Cantidad, Estadistica.Tipo"
                ZSql = ZSql + " FROM Estadistica"
                ZSql = ZSql + " Where Estadistica.Articulo = " + "'" + ZProducto + "'"
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
                
                Vector(ZCiclo, 5) = VentaMes(LugarMes, 2)
                Vector(ZCiclo, 6) = Str$(VentaMayor)
                Vector(ZCiclo, 7) = Str$(SumaVenta / LugarMes)
                
            Next ZCiclo
        
        End If
        
        Call Conecta_Empresa
        
        For Ciclo = 1 To Suma
        
            Terminado = Vector(Ciclo, 1)
            Articulo = ""
            WCodigo = Terminado
            Descri1 = Left$(Vector(Ciclo, 2), 50)
            Stock = Vector(Ciclo, 3)
            Minimo = Vector(Ciclo, 4)
            Venta1 = Vector(Ciclo, 5)
            Venta2 = Vector(Ciclo, 6)
            Venta3 = Vector(Ciclo, 7)
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
                    ZVenta1 = Str$(rstMinimo!Venta1 + Val(Venta1))
                    ZVenta2 = Str$(rstMinimo!Venta2 + Val(Venta2))
                    ZVenta3 = Str$(rstMinimo!Venta3 + Val(Venta3))
           
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
                    
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Minimo SET "
                    ZSql = ZSql & "Stock1 = " + "'" + Stock1 + "',"
                    ZSql = ZSql & "Stock2 = " + "'" + Stock2 + "',"
                    ZSql = ZSql & "Stock3 = " + "'" + Stock3 + "',"
                    ZSql = ZSql & "Stock4 = " + "'" + Stock4 + "',"
                    ZSql = ZSql & "Stock5 = " + "'" + Stock5 + "',"
                    ZSql = ZSql & "Venta1 = " + "'" + ZVenta1 + "',"
                    ZSql = ZSql & "Venta2 = " + "'" + ZVenta2 + "',"
                    ZSql = ZSql & "Venta3 = " + "'" + ZVenta3 + "'"
                    ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                
                    spMinimo = ZSql
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
                    ZSql = ZSql + "Venta1 ,"
                    ZSql = ZSql + "Venta2 ,"
                    ZSql = ZSql + "Venta3 )"
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
                    ZSql = ZSql + "'" + Venta1 + "',"
                    ZSql = ZSql + "'" + Venta2 + "',"
                    ZSql = ZSql + "'" + Venta3 + "')"
        
                    spMinimo = ZSql
                    Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
        Next Ciclo
                
    Next a
    
    Call Conecta_Empresa

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
    
    Select Case Tipo.ListIndex
        Case 0
            Listado.ReportFileName = "WAnalisisPt.rpt"
            Listado.SQLQuery = "SELECT Minimo.Terminado, Minimo.Descripcion, Minimo.Stock1, Minimo.Stock2, Minimo.Stock3, Minimo.Stock4, Minimo.Stock5, Minimo.Minimo, Minimo.Dife, Minimo.Venta1, Minimo.Venta2, Minimo.Venta3 " _
                        + "From " _
                        + DSQ + ".dbo.Minimo Minimo " _
                        + "Where " _
                        + "Minimo.Dife < 0."
        Case Else
            Listado.ReportFileName = "WAnalisisPtII.rpt"
            Listado.SQLQuery = "SELECT Minimo.Terminado, Minimo.Descripcion, Minimo.Stock1, Minimo.Stock2, Minimo.Stock3, Minimo.Stock4, Minimo.Stock5, Minimo.Minimo, Minimo.Dife, Minimo.Venta1, Minimo.Venta2, Minimo.Venta3 " _
                        + "From " _
                        + DSQ + ".dbo.Minimo Minimo " _
                        + "Where " _
                        + "Minimo.Minimo <> 0."
    End Select
    
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
            Desde.SetFocus
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

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Debajo Minimo"
    Tipo.AddItem "Minimo Informado"
    
    Tipo.ListIndex = 0

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgAnalisisPt.Caption = "Listado de Analisis de Productos Terminado :  " + !Nombre
        End If
    End With
    
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
End Sub


