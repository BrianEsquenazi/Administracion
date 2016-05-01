VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgOrdenAnual 
   Caption         =   "Listado de Ordenes de Compra Anuales"
   ClientHeight    =   3735
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3735
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      Begin MSMask.MaskEdBox Hastafecha 
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
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
      Begin MSMask.MaskEdBox Desdefecha 
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
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
         Left            =   2520
         TabIndex        =   7
         Top             =   2280
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
         Left            =   840
         TabIndex        =   6
         Top             =   2280
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
         Left            =   3360
         TabIndex        =   5
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
         Left            =   3360
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1680
         TabIndex        =   12
         Top             =   720
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
         Left            =   1680
         TabIndex        =   0
         Top             =   240
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
      Begin VB.Label Label4 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6600
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WOrdenAnual.rpt"
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
Attribute VB_Name = "PrgOrdenAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstOrdenAnual As Recordset
Dim spOrdenAnual As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim XParam As String
Dim Empe(100, 10) As String
Dim Vector(10000, 50) As String
Dim LugarVector As Integer
Dim WAno As String
Dim WMes As String
Dim WDia As String
Dim XMeses(12) As String
Dim CantiPedida(12) As Double

Private Sub Acepta_Click()

    WAno = Right$(DesdeFecha.Text, 4)
    WMes = Mid$(DesdeFecha.Text, 4, 2)
    WDia = Left$(DesdeFecha.Text, 2)
    WDesdeFecha = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHastaFecha = WAno + WMes + WDia

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    If WDesdeFecha > WHastaFecha Then
        Exit Sub
    End If
    
    Meses = 0
    
    WAno = Right$(DesdeFecha.Text, 4)
    WMes = Mid$(DesdeFecha.Text, 4, 2)
    WCompara = Left$(WHastaFecha, 6)

    Do
    
        Meses = Meses + 1
    
        WMes = Str$(Val(WMes) + 1)
        Call Ceros(WMes, 2)
        
        If Val(WMes) > 12 Then
            WMes = "01"
            WAno = Str$(Val(WAno) + 1)
            Call Ceros(WAno, 4)
        End If
        
        WCompara1 = WAno + WMes
        
        If WCompara1 > WCompara Then
            Exit Do
        End If
        
        If Meses = 13 Then
            Exit Do
        End If
    
    Loop
        
    If Meses > 12 Then
        Exit Sub
    End If
    
    XEmpresa = WEmpresa
    
    Erase Vector
    LugarVector = 0
    
    Sql1 = "DELETE OrdenAnual"
    spOrdenAnual = Sql1
    Set rstOrdenAnual = db.OpenRecordset(spOrdenAnual, dbOpenSnapshot, dbSQLPassThrough)
    
    Sql1 = "Select *"
    Sql2 = " FROM Articulo"
    Sql3 = " Where Articulo.Codigo >= " + "'" + Desde.Text + "'"
    Sql4 = " and Articulo.Codigo <= " + "'" + Hasta.Text + "'"
    spArticulo = Sql1 + Sql2 + Sql3 + Sql4
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    LugarVector = LugarVector + 1
                    
                    Vector(LugarVector, 1) = rstArticulo!Codigo
                    Vector(LugarVector, 2) = rstArticulo!Descripcion
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
    End If
    
    
    For CicloEmpresa = 1 To 2
            
        If CicloEmpresa = 1 Then
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Else
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    
        For Ciclo = 1 To LugarVector
    
            WFamilia = Mid$(Vector(Ciclo, 1), 4, 3)
            WArticulo = Vector(Ciclo, 1)
            WDescripcion = Vector(Ciclo, 2)
        
            Erase CantiPedida
        
            Sql1 = "Select *"
            Sql2 = " FROM Orden"
            Sql3 = " Where Orden.Articulo = " + "'" + WArticulo + "'"
            Sql4 = " and Orden.FechaOrd >= " + "'" + WDesdeFecha + "'"
            Sql5 = " and Orden.FechaOrd <= " + "'" + WHastaFecha + "'"
            spOrden = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
            
                With rstOrden
                    .MoveFirst
                    Do
                        If .EOF = False Then
                    
                            Meses = 0
    
                            WAno = Right$(DesdeFecha.Text, 4)
                            WMes = Mid$(DesdeFecha.Text, 4, 2)
                            WCompara = Left$(!FechaOrd, 6)

                            Do
    
                                Meses = Meses + 1
    
                                WMes = Str$(Val(WMes) + 1)
                                Call Ceros(WMes, 2)
            
                                If Val(WMes) > 12 Then
                                    WMes = "01"
                                    WAno = Str$(Val(WAno) + 1)
                                    Call Ceros(WAno, 4)
                                End If
            
                                WCompara1 = WAno + WMes
        
                                If WCompara1 > WCompara Then
                                    Exit Do
                                End If
        
                            Loop
                        
                            CantiPedida(Meses) = CantiPedida(Meses) + !Cantidad
                    
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstOrden.Close
            End If
        
            WMes = Mid$(DesdeFecha.Text, 4, 2)
            For XCiclo = 1 To 12
                XMes = Val(WMes)
                Select Case XCiclo
                    Case 1
                        WImpre1 = XMeses(XMes)
                    Case 2
                        WImpre2 = XMeses(XMes)
                    Case 3
                        WImpre3 = XMeses(XMes)
                    Case 4
                        WImpre4 = XMeses(XMes)
                    Case 5
                        WImpre5 = XMeses(XMes)
                    Case 6
                        WImpre6 = XMeses(XMes)
                    Case 7
                        WImpre7 = XMeses(XMes)
                    Case 8
                        WImpre8 = XMeses(XMes)
                    Case 9
                        WImpre9 = XMeses(XMes)
                    Case 10
                        WImpre10 = XMeses(XMes)
                    Case 11
                        WImpre11 = XMeses(XMes)
                    Case 12
                        WImpre12 = XMeses(XMes)
                    Case Else
                End Select
            
                WMes = Str$(Val(WMes) + 1)
                If Val(WMes) > 12 Then
                    WMes = "01"
                End If
                Call Ceros(WMes, 2)
            
            Next XCiclo
            
            WCanti1 = Str$(CantiPedida(1))
            WCanti2 = Str$(CantiPedida(2))
            WCanti3 = Str$(CantiPedida(3))
            WCanti4 = Str$(CantiPedida(4))
            WCanti5 = Str$(CantiPedida(5))
            WCanti6 = Str$(CantiPedida(6))
            WCanti7 = Str$(CantiPedida(7))
            WCanti8 = Str$(CantiPedida(8))
            WCanti9 = Str$(CantiPedida(9))
            WCanti10 = Str$(CantiPedida(10))
            WCanti11 = Str$(CantiPedida(11))
            WCanti12 = Str$(CantiPedida(12))
            
            Vector(Ciclo, 3) = WImpre1
            Vector(Ciclo, 4) = WImpre2
            Vector(Ciclo, 5) = WImpre3
            Vector(Ciclo, 6) = WImpre4
            Vector(Ciclo, 7) = WImpre5
            Vector(Ciclo, 8) = WImpre6
            Vector(Ciclo, 9) = WImpre7
            Vector(Ciclo, 10) = WImpre8
            Vector(Ciclo, 11) = WImpre9
            Vector(Ciclo, 12) = WImpre10
            Vector(Ciclo, 13) = WImpre11
            Vector(Ciclo, 14) = WImpre12
            
            Vector(Ciclo, 15) = Str$(Val(Vector(Ciclo, 15)) + CantiPedida(1))
            Vector(Ciclo, 16) = Str$(Val(Vector(Ciclo, 16)) + CantiPedida(2))
            Vector(Ciclo, 17) = Str$(Val(Vector(Ciclo, 17)) + CantiPedida(3))
            Vector(Ciclo, 18) = Str$(Val(Vector(Ciclo, 18)) + CantiPedida(4))
            Vector(Ciclo, 19) = Str$(Val(Vector(Ciclo, 19)) + CantiPedida(5))
            Vector(Ciclo, 20) = Str$(Val(Vector(Ciclo, 20)) + CantiPedida(6))
            Vector(Ciclo, 21) = Str$(Val(Vector(Ciclo, 21)) + CantiPedida(7))
            Vector(Ciclo, 22) = Str$(Val(Vector(Ciclo, 22)) + CantiPedida(8))
            Vector(Ciclo, 23) = Str$(Val(Vector(Ciclo, 23)) + CantiPedida(9))
            Vector(Ciclo, 24) = Str$(Val(Vector(Ciclo, 24)) + CantiPedida(10))
            Vector(Ciclo, 25) = Str$(Val(Vector(Ciclo, 25)) + CantiPedida(11))
            Vector(Ciclo, 26) = Str$(Val(Vector(Ciclo, 26)) + CantiPedida(12))
            
        Next Ciclo
        
    Next CicloEmpresa
    
    
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
    
    
            
    For Ciclo = 1 To LugarVector
            
        WFamilia = Mid$(Vector(Ciclo, 1), 4, 3)
        WArticulo = Vector(Ciclo, 1)
        WDescripcion = Vector(Ciclo, 2)
            
        WImpre1 = Vector(Ciclo, 3)
        WImpre2 = Vector(Ciclo, 4)
        WImpre3 = Vector(Ciclo, 5)
        WImpre4 = Vector(Ciclo, 6)
        WImpre5 = Vector(Ciclo, 7)
        WImpre6 = Vector(Ciclo, 8)
        WImpre7 = Vector(Ciclo, 9)
        WImpre8 = Vector(Ciclo, 10)
        WImpre9 = Vector(Ciclo, 11)
        WImpre10 = Vector(Ciclo, 12)
        WImpre11 = Vector(Ciclo, 13)
        WImpre12 = Vector(Ciclo, 14)
            
        WCanti1 = Vector(Ciclo, 15)
        WCanti2 = Vector(Ciclo, 16)
        WCanti3 = Vector(Ciclo, 17)
        WCanti4 = Vector(Ciclo, 18)
        WCanti5 = Vector(Ciclo, 19)
        WCanti6 = Vector(Ciclo, 20)
        WCanti7 = Vector(Ciclo, 21)
        WCanti8 = Vector(Ciclo, 22)
        WCanti9 = Vector(Ciclo, 23)
        WCanti10 = Vector(Ciclo, 24)
        WCanti11 = Vector(Ciclo, 25)
        WCanti12 = Vector(Ciclo, 26)
        
        Sql1 = "INSERT INTO OrdenAnual ("
        Sql2 = "Familia ,"
        Sql3 = "Codigo ,"
        Sql4 = "Descripcion ,"
        Sql5 = "Cantidad1 ,"
        Sql6 = "Cantidad2 ,"
        Sql7 = "Cantidad3 ,"
        Sql8 = "Cantidad4 ,"
        Sql9 = "Cantidad5 ,"
        Sql10 = "Cantidad6 ,"
        Sql11 = "Cantidad7 ,"
        Sql12 = "Cantidad8 ,"
        Sql13 = "Cantidad9 ,"
        Sql14 = "Cantidad10 ,"
        Sql15 = "Cantidad11 ,"
        Sql16 = "Cantidad12 ,"
        Sql17 = "Impre1 ,"
        Sql18 = "Impre2 ,"
        Sql19 = "Impre3 ,"
        Sql20 = "Impre4 ,"
        Sql21 = "Impre5 ,"
        Sql22 = "Impre6 ,"
        Sql23 = "Impre7 ,"
        Sql24 = "Impre8 ,"
        Sql25 = "Impre9 ,"
        Sql26 = "Impre10 ,"
        Sql27 = "Impre11 ,"
        Sql28 = "Impre12 )"
        Sql29 = "Values ("
        Sql30 = "'" + WFamilia + "',"
        Sql31 = "'" + WArticulo + "',"
        Sql32 = "'" + WDescripcion + "',"
        Sql33 = "'" + WCanti1 + "',"
        Sql34 = "'" + WCanti2 + "',"
        Sql35 = "'" + WCanti3 + "',"
        Sql36 = "'" + WCanti4 + "',"
        Sql37 = "'" + WCanti5 + "',"
        Sql38 = "'" + WCanti6 + "',"
        Sql39 = "'" + WCanti7 + "',"
        Sql40 = "'" + WCanti8 + "',"
        Sql41 = "'" + WCanti9 + "',"
        Sql42 = "'" + WCanti10 + "',"
        Sql43 = "'" + WCanti11 + "',"
        Sql44 = "'" + WCanti12 + "',"
        Sql45 = "'" + WImpre1 + "',"
        Sql46 = "'" + WImpre2 + "',"
        Sql47 = "'" + WImpre3 + "',"
        Sql48 = "'" + WImpre4 + "',"
        Sql49 = "'" + WImpre5 + "',"
        Sql50 = "'" + WImpre6 + "',"
        Sql51 = "'" + WImpre7 + "',"
        Sql52 = "'" + WImpre8 + "',"
        Sql53 = "'" + WImpre9 + "',"
        Sql54 = "'" + WImpre10 + "',"
        Sql55 = "'" + WImpre11 + "',"
        Sql56 = "'" + WImpre12 + "')"
        
        spOrdenAnual = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                       Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                       Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                       Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 + _
                       Sql41 + Sql42 + Sql43 + Sql44 + Sql45 + Sql46 + Sql47 + Sql48 + Sql49 + Sql50 + _
                       Sql51 + Sql52 + Sql53 + Sql54 + Sql55 + Sql56

        Set rstOrdenAnual = db.OpenRecordset(spOrdenAnual, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    Listado.WindowTitle = "Listado de Ordenes de Compra Anuales"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT OrdenAnual.Familia, OrdenAnual.Codigo, OrdenAnual.Descripcion, OrdenAnual.Cantidad1, OrdenAnual.Cantidad2, OrdenAnual.Cantidad3, OrdenAnual.Cantidad4, OrdenAnual.Cantidad5, OrdenAnual.Cantidad6, OrdenAnual.Cantidad7, OrdenAnual.Cantidad8, OrdenAnual.Cantidad9, OrdenAnual.Cantidad10, OrdenAnual.Cantidad11, OrdenAnual.Cantidad12, OrdenAnual.Impre1, OrdenAnual.Impre2, OrdenAnual.Impre3, OrdenAnual.Impre4, OrdenAnual.Impre5, OrdenAnual.Impre6, OrdenAnual.Impre7, OrdenAnual.Impre8, OrdenAnual.Impre9, OrdenAnual.Impre10, OrdenAnual.Impre11, OrdenAnual.Impre12 " _
                + "From " _
                + DSQ + ".dbo.OrdenAnual OrdenAnual " _
                + "Where " _
                + "OrdenAnual.Familia >= '000' AND " _
                + "OrdenAnual.Familia <= '999'"
    
    Rem Listado.GroupSelectionFormula = "{Orden.Articulo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Rem Listado.DataFiles(3) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgOrdenAnual.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        DesdeFecha.SetFocus
    End If
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
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFecha.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            HastaFecha.SetFocus
        End If
    End If
End Sub


Sub Form_Load()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgOrdenAnual.Caption = "Listado de Ordenes de Compra Anuales :  " + !Nombre
        End If
    End With
    
    XMeses(1) = " Enero"
    XMeses(2) = "Febrero"
    XMeses(3) = " Marzo"
    XMeses(4) = " Abril"
    XMeses(5) = " Mayo"
    XMeses(6) = " Junio"
    XMeses(7) = " Julio"
    XMeses(8) = " Agosto"
    XMeses(9) = "Septiem."
    XMeses(10) = "Octubre"
    XMeses(11) = "Noviem."
    XMeses(12) = "Diciem"
    
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
End Sub


