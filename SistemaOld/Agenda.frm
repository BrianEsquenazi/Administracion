VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAgenda 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Agenda de Vencimientos"
   ClientHeight    =   3825
   ClientLeft      =   2790
   ClientTop       =   1320
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   3825
   ScaleWidth      =   5655
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox TipoII 
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
         Left            =   1800
         TabIndex        =   11
         Top             =   2040
         Width           =   2175
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
         Left            =   1800
         TabIndex        =   10
         Top             =   1560
         Width           =   2175
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
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
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   2520
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
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   2520
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
         Left            =   3480
         TabIndex        =   6
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
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo Letras"
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
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Listado"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
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
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
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
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4200
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wivacomp.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva Compras"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOrden As Recordset
Dim spOrden As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstAgenda As Recordset
Dim spAgenda As String

Dim CargaEmpresa(10, 2) As String
Dim XEmpresa As String
Dim ZVector(1000, 15) As String
Dim ZTipoPago As Integer

Dim XParam As String

Private Sub Acepta_Click()


    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    ZSql = "DELETE Agenda"
    spAgenda = ZSql
    Set rstAgenda = db.OpenRecordset(spAgenda, dbOpenSnapshot, dbSQLPassThrough)
    
    
    OPEN_FILE_Empresa
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WDesEmpresa = !Nombre
        End If
    End With
    WPeriodo = "Del " + Desde.Text + " al " + Hasta.Text
    
    XEmpresa = WEmpresa
        
    Select Case Val(XEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            CargaEmpresa(1, 1) = "0001"
            CargaEmpresa(1, 2) = "Empresa01"
            CargaEmpresa(2, 1) = "0003"
            CargaEmpresa(2, 2) = "Empresa03"
            CargaEmpresa(3, 1) = "0005"
            CargaEmpresa(3, 2) = "Empresa05"
            CargaEmpresa(4, 1) = "0006"
            CargaEmpresa(4, 2) = "Empresa06"
            CargaEmpresa(5, 1) = "0007"
            CargaEmpresa(5, 2) = "Empresa07"
            CargaEmpresa(6, 1) = "0010"
            CargaEmpresa(6, 2) = "Empresa10"
            CargaEmpresa(7, 1) = "0011"
            CargaEmpresa(7, 2) = "Empresa11"
            ZHasta = 7
                    
        Case Else
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
            ZHasta = 4
                    
    End Select
    
    Erase ZVector
    ZLugar = 0
                    
    If Tipo.ListIndex = 0 Or Tipo.ListIndex = 2 Then
                    
        For Cicla = 1 To ZHasta
            If CargaEmpresa(Cicla, 1) <> "" Then
                
                WEmpresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Orden"
                ZSql = ZSql + " Where Orden.PagoDespacho = 0"
                ZSql = ZSql + " and Orden.ImpoDespacho <> 0"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and Orden.Renglon = 1"
                ZSql = ZSql + " Order by Orden.Clave"
        
                spOrden = ZSql
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                If rstOrden.RecordCount > 0 Then
                    With rstOrden
                        .MoveFirst
                        Do
                            If .EOF = False Then
                
                                If rstOrden!OrdVtoDespacho <= WHasta Then
                
                                    ZLugar = ZLugar + 1
                    
                                    ZVector(ZLugar, 1) = rstOrden!VtoDespacho
                                    ZVector(ZLugar, 2) = Str$(rstOrden!Carpeta)
                                    ZVector(ZLugar, 3) = rstOrden!Proveedor
                                    ZVector(ZLugar, 4) = "Despacho"
                                    ZVector(ZLugar, 5) = Str$(rstOrden!ImpoDespacho)
                                    ZVector(ZLugar, 6) = "0"
                                    ZVector(ZLugar, 7) = rstOrden!Orden
                                    ZVector(ZLugar, 8) = rstOrden!OrdVtoDespacho
                                    ZVector(ZLugar, 9) = "0"
                                    ZVector(ZLugar, 10) = ""
                                    ZVector(ZLugar, 11) = ""
                                
                                End If
                            
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstOrden.Close
                End If
            End If
        Next Cicla
    
    End If
    
    If Tipo.ListIndex = 0 Or Tipo.ListIndex = 1 Then
    
        For Cicla = 1 To ZHasta
            If CargaEmpresa(Cicla, 1) <> "" Then
                
                WEmpresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Orden"
                ZSql = ZSql + " Where Orden.PagoLetra = 0"
                ZSql = ZSql + " and Orden.ImpoLetra <> 0"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and Orden.Renglon = 1"
                ZSql = ZSql + " Order by Orden.Clave"
    
                spOrden = ZSql
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                If rstOrden.RecordCount > 0 Then
                    With rstOrden
                        .MoveFirst
                        Do
                            If .EOF = False Then
                
                                ZZVtoLetra = IIf(IsNull(rstOrden!VtoLetra), "", rstOrden!VtoLetra)
                                ZZOrdVtoLetra = IIf(IsNull(rstOrden!OrdVtoLetra), "", rstOrden!OrdVtoLetra)
                                If ZZOrdVtoLetra <= WHasta Then
                                    
                                    ZTipoPago = IIf(IsNull(rstOrden!TipoPago), "0", rstOrden!TipoPago)
                                    If TipoII.ListIndex = 0 Or TipoII.ListIndex = ZTipoPago Then
                                
                                        ZLugar = ZLugar + 1
                    
                                        ZVector(ZLugar, 1) = ZZVtoLetra
                                        ZVector(ZLugar, 2) = Str$(rstOrden!Carpeta)
                                        ZVector(ZLugar, 3) = rstOrden!Proveedor
                                        ZVector(ZLugar, 4) = "Letra"
                                        ZVector(ZLugar, 5) = "0"
                                        ZVector(ZLugar, 6) = Str$(rstOrden!ImpoLetra)
                                        ZVector(ZLugar, 7) = rstOrden!Orden
                                        ZVector(ZLugar, 8) = ZZOrdVtoLetra
                                        ZVector(ZLugar, 9) = IIf(IsNull(rstOrden!TipoPago), "0", rstOrden!TipoPago)
                                        ZVector(ZLugar, 10) = IIf(IsNull(rstOrden!VtoLetraII), "", rstOrden!VtoLetraII)
                                        ZVector(ZLugar, 11) = IIf(IsNull(rstOrden!OrdVtoLetraII), "00000000", rstOrden!OrdVtoLetraII)
                                    
                                    End If
                        
                                End If
                    
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstOrden.Close
                End If
            End If
        Next Cicla
        
    End If
    
    Call Conecta_Empresa
        
    For Ciclo = 1 To ZLugar
    
        ZZVtoLetra = ZVector(Ciclo, 1)
        ZZCarpeta = ZVector(Ciclo, 2)

        ZZProveedor = ZVector(Ciclo, 3)
        ZZTipo = ZVector(Ciclo, 4)
        ZZImporteI = ZVector(Ciclo, 5)
        ZZImporteII = ZVector(Ciclo, 6)
        ZZOrden = ZVector(Ciclo, 7)
        ZZOrdVtoLetra = ZVector(Ciclo, 8)
        ZZTipoPago = ZVector(Ciclo, 9)
        zzdescripcion = ""
        ZZClave = ZZOrdVtoLetra
        ZZFechaII = ZVector(Ciclo, 10)
        ZZOrdFechaII = ZVector(Ciclo, 11)
        
        If ZZOrdVtoLetra < WDesde Then
            ZZClave = WDesde
                Else
            ZZClave = ZZOrdVtoLetra
        End If
        
        spProveedor = "ConsultaProveedores " + "'" + ZZProveedor + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            zzdescripcion = RstProveedor!Nombre
            RstProveedor.Close
        End If
        
        If Val(ZZImporteI) <> 0 Or Val(ZZImporteII) <> 0 Then

            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pagos"
            ZSql = ZSql + " Where Pagos.Carpeta = " + "'" + ZZCarpeta + "'"
            ZSql = ZSql + " Or Pagos.Carpeta1 = " + "'" + ZZCarpeta + "'"
            ZSql = ZSql + " Or Pagos.Carpeta2 = " + "'" + ZZCarpeta + "'"
            ZSql = ZSql + " Or Pagos.Carpeta3 = " + "'" + ZZCarpeta + "'"
            ZSql = ZSql + " Or Pagos.Carpeta4 = " + "'" + ZZCarpeta + "'"
            spPagos = ZSql
            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
            If rstPagos.RecordCount > 0 Then
                With rstPagos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            aa = rstPagos!Clave
                            ZZZProveedor = rstPagos!Proveedor
                            If ZZZProveedor = "10167878480" Or ZZZProveedor = "10000000100" Or ZZZProveedor = "10071081483" Then
                                ZZImporteI = "0"
                                        Else
                                ZZImporteII = "0"
                            End If
                            
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPagos.Close
            End If
            
        End If
        
        If Val(ZZImporteI) <> 0 Or Val(ZZImporteII) <> 0 Then
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Agenda ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Carpeta ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "TipoPago ,"
            ZSql = ZSql + "ImporteI ,"
            ZSql = ZSql + "ImporteII ,"
            ZSql = ZSql + "DesEmpresa ,"
            ZSql = ZSql + "Periodo ,"
            ZSql = ZSql + "OrdFechaII ,"
            ZSql = ZSql + "FechaII ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Orden )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + ZZOrdVtoLetra + "',"
            ZSql = ZSql + "'" + ZZVtoLetra + "',"
            ZSql = ZSql + "'" + ZZCarpeta + "',"
            ZSql = ZSql + "'" + Left(zzdescripcion, 50) + "',"
            ZSql = ZSql + "'" + ZZTipo + "',"
            ZSql = ZSql + "'" + ZZTipoPago + "',"
            ZSql = ZSql + "'" + ZZImporteI + "',"
            ZSql = ZSql + "'" + ZZImporteII + "',"
            ZSql = ZSql + "'" + WDesEmpresa + "',"
            ZSql = ZSql + "'" + WPeriodo + "',"
            ZSql = ZSql + "'" + ZZOrdFechaII + "',"
            ZSql = ZSql + "'" + ZZFechaII + "',"
            ZSql = ZSql + "'" + ZZProveedor + "',"
            ZSql = ZSql + "'" + ZZOrden + "')"
                
            spAgenda = ZSql
            Set rstAgenda = db.OpenRecordset(spAgenda, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
    
    
    
    
    
    Listado.WindowTitle = "Agenda de Vencimientos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    If Tipo.ListIndex = 0 Then
    
        Listado.SQLQuery = "SELECT Agenda.Clave, Agenda.Fecha, Agenda.Carpeta, Agenda.Descripcion, Agenda.Tipo, Agenda.ImporteI, Agenda.ImporteII, Agenda.Orden, Agenda.TipoPago " _
            + "From " _
            + DSQ + ".dbo.Agenda Agenda " _
            + "Where " _
            + "Agenda.Clave >= ' ' AND " _
            + "Agenda.Clave <= 'ZZZZZZZZZZ'"
    
        Listado.ReportFileName = "Agenda.rpt"
    
        Listado.GroupSelectionFormula = "{Agenda.Clave} in " + Chr$(34) + "" + Chr$(34) + " to " + Chr$(34) + "ZZZZZZZZZZ" + Chr$(34)
        Listado.SelectionFormula = "{Agenda.Clave} in " + Chr$(34) + "" + Chr$(34) + " to " + Chr$(34) + "ZZZZZZZZZZ" + Chr$(34)
        
            Else
            
        Listado.SQLQuery = "SELECT Agenda.Clave, Agenda.Fecha, Agenda.Carpeta, Agenda.Descripcion, Agenda.Tipo, Agenda.ImporteI, Agenda.ImporteII, Agenda.Orden, Agenda.TipoPago " _
            + "From " _
            + DSQ + ".dbo.Agenda Agenda " _
            + "Where " _
            + "Agenda.Clave >= ' ' AND " _
            + "Agenda.Clave <= 'ZZZZZZZZZZ'"
    
        Listado.ReportFileName = "AgendaLetra.rpt"
    
        Listado.GroupSelectionFormula = "{Agenda.Clave} in " + Chr$(34) + "" + Chr$(34) + " to " + Chr$(34) + "ZZZZZZZZZZ" + Chr$(34)
        Listado.SelectionFormula = "{Agenda.Clave} in " + Chr$(34) + "" + Chr$(34) + " to " + Chr$(34) + "ZZZZZZZZZZ" + Chr$(34)
            
    End If

    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.Connect = Connect()
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    PrgAgenda.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Sub Form_Load()


    Tipo.Clear
    
    Tipo.AddItem "Completo"
    Tipo.AddItem "Letras"
    Tipo.AddItem "Despacho"
    
    Tipo.ListIndex = 0
    
    
    TipoII.Clear
    
    TipoII.AddItem "Completo"
    TipoII.AddItem "Pago Anticipado"
    TipoII.AddItem "A la vista"
    TipoII.AddItem "Cuenta Corriente"
    
    TipoII.ListIndex = 0
    

    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Conecta_Empresa()

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
        Case Else
    End Select

End Sub










