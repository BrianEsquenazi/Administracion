VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCalificaEnvase 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Evaluacion Semestral Actual de Proveedores"
   ClientHeight    =   3570
   ClientLeft      =   2025
   ClientTop       =   1050
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3570
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   4575
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
         TabIndex        =   8
         Top             =   2520
         Width           =   1215
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
         TabIndex        =   7
         Top             =   2520
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2280
         TabIndex        =   6
         Top             =   840
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   2280
         TabIndex        =   0
         Top             =   480
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
         Left            =   2400
         TabIndex        =   5
         Top             =   1920
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
         Left            =   840
         TabIndex        =   4
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7080
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WCalificaEnvase.rpt"
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
Attribute VB_Name = "PrgCalificaEnvase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WTerminado As String
Private WInicial As Double
Private WEntrada As Double
Private WSalida As Double
Private WTipo As Integer
Private WWtipo As Integer
Private WNumero As String
Private Impre1 As String
Private Impre2 As String
Private WFecha As String
Dim WVector(1000, 8) As String
Dim WDevuelta As String
Dim WLiberada As String
Dim WPartida1 As String
Dim WPartida2 As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim ZProveedor(5000, 10) As String
Dim CargaEmpresa(10, 2) As String

Private Sub Acepta_Click()

    XEmpresa = WEmpresa
        
    CargaEmpresa(1, 1) = "0001"
    CargaEmpresa(1, 2) = "Empresa01"
    CargaEmpresa(2, 1) = "0002"
    CargaEmpresa(2, 2) = "Empresa02"
    CargaEmpresa(3, 1) = "0003"
    CargaEmpresa(3, 2) = "Empresa03"
    CargaEmpresa(4, 1) = "0004"
    CargaEmpresa(4, 2) = "Empresa04"
    CargaEmpresa(5, 1) = "0005"
    CargaEmpresa(5, 2) = "Empresa05"
    CargaEmpresa(6, 1) = "0006"
    CargaEmpresa(6, 2) = "Empresa06"
    CargaEmpresa(7, 1) = "0007"
    CargaEmpresa(7, 2) = "Empresa07"
    CargaEmpresa(8, 1) = "0008"
    CargaEmpresa(8, 2) = "Empresa08"
    CargaEmpresa(9, 1) = "0009"
    CargaEmpresa(9, 2) = "Empresa09"

    Erase ZProveedor
    ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Where TipoProv = 2"
    ZSql = ZSql + " Order by Nombre"
    spProveedor = ZSql
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        With RstProveedor
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZLugar = ZLugar + 1
                    ZProveedor(ZLugar, 1) = !Proveedor
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        RstProveedor.Close
    End If
    
    Rem ZLugar = 20
    
    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    XEmpresa = WEmpresa
    
    For ZCiclo = 1 To 9
    
        WEmpresa = CargaEmpresa(ZCiclo, 1)
        txtOdbc = CargaEmpresa(ZCiclo, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        For WCiclo = 1 To ZLugar
        
            Erase WVector
            Lugar = 0
                    
            XParam = "'" + WDesde + "','" _
                         + WHasta + "','" _
                         + ZProveedor(WCiclo, 1) + "','" _
                         + ZProveedor(WCiclo, 1) + "'"

            spInforme = "ListaInformeListado" + XParam
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
            
                With rstInforme
    
                    .MoveFirst
            
                    If .NoMatch = False Then
                        Do
                
                            If .EOF = True Then
                                Exit Do
                            End If
                
                            Lugar = Lugar + 1
                
                            WCertificado = IIf(IsNull(rstInforme!Certificado1), "1", rstInforme!Certificado1)
                            WEstado = IIf(IsNull(rstInforme!Estado1), "1", rstInforme!Estado1)
                            
                            WVector(Lugar, 1) = rstInforme!Articulo
                            WVector(Lugar, 2) = rstInforme!Orden
                            WVector(Lugar, 3) = rstInforme!FechaOrd
                            WVector(Lugar, 4) = rstInforme!Clave
                            WVector(Lugar, 5) = WCertificado
                            WVector(Lugar, 6) = WEstado
                            WVector(Lugar, 7) = rstInforme!Informe
                            WVector(Lugar, 8) = rstInforme!Cantidad
                            
                
                            .MoveNext
                
                            If .EOF = True Then
                                Exit Do
                            End If
                
                        Loop
                    End If
                End With
        
                rstInforme.Close
        
            End If
    
            For Ciclo = 1 To Lugar
            
                Rem Calcula las diferencias de fecha entre la
                Rem Orden de compra y el informe de recepcion
    
                WArticulo = WVector(Ciclo, 1)
                WOrden = WVector(Ciclo, 2)
                WFecha = WVector(Ciclo, 3)
                WClave = WVector(Ciclo, 4)
                WCertificado = WVector(Ciclo, 5)
                WEstado = WVector(Ciclo, 6)
                WInforme = WVector(Ciclo, 7)
                WCantidad = WVector(Ciclo, 8)
                
                XFecha = "  /  /    "
                XOrdFecha = "00000000"
        
                XParam = "'" + WOrden + "','" _
                             + WArticulo + "'"

                spOrden = "ListaOrdenArticulo" + XParam
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                If rstOrden.RecordCount > 0 Then
        
                    XOrdFecha = Right$(rstOrden!Fecha2, 4) + Mid$(rstOrden!Fecha2, 4, 2) + Left$(rstOrden!Fecha2, 2)
                    XFecha = rstOrden!Fecha2
            
                    rstOrden.Close
            
                End If
        
                BAse1 = (Val(Left$(XOrdFecha, 4)) * 365) + (Val(Mid$(XOrdFecha, 5, 2)) * 30) + (Val(Right$(XOrdFecha, 2)) * 1)
                Base2 = (Val(Left$(WFecha, 4)) * 365) + (Val(Mid$(WFecha, 5, 2)) * 30) + (Val(Right$(WFecha, 2)) * 1)
        
                Dife = Base2 - BAse1
                If Dife < 0 Then
                    Dife = 0
                End If
                
                ZProveedor(WCiclo, 2) = Str$(Val(ZProveedor(WCiclo, 2)) + 1)
                
                If Val(WCertificado) = 1 Then
                    ZProveedor(WCiclo, 6) = Str$(Val(ZProveedor(WCiclo, 6)) + 1)
                End If
                
                If Val(WEstado) = 1 Then
                    ZProveedor(WCiclo, 7) = Str$(Val(ZProveedor(WCiclo, 7)) + 1)
                End If
                
                aa = Dife
                If Dife > 100 Then Dife = 0
                ZProveedor(WCiclo, 8) = Str$(Val(ZProveedor(WCiclo, 8)) + Dife)
                
                
                
                Rem Calcula las diferencias de fecha entre la
                Rem Orden de compra y el informe de recepcion
                
                WArticulo = WVector(Ciclo, 1)
                WOrden = WVector(Ciclo, 2)
                WFecha = WVector(Ciclo, 3)
                WClave = WVector(Ciclo, 4)
                WCertificado = WVector(Ciclo, 5)
                WEstado = WVector(Ciclo, 6)
                WInforme = WVector(Ciclo, 7)
                WCantidad = Val(WVector(Ciclo, 8))
                
                WLiberada = ""
                WDevuelta = ""
                WPartida1 = ""
                WPartida2 = ""
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Informe"
                ZSql = ZSql + " Where Informe.Clave = " + "'" + WClave + "'"
                spInforme = ZSql
                Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                If rstInforme.RecordCount > 0 Then
                
                    ZZCanti = rstInforme!Cantidad
                    ZZCantienv = rstInforme!CantidadEnv
                    ZZLiberada = ZZCanti - ZZCantienv
                    ZZDevuelta = ZZCantienv
                    ZZDesvio = 0
                    
                    If rstInforme!EstadoEnvII = 1 Or rstInforme!EstadoEnvIV = 1 Or rstInforme!EstadoEnvVI = 1 Or rstInforme!EstadoEnvVIII = 1 Or rstInforme!EstadoEnvX = 1 Then
                        ZZDesvio = ZZLiberada
                        ZZLiberada = 0
                    End If
                
                    rstInforme.Close
                End If
                
                If ZZDevuelta > 0 Then
                    ZProveedor(WCiclo, 5) = Str$(Val(ZProveedor(WCiclo, 5)) + 1)
                End If
                
                If ZZDesvio > 0 Then
                    ZProveedor(WCiclo, 4) = Str$(Val(ZProveedor(WCiclo, 4)) + 1)
                End If
                
                If ZZLiberada > 0 Then
                    ZProveedor(WCiclo, 3) = Str$(Val(ZProveedor(WCiclo, 3)) + 1)
                End If
                
            Next Ciclo
        
        Next WCiclo
        
    Next ZCiclo
    
    Call Conecta_Empresa
    
    For WCiclo = 1 To ZLugar
        
        ZProve = ZProveedor(WCiclo, 1)
        
        Rem total de movimientos
        ZImpre1 = ZProveedor(WCiclo, 2)
        
        Rem item aprobados
        ZImpre2 = ZProveedor(WCiclo, 3)
        
        Rem item desvios
        ZImpre3 = ZProveedor(WCiclo, 4)
        
        Rem item rechazados
        ZImpre4 = ZProveedor(WCiclo, 5)
        
        Rem cantidad de certificado ok
        ZImpre5 = ZProveedor(WCiclo, 6)
        
        Rem cantidad de estados de envases ok
        ZImpre6 = ZProveedor(WCiclo, 7)
        
        Rem cantidad de estados de envases ok
        ZRetrazo = ZProveedor(WCiclo, 8)
        
        If Val(ZImpre1) <> 0 Then
            ZImpre7 = Str$((Val(ZImpre5) / Val(ZImpre1)) * 100)
                Else
            ZImpre7 = ""
        End If
        
        If Val(ZImpre1) <> 0 Then
            ZImpre8 = Str$((Val(ZImpre6) / Val(ZImpre1)) * 100)
                Else
            ZImpre8 = ""
        End If
        
        If Val(ZImpre1) <> 0 Then
            ZImpre9 = Str$(((Val(ZImpre5) + Val(ZImpre6)) / (Val(ZImpre1) * 2)) * 100)
                Else
            ZImpre9 = ""
        End If
            
         If Val(ZImpre1) <> 0 Then
            ZImpre10 = Str$(Val(ZRetrazo) / Val(ZImpre1))
                Else
            ZImpre10 = ""
        End If
        ZImpre10 = Str$(Int(Val(ZImpre10)))
                         
        If Val(ZImpre10) <= 1 Then
            ZImpre11 = "Muy Bueno"
                Else
            If Val(ZImpre10) <= 2 Then
                ZImpre11 = "Bueno"
                    Else
                If Val(ZImpre10) <= 7 Then
                    ZImpre11 = "Regular"
                        Else
                    ZImpre11 = "Malo"
                End If
            End If
        End If
        
        If Val(ZImpre4) = 0 Then
            ZImpre12 = "A"
                Else
            If Val(ZImpre4) = 1 Then
                ZImpre12 = "B"
                    Else
                ZImpre12 = "C"
            End If
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Impre1 = " + "'" + ZImpre1 + "',"
        ZSql = ZSql + " Impre2 = " + "'" + ZImpre2 + "',"
        ZSql = ZSql + " Impre3 = " + "'" + ZImpre3 + "',"
        ZSql = ZSql + " Impre4 = " + "'" + ZImpre4 + "',"
        ZSql = ZSql + " Impre5 = " + "'" + ZImpre5 + "',"
        ZSql = ZSql + " Impre6 = " + "'" + ZImpre6 + "',"
        ZSql = ZSql + " Impre7 = " + "'" + ZImpre7 + "',"
        ZSql = ZSql + " Impre8 = " + "'" + ZImpre8 + "',"
        ZSql = ZSql + " Impre9 = " + "'" + ZImpre9 + "',"
        ZSql = ZSql + " Impre10 = " + "'" + ZImpre10 + "',"
        ZSql = ZSql + " Impre11 = " + "'" + ZImpre11 + "',"
        ZSql = ZSql + " Impre12 = " + "'" + ZImpre12 + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + ZProve + "'"
        spProveedor = ZSql
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        ZPeriodo = "Del " + Desde.Text + " al " + Hasta.Text
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Periodo = " + "'" + ZPeriodo + "'"
        spProveedor = ZSql
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
    Next WCiclo
    
    Listado.WindowTitle = "Calificacion Semestral de Proveedores"
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
    
    Listado.SQLQuery = "SELECT Proveedor.Proveedor, Proveedor.Nombre, Proveedor.TipoProv, Proveedor.CategoriaI, Proveedor.CategoriaII, Proveedor.Iso, Proveedor.VtoIso, Proveedor.Impre1, Proveedor.Impre2, Proveedor.Impre3, Proveedor.Impre4, Proveedor.Impre5, Proveedor.Impre6, Proveedor.Impre7, Proveedor.Impre8, Proveedor.Impre9, Proveedor.Impre10, Proveedor.Impre11, Proveedor.Impre12, Proveedor.Periodo " _
        + "From " _
        + DSQ + ".dbo.Proveedor Proveedor " _
        + "Where " _
        + "Proveedor.Proveedor >= '0' AND " _
        + "Proveedor.Proveedor <= '999999999999' AND " _
        + "Proveedor.TipoProv = 2"
    
      Listado.Connect = Connect()
      Listado.Action = 1
    
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub Cancela_click()
    PrgCalificaEnvase.Hide
    Unload Me
    Menu.Show
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



