VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgIvacomp 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Iva Compras"
   ClientHeight    =   4920
   ClientLeft      =   3240
   ClientTop       =   2025
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   4920
   ScaleWidth      =   5655
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   4080
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1935
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Width           =   3735
      Begin VB.ComboBox TipoListado 
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   1000
         Width           =   1575
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4680
      Top             =   3120
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
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "ivacomp.frx":0000
      Left            =   480
      List            =   "ivacomp.frx":0007
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2760
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgIvacomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstIvaComp As Recordset
Dim spIvaComp As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim XParam As String
Dim ZZPunto As String
Dim ZZNumero As String
Dim ZZib As String
Dim ZVector(8000, 25) As String


Private Sub Acepta_Click()

    Listado.WindowTitle = "Listado de Iva Compras"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With
    
    WTituloII = "del " + Desde.Text + " al " + Hasta.Text
    
    da = 0
    With rstIva
        .Index = "IvaComp"
        .Seek ">=", da
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
            
    Erase ZVector
    ZLugar = 0
            
    Rem XParam = "'" + WDesde + "','" _
    rem              + WHasta + "'"
    Rem spIvaComp = "ListaIvacompDesdeHasta " + XParam
    spIvaComp = "ListaIvacomp"
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
                
        With rstIvaComp
                .MoveFirst
                Do
                
                    WFecha = Right$(!Periodo, 4) + Mid$(!Periodo, 4, 2) + Left$(!Periodo, 2)
                    WLetra = !Letra
                    
                    If WDesde <= WFecha And WFecha <= WHasta Then
                    
                                        
                    
                        ZLugar = ZLugar + 1
                        
                        ZIva105 = IIf(IsNull(!Iva105), "0", !Iva105)
                        
                        ZVector(ZLugar, 1) = !Proveedor
                        ZVector(ZLugar, 2) = !Tipo
                        ZVector(ZLugar, 3) = !Letra
                        ZVector(ZLugar, 4) = !Punto
                        ZVector(ZLugar, 5) = !Numero
                        ZVector(ZLugar, 6) = !Fecha
                        ZVector(ZLugar, 7) = !Vencimiento
                        ZVector(ZLugar, 8) = !Periodo
                        ZVector(ZLugar, 9) = Str$(!Neto)
                        ZVector(ZLugar, 10) = Str$(!Iva21)
                        ZVector(ZLugar, 11) = Str$(!Iva5)
                        ZVector(ZLugar, 12) = Str$(!Iva27)
                        ZVector(ZLugar, 13) = Str$(!Ib)
                        ZVector(ZLugar, 14) = Str$(ZIva105)
                        ZVector(ZLugar, 15) = Str$(!Exento)
                        ZVector(ZLugar, 16) = !Impre
                        ZVector(ZLugar, 17) = !OrdFecha
                        ZVector(ZLugar, 18) = !Contado
                        ZVector(ZLugar, 19) = !Empresa
                        ZVector(ZLugar, 20) = !NroInterno
                        
                        ZSoloIva = IIf(IsNull(rstIvaComp!SoloIva), "0", rstIvaComp!SoloIva)
                        
                        If ZSoloIva = 1 Then
                            ZVector(ZLugar, 9) = "0"
                        End If
                    
                    End If
                    
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
        End With
        rstIvaComp.Close
    End If
    
    
    ZZLugar = 0
    
    For Ciclo = 1 To ZLugar
        
        ZZProveedor = ZVector(Ciclo, 1)
        ZZTipo = ZVector(Ciclo, 2)
        ZZLetra = ZVector(Ciclo, 3)
        ZZPunto = ZVector(Ciclo, 4)
        ZZNumero = ZVector(Ciclo, 5)
        ZZFecha = ZVector(Ciclo, 6)
        ZZvencimiento = ZVector(Ciclo, 7)
        ZZPeriodo = ZVector(Ciclo, 8)
        ZZNeto = Val(ZVector(Ciclo, 9))
        ZZIva21 = Val(ZVector(Ciclo, 10))
        ZZIva5 = Val(ZVector(Ciclo, 11))
        ZZIva27 = Val(ZVector(Ciclo, 12))
        ZZib = Val(ZVector(Ciclo, 13))
        ZZIva105 = Val(ZVector(Ciclo, 14))
        ZZExento = Val(ZVector(Ciclo, 15))
        ZZImpre = ZVector(Ciclo, 16)
        ZZOrdFecha = ZVector(Ciclo, 17)
        ZZContado = ZVector(Ciclo, 18)
        ZZEmpresa = ZVector(Ciclo, 19)
        ZZNroInterno = ZVector(Ciclo, 20)
        
        Rem If ZZIva105 <> 0 Then Stop
        ZZGraba = "S"
        
        If TipoListado.ListIndex = 0 Then
        
            For A = 1 To 50
                
                Auxi = ZZNroInterno
                Call Ceros(Auxi, 8)
                    
                Renglon = A
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                    
                ZZClave = Auxi + Auxi1
                    
                ZSql = "Select *"
                ZSql = ZSql + " FROM IvaCompAdicional"
                ZSql = ZSql + " Where IvaCompAdicional.Clave = " + "'" + ZZClave + "'"
                spIvaCompAdicional = ZSql
                Set rstIvaCompAdicional = db.OpenRecordset(spIvaCompAdicional, dbOpenSnapshot, dbSQLPassThrough)
                If rstIvaCompAdicional.RecordCount > 0 Then
                
                    ZZTipoFac = rstIvaCompAdicional!Tipo
                    
                    Select Case ZZTipoFac
                        Case "NC", "C"
                            ZZTipo = "03"
                            ZZImpre = "NC"
                        Case "ND", "D"
                            ZZTipo = "02"
                            ZZImpre = "NF"
                        Case Else
                            ZZTipo = "01"
                            ZZImpre = "FC"
                    End Select
    
                    
                    ZZLetra = rstIvaCompAdicional!Letra
                    ZZPunto = Trim(rstIvaCompAdicional!Punto)
                    ZZNumero = Trim(rstIvaCompAdicional!Numero)
                    ZZFecha = rstIvaCompAdicional!Fecha
                    ZZOrdFecha = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
                    ZZvencimiento = rstIvaCompAdicional!Fecha
                    ZZPeriodo = rstIvaCompAdicional!Fecha
                    ZZNeto = rstIvaCompAdicional!Neto
                    ZZIva21 = rstIvaCompAdicional!Iva21
                    ZZIva5 = rstIvaCompAdicional!perceiva
                    ZZIva27 = rstIvaCompAdicional!Iva27
                    ZZib = rstIvaCompAdicional!perceib
                    ZZIva105 = rstIvaCompAdicional!Iva105
                    ZZExento = rstIvaCompAdicional!Exento
                    
                    ZZNombre = rstIvaCompAdicional!Razon
                    ZZCuit = rstIvaCompAdicional!Cuit
        
                    rstIvaCompAdicional.Close
                    
                    Call Ceros(ZZPunto, 4)
                    Call Ceros(ZZNumero, 8)
                    
                    ZZLugar = ZZLugar + 1
                    Auxi = Str$(ZZLugar)
                    Call Ceros(Auxi, 2)
        
                    With rstIva
                        .AddNew
                        !NroInterno = ZZNroInterno + Auxi
                        !Proveedor = ZZProveedor
                        
                        !Tipo = ZZTipo
                        !Letra = ZZLetra
                        !Punto = ZZPunto
                        !Numero = ZZNumero
                        !Fecha = ZZFecha
                        !Vencimiento = ZZvencimiento
                        !Periodo = ZZPeriodo
                        !Concepto = 0
                        !Neto = ZZNeto
                        !Iva21 = ZZIva21
                        !Iva5 = ZZIva5
                        !Iva27 = ZZIva27 + ZZIva105
                        !Ib = ZZib
                        !Exento = ZZExento
                        !Impre = ZZImpre
                        !OrdFecha = ZZOrdFecha
                        !Contado = ZZContado
                        !Empresa = Val(XEmpresa)
                        !Titulo = WTitulo
                        !TituloII = WTituloII
                        !Nombre = ZZNombre
                        !Cuit = ZZCuit
                        .Update
                    End With
                        
                    ZZGraba = "N"
                        
                        Else
                        
                    Exit For
                        
                End If
                    
            Next A
    
        End If
        
        If ZZGraba = "S" Then
            
        
            ZZNombre = ""
            ZZCuit = ""
            spProveedor = "ConsultaProveedores " + "'" + ZZProveedor + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                ZZNombre = RstProveedor!Nombre
                ZZCuit = RstProveedor!Cuit
                RstProveedor.Close
            End If
                    
            With rstIva
                .AddNew
                !NroInterno = ZZNroInterno
                !Proveedor = ZZProveedor
                !Tipo = ZZTipo
                !Letra = ZZLetra
                !Punto = ZZPunto
                !Numero = ZZNumero
                !Fecha = ZZFecha
                !Vencimiento = ZZvencimiento
                !Periodo = ZZPeriodo
                !Concepto = ZZConcepto
                !Neto = ZZNeto
                !Iva21 = ZZIva21
                !Iva5 = ZZIva5
                !Iva27 = ZZIva27 + ZZIva105
                !Ib = ZZib
                !Exento = ZZExento
                !Impre = ZZImpre
                !OrdFecha = ZZOrdFecha
                !Contado = ZZContado
                !Empresa = Val(XEmpresa)
                !Titulo = WTitulo
                !TituloII = WTituloII
                !Nombre = ZZNombre
                !Cuit = ZZCuit
                .Update
            End With
                        
        End If
        
    Next Ciclo
    

    Rem Listado.GroupSelectionFormula = "{Ivacomp.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
    
    Listado.Action = 1
End Sub




Private Sub Command1_Click()

    Listado.WindowTitle = "Listado de Iva Compras"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With
    
    WTituloII = "del " + Desde.Text + " al " + Hasta.Text
    
    da = 0
    With rstIva
        .Index = "IvaComp"
        .Seek ">=", da
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
            
    Erase ZVector
    ZLugar = 0
            
    Rem XParam = "'" + WDesde + "','" _
    rem              + WHasta + "'"
    Rem spIvaComp = "ListaIvacompDesdeHasta " + XParam
    spIvaComp = "ListaIvacomp"
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
                
        With rstIvaComp
                .MoveFirst
                Do
                
                    WFecha = Right$(!Periodo, 4) + Mid$(!Periodo, 4, 2) + Left$(!Periodo, 2)
                    WLetra = !Letra
                    
                    If WDesde <= WFecha And WFecha <= WHasta Then
                    
                    If !Proveedor = "10053988991" Then
                    
                        ZLugar = ZLugar + 1
                        
                        ZIva105 = IIf(IsNull(!Iva105), "0", !Iva105)
                        
                        ZVector(ZLugar, 1) = !Proveedor
                        ZVector(ZLugar, 2) = !Tipo
                        ZVector(ZLugar, 3) = !Letra
                        ZVector(ZLugar, 4) = !Punto
                        ZVector(ZLugar, 5) = !Numero
                        ZVector(ZLugar, 6) = !Fecha
                        ZVector(ZLugar, 7) = !Vencimiento
                        ZVector(ZLugar, 8) = !Periodo
                        ZVector(ZLugar, 9) = Str$(!Neto)
                        ZVector(ZLugar, 10) = Str$(!Iva21)
                        ZVector(ZLugar, 11) = Str$(!Iva5)
                        ZVector(ZLugar, 12) = Str$(!Iva27)
                        ZVector(ZLugar, 13) = Str$(!Ib)
                        ZVector(ZLugar, 14) = Str$(ZIva105)
                        ZVector(ZLugar, 15) = Str$(!Exento)
                        ZVector(ZLugar, 16) = !Impre
                        ZVector(ZLugar, 17) = !OrdFecha
                        ZVector(ZLugar, 18) = !Contado
                        ZVector(ZLugar, 19) = !Empresa
                        ZVector(ZLugar, 20) = !NroInterno
                        ZVector(ZLugar, 21) = IIf(IsNull(!Remito), "0", !Remito)
                        
                        Rem If Val(ZVector(ZLugar, 21)) <> 0 Then Stop
                        
                        
                        ZSoloIva = IIf(IsNull(rstIvaComp!SoloIva), "0", rstIvaComp!SoloIva)
                        
                        If ZSoloIva = 1 Then
                            ZVector(ZLugar, 9) = "0"
                        End If
                    
                    End If
                    
                    End If
                    
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
        End With
        rstIvaComp.Close
    End If
    
            
        
    XEmpresa = Wempresa
    
    For CiclaEmpresa = 1 To 6
    
        Select Case CiclaEmpresa
            Case 1
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                Wempresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                Wempresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                Wempresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 5
                Wempresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 6
                Wempresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 7
                Wempresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
    
        For Ciclo = 1 To ZLugar
        
            Rem ZZRemito = ZNroRemito(CicloII)
        
            ZZProveedor = ZVector(Ciclo, 1)
            ZZRemito = ZVector(Ciclo, 21)
            ZZOrden = ZVector(Ciclo, 22)
            
            If Val(ZZRemito) <> 0 And Val(ZZOrden) = 0 Then
        
                ZZArticulo = ""
        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Informe"
                ZSql = ZSql + " Where Informe.Remito = " + "'" + ZZRemito + "'"
                ZSql = ZSql + " and Informe.Proveedor = " + "'" + ZZProveedor + "'"
                spInforme = ZSql
                Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                If rstInforme.RecordCount > 0 Then
                    With rstInforme
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                
                                ZZOrden = Str$(rstInforme!Orden)
                                ZVector(Ciclo, 22) = ZZOrden
                                ZZArticulo = rstInforme!Articulo
                                
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstInforme.Close
                End If
                
                If Trim(ZZArticulo) <> "" Then
                    spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        ZVector(Ciclo, 23) = Trim(rstArticulo!Descripcion)
                        rstArticulo.Close
                    End If
                End If
            
            End If
            
        Next Ciclo

    Next CiclaEmpresa
    
    Call Conecta_Empresa
        
            
            
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ZZLugar = 0
    
    For Ciclo = 1 To ZLugar
        
        ZZProveedor = ZVector(Ciclo, 1)
        ZZTipo = ZVector(Ciclo, 2)
        ZZLetra = ZVector(Ciclo, 3)
        ZZPunto = ZVector(Ciclo, 4)
        ZZNumero = ZVector(Ciclo, 5)
        ZZFecha = ZVector(Ciclo, 6)
        ZZvencimiento = ZVector(Ciclo, 7)
        ZZPeriodo = ZVector(Ciclo, 8)
        ZZNeto = Val(ZVector(Ciclo, 9))
        ZZIva21 = Val(ZVector(Ciclo, 10))
        ZZIva5 = Val(ZVector(Ciclo, 11))
        ZZIva27 = Val(ZVector(Ciclo, 12))
        ZZib = Val(ZVector(Ciclo, 13))
        ZZIva105 = Val(ZVector(Ciclo, 14))
        ZZExento = Val(ZVector(Ciclo, 15))
        ZZImpre = ZVector(Ciclo, 16)
        ZZOrdFecha = ZVector(Ciclo, 17)
        ZZContado = ZVector(Ciclo, 18)
        ZZEmpresa = ZVector(Ciclo, 19)
        ZZNroInterno = ZVector(Ciclo, 20)
        ZZRemito = ZVector(Ciclo, 21)
        ZZOrden = ZVector(Ciclo, 22)
        ZZProducto = ZVector(Ciclo, 23)
        
        Rem If Val(ZZRemito) <> 0 Then Stop
        
        Rem If ZZIva105 <> 0 Then Stop
        ZZGraba = "S"
        
        If TipoListado.ListIndex = 0 Then
        
            For A = 1 To 50
                
                Auxi = ZZNroInterno
                Call Ceros(Auxi, 8)
                    
                Renglon = A
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                    
                ZZClave = Auxi + Auxi1
                    
                ZSql = "Select *"
                ZSql = ZSql + " FROM IvaCompAdicional"
                ZSql = ZSql + " Where IvaCompAdicional.Clave = " + "'" + ZZClave + "'"
                spIvaCompAdicional = ZSql
                Set rstIvaCompAdicional = db.OpenRecordset(spIvaCompAdicional, dbOpenSnapshot, dbSQLPassThrough)
                If rstIvaCompAdicional.RecordCount > 0 Then
                
                    ZZTipoFac = rstIvaCompAdicional!Tipo
                    
                    Select Case ZZTipoFac
                        Case "NC", "C"
                            ZZTipo = "03"
                            ZZImpre = "NC"
                        Case "ND", "D"
                            ZZTipo = "02"
                            ZZImpre = "NF"
                        Case Else
                            ZZTipo = "01"
                            ZZImpre = "FC"
                    End Select
    
                    
                    ZZLetra = rstIvaCompAdicional!Letra
                    ZZPunto = Trim(rstIvaCompAdicional!Punto)
                    ZZNumero = Trim(rstIvaCompAdicional!Numero)
                    ZZFecha = rstIvaCompAdicional!Fecha
                    ZZOrdFecha = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
                    ZZvencimiento = rstIvaCompAdicional!Fecha
                    ZZPeriodo = rstIvaCompAdicional!Fecha
                    ZZNeto = rstIvaCompAdicional!Neto
                    ZZIva21 = rstIvaCompAdicional!Iva21
                    ZZIva5 = rstIvaCompAdicional!perceiva
                    ZZIva27 = rstIvaCompAdicional!Iva27
                    ZZib = rstIvaCompAdicional!perceib
                    ZZIva105 = rstIvaCompAdicional!Iva105
                    ZZExento = rstIvaCompAdicional!Exento
                    
                    ZZNombre = rstIvaCompAdicional!Razon
                    ZZCuit = rstIvaCompAdicional!Cuit
        
                    rstIvaCompAdicional.Close
                    
                    Call Ceros(ZZPunto, 4)
                    Call Ceros(ZZNumero, 8)
                    
                    ZZLugar = ZZLugar + 1
                    Auxi = Str$(ZZLugar)
                    Call Ceros(Auxi, 2)
        
                    With rstIva
                        .AddNew
                        !NroInterno = ZZNroInterno + Auxi
                        !Proveedor = ZZProveedor
                        
                        !Tipo = ZZTipo
                        !Letra = ZZLetra
                        !Punto = ZZPunto
                        !Numero = ZZNumero
                        !Fecha = ZZFecha
                        !Vencimiento = ZZvencimiento
                        !Periodo = ZZPeriodo
                        !Concepto = 0
                        !Neto = ZZNeto
                        !Iva21 = ZZIva21
                        !Iva5 = ZZIva5
                        !Iva27 = ZZIva27 + ZZIva105
                        !Ib = ZZib
                        !Exento = ZZExento
                        !Impre = ZZImpre
                        !OrdFecha = ZZOrdFecha
                        !Contado = ZZContado
                        !Empresa = Val(XEmpresa)
                        !Titulo = WTitulo
                        !TituloII = WTituloII
                        !Nombre = ZZNombre
                        !Cuit = ZZCuit
                        !Remito = ZZRemito
                        !Orden = ZZOrden
                        
                        .Update
                    End With
                        
                    ZZGraba = "N"
                        
                        Else
                        
                    Exit For
                        
                End If
                    
            Next A
    
        End If
        
        If ZZGraba = "S" Then
            
        
            ZZNombre = ""
            ZZCuit = ""
            spProveedor = "ConsultaProveedores " + "'" + ZZProveedor + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                ZZNombre = RstProveedor!Nombre
                ZZCuit = RstProveedor!Cuit
                RstProveedor.Close
            End If
            
            
            
            
            
            
                    
            With rstIva
                .AddNew
                !NroInterno = ZZNroInterno
                !Proveedor = ZZProveedor
                !Tipo = ZZTipo
                !Letra = ZZLetra
                !Punto = ZZPunto
                !Numero = ZZNumero
                !Fecha = ZZFecha
                !Vencimiento = ZZvencimiento
                !Periodo = ZZPeriodo
                !Concepto = ZZConcepto
                !Neto = ZZNeto
                !Iva21 = ZZIva21
                !Iva5 = ZZIva5
                !Iva27 = ZZIva27 + ZZIva105
                !Ib = ZZib
                !Exento = ZZExento
                !Impre = ZZImpre
                !OrdFecha = ZZOrdFecha
                !Contado = ZZContado
                !Empresa = Val(XEmpresa)
                !Titulo = WTitulo
                !TituloII = WTituloII
                !Nombre = ZZProducto
                !Cuit = ZZCuit
                !Remito = Val(ZZRemito)
                !Orden = Val(ZZOrden)
                .Update
            End With
                        
        End If
        
    Next Ciclo
    
    Listado.ReportFileName = "WIvacompotro.rpt"

    Rem Listado.GroupSelectionFormula = "{Ivacomp.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
    
    Listado.Action = 1
End Sub




Private Sub Cancela_Click()
    With rstEmpresa
        .Close
    End With
    With rstIva
        .Close
    End With
    Desde.SetFocus
    PrgIvacomp.Hide
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
    Rem If KeyAscii >= 48 And KeyAscii <= 57 Then
    Rem     If Desde.SelStart = 1 Then
    Rem         Desde.Text = Mid$(Desde.Text, 1, Desde.SelStart) + Chr$(KeyAscii) + Mid$(Desde.Text, Desde.SelStart + 1, 10)
    Rem         If Mid$(Desde.Text, 3, 1) <> "/" Then
    Rem             Desde.Text = Desde.Text + "/"
    Rem         End If
    Rem         KeyAscii = 0
    Rem         Desde.SelStart = 3
    Rem     End If
    Rem     If Desde.SelStart = 4 Then
    Rem         Desde.Text = Mid$(Desde.Text, 1, Desde.SelStart) + Chr$(KeyAscii) + Mid$(Desde.Text, Desde.SelStart + 1, 10)
    Rem         If Mid$(Desde.Text, 6, 1) <> "/" Then
    Rem             Desde.Text = Desde.Text + "/"
    Rem         End If
    Rem         KeyAscii = 0
    Rem         Desde.SelStart = 6
    Rem     End If
    Rem End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Iva
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

    TipoListado.Clear
    
    TipoListado.AddItem "C/Apertura"
    TipoListado.AddItem "S/Apertura"
    
    TipoListado.ListIndex = 0

    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub
