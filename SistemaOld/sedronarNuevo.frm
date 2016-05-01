VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgSedronarNuevo 
   AutoRedraw      =   -1  'True
   Caption         =   "Declaracion Jurada (Sedronar)"
   ClientHeight    =   7365
   ClientLeft      =   450
   ClientTop       =   825
   ClientWidth     =   11100
   LinkTopic       =   "Form2"
   ScaleHeight     =   7365
   ScaleWidth      =   11100
   Begin MSFlexGridLib.MSFlexGrid IngresoDatos 
      Height          =   2535
      Left            =   360
      TabIndex        =   15
      Top             =   3000
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4471
      _Version        =   327680
      Rows            =   1000
      Cols            =   3
   End
   Begin VB.TextBox Ayuda 
      Height          =   375
      Left            =   6840
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Procesogg 
      Caption         =   "Proceso"
      Height          =   300
      Left            =   480
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   4335
      Begin VB.CheckBox Limpia 
         Caption         =   "Limpia"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1320
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1320
         TabIndex        =   0
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Sedronar.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cuenta Corriente de Proveedores"
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
      Left            =   5280
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   6495
      ItemData        =   "sedronarNuevo.frx":0000
      Left            =   6840
      List            =   "sedronarNuevo.frx":0007
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox WProducto 
      Height          =   300
      Left            =   2640
      TabIndex        =   17
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      Caption         =   "Ingreso de Producto"
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
      Left            =   480
      TabIndex        =   16
      Top             =   5640
      Width           =   1935
   End
End
Attribute VB_Name = "PrgSedronarNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Vector(1000, 10) As String
Private WVectorII(1000, 20) As String
Private WProceso(1000, 10) As String
Private ProveCompras(1000, 10) As String
Private OrdenCompras(1000, 10) As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstSedronar As Recordset
Dim spSedronar As String
Dim XParam As String
Dim WArticulo As String
Dim WEntradas As Double
Dim WSalidas As Double
Dim Stock1 As Double
Dim Stock2 As Double
Dim WCompras As Double
Dim WDesde As String
Dim WHasta As String
Dim WFechaord As String
Dim Lugar As Integer
Dim LugarProve As Integer
Dim LugarOrden As Integer
Dim WEmpre(10) As String
Dim LugarVectorII As Integer
Dim LugarProceso As Integer

Private Sub Acepta_Click()

    Dim ZZCufe(100) As String
    
    ZZCufe(1) = "9980334210003"
    ZZCufe(2) = ""
    ZZCufe(3) = "9980396510004"
    ZZCufe(4) = "9980401950009"
    ZZCufe(5) = "9980396350006"
    ZZCufe(6) = ""
    ZZCufe(7) = "9980396360005"
    ZZCufe(8) = "9980307940005"
    ZZCufe(9) = ""
    ZZCufe(10) = "9980396370004"
    ZZCufe(11) = "9980396380003"

    OPEN_FILE_SedronarProceso
    If Limpia.Value = 1 Then
        da = 0
        With rstSedronarProceso
            .Index = "Clave"
            .Seek ">=", ""
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
    End If

    For A = 1 To 999
    
        iRow = A
    
        IngresoDatos.Col = 1
        IngresoDatos.Row = iRow
        WArticulo = IngresoDatos.Text
        XCodigo = IngresoDatos.Text
        XXDescripcion = ""
                
        If WArticulo <> "" Then
                
            XEmpresa = Wempresa
            Select Case Val(XEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2, 4, 8, 9
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
                
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZCodsedronar = IIf(IsNull(rstArticulo!CodSedronar), "", rstArticulo!CodSedronar)
                XXDescripcion = rstArticulo!Descripcion
                rstArticulo.Close
            End If
            
            Select Case Val(XEmpresa)
                Case 1
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2
                    Wempresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 3
                    Wempresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 4
                    Wempresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 5
                    Wempresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 6
                    Wempresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 7
                    Wempresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 8
                    Wempresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 9
                    Wempresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 10
                    Wempresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 11
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
                
            WAno = Right$(Desde.Text, 4)
            WMes = Mid$(Desde.Text, 4, 2)
            WDia = Left$(Desde.Text, 2)
            WFechaord = WAno + WMes + WDia
                        
            WAno = Right$(Desde.Text, 4)
            WMes = Mid$(Desde.Text, 4, 2)
            WDia = Left$(Desde.Text, 2)
            WDesde = WAno + WMes + WDia
                    
            WAno = Right$(Hasta.Text, 4)
            WMes = Mid$(Hasta.Text, 4, 2)
            WDia = Left$(Hasta.Text, 2)
            WHasta = WAno + WMes + WDia
                
            Erase WVectorII
            LugarVectorII = 0
            
            Call Proceso
                    
            For Ciclo = 1 To LugarVectorII
                
                
                aa = WArticulo
                If Trim(ZZCodsedronar) = "" Then Stop
                
                ZZTipo = WVectorII(Ciclo, 1)
                ZZFecha = WVectorII(Ciclo, 2)
                ZZCantidad = WVectorII(Ciclo, 3)
                ZZCodigo = WVectorII(Ciclo, 4)
                ZZMovi = WVectorII(Ciclo, 5)
                ZZDestino = WVectorII(Ciclo, 6)
                ZZtipomov = WVectorII(Ciclo, 7)
                ZZCufeI = WVectorII(Ciclo, 8)
                ZZCufeII = WVectorII(Ciclo, 9)
                ZZCufeIII = WVectorII(Ciclo, 10)
                ZZTerminado = WVectorII(Ciclo, 11)
                ZZFechaOrd = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
                
                If Trim(ZZCufeI) <> "" And Trim(ZZCufeII) = "" And Trim(ZZCufeIII) = "" Then
                    zzcufeok = ZZCufeI
                        Else
                    zzcufeok = ZZtipomov
                End If
                
                Select Case Val(ZZTipo)
                    Case 1
                        ZLaudo = 0
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArticulo + "'"
                        ZSql = ZSql + " and Laudo.Informe = " + "'" + ZZMovi + "'"
                        spLaudo = ZSql
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                        
                            With rstLaudo
                        
                                .MoveFirst
                                
                                If .NoMatch = False Then
                                Do
                                
                                    If .EOF = True Then
                                        Exit Do
                                    End If
                                    
                                    ZZNroDespacho = IIf(IsNull(rstLaudo!NroDespacho), "", rstLaudo!NroDespacho)
                                    WLiberada = IIf(IsNull(rstLaudo!Liberadaant), 0, rstLaudo!Liberadaant)
                                    
                                    If WLiberada = 0 Then
                                        WLiberada = rstLaudo!Liberada
                                    End If
                                    ZLaudo = ZLaudo + WLiberada
                                    
                                    .MoveNext
                                    
                                    If .EOF = True Then
                                        Exit Do
                                    End If
                                    
                                Loop
                                End If
                            End With
                            
                            rstLaudo.Close
                            
                        End If
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden.Orden = " + "'" + ZZDestino + "'"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            ZZTipoOrden = rstOrden!Tipo
                            rstOrden.Close
                        End If
                        
                    
                    
                        If ZLaudo > 0 Then
                        
                            If ZZTipoOrden = 1 Then
                                ZZEvento = "45"
                                    Else
                                If Left$(UCase(WArticulo), 2) = "XX" Or ZZtipomov = "10071011210" Then
                                    ZZEvento = "68"
                                        Else
                                    ZZEvento = "43"
                                End If
                            End If
                            With rstSedronarProceso
                                .AddNew
                                !Fecha = ZZFecha
                                !Evento = ZZEvento
                                !Gtin = Trim(ZZCodsedronar)
                                !Cantidad = ZLaudo
                                !Analitica = ""
                                !Parcial = ""
                                !Tipo = 2
                                !Numero = ZZCodigo
                                !CufeOrigen = Trim(zzcufeok)
                                !CufeDestino = Trim(ZZCufe(Val(Wempresa)))
                                !CufeTransportista = ""
                                !Permiso = ""
                                !PermisoII = ""
                                !Dominio = ""
                                !TipoDoc = ""
                                !NroDoc = ""
                                !TipoTransporte = ""
                                If ZZTipoOrden = 1 Then
                                    !Plaza = ZZNroDespacho
                                    !DJai = "falta djai"
                                    !Paso = "219"
                                        Else
                                    !Plaza = ""
                                    !DJai = ""
                                    !Paso = ""
                                End If
                                !NroCertificado = ""
                                !Clave = "M" + Trim(ZZCodsedronar) + "E" + ZZEvento + ZZFechaOrd
                                !Suma = ZLaudo
                                .Update
                            End With
                        
                            
                        End If
                
                    Case 2
                        ZZCodTer = Val(Mid$(ZZTerminado, 4, 5))
                        If ZZCodTer >= 40000 And ZZCodTer <= 41999 Then
                            ZZEvento = "69"
                                Else
                            ZZEvento = "54"
                        End If
                        With rstSedronarProceso
                            .AddNew
                            !Fecha = ZZFecha
                            !Evento = ZZEvento
                            !Gtin = Trim(ZZCodsedronar)
                            !Cantidad = Val(ZZCantidad)
                            !Analitica = ""
                            !Parcial = ""
                            !Tipo = 3
                            !Numero = ZZCodigo
                            !CufeOrigen = Trim(ZZCufe(Val(Wempresa)))
                            !CufeDestino = ""
                            !CufeTransportista = ""
                            !Permiso = ""
                            !PermisoII = ""
                            !Dominio = ""
                            !TipoDoc = ""
                            !NroDoc = ""
                            !TipoTransporte = ""
                            !Plaza = ""
                            !DJai = ""
                            !Paso = ""
                            !NroCertificado = ""
                            !Clave = "M" + Trim(ZZCodsedronar) + "S" + ZZEvento + ZZFechaOrd
                            !Suma = Val(ZZCantidad) * -1
                            .Update
                        End With
                    
                        
                        
                    Case 3
                        If ZZMovi = "S" Then
                            ZZEvento = "66"
                                Else
                            ZZEvento = "58"
                        End If
                        With rstSedronarProceso
                            .AddNew
                            !Fecha = ZZFecha
                            !Evento = ZZEvento
                            !Gtin = Trim(ZZCodsedronar)
                            !Cantidad = Val(ZZCantidad)
                            !Analitica = ""
                            !Parcial = ""
                            !Tipo = ""
                            !Numero = ""
                            !CufeOrigen = Trim(ZZCufe(Val(Wempresa)))
                            !CufeDestino = ""
                            !CufeTransportista = ""
                            !Permiso = ""
                            !PermisoII = ""
                            !Dominio = ""
                            !TipoDoc = ""
                            !NroDoc = ""
                            !TipoTransporte = ""
                            !Plaza = ""
                            !DJai = ""
                            !Paso = ""
                            !NroCertificado = ""
                            !Clave = "M" + Trim(ZZCodsedronar) + ZZMovi + ZZEvento + ZZFechaOrd
                            If ZZMovi = "S" Then
                                !Suma = Val(ZZCantidad) * -1
                                    Else
                                !Suma = Val(ZZCantidad)
                            End If
                            .Update
                        End With
                        
                        
                    Case 4
                        ZZTipo = WVectorII(Ciclo, 1)
                        ZZFecha = WVectorII(Ciclo, 2)
                        ZZCantidad = WVectorII(Ciclo, 3)
                        ZZCodigo = WVectorII(Ciclo, 4)
                        ZZMovi = WVectorII(Ciclo, 5)
                        ZZDestino = WVectorII(Ciclo, 6)
                        ZZtipomov = WVectorII(Ciclo, 7)
                        ZZCufeI = WVectorII(Ciclo, 8)
                        ZZCufeII = WVectorII(Ciclo, 9)
                        ZZCufeIII = WVectorII(Ciclo, 10)
                    
                        If ZZMovi = "S" Then
                            ZZEvento = "48"
                            ZZLugarCufe = ZZDestino
                                Else
                            ZZLugarCufe = ZZtipomov
                            ZZEvento = "47"
                        End If
                        If ZZMovi = "S" Then
                            ZZCufeOrigen = ZZCufe(Val(Wempresa))
                            ZZCufeDestino = ZZCufe(ZZLugarCufe)
                                Else
                            ZZCufeOrigen = ZZCufe(ZZLugarCufe)
                            ZZCufeDestino = ZZCufe(Val(Wempresa))
                        End If
                        With rstSedronarProceso
                            .AddNew
                            !Fecha = ZZFecha
                            !Evento = ZZEvento
                            !Gtin = Trim(ZZCodsedronar)
                            !Cantidad = Val(ZZCantidad)
                            !Analitica = ""
                            !Parcial = ""
                            !Tipo = ""
                            !Numero = ""
                            !CufeOrigen = Trim(ZZCufeOrigen)
                            !CufeDestino = Trim(ZZCufeDestino)
                            !CufeTransportista = ""
                            !Permiso = ""
                            !PermisoII = ""
                            !Dominio = ""
                            !TipoDoc = ""
                            !NroDoc = ""
                            !TipoTransporte = ""
                            !Plaza = ""
                            !DJai = ""
                            !Paso = ""
                            !NroCertificado = ""
                            !Clave = "M" + Trim(ZZCodsedronar) + ZZMovi + ZZEvento + ZZFechaOrd
                            If ZZMovi = "S" Then
                                !Suma = Val(ZZCantidad) * -1
                                    Else
                                !Suma = Val(ZZCantidad)
                            End If
                            .Update
                        End With
                        
                    Case Else
                End Select
                
                
                
            Next Ciclo
                
        End If
        
    Next A
    
    Call Cancela_click
    
End Sub



Private Sub AceptaAnterior_Click()

    Dim ZZCufe(100) As String
    
    ZZCufe(1) = "9980334210003"
    ZZCufe(2) = ""
    ZZCufe(3) = "9980396510004"
    ZZCufe(4) = "9980401950009"
    ZZCufe(5) = "9980396350006"
    ZZCufe(6) = ""
    ZZCufe(7) = "9980396360005"
    ZZCufe(8) = "9980307940005"
    ZZCufe(9) = ""
    ZZCufe(10) = "9980396370004"
    ZZCufe(11) = "9980396380003"

    Set appExcel = CreateObject("Excel.application")
    
    Select Case Val(Wempresa)
        Case 1
            ruta = "C:\sedronar\pasasedrosi.xls"
        Case 2
            ruta = "C:\sedronar\pasasedropi.xls"
        Case 3
            ruta = "C:\sedronar\pasasedrosii.xls"
        Case 4
            ruta = "C:\sedronar\pasasedropii.xls"
        Case 5
            ruta = "C:\sedronar\pasasedrosiii.xls"
        Case 6
            ruta = "C:\sedronar\pasasedrosiv.xls"
        Case 7
            ruta = "C:\sedronar\pasasedrosv.xls"
        Case 8
            ruta = "C:\sedronar\pasasedropiii.xls"
        Case 9
            ruta = "C:\sedronar\pasasedropv.xls"
        Case 10
            ruta = "C:\sedronar\pasasedrosvi.xls"
        Case Else
            ruta = "C:\sedronar\pasasedrosvii.xls"
    End Select

    Rem ruta = "C:\sedronar\prueba.xls"


    If Len(Dir(ruta)) > 0 Then
    
    
        Set objLibro = appExcel.workbooks.Open(ruta)
        LugarPlanilla = 1
    
        For Ciclo = 2 To 500
            appExcel.cells(Ciclo, 1).Value = ""
            appExcel.cells(Ciclo, 2).Value = ""
            appExcel.cells(Ciclo, 3).Value = ""
            appExcel.cells(Ciclo, 4).Value = ""
            appExcel.cells(Ciclo, 5).Value = ""
            appExcel.cells(Ciclo, 6).Value = ""
            appExcel.cells(Ciclo, 7).Value = ""
            appExcel.cells(Ciclo, 8).Value = ""
            appExcel.cells(Ciclo, 9).Value = ""
            appExcel.cells(Ciclo, 10).Value = ""
            appExcel.cells(Ciclo, 11).Value = ""
            appExcel.cells(Ciclo, 12).Value = ""
            appExcel.cells(Ciclo, 13).Value = ""
            appExcel.cells(Ciclo, 14).Value = ""
            appExcel.cells(Ciclo, 15).Value = ""
            appExcel.cells(Ciclo, 16).Value = ""
            appExcel.cells(Ciclo, 17).Value = ""
            appExcel.cells(Ciclo, 18).Value = ""
            appExcel.cells(Ciclo, 19).Value = ""
            appExcel.cells(Ciclo, 20).Value = ""
            appExcel.cells(Ciclo, 21).Value = ""
        Next Ciclo
    
        For A = 1 To 999
        
            iRow = A
        
            IngresoDatos.Col = 1
            IngresoDatos.Row = iRow
            WArticulo = IngresoDatos.Text
            XCodigo = IngresoDatos.Text
            XXDescripcion = ""
                    
            If WArticulo <> "" Then
                    
                XEmpresa = Wempresa
                    
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZCodsedronar = IIf(IsNull(rstArticulo!CodSedronar), "", rstArticulo!CodSedronar)
                    XXDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                
                Select Case Val(XEmpresa)
                    Case 1
                        Wempresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 2
                        Wempresa = "0002"
                        txtOdbc = "Empresa02"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 3
                        Wempresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 4
                        Wempresa = "0004"
                        txtOdbc = "Empresa04"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 5
                        Wempresa = "0005"
                        txtOdbc = "Empresa05"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 6
                        Wempresa = "0006"
                        txtOdbc = "Empresa06"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 7
                        Wempresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 8
                        Wempresa = "0008"
                        txtOdbc = "Empresa08"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 9
                        Wempresa = "0009"
                        txtOdbc = "Empresa09"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 10
                        Wempresa = "0010"
                        txtOdbc = "Empresa10"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 11
                        Wempresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                End Select
                    
                WAno = Right$(Desde.Text, 4)
                WMes = Mid$(Desde.Text, 4, 2)
                WDia = Left$(Desde.Text, 2)
                WFechaord = WAno + WMes + WDia
                            
                WAno = Right$(Desde.Text, 4)
                WMes = Mid$(Desde.Text, 4, 2)
                WDia = Left$(Desde.Text, 2)
                WDesde = WAno + WMes + WDia
                        
                WAno = Right$(Hasta.Text, 4)
                WMes = Mid$(Hasta.Text, 4, 2)
                WDia = Left$(Hasta.Text, 2)
                WHasta = WAno + WMes + WDia
                    
                Erase WVectorII
                LugarVectorII = 0
                
                Call Proceso
                        
                For Ciclo = 1 To LugarVectorII
                    
                    
                     aa = WArticulo
                    If Trim(ZZCodsedronar) = "" Then Stop
                    
                    ZZTipo = WVectorII(Ciclo, 1)
                    ZZFecha = WVectorII(Ciclo, 2)
                    ZZCantidad = WVectorII(Ciclo, 3)
                    ZZCodigo = WVectorII(Ciclo, 4)
                    ZZMovi = WVectorII(Ciclo, 5)
                    ZZDestino = WVectorII(Ciclo, 6)
                    ZZtipomov = WVectorII(Ciclo, 7)
                    ZZCufeI = WVectorII(Ciclo, 8)
                    ZZCufeII = WVectorII(Ciclo, 9)
                    ZZCufeIII = WVectorII(Ciclo, 10)
                    ZZTerminado = WVectorII(Ciclo, 11)
                    
                    If Trim(ZZCufeI) <> "" And Trim(ZZCufeII) = "" And Trim(ZZCufeIII) = "" Then
                        zzcufeok = ZZCufeI
                            Else
                        zzcufeok = ZZtipomov
                    End If
                    
                    Select Case Val(ZZTipo)
                        Case 1
                            ZLaudo = 0
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Laudo"
                            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArticulo + "'"
                            ZSql = ZSql + " and Laudo.Informe = " + "'" + ZZMovi + "'"
                            spLaudo = ZSql
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                            
                                With rstLaudo
                            
                                    .MoveFirst
                                    
                                    If .NoMatch = False Then
                                    Do
                                    
                                        If .EOF = True Then
                                            Exit Do
                                        End If
                                        
                                        ZZNroDespacho = IIf(IsNull(rstLaudo!NroDespacho), "", rstLaudo!NroDespacho)
                                        WLiberada = IIf(IsNull(rstLaudo!Liberadaant), 0, rstLaudo!Liberadaant)
                                        
                                        If WLiberada = 0 Then
                                            WLiberada = rstLaudo!Liberada
                                        End If
                                        ZLaudo = ZLaudo + WLiberada
                                        
                                        .MoveNext
                                        
                                        If .EOF = True Then
                                            Exit Do
                                        End If
                                        
                                    Loop
                                    End If
                                End With
                                
                                rstLaudo.Close
                                
                            End If
                            
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Orden"
                            ZSql = ZSql + " Where Orden.Orden = " + "'" + ZZDestino + "'"
                            spOrden = ZSql
                            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                            If rstOrden.RecordCount > 0 Then
                                ZZTipoOrden = rstOrden!Tipo
                                rstOrden.Close
                            End If
                            
                        
                        
                            If ZLaudo > 0 Then
                            
                                LugarPlanilla = LugarPlanilla + 1
                                appExcel.cells(LugarPlanilla, 1).Value = ZZFecha
                                If ZZTipoOrden = 1 Then
                                    appExcel.cells(LugarPlanilla, 2).Value = "45"
                                        Else
                                    If Left$(UCase(WArticulo), 2) = "XX" Or ZZtipomov = "10071011210" Then
                                        appExcel.cells(LugarPlanilla, 2).Value = "68"
                                            Else
                                        appExcel.cells(LugarPlanilla, 2).Value = "43"
                                    End If
                                End If
                                appExcel.cells(LugarPlanilla, 3).Value = ZZCodsedronar
                                appExcel.cells(LugarPlanilla, 4).Value = Val(ZLaudo)
                                appExcel.cells(LugarPlanilla, 5).Value = ""
                                appExcel.cells(LugarPlanilla, 6).Value = ""
                                appExcel.cells(LugarPlanilla, 7).Value = "2"
                                appExcel.cells(LugarPlanilla, 8).Value = ZZCodigo
                                appExcel.cells(LugarPlanilla, 9).Value = zzcufeok
                                appExcel.cells(LugarPlanilla, 10).Value = ZZCufe(Val(Wempresa))
                                appExcel.cells(LugarPlanilla, 11).Value = ""
                                appExcel.cells(LugarPlanilla, 13).Value = ""
                                appExcel.cells(LugarPlanilla, 14).Value = ""
                                appExcel.cells(LugarPlanilla, 15).Value = ""
                                appExcel.cells(LugarPlanilla, 16).Value = ""
                                appExcel.cells(LugarPlanilla, 17).Value = ""
                                appExcel.cells(LugarPlanilla, 18).Value = ""
                                If ZZTipoOrden = 1 Then
                                    appExcel.cells(LugarPlanilla, 19).Value = ZZNroDespacho
                                    appExcel.cells(LugarPlanilla, 20).Value = "falta djai"
                                    appExcel.cells(LugarPlanilla, 12).Value = "219"
                                        Else
                                    appExcel.cells(LugarPlanilla, 19).Value = ""
                                    appExcel.cells(LugarPlanilla, 20).Value = ""
                                    appExcel.cells(LugarPlanilla, 12).Value = ""
                                End If
                                appExcel.cells(LugarPlanilla, 21).Value = ""
                                appExcel.cells(LugarPlanilla, 22).Value = ZZFecha
                                
                            End If
                    
                        Case 2
                            LugarPlanilla = LugarPlanilla + 1
                            appExcel.cells(LugarPlanilla, 1).Value = ZZFecha
                            ZZCodTer = Val(Mid$(ZZTerminado, 4, 5))
                            If ZZCodTer >= 40000 And ZZCodTer <= 41999 Then
                                appExcel.cells(LugarPlanilla, 2).Value = "69"
                                    Else
                                appExcel.cells(LugarPlanilla, 2).Value = "54"
                            End If
                            appExcel.cells(LugarPlanilla, 3).Value = ZZCodsedronar
                            appExcel.cells(LugarPlanilla, 4).Value = Val(ZZCantidad)
                            appExcel.cells(LugarPlanilla, 5).Value = ""
                            appExcel.cells(LugarPlanilla, 6).Value = ""
                            appExcel.cells(LugarPlanilla, 7).Value = "3"
                            appExcel.cells(LugarPlanilla, 8).Value = ZZCodigo
                            appExcel.cells(LugarPlanilla, 9).Value = ZZCufe(Val(Wempresa))
                            appExcel.cells(LugarPlanilla, 10).Value = ""
                            appExcel.cells(LugarPlanilla, 11).Value = ""
                            appExcel.cells(LugarPlanilla, 12).Value = ""
                            appExcel.cells(LugarPlanilla, 13).Value = ""
                            appExcel.cells(LugarPlanilla, 14).Value = ""
                            appExcel.cells(LugarPlanilla, 15).Value = ""
                            appExcel.cells(LugarPlanilla, 16).Value = ""
                            appExcel.cells(LugarPlanilla, 17).Value = ""
                            appExcel.cells(LugarPlanilla, 18).Value = ""
                            appExcel.cells(LugarPlanilla, 19).Value = ""
                            appExcel.cells(LugarPlanilla, 20).Value = ""
                            appExcel.cells(LugarPlanilla, 21).Value = ""
                            appExcel.cells(LugarPlanilla, 22).Value = ZZFecha
                            
                        Case 3
                            LugarPlanilla = LugarPlanilla + 1
                            appExcel.cells(LugarPlanilla, 1).Value = ZZFecha
                            If ZZMovi = "S" Then
                                appExcel.cells(LugarPlanilla, 2).Value = "66"
                                    Else
                                appExcel.cells(LugarPlanilla, 2).Value = "58"
                            End If
                            appExcel.cells(LugarPlanilla, 3).Value = ZZCodsedronar
                            appExcel.cells(LugarPlanilla, 4).Value = Val(ZZCantidad)
                            appExcel.cells(LugarPlanilla, 5).Value = ""
                            appExcel.cells(LugarPlanilla, 6).Value = ""
                            appExcel.cells(LugarPlanilla, 7).Value = ""
                            appExcel.cells(LugarPlanilla, 8).Value = ""
                            appExcel.cells(LugarPlanilla, 9).Value = ZZCufe(Val(Wempresa))
                            appExcel.cells(LugarPlanilla, 10).Value = ""
                            appExcel.cells(LugarPlanilla, 11).Value = ""
                            appExcel.cells(LugarPlanilla, 12).Value = ""
                            appExcel.cells(LugarPlanilla, 13).Value = ""
                            appExcel.cells(LugarPlanilla, 14).Value = ""
                            appExcel.cells(LugarPlanilla, 15).Value = ""
                            appExcel.cells(LugarPlanilla, 16).Value = ""
                            appExcel.cells(LugarPlanilla, 17).Value = ""
                            appExcel.cells(LugarPlanilla, 18).Value = ""
                            appExcel.cells(LugarPlanilla, 19).Value = ""
                            appExcel.cells(LugarPlanilla, 20).Value = ""
                            appExcel.cells(LugarPlanilla, 21).Value = ""
                            appExcel.cells(LugarPlanilla, 22).Value = ZZFecha
                            
                            
                        Case 4
                            ZZTipo = WVectorII(Ciclo, 1)
                            ZZFecha = WVectorII(Ciclo, 2)
                            ZZCantidad = WVectorII(Ciclo, 3)
                            ZZCodigo = WVectorII(Ciclo, 4)
                            ZZMovi = WVectorII(Ciclo, 5)
                            ZZDestino = WVectorII(Ciclo, 6)
                            ZZtipomov = WVectorII(Ciclo, 7)
                            ZZCufeI = WVectorII(Ciclo, 8)
                            ZZCufeII = WVectorII(Ciclo, 9)
                            ZZCufeIII = WVectorII(Ciclo, 10)
                        
                        
                        
                            LugarPlanilla = LugarPlanilla + 1
                            appExcel.cells(LugarPlanilla, 1).Value = ZZFecha
                            If ZZMovi = "S" Then
                                appExcel.cells(LugarPlanilla, 2).Value = "48"
                                ZZLugarCufe = ZZDestino
                                    Else
                                ZZLugarCufe = ZZtipomov
                                appExcel.cells(LugarPlanilla, 2).Value = "47"
                            End If
                            appExcel.cells(LugarPlanilla, 3).Value = ZZCodsedronar
                            appExcel.cells(LugarPlanilla, 4).Value = Val(ZZCantidad)
                            appExcel.cells(LugarPlanilla, 5).Value = ""
                            appExcel.cells(LugarPlanilla, 6).Value = ""
                            appExcel.cells(LugarPlanilla, 7).Value = ""
                            appExcel.cells(LugarPlanilla, 8).Value = ""
                            If ZZMovi = "S" Then
                                appExcel.cells(LugarPlanilla, 9).Value = ZZCufe(Val(Wempresa))
                                appExcel.cells(LugarPlanilla, 10).Value = ZZCufe(ZZLugarCufe)
                                    Else
                                appExcel.cells(LugarPlanilla, 9).Value = ZZCufe(ZZLugarCufe)
                                appExcel.cells(LugarPlanilla, 10).Value = ZZCufe(Val(Wempresa))
                            End If
                            appExcel.cells(LugarPlanilla, 11).Value = ""
                            appExcel.cells(LugarPlanilla, 12).Value = ""
                            appExcel.cells(LugarPlanilla, 13).Value = ""
                            appExcel.cells(LugarPlanilla, 14).Value = ""
                            appExcel.cells(LugarPlanilla, 15).Value = ""
                            appExcel.cells(LugarPlanilla, 16).Value = ""
                            appExcel.cells(LugarPlanilla, 17).Value = ""
                            appExcel.cells(LugarPlanilla, 18).Value = ""
                            appExcel.cells(LugarPlanilla, 19).Value = ""
                            appExcel.cells(LugarPlanilla, 20).Value = ""
                            appExcel.cells(LugarPlanilla, 21).Value = ""
                            appExcel.cells(LugarPlanilla, 22).Value = ZZFecha
                            
                            
                            
                        Case Else
                    End Select
                    
                    
                    
                Next Ciclo
                    
            End If
            
        Next A
            
        appExcel.Quit
        Set appExcel = Nothing
        
    End If
    
    Call Cancela_click
    
End Sub



Private Sub Cancela_click()

    For A = 1 To 999
        With rstSedro
            .Index = "Clave"
            .Seek "=", A
            If .NoMatch = False Then
                .Delete
            End If
        End With
    Next A

    Lugar = 0
    For A = 1 To 999
        WProd = Vector(A, 1)
        If WProd <> "" Then
            Lugar = Lugar + 1
            With rstSedro
                .AddNew
                !Clave = Lugar
                !Producto = WProd
                .Update
            End With
        End If
    Next A

    With rstEmpresa
        .Close
    End With
    PrgSedronar.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Sedro
    OPEN_FILE_Empresa
End Sub

Private Sub IngresoDatos_DblClick()
    IngresoDatos.Col = 1
    IngresoDatos.Text = ""
    IngresoDatos.Col = 2
    IngresoDatos.Text = ""
    Lugar = IngresoDatos.Row
    Vector(Lugar, 1) = ""
    WProducto.SetFocus
End Sub

Private Sub WProducto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WProducto.Text = UCase(WProducto.Text)
    
        Ingre = "S"
        Lugar = 0
        For A = 1 To 1000
            If Vector(A, 1) = "" And Lugar = 0 Then
                Lugar = A
            End If
            If Vector(A, 1) = WProducto.Text Then
                Ingre = "N"
                Exit For
            End If
        Next A
                            
        If Ingre = "S" Then
            IngresoDatos.Row = Lugar
            Vector(Lugar, 1) = WProducto.Text
            IngresoDatos.Col = 1
            IngresoDatos.Text = WProducto.Text
            WArticulo = WProducto.Text
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                IngresoDatos.Col = 2
                IngresoDatos.Text = rstArticulo!Descripcion
                rstArticulo.Close
                    Else
                IngresoDatos.Col = 2
                IngresoDatos.Text = ""
            End If
            WProducto.Text = "  -   -   "
            WProducto.SetFocus
        End If
    End If
    
End Sub

Private Sub Form_Load()

    IngresoDatos.Clear
    Erase Vector
    
    IngresoDatos.ColWidth(0) = 150
    IngresoDatos.ColWidth(1) = 1600
    IngresoDatos.ColWidth(2) = 3500
    
    IngresoDatos.Row = 0
    
    IngresoDatos.Col = 1
    IngresoDatos.Text = "Articulo"
    
    IngresoDatos.Col = 2
    IngresoDatos.Text = "Descripcion"
    
    Lugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Sedronar = 1"
    ZSql = ZSql + " Order by Articulo.Codigo"
    
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Lugar = Lugar + 1
                
                    IngresoDatos.Row = Lugar
                
                    IngresoDatos.Col = 1
                    IngresoDatos.Text = rstArticulo!Codigo
                
                    IngresoDatos.Col = 2
                    IngresoDatos.Text = rstArticulo!Descripcion
                
                    .MoveNext
                    
                        Else
                    
                    Exit Do
                
                End If
            
            Loop
        End With
        rstArticulo.Close
    End If
    
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True

    IngresoDatos.Col = 1
    IngresoDatos.Row = 1
    
    Desde.Text = "01/01/2016"
    Hasta.Text = "31/03/2016"
    
    Limpia.Value = 1
    
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                    Auxi = !Codigo
                    IngresaItem = Auxi + "      " + !Descripcion
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Codigo
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
    End If
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus

End Sub

Private Sub Pantalla_Click()
    Rem Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spArticulo = "ConsultaArticulo " + "'" + Claveven$ + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Ingre = "S"
        Lugar = 0
        For A = 1 To 1000
            If Vector(A, 1) = "" And Lugar = 0 Then
                Lugar = A
            End If
            If Vector(A, 1) = rstArticulo!Codigo Then
                Ingre = "N"
                Exit For
            End If
        Next A
        If Ingre = "S" Then
            IngresoDatos.Row = Lugar
            Vector(Lugar, 1) = rstArticulo!Codigo
            IngresoDatos.Col = 1
            IngresoDatos.Text = rstArticulo!Codigo
            WArticulo = rstArticulo!Codigo
            IngresoDatos.Col = 2
            IngresoDatos.Text = rstArticulo!Descripcion
            rstArticulo.Close
        End If
    End If
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
    
    With rstArticulo
        .MoveFirst
        Do
            If .EOF = False Then
            
                da = Len(!Descripcion) - WEspacios
                
                For aa = 1 To da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                        Auxi = !Codigo
                        IngresaItem = Auxi + "    " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Codigo
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
    
    rstArticulo.Close
    
    End If
    
    End If

End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WProducto.SetFocus
    End If
End Sub

Private Sub calcula_Compras()
                    
    Erase OrdenCompras
    LugarOrden = 0
    
    Rem If WArticulo = "PC-013-100" Then Stop

    Rem PROCESA LOS LAUDOS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spLaudo = "ListaLaudoArticuloDesdeHasta " + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
        With rstLaudo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                WLiberadaAnt = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)

                WAno = Right$(rstLaudo!Fecha, 4)
                WMes = Mid$(rstLaudo!Fecha, 4, 2)
                WDia = Left$(rstLaudo!Fecha, 2)
                WCompara = WAno + WMes + WDia
                        
                If WCompara >= WDesde And WCompara <= WHasta Then
                    If rstLaudo!Articulo = WArticulo Then
                    
                        If WLiberadaAnt <> 0 Then
                            WSuma = WLiberadaAnt
                                Else
                            WSuma = WLiberada
                        End If
                    
                        LugarOrden = LugarOrden + 1
                        OrdenCompras(LugarOrden, 1) = rstLaudo!Orden
                        OrdenCompras(LugarOrden, 2) = Str$(WSuma)
                        OrdenCompras(LugarOrden, 3) = rstLaudo!Articulo
                        OrdenCompras(LugarOrden, 4) = rstLaudo!informe
                        OrdenCompras(LugarOrden, 5) = Wempresa
                    
                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        
        rstLaudo.Close
        
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Orden"
    ZSql = ZSql + " Where Orden.Tipo = 4"
    ZSql = ZSql + " and Orden.Articulo = " + "'" + WArticulo + "'"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then

        With rstOrden
    
            .MoveFirst
            
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    WAno = Right$(rstOrden!Fecha, 4)
                    WMes = Mid$(rstOrden!Fecha, 4, 2)
                    WDia = Left$(rstOrden!Fecha, 2)
                    WCompara = WAno + WMes + WDia
                            
                    If WCompara >= WDesde And WCompara <= WHasta Then
                        If rstOrden!Articulo = WArticulo Then
                            
                            LugarOrden = LugarOrden + 1
                            OrdenCompras(LugarOrden, 1) = rstOrden!Orden
                            OrdenCompras(LugarOrden, 2) = Str$(rstOrden!Cantidad)
                            OrdenCompras(LugarOrden, 3) = rstOrden!Articulo
                            OrdenCompras(LugarOrden, 4) = ""
                            OrdenCompras(LugarOrden, 5) = Wempresa
                            
                        End If
                    End If
                        
                    .MoveNext
            
                    If .EOF = True Then
                        Exit Do
                    End If
                                                                            
                Loop
            End If
            
        End With
                
        rstOrden.Close
    End If
    
    For CicloProve = 1 To LugarOrden
    
        WOrden = OrdenCompras(CicloProve, 1)
        WCantidad = OrdenCompras(CicloProve, 2)
        WArticulo = OrdenCompras(CicloProve, 3)
        WInforme = OrdenCompras(CicloProve, 4)
        XXEmpresa = OrdenCompras(CicloProve, 5)
        
        WFechaFactura = ""
        WNumeroFactura = ""
        WNroRemito = ""
        WProve = ""
        WFechaOrden = ""
        
        Rem If WArticulo = "PC-013-100" Then Stop
        
        XEmpresa = Wempresa
        Select Case Val(XXEmpresa)
            Case 1
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                Wempresa = "0002"
                txtOdbc = "Empresa02"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                Wempresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                Wempresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 5
                Wempresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 6
                Wempresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 7
                Wempresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 8
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 9
                Wempresa = "0009"
                txtOdbc = "Empresa09"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 10
                Wempresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 11
                Wempresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
    
        spOrden = "ListaOrden " + "'" + WOrden + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WProve = rstOrden!proveedor
            WFechaOrden = rstOrden!Fecha
            rstOrden.Close
        End If
                                                
        If Trim(WInforme) <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Informe = " + "'" + WInforme + "'"
            spinforme = ZSql
            Set rstInforme = db.OpenRecordset(spinforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                WNroRemito = rstInforme!remito
                rstInforme.Close
            End If
        End If
        
        Auxi = WNroRemito
        Auxi = Trim(Auxi)
        
        Select Case Val(XEmpresa)
            Case 1
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                Wempresa = "0002"
                txtOdbc = "Empresa02"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                Wempresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                Wempresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 5
                Wempresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 6
                Wempresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 7
                Wempresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 8
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 9
                Wempresa = "0009"
                txtOdbc = "Empresa09"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 10
                Wempresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 11
                Wempresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
        
        
        XEmpresa = Wempresa
        
        Select Case Val(Wempresa)
            Case 1, 3, 5, 6, 7, 10, 11
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            Case Else
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
        End Select
                                                
        If Trim(WNroRemito) <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ivacomp"
            ZSql = ZSql + " Where Ivacomp.Remito LIKE " + "'" + "%" + Auxi + "%" + "'"
            ZSql = ZSql + " Order by Ivacomp.Proveedor,Ivacomp.OrdFecha"
            Rem ZSql = ZSql + " and Ivacomp.Proveedor = " + "'" + WProve + "'"
            spIvacomp = ZSql
            Set rstIvaComp = db.OpenRecordset(spIvacomp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount > 0 Then
            
                With rstIvaComp
            
                    .MoveFirst
                    
                    If .NoMatch = False Then
                        Do
                    
                            If .EOF = True Then
                                Exit Do
                            End If
                            
                            If WProve = rstIvaComp!proveedor Then
                                WFechaFactura = rstIvaComp!Fecha
                                WNumeroFactura = rstIvaComp!Numero
                            End If
                                
                            .MoveNext
                    
                            If .EOF = True Then
                                Exit Do
                            End If
                                                                                    
                        Loop
                    End If
                    
                End With
            
                rstIvaComp.Close
            End If
        End If
        
        Select Case Val(XEmpresa)
            Case 1
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                Wempresa = "0002"
                txtOdbc = "Empresa02"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                Wempresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                Wempresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 5
                Wempresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 6
                Wempresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 7
                Wempresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 8
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 9
                Wempresa = "0009"
                txtOdbc = "Empresa09"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 10
                Wempresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 11
                Wempresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
                                                
        Entra = "S"
                        
        For Ciclo = 1 To LugarProve
            If ProveCompras(Ciclo, 1) = WProve Then
                ProveCompras(Ciclo, 2) = Str$(Val(ProveCompras(Ciclo, 2)) + Val(WCantidad))
                Entra = "N"
                Exit For
            End If
        Next Ciclo
                        
        If Entra = "S" Then
            LugarProve = LugarProve + 1
            ProveCompras(LugarProve, 1) = WProve
            ProveCompras(LugarProve, 2) = WCantidad
        End If
        
        LugarVectorII = LugarVectorII + 1
        
        WAno = Right$(WFechaFactura, 4)
        WMes = Mid$(WFechaFactura, 4, 2)
        WDia = Left$(WFechaFactura, 2)
        WComparaI = WAno + WMes + WDia
        
        WAno = Right$(Desde.Text, 4)
        WMes = Mid$(Desde.Text, 4, 2)
        WDia = Left$(Desde.Text, 2)
        WComparaII = WAno + WMes + WDia
        
        If WComparaI < WComparaII Then
            WFechaFactura = Desde.Text
        End If
        
        
        
        
        WVectorII(LugarVectorII, 1) = WProve
        WVectorII(LugarVectorII, 2) = WFechaFactura
        WVectorII(LugarVectorII, 3) = WNumeroFactura
        WVectorII(LugarVectorII, 4) = WCantidad
        WVectorII(LugarVectorII, 5) = WNroRemito
        WVectorII(LugarVectorII, 6) = WFechaOrden
        WVectorII(LugarVectorII, 7) = WOrden
        
    Next CicloProve
    
End Sub






Private Sub Proceso()

    Erase WVectorII
    LugarVectorII = 0

    
    WSalidaError = ""
    On Error GoTo Control_error
    

                
    Rem PROCESA Las compras
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Informe"
    ZSql = ZSql + " Where Informe.Articulo = " + "'" + WArticulo + "'"
    ZSql = ZSql + " and Informe.FechaOrd >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Informe.FechaOrd <= " + "'" + WHasta + "'"
    spinforme = ZSql
    Set rstInforme = db.OpenRecordset(spinforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
    
        With rstInforme
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                LugarVectorII = LugarVectorII + 1
                
                WVectorII(LugarVectorII, 1) = "1"
                WVectorII(LugarVectorII, 2) = rstInforme!Fecha
                WVectorII(LugarVectorII, 3) = Str$(rstInforme!Cantidad)
                WVectorII(LugarVectorII, 4) = rstInforme!remito
                WVectorII(LugarVectorII, 5) = rstInforme!informe
                WVectorII(LugarVectorII, 6) = rstInforme!Orden
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstInforme.Close
    End If
    
    For Ciclo = 1 To LugarVectorII
    
        WOrden = WVectorII(LugarVectorII, 6)
        
        WProveedor = ""
        spOrden = "ListaOrden" + "'" + WOrden + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WProveedor = rstOrden!proveedor
            WVectorII(Ciclo, 7) = rstOrden!proveedor
            rstOrden.Close
        End If
        
        WDEsProveedor = ""

                    
        XEmpresa = Wempresa
            
        Wempresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
        
                
        spProveedor = "ConsultaProveedores" + "'" + WProveedor + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            WVectorII(Ciclo, 8) = IIf(IsNull(RstProveedor!cufe), "", RstProveedor!cufe)
            WVectorII(Ciclo, 9) = IIf(IsNull(RstProveedor!cufeii), "", RstProveedor!cufeii)
            WVectorII(Ciclo, 10) = IIf(IsNull(RstProveedor!cufeiii), "", RstProveedor!cufeiii)
            RstProveedor.Close
        End If
        
                
        Select Case Val(XEmpresa)
            Case 1
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                Wempresa = "0002"
                txtOdbc = "Empresa02"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                Wempresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                Wempresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 5
                Wempresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 6
                Wempresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 7
                Wempresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 8
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 9
                Wempresa = "0009"
                txtOdbc = "Empresa09"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 10
                Wempresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 11
                Wempresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
            
        
        
    Next Ciclo
    
    
    
    
    
    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spHoja = "ListaHojaArticuloDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                
                If rstHoja!Tipo = "M" And rstHoja!Articulo = WArticulo Then
                
                    If WDesde <= XFec And WHasta >= XFec Then
                    
                        LugarVectorII = LugarVectorII + 1
                        
                        WVectorII(LugarVectorII, 1) = "2"
                        WVectorII(LugarVectorII, 2) = rstHoja!Fecha
                        WVectorII(LugarVectorII, 3) = Str$(rstHoja!Cantidad)
                        WVectorII(LugarVectorII, 4) = rstHoja!hoja
                        WVectorII(LugarVectorII, 11) = rstHoja!Producto

                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        
        End With
        rstHoja.Close
    End If
    
    
    
    
    
    
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovvar = "ListaMovvarArticuloDesdeHasta" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then

        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                
                XFec = Right$(rstMovvar!Fecha, 4) + Mid$(rstMovvar!Fecha, 4, 2) + Left$(rstMovvar!Fecha, 2)
                
                If rstMovvar!Tipo = "M" And rstMovvar!Articulo = WArticulo Then
                
                    If WDesde <= XFec And WHasta >= XFec And rstMovvar!Cantidad > 0 Then
                    
                        LugarVectorII = LugarVectorII + 1
                        
                        WVectorII(LugarVectorII, 1) = "3"
                        WVectorII(LugarVectorII, 2) = rstMovvar!Fecha
                        WVectorII(LugarVectorII, 3) = Str$(rstMovvar!Cantidad)
                        WVectorII(LugarVectorII, 4) = rstMovvar!Codigo
                        WVectorII(LugarVectorII, 5) = rstMovvar!Movi

                    End If
                    
                End If
                
                .MoveNext
            
                If .EOF = True Then
                    Exit Do
                End If
                                                                            
            Loop
            End If
            
        End With
        rstMovvar.Close
    End If
    
    
    
    
    
    
    Rem PROCESA LAS GUIAS DE TRASLADO INTERNOS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovguia = "ListaMovguiaArticuloDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then

        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstMovguia!Fecha, 4) + Mid$(rstMovguia!Fecha, 4, 2) + Left$(rstMovguia!Fecha, 2)
                
                If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArticulo Then
                
                    If WDesde <= XFec And WHasta >= XFec And rstMovguia!Codigo < 900000 Then
                    
                        LugarVectorII = LugarVectorII + 1
                        
                        WVectorII(LugarVectorII, 1) = "4"
                        WVectorII(LugarVectorII, 2) = rstMovguia!Fecha
                        WVectorII(LugarVectorII, 3) = Str$(rstMovguia!Cantidad)
                        WVectorII(LugarVectorII, 4) = rstMovguia!Codigo
                        WVectorII(LugarVectorII, 5) = rstMovguia!Movi
                        WVectorII(LugarVectorII, 6) = rstMovguia!Destino
                        WVectorII(LugarVectorII, 7) = rstMovguia!Tipomov

                    End If
                    
                End If
                
                .MoveNext
            
                If .EOF = True Then
                    Exit Do
                End If
                                                                            
            Loop
            End If
            
        End With
        rstMovguia.Close
    End If
    
    
    
    
    
    
    
    Rem PROCESA LAS HOJAS DE LABORATORIO
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spMovlab = "ListaMovlabArticuloDesdeHasta" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                
                XFec = Right$(rstMovlab!Fecha, 4) + Mid$(rstMovlab!Fecha, 4, 2) + Left$(rstMovlab!Fecha, 2)
                
                If rstMovlab!Tipo = "M" And rstMovlab!Articulo = WArticulo Then
                
                    If WDesde <= XFec And WHasta >= XFec Then
                    
                        LugarVectorII = LugarVectorII + 1
                        
                        WVectorII(LugarVectorII, 1) = "3"
                        WVectorII(LugarVectorII, 2) = rstMovlab!Fecha
                        WVectorII(LugarVectorII, 3) = Str$(rstMovlab!Cantidad)
                        WVectorII(LugarVectorII, 4) = rstMovlab!Codigo
                        WVectorII(LugarVectorII, 5) = rstMovlab!Movi

                    End If
                    
                End If
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
            
        End With
    End If
    
    Exit Sub
    
Control_error:
    Rem MsgBox Err.Description
    WSalidaError = "N"
    Resume Next
    
End Sub


