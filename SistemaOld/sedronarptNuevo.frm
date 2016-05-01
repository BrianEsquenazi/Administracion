VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgSedronarPtNuevo 
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
      TabIndex        =   14
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
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
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
         Left            =   3000
         TabIndex        =   17
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   1440
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
         TabIndex        =   11
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
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   12
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
      ItemData        =   "sedronarptNuevo.frx":0000
      Left            =   6840
      List            =   "sedronarptNuevo.frx":0007
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
      TabIndex        =   16
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   12
      Mask            =   "AA-#####-###"
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
      TabIndex        =   15
      Top             =   5640
      Width           =   1935
   End
End
Attribute VB_Name = "PrgSedronarPtNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Vector(1000, 10) As String
Private WVectorII(1000, 10) As String
Private Clieventas(1000, 10) As String
Private OrdenCompras(1000, 10) As String

Dim rstTerminado As Recordset
Dim spTerminado As String

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
Dim WTerminado As String
Dim WEntradas As Double
Dim WSalidas As Double
Dim Stock1 As Double
Dim Stock2 As Double
Dim WCompras As Double
Dim WDesde As String
Dim WHasta As String
Dim WFechaord As String
Dim Lugar As Integer
Dim LugarClie As Integer
Dim LugarOrden As Integer
Dim WEmpre(10) As String
Dim LugarVectorII As Integer

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
        WTerminado = IngresoDatos.Text
        XCodigo = IngresoDatos.Text
        XXDescripcion = ""
                
        If WTerminado <> "" Then
                
            XEmpresa = Wempresa
                
            Select Case Val(XEmpresa)
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
                
            spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                ZZCodsedronar = IIf(IsNull(rstTerminado!CodSedronar), "", rstTerminado!CodSedronar)
                XXDescripcion = rstTerminado!Descripcion
                XXLinea = rstTerminado!linea
                rstTerminado.Close
            End If
            
            If Trim(ZZCodsedronar) = "" Then
                ZZCodsedronar = WTerminado
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
                
                Rem If Trim(ZZCodsedronar) = "" Then Stop
                
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
                ZZFechaOrd = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
                
                If Trim(ZZCufeI) <> "" And Trim(ZZCufeII) = "" And Trim(ZZCufeIII) = "" Then
                    zzcufeok = ZZCufeI
                        Else
                    zzcufeok = ZZtipomov
                End If
                
                Select Case Val(ZZTipo)
                    Case 1
                        If XXLinea = 13 Or XXLinea = 20 Or XXLinea = 21 Then
                            ZZEvento = "69"
                                Else
                            ZZEvento = "44"
                        End If
                        With rstSedronarProceso
                            .AddNew
                            !Fecha = ZZFecha
                            !Evento = ZZEvento
                            !Gtin = Trim(ZZCodsedronar)
                            !Cantidad = Val(ZZCantidad)
                            !Analitica = ""
                            !Parcial = ""
                            !Tipo = 1
                            !Numero = ZZCodigo
                            !CufeOrigen = Trim(ZZCufe(Val(Wempresa)))
                            !CufeDestino = Trim(zzcufeok)
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
                            !Clave = "T" + Trim(ZZCodsedronar) + "S" + ZZEvento + ZZFechaOrd
                            !Suma = Val(ZZCantidad) * -1
                            .Update
                        End With
                        
                
                    Case 2
                        ZZEvento = "54"
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
                            !Clave = "T" + Trim(ZZCodsedronar) + "S" + ZZEvento + ZZFechaOrd
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
                            !Clave = "T" + Trim(ZZCodsedronar) + ZZMovi + ZZEvento + ZZFechaOrd
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
                            !Tipo = 3
                            !Numero = ZZCodigo
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
                            !Clave = "T" + Trim(ZZCodsedronar) + ZZMovi + ZZEvento + ZZFechaOrd
                            If ZZMovi = "S" Then
                                !Suma = Val(ZZCantidad) * -1
                                    Else
                                !Suma = Val(ZZCantidad)
                            End If
                            .Update
                        End With
                        
                
                    Case 5
                        ZZEvento = "40"
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
                            !Clave = "T" + Trim(ZZCodsedronar) + "E" + ZZEvento + ZZFechaOrd
                            !Suma = Val(ZZCantidad)
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
            ruta = "C:\sedronar\pasasedroptsi.xls"
        Case 2
            ruta = "C:\sedronar\pasasedroptpi.xls"
        Case 3
            ruta = "C:\sedronar\pasasedroptsii.xls"
        Case 4
            ruta = "C:\sedronar\pasasedroptpii.xls"
        Case 5
            ruta = "C:\sedronar\pasasedroptsiii.xls"
        Case 6
            ruta = "C:\sedronar\pasasedroptsiv.xls"
        Case 7
            ruta = "C:\sedronar\pasasedroptsv.xls"
        Case 8
            ruta = "C:\sedronar\pasasedroptpiii.xls"
        Case 9
            ruta = "C:\sedronar\pasasedroptpv.xls"
        Case 10
            ruta = "C:\sedronar\pasasedroptsvi.xls"
        Case Else
            ruta = "C:\sedronar\pasasedroptsvii.xls"
    End Select

    If Len(Dir(ruta)) > 0 Then
    
    
        Set objLibro = appExcel.workbooks.Open(ruta)
        LugarPlanilla = 1
    
        For Ciclo = 2 To 5000
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
            WTerminado = IngresoDatos.Text
            XCodigo = IngresoDatos.Text
            XXDescripcion = ""
                    
            If WTerminado <> "" Then
                    
                XEmpresa = Wempresa
                    
                Select Case Val(XEmpresa)
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
                    
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZZCodsedronar = IIf(IsNull(rstTerminado!CodSedronar), "", rstTerminado!CodSedronar)
                    XXDescripcion = rstTerminado!Descripcion
                    XXLinea = rstTerminado!linea
                    rstTerminado.Close
                End If
                
                If Trim(ZZCodsedronar) = "" Then
                    ZZCodsedronar = WTerminado
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
                    
                    Rem If Trim(ZZCodsedronar) = "" Then Stop
                    
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
                    ZZFechaOrd = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
                    
                    If Trim(ZZCufeI) <> "" And Trim(ZZCufeII) = "" And Trim(ZZCufeIII) = "" Then
                        zzcufeok = ZZCufeI
                            Else
                        zzcufeok = ZZtipomov
                    End If
                    
                    Select Case Val(ZZTipo)
                        Case 1
                            LugarPlanilla = LugarPlanilla + 1
                            appExcel.cells(LugarPlanilla, 1).Value = ZZFecha
                            If XXLinea = 13 Or XXLinea = 20 Or XXLinea = 21 Then
                                appExcel.cells(LugarPlanilla, 2).Value = "69"
                                    Else
                                appExcel.cells(LugarPlanilla, 2).Value = "44"
                            End If
                            appExcel.cells(LugarPlanilla, 3).Value = ZZCodsedronar
                            appExcel.cells(LugarPlanilla, 4).Value = Val(ZZCantidad)
                            appExcel.cells(LugarPlanilla, 5).Value = ""
                            appExcel.cells(LugarPlanilla, 6).Value = ""
                            appExcel.cells(LugarPlanilla, 7).Value = "1"
                            appExcel.cells(LugarPlanilla, 8).Value = ZZCodigo
                            appExcel.cells(LugarPlanilla, 9).Value = ZZCufe(Val(Wempresa))
                            appExcel.cells(LugarPlanilla, 10).Value = zzcufeok
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
                    
                        Case 2
                            LugarPlanilla = LugarPlanilla + 1
                            appExcel.cells(LugarPlanilla, 1).Value = ZZFecha
                            appExcel.cells(LugarPlanilla, 2).Value = "54"
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
                    
                        Case 5
                            LugarPlanilla = LugarPlanilla + 1
                            appExcel.cells(LugarPlanilla, 1).Value = ZZFecha
                            appExcel.cells(LugarPlanilla, 2).Value = "40"
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
    PrgSedronarPtNuevo.Hide
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
            WTerminado = WProducto.Text
            spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                IngresoDatos.Col = 2
                IngresoDatos.Text = rstTerminado!Descripcion
                rstTerminado.Close
                    Else
                IngresoDatos.Col = 2
                IngresoDatos.Text = ""
            End If
            WProducto.Text = "  -     -   "
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
    IngresoDatos.Text = "Producto"
    
    IngresoDatos.Col = 2
    IngresoDatos.Text = "Descripcion"
    
    Lugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Terminado"
    ZSql = ZSql + " Where Terminado.Sedronar = 1"
    ZSql = ZSql + " Order by Terminado.Codigo"
    
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        With rstTerminado
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Lugar = Lugar + 1
                
                    IngresoDatos.Row = Lugar
                
                    IngresoDatos.Col = 1
                    IngresoDatos.Text = rstTerminado!Codigo
                
                    IngresoDatos.Col = 2
                    IngresoDatos.Text = rstTerminado!Descripcion
                
                    .MoveNext
                    
                        Else
                    
                    Exit Do
                
                End If
            
            Loop
        End With
        rstTerminado.Close
    End If
    
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True

    IngresoDatos.Col = 1
    IngresoDatos.Row = 1
    
    Desde.Text = "01/01/2016"
    Hasta.Text = "31/03/2016"
    
End Sub


Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        With rstTerminado
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
        rstTerminado.Close
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
    spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        Ingre = "S"
        Lugar = 0
        For A = 1 To 1000
            If Vector(A, 1) = "" And Lugar = 0 Then
                Lugar = A
            End If
            If Vector(A, 1) = rstTerminado!Codigo Then
                Ingre = "N"
                Exit For
            End If
        Next A
        If Ingre = "S" Then
            IngresoDatos.Row = Lugar
            Vector(Lugar, 1) = rstTerminado!Codigo
            IngresoDatos.Col = 1
            IngresoDatos.Text = rstTerminado!Codigo
            WTerminado = rstTerminado!Codigo
            IngresoDatos.Col = 2
            IngresoDatos.Text = rstTerminado!Descripcion
            rstTerminado.Close
        End If
    End If
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
    With rstTerminado
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
    
    rstTerminado.Close
    
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

Private Sub calcula_datos()

    Rem PROCESA LOS LAUDOS
    
    Rem If WTerminado = "PC-013-100" Then Stop
    
    WEntradas = 0
    WSalidas = 0
    
                
    Rem dada
    Rem PROCESA LAS ESTADISTICAS
    Rem dada
    
    Sql1 = "Select Estadistica.Marca, Estadistica.Tipo, Estadistica.Articulo, Estadistica.Cantidad, Estadistica.Fecha, Estadistica.Numero, Estadistica.Cliente, Estadistica.Lote1, Estadistica.Lote2, Estadistica.Lote3, Estadistica.Lote4, Estadistica.Lote5, Estadistica.Canti1, Estadistica.Canti2, Estadistica.Canti3, Estadistica.Canti4, Estadistica.Canti5, Estadistica.Remito, Estadistica.LoteAdicional"
    Sql2 = " FROM Estadistica"
    Sql3 = " Where Estadistica.Articulo = " + "'" + WTerminado + "'"
    Sql4 = " and Estadistica.OrdFecha > " + "'" + WFechaord + "'"
    Sql5 = " and Estadistica.Marca <> " + "'" + "X" + "'"
    spEstadistica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then

        With rstEstadistica
            .MoveFirst
            If .NoMatch = False Then
                Do
    
                    If .EOF = True Then
                        Exit Do
                    End If
        
                    WSalidas = WSalidas + rstEstadistica!Cantidad
        
                    .MoveNext
        
                    If .EOF = True Then
                        Exit Do
                    End If
        
                Loop
            End If
    
        End With

        rstEstadistica.Close

    End If
    
    
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM Hoja"
    Sql3 = " Where Hoja.Terminado = " + "'" + WTerminado + "'"
    Sql4 = " and Hoja.FechaOrd > " + "'" + WFechaord + "'"
    Sql5 = " and Hoja.Tipo = " + "'" + "T" + "'"
    spHoja = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                ADADA = rstHoja!hoja
                aADADA = rstHoja!Fecha
                asdfaADADA = rstHoja!Fechaord
                
                
                WSalidas = WSalidas + rstHoja!Cantidad
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
            
        End With
        rstHoja.Close
    End If
    
    
    
    
    Rem dada
    Rem PROCESA LAS HOJAS
    Rem dada
    
    
    Sql1 = "Select *"
    Sql2 = " FROM Hoja"
    Sql3 = " Where Hoja.Producto = " + "'" + WTerminado + "'"
    Sql4 = " and Hoja.FechaOrd > " + "'" + WFechaord + "'"
    Sql5 = " and Hoja.Renglon = " + "'" + "1" + "'"
    spHoja = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    If rstHoja!realant > 0 Then
                        WEntradas = WEntradas + rstHoja!realant
                            Else
                        WEntradas = WEntradas + rstHoja!Real
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
    
    
    
    Rem dada
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    Rem dada
    
    Sql1 = "Select Movvar.Marca, Movvar.Tipo, Movvar.Terminado, Movvar.Cantidad, Movvar.Fecha, Movvar.Codigo, Movvar.Movi, Movvar.Lote, Movvar.TipoMov, Movvar.Observaciones"
    Sql2 = " FROM Movvar"
    Sql3 = " Where Movvar.Terminado = " + "'" + WTerminado + "'"
    Sql4 = " and Movvar.FechaOrd > " + "'" + WFechaord + "'"
    Sql5 = " and Movvar.Tipo = " + "'" + "T" + "'"
    spMovvar = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
    
        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WTerminado = rstMovvar!terminado
                WCantidad = rstMovvar!Cantidad
                WFecha = rstMovvar!Fecha
                WCodigo = rstMovvar!Codigo
                WMovi = rstMovvar!Movi
                WLote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                If WMovi = "E" Then
                    WEntradas = WEntradas + WCantidad
                        Else
                    WSalidas = WSalidas + WCantidad
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
    
    
    
    
    Rem dada
    Rem PROCESA LAS GUIAS DE TRASLADO INTERNO
    Rem dada
    
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM Guia"
    Sql3 = " Where Guia.Terminado = " + "'" + WTerminado + "'"
    Sql4 = " and Guia.FechaOrd > " + "'" + WFechaord + "'"
    Sql5 = " and Guia.Tipo = " + "'" + "T" + "'"
    spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
    
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    WTerminado = rstMovguia!terminado
                    WCantidad = rstMovguia!Cantidad
                    WFecha = rstMovguia!Fecha
                    WCodigo = rstMovguia!Codigo
                    WMovi = rstMovguia!Movi
                    WDestino = rstMovguia!Destino
                    WTipomov = rstMovguia!Tipomov

                    If WMovi = "E" Then
                        WEntradas = WEntradas + WCantidad
                            Else
                        WSalidas = WSalidas + WCantidad
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
    
    
    
    Rem dada
    Rem PROCESA LOS MOVIMIENTOS DE LABORATORIO
    Rem dada
    
    
    Sql1 = "Select *"
    Sql2 = " FROM MovLab"
    Sql3 = " Where MovLab.Terminado = " + "'" + WTerminado + "'"
    Sql4 = " and MovLab.FechaOrd > " + "'" + WFechaord + "'"
    Sql5 = " and MovLab.Tipo = " + "'" + "T" + "'"
    spMovlab = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WTerminado = rstMovlab!terminado
                WCantidad = rstMovlab!Cantidad
                WFecha = rstMovlab!Fecha
                WCodigo = rstMovlab!Codigo
                WMovi = rstMovlab!Movi
                
                If WMovi = "E" Then
                    WEntradas = WEntradas + WCantidad
                        Else
                    WSalidas = WSalidas + WCantidad
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
            
        End With
        rstMovlab.Close
    End If
    
        
    
    
    
    
        
        
        
        
    
    
End Sub


Private Sub Calcula_Ventas()
                    
    
    Sql1 = "Select Estadistica.Tipo, Estadistica.Articulo, Estadistica.Cantidad, Estadistica.Fecha, Estadistica.Numero, Estadistica.Cliente, Estadistica.Lote1, Estadistica.Lote2, Estadistica.Lote3, Estadistica.Lote4, Estadistica.Lote5, Estadistica.Canti1, Estadistica.Canti2, Estadistica.Canti3, Estadistica.Canti4, Estadistica.Canti5, Estadistica.Remito, Estadistica.LoteAdicional"
    Sql2 = " FROM Estadistica"
    Sql3 = " Where Estadistica.Articulo = " + "'" + WTerminado + "'"
    Sql4 = " and Estadistica.OrdFecha >= " + "'" + WDesde + "'"
    Sql5 = " and Estadistica.OrdFecha <= " + "'" + WHasta + "'"
    spEstadistica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then

        With rstEstadistica
            .MoveFirst
            If .NoMatch = False Then
                Do
    
                    If .EOF = True Then
                        Exit Do
                    End If
        
                    WFechaFactura = rstEstadistica!Fecha
                    WNumeroFactura = rstEstadistica!Numero
                    WNroRemito = rstEstadistica!remito
                    WCliente = rstEstadistica!cliente
                    WCantidad = rstEstadistica!Cantidad
                                                            
                    Entra = "S"
                                    
                    For Ciclo = 1 To LugarClie
                        If Clieventas(Ciclo, 1) = WCliente Then
                            Clieventas(Ciclo, 2) = Str$(Val(Clieventas(Ciclo, 2)) + Val(WCantidad))
                            Entra = "N"
                            Exit For
                        End If
                    Next Ciclo
                                    
                    If Entra = "S" Then
                        LugarClie = LugarClie + 1
                        Clieventas(LugarClie, 1) = WCliente
                        Clieventas(LugarClie, 2) = WCantidad
                    End If
                    
                    LugarVectorII = LugarVectorII + 1
                    WVectorII(LugarVectorII, 1) = WCliente
                    WVectorII(LugarVectorII, 2) = WFechaFactura
                    WVectorII(LugarVectorII, 3) = WNumeroFactura
                    WVectorII(LugarVectorII, 4) = WCantidad
                    WVectorII(LugarVectorII, 5) = WNumeroFactura
                    WVectorII(LugarVectorII, 6) = WFechaFactura
                    WVectorII(LugarVectorII, 7) = ""
        
                    .MoveNext
        
                    If .EOF = True Then
                        Exit Do
                    End If
        
                Loop
            End If
    
        End With

        rstEstadistica.Close

    End If
    
End Sub







Private Sub Proceso()

    Erase WVectorII
    LugarVectorII = 0

    
    WSalidaError = ""
    On Error GoTo Control_error
    

                
    
    
    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spHoja = "ListaHojaTerminadoDesdeHasta" + XParam
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
                
                If rstHoja!Tipo = "T" And rstHoja!terminado = WTerminado Then
                
                    If WDesde <= XFec And WHasta >= XFec Then
                    
                        LugarVectorII = LugarVectorII + 1
                        
                        WVectorII(LugarVectorII, 1) = "2"
                        WVectorII(LugarVectorII, 2) = rstHoja!Fecha
                        WVectorII(LugarVectorII, 3) = Str$(rstHoja!Cantidad)
                        WVectorII(LugarVectorII, 4) = rstHoja!hoja

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
    
    
    
    
    
    
    
    
    
    
    
    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spHoja = "ListaHojaProductoDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    If Val(rstHoja!Renglon) = 1 Then
                        
                        XFec = Right$(rstHoja!Fechaing, 4) + Mid$(rstHoja!Fechaing, 4, 2) + Left$(rstHoja!Fechaing, 2)
                        If WDesde <= XFec And WHasta >= XFec Then
                            
                            LugarVectorII = LugarVectorII + 1
                            
                            WCantidad = rstHoja!Real
                            WCantidadII = IIf(IsNull(rstHoja!realant), "0", rstHoja!realant)
                            WCantidadIII = WCantidad + WCantidadII
                            
                            WVectorII(LugarVectorII, 1) = "5"
                            WVectorII(LugarVectorII, 2) = rstHoja!Fecha
                            WVectorII(LugarVectorII, 3) = Str$(WCantidadIII)
                            WVectorII(LugarVectorII, 4) = rstHoja!hoja
                            
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
    
    XParam = "'" + WTerminado + "','" _
                + WTerminado + "'"
    spMovvar = "ListaMovvarTerminadoDesdeHasta" + XParam
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
                
                If rstMovvar!Tipo = "T" And rstMovvar!terminado = WTerminado Then
                
                    If WDesde <= XFec And WHasta >= XFec Then
                    
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
    
    XParam = "'" + WTerminado + "','" _
                + WTerminado + "'"
    spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
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
                
                If rstMovguia!Tipo = "T" And rstMovguia!terminado = WTerminado Then
                
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
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    
    spMovlab = "ListaMovlabTerminadoDesdeHasta" + XParam
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
                
                If rstMovlab!Tipo = "T" And rstMovlab!terminado = WTerminado Then
                
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
    
    
    
    
    
    
    Rem dada
    Rem PROCESA LAS ESTADISTICAS
    Rem dada
    
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Where Estadistica.Articulo = " + "'" + WTerminado + "'"
    ZSql = ZSql + " and Estadistica.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Estadistica.OrdFecha <= " + "'" + WHasta + "'"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then

        With rstEstadistica
            .MoveFirst
            If .NoMatch = False Then
                Do
    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    If rstEstadistica!Marca <> "X" Then
                        LugarVectorII = LugarVectorII + 1
                        
                        WVectorII(LugarVectorII, 1) = "1"
                        WVectorII(LugarVectorII, 2) = rstEstadistica!Fecha
                        WVectorII(LugarVectorII, 3) = rstEstadistica!Cantidad
                        If rstEstadistica!Numero < 200000 Then
                            WVectorII(LugarVectorII, 4) = rstEstadistica!Numero - 100000
                                Else
                            WVectorII(LugarVectorII, 4) = rstEstadistica!Numero
                        End If
                        WVectorII(LugarVectorII, 5) = rstEstadistica!cliente
                    End If
        
                    .MoveNext
        
                    If .EOF = True Then
                        Exit Do
                    End If
        
                Loop
            End If
    
        End With

        rstEstadistica.Close

    End If
    
    Select Case Val(XEmpresa)
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
    
    For Ciclo = 1 To LugarVectorII
    
        WCliente = WVectorII(Ciclo, 5)
        
        spCliente = "ConsultaCliente " + "'" + WCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WVectorII(Ciclo, 8) = IIf(IsNull(rstCliente!cufe), "", rstCliente!cufe)
            WVectorII(Ciclo, 9) = IIf(IsNull(rstCliente!cufeii), "", rstCliente!cufeii)
            WVectorII(Ciclo, 10) = IIf(IsNull(rstCliente!cufeiii), "", rstCliente!cufeiii)
            rstCliente.Close
        End If
        
    Next Ciclo
    
    
    
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
    
    
    
    
    
    Exit Sub
    
Control_error:
    Rem MsgBox Err.Description
    WSalidaError = "N"
    Resume Next
    
End Sub




