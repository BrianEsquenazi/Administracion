VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgVerificaLoteArti 
   AutoRedraw      =   -1  'True
   Caption         =   "Verificacion de Partidas Vencidas"
   ClientHeight    =   7875
   ClientLeft      =   150
   ClientTop       =   690
   ClientWidth     =   15075
   LinkTopic       =   "Form2"
   ScaleHeight     =   7875
   ScaleWidth      =   15075
   Begin VB.Frame Clave1 
      Caption         =   "  Ingreso de Clave de Seguridad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3600
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Cancelagraba 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Ingrese su Password"
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton AutorizoClave 
      Caption         =   "Ajuste"
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
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox TipoDatos 
      Height          =   315
      Left            =   6720
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Width           =   2415
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WImpreEtiDy.rpt"
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Lee datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   11456
      _Version        =   327680
      Rows            =   4000
      Cols            =   16
      BackColor       =   16777215
   End
End
Attribute VB_Name = "PrgVerificaLoteArti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZDesplanta(100) As String
Dim ZZDesplantaII(100) As String

Dim ZZVector(4000, 40)
Dim ZZLugar As Integer

Dim rstVerificaVtoArti As Recordset
Dim spVerificaVtoArti As String
Dim rstArticulo As Recordset
Dim spArticulo As String

Dim ZZZFechaVto As String
Dim XMes As String
Dim XAno As String
Dim ZZZVto As String
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer

Dim ZZZArticulo As String
Dim ZZZSaldo As Double
Dim ZZZSaldoII As Double
Dim ZZZEmpresa As String
Dim ZZZPartida As String
Dim ZZZLugares As Integer

Dim ZZZLee As String
Dim WGraba As String

Dim XParam As String




Private Sub cmdClose_Click()
    PrgVerificaLoteArti.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Load()

    XEmpresa = WEmpresaVerifica
    Call Conecta_Empresa

    Call Limpia_Vector
    
    Muestra.Font.Bold = True
    
    Muestra.ColWidth(0) = 200
    Muestra.ColWidth(1) = 1300
    Muestra.ColWidth(2) = 2000
    Muestra.ColWidth(3) = 800
    Muestra.ColWidth(4) = 900
    Muestra.ColWidth(5) = 800
    Muestra.ColWidth(6) = 900
    Muestra.ColWidth(7) = 1200
    Muestra.ColWidth(8) = 800
    Muestra.ColWidth(9) = 800
    Muestra.ColWidth(10) = 800
    Muestra.ColWidth(11) = 800
    Muestra.ColWidth(12) = 800
    Muestra.ColWidth(13) = 800
    Muestra.ColWidth(14) = 800
    Muestra.ColWidth(15) = 800
    
    Muestra.ColAlignment(1) = flexAlignLeftCenter
    Muestra.ColAlignment(2) = flexAlignLeftCenter
    Muestra.ColAlignment(3) = flexAlignLeftCenter
    Muestra.ColAlignment(4) = flexAlignRightCenter
    Muestra.ColAlignment(5) = flexAlignLeftCenter
    Muestra.ColAlignment(6) = flexAlignRightCenter
    Muestra.ColAlignment(7) = flexAlignLeftCenter
    Muestra.ColAlignment(8) = flexAlignRightCenter
    Muestra.ColAlignment(9) = flexAlignRightCenter
    Muestra.ColAlignment(10) = flexAlignRightCenter
    Muestra.ColAlignment(11) = flexAlignRightCenter
    Muestra.ColAlignment(12) = flexAlignRightCenter
    Muestra.ColAlignment(13) = flexAlignRightCenter
    Muestra.ColAlignment(14) = flexAlignRightCenter
    Muestra.ColAlignment(15) = flexAlignRightCenter
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "M.Prima"
    
    Muestra.Col = 2
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 3
    Muestra.Text = "Planta"
    
    Muestra.Col = 4
    Muestra.Text = "Partida"
    
    Muestra.Col = 5
    Muestra.Text = "Planta"
    
    Muestra.Col = 6
    Muestra.Text = "Hoja"
    
    Muestra.Col = 7
    Muestra.Text = "Fecha"
    
    Muestra.Col = 8
    Muestra.Text = "Stock"
    
    Muestra.Col = 9
    Muestra.Text = "SI"
    
    Muestra.Col = 10
    Muestra.Text = "SII"
    
    Muestra.Col = 11
    Muestra.Text = "SIII"
    
    Muestra.Col = 12
    Muestra.Text = "SIV"
    
    Muestra.Col = 13
    Muestra.Text = "SV"
    
    Muestra.Col = 14
    Muestra.Text = "SVI"
    
    Muestra.Col = 15
    Muestra.Text = "SVII"
    
    ZZDesplanta(1) = "SI"
    ZZDesplanta(3) = "SII"
    ZZDesplanta(5) = "SIII"
    ZZDesplanta(6) = "SIV"
    ZZDesplanta(7) = "SV"
    ZZDesplanta(10) = "SVI"
    ZZDesplanta(11) = "SVII"
    
    ZZDesplantaII(1) = "SI"
    ZZDesplantaII(2) = "SII"
    ZZDesplantaII(3) = "SIII"
    ZZDesplantaII(4) = "SIV"
    ZZDesplantaII(5) = "SV"
    ZZDesplantaII(6) = "SVI"
    ZZDesplantaII(7) = "SVII"
    
    
    TipoDatos.Clear
    
    TipoDatos.AddItem "Pendientes"
    TipoDatos.AddItem "Finalizados"
    TipoDatos.AddItem "Todos"
    
    TipoDatos.ListIndex = 0
    Rem Call Proceso_Click
    
End Sub

Private Sub Proceso_Click()

    WSalida = "N"
    DoEvents
    PrgVerificaLoteArti.Show
    
    Call Limpia_Vector
    Erase ZZVector
    ZZLugar = 0
    Lugar = 0
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM VerificaVtoArti"
    ZSql = ZSql + " Where VerificaVtoArti.Estado = 0"
    ZSql = ZSql + " Order by Verificavtoarti.codigo"
    spVerificaVtoArti = ZSql
    Set rstVerificaVtoArti = db.OpenRecordset(spVerificaVtoArti, dbOpenSnapshot, dbSQLPassThrough)
    If rstVerificaVtoArti.RecordCount > 0 Then
        With rstVerificaVtoArti
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    Lugar = Lugar + 1
                    
                    
                    Rem If rstVerificaVtoArti!Codigo = 13 Then Stop
                    
                    ZZZTipoMov = IIf(IsNull(rstVerificaVtoArti!Tipomov), "", rstVerificaVtoArti!Tipomov)
                    
                    If ZZZTipoMov = "T" Then
                        ZZVector(Lugar, 1) = rstVerificaVtoArti!Terminado
                            Else
                        ZZVector(Lugar, 1) = rstVerificaVtoArti!Articulo
                    End If
                    ZZVector(Lugar, 2) = ""
                    ZZVector(Lugar, 3) = ZZDesplantaII(rstVerificaVtoArti!Empresapartida)
                    ZZVector(Lugar, 4) = Trim(rstVerificaVtoArti!Partida)
                    
                    If UCase(Trim(rstVerificaVtoArti!Tipo)) = "PED." Then
                        ZZVector(Lugar, 5) = "PED"
                            Else
                        ZZVector(Lugar, 5) = ZZDesplanta(rstVerificaVtoArti!Empresatipo)
                    End If
                    ZZVector(Lugar, 6) = rstVerificaVtoArti!Numero
                    ZZVector(Lugar, 7) = rstVerificaVtoArti!Fecha
                    ZZVector(Lugar, 8) = Str$(rstVerificaVtoArti!stock)
                    ZZVector(Lugar, 9) = Str$(rstVerificaVtoArti!stockI)
                    ZZVector(Lugar, 10) = Str$(rstVerificaVtoArti!stockII)
                    ZZVector(Lugar, 11) = Str$(rstVerificaVtoArti!stockIII)
                    ZZVector(Lugar, 12) = Str$(rstVerificaVtoArti!stockIV)
                    ZZVector(Lugar, 13) = Str$(rstVerificaVtoArti!stockV)
                    ZZVector(Lugar, 14) = Str$(rstVerificaVtoArti!stockVI)
                    ZZVector(Lugar, 15) = Str$(rstVerificaVtoArti!stockVII)
                    ZZVector(Lugar, 16) = rstVerificaVtoArti!Tipo
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            End If
            
        End With
    
        rstVerificaVtoArti.Close
    
    End If
    
    For dada = 1 To Lugar
    Rem by nan
     Rem   If ZZVector(dada, 1) = "PT-03000-100" Then Stop
        
    
        If Len(Trim(ZZVector(dada, 1))) = 12 Then
            ZZZTipoMov = "T"
                Else
            ZZZTipoMov = "M"
        End If
    
    
        Select Case ZZZTipoMov
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + ZZVector(dada, 1) + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZVector(dada, 2) = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                XEmpresa = WEmpresa
                
                Rem verifica el vencimiento
                
                Select Case ZZVector(dada, 3)
                    Case "SI"
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SII"
                        WEmpresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SIII"
                        WEmpresa = "0005"
                        txtOdbc = "Empresa05"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SV"
                        WEmpresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SVI"
                        WEmpresa = "0010"
                        txtOdbc = "Empresa10"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SVII"
                        WEmpresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case Else
                End Select
                        
                
                
                        
                ZZZArticulo = ZZVector(dada, 1)
                zzzLote = ZZVector(dada, 4)
                        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Laudo"
                ZSql = ZSql + " Where Laudo.Articulo = " + "'" + ZZZArticulo + "'"
                ZSql = ZSql + " AND Laudo.Laudo = " + "'" + zzzLote + "'"
                spLaudo = ZSql
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    ZZZFecha = rstLaudo!Fecha
                    ZZZFechaVto = IIf(IsNull(rstLaudo!fechavencimiento), "", rstLaudo!fechavencimiento)
                    rstLaudo.Close
                End If
        
                ZZZVto = ""
                ZZZMarcaVencida = ""
                        
                ZZZOrdFecha = Right$(ZZZFecha, 4) + Mid$(ZZZFecha, 4, 2) + Left$(ZZZFecha, 2)
                
                If ZZZFechaVto <> "" And ZZZFechaVto <> "  /  /    " And ZZZFechaVto <> "00/00/0000" Then
                    Call Valida_fecha(ZZZFechaVto, Auxi)
                    If Auxi = "S" Then
                        ZZZVto = ZZZFechaVto
                    End If
                End If
                        
                If ZZZVto = "" Then
                        
                    ZZZMeses = 0
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Articulo"
                    ZSql = ZSql + " Where Codigo = " + "'" + ZZZArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        ZZZMeses = rstArticulo!Meses
                        rstArticulo.Close
                    End If
                        
                    WMes = Val(Mid$(ZZZFecha, 4, 2))
                    WAno = Val(Right$(ZZZFecha, 4))
                    For ZCiclo = 1 To ZZZMeses
                        WMes = WMes + 1
                        If WMes > 12 Then
                            WAno = WAno + 1
                            WMes = 1
                        End If
                    Next ZCiclo
                        
                    XMes = Str$(WMes)
                    XAno = Str$(WAno)
                    Call Ceros(XMes, 2)
                    Call Ceros(XAno, 4)
                    If Val(Left$(ZZZFecha, 2)) <= 30 Then
                        If Val(XMes) = 2 And Val(Left$(ZZZFecha, 2)) > 28 Then
                            ZZZVto = "28/" + XMes + "/" + XAno
                                Else
                            ZZZVto = Left$(ZZZFecha, 3) + XMes + "/" + XAno
                        End If
                            Else
                        If Val(XMes) = 2 Then
                            ZZZVto = "28/" + XMes + "/" + XAno
                                Else
                            ZZZVto = "30/" + XMes + "/" + XAno
                        End If
                    End If
                       
                End If
                    
                If ZZZVto <> "" Then
                    
                    Do
                        Call Valida_fecha(ZZZVto, Auxi)
                        If Auxi = "S" Then
                            Exit Do
                                Else
                            XFec1 = ZZZVto
                            SumaDia = 1
                            Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                            ZZZVto = XFec2
                        End If
                    Loop
                    
                    XXXFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                    ZZComparaI = XXXFecha
                    If Left$(ZZZVto, 2) > "28" Then
                        ZZComparaII = "28" + Mid$(ZZZVto, 3, 8)
                            Else
                        ZZComparaII = ZZZVto
                    End If
                    
                    ZDias = DateDiff("d", ZZComparaI, ZZComparaII)
                    
                    ZZVector(dada, 21) = ZDias
                    
                End If
                    
                    
                    
                    
                
                
                
                
                Rem verifica la hoja de produccion
                
                Select Case ZZVector(dada, 5)
                    Case "SI"
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SII"
                        WEmpresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SIII"
                        WEmpresa = "0005"
                        txtOdbc = "Empresa05"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SV"
                        WEmpresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SVI"
                        WEmpresa = "0010"
                        txtOdbc = "Empresa10"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SVII"
                        WEmpresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case Else
                End Select
                    
                        
                ZZZArticulo = ZZVector(dada, 1)
                ZZZPartida = ZZVector(dada, 4)
                ZZZHoja = ZZVector(dada, 6)
                ZZZLote1 = 0
                ZZZLote2 = 0
                ZZZLote3 = 0
                ZZZCanti1 = 0
                ZZZCanti2 = 0
                ZZZCanti3 = 0
                ZZZReal = 0
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Hoja"
                ZSql = ZSql + " Where Hoja.Articulo = " + "'" + ZZZArticulo + "'"
                ZSql = ZSql + " AND Hoja.Hoja = " + "'" + Str$(ZZZHoja) + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    ZZZLote1 = rstHoja!lote1
                    ZZZLote2 = rstHoja!lote2
                    ZZZLote3 = rstHoja!lote3
                    ZZZCanti1 = rstHoja!Canti1
                    ZZZCanti2 = rstHoja!Canti2
                    ZZZCanti3 = rstHoja!Canti3
                    ZZZReal = rstHoja!Real
                    rstHoja.Close
                End If
                
                ZZVector(dada, 22) = ZZZReal
                ZZVector(dada, 23) = ZZZLote1
                ZZVector(dada, 24) = ZZZLote2
                ZZVector(dada, 25) = ZZZLote3
                ZZVector(dada, 26) = ZZZCanti1
                ZZVector(dada, 27) = ZZZCanti2
                ZZVector(dada, 28) = ZZZCanti3
                
                ZZVector(dada, 31) = ""
                ZZVector(dada, 32) = ""
                ZZVector(dada, 33) = ""
                ZZVector(dada, 34) = ""
                ZZVector(dada, 35) = ""
                ZZVector(dada, 36) = ""
                ZZVector(dada, 37) = ""
                
                If Val(ZZVector(dada, 9)) <> 0 Then
                    ZZZEmpresa = "SI"
                    Call Calcula_Stock
                    ZZVector(dada, 31) = Str$(ZZZSaldo)
                End If
                            
                If Val(ZZVector(dada, 10)) <> 0 Then
                    ZZZEmpresa = "SII"
                    Call Calcula_Stock
                    ZZVector(dada, 32) = Str$(ZZZSaldo)
                End If
                            
                If Val(ZZVector(dada, 11)) <> 0 Then
                    ZZZEmpresa = "SIII"
                    Call Calcula_Stock
                    ZZVector(dada, 33) = Str$(ZZZSaldo)
                End If
                            
                If Val(ZZVector(dada, 12)) <> 0 Then
                    ZZZEmpresa = "SIV"
                    Call Calcula_Stock
                    ZZVector(dada, 34) = Str$(ZZZSaldo)
                End If
                            
                If Val(ZZVector(dada, 13)) <> 0 Then
                    ZZZEmpresa = "SV"
                    Call Calcula_Stock
                    ZZVector(dada, 35) = Str$(ZZZSaldo)
                End If
                            
                If Val(ZZVector(dada, 14)) <> 0 Then
                    ZZZEmpresa = "SVI"
                    Call Calcula_Stock
                    ZZVector(dada, 36) = Str$(ZZZSaldo)
                End If
                            
                If Val(ZZVector(dada, 15)) <> 0 Then
                    ZZZEmpresa = "SVII"
                    Call Calcula_Stock
                    ZZVector(dada, 37) = Str$(ZZZSaldo)
                End If
                
                ZZZSaldo = Val(ZZVector(dada, 31)) + Val(ZZVector(dada, 32)) + Val(ZZVector(dada, 33)) + Val(ZZVector(dada, 34)) + Val(ZZVector(dada, 35)) + Val(ZZVector(dada, 36)) + Val(ZZVector(dada, 37))
            
                ZZZEntra = "N"
                Select Case TipoDatos.ListIndex
                    Case 0
                        If ZZZSaldo > 0 Then
                            ZZZEntra = "S"
                        End If
                    Case 1
                        If ZZZSaldo <= 0 Then
                            ZZZEntra = "S"
                        End If
                    Case Else
                        ZZZEntra = "S"
                End Select
            
            
                
                
            Case Else
                spTerminado = "ConsultaTerminado " + "'" + ZZVector(dada, 1) + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZZVector(dada, 2) = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
                
                XEmpresa = WEmpresa
                
                Rem verifica el vencimiento
                
                Select Case ZZVector(dada, 3)
                    Case "SI"
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SII"
                        WEmpresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SIII"
                        WEmpresa = "0005"
                        txtOdbc = "Empresa05"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SV"
                        WEmpresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SVI"
                        WEmpresa = "0010"
                        txtOdbc = "Empresa10"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SVII"
                        WEmpresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case Else
                End Select
                        
                
                
                        
                ZZZArticulo = ZZVector(dada, 1)
                zzzLote = ZZVector(dada, 4)
                
                spHoja = "ListaHoja " + "'" + zzzLote + "'"
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    ZZZFecha = rstHoja!Fecha
                    ZZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "", rstHoja!FechaRevalida)
                    ZZZRevalida = IIf(IsNull(rstHoja!Revalida), "", rstHoja!Revalida)
                    ZZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "", rstHoja!MesesRevalida)
                    rstHoja.Close
                End If
                
                
                spTerminado = "ConsultaTerminado " + "'" + ZZZArticulo + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WVida = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
                    rstTerminado.Close
                End If
                
                
                WMes = Val(Mid$(ZZZFecha, 4, 2))
                WAno = Val(Right$(ZZZFecha, 4))
            
                If Val(ZZZRevalida) <> 0 Then
                    WMes = Val(Mid$(ZZZFechaRevalida, 4, 2))
                    WAno = Val(Right$(ZZZFechaRevalida, 4))
                    WVida = Val(ZZZMesesRevalida)
                End If
            
                For Ciclo = 1 To WVida
                    WMes = WMes + 1
                    If WMes > 12 Then
                        WAno = WAno + 1
                        WMes = 1
                    End If
                Next Ciclo
                WElaboracion = ZZZFecha
                Rem XFec1 = WElaboracion
                Rem SumaDia = WVida + 1
                Rem Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                If WVida <> 0 Then
                    XMes = Str$(WMes)
                    XAno = Str$(WAno)
                    Call Ceros(XMes, 2)
                    Call Ceros(XAno, 4)
                    ZZZVto = "01/" + XMes + "/" + XAno
                        Else
                    ZZZVto = XXXFecha
                End If
            
                XXXFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZComparaI = XXXFecha
                ZZComparaII = ZZZVto
                ZDias = DateDiff("d", ZZComparaI, ZZComparaII)
                ZZVector(dada, 21) = ZDias
                
                
                
                
                Rem verifica la hoja de produccion
                
                Select Case ZZVector(dada, 5)
                    Case "SI"
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SII"
                        WEmpresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SIII"
                        WEmpresa = "0005"
                        txtOdbc = "Empresa05"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SV"
                        WEmpresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SVI"
                        WEmpresa = "0010"
                        txtOdbc = "Empresa10"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case "SVII"
                        WEmpresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case Else
                End Select
                    
                        
                ZZZArticulo = ZZVector(dada, 1)
                ZZZPartida = ZZVector(dada, 4)
                ZZZHoja = ZZVector(dada, 6)
                ZZZLote1 = 0
                ZZZLote2 = 0
                ZZZLote3 = 0
                ZZZCanti1 = 0
                ZZZCanti2 = 0
                ZZZCanti3 = 0
                ZZZReal = 0
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Hoja"
                ZSql = ZSql + " Where Hoja.Terminado = " + "'" + ZZZArticulo + "'"
                ZSql = ZSql + " AND Hoja.Hoja = " + "'" + Str$(ZZZHoja) + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    ZZZLote1 = rstHoja!lote1
                    ZZZLote2 = rstHoja!lote2
                    ZZZLote3 = rstHoja!lote3
                    ZZZCanti1 = rstHoja!Canti1
                    ZZZCanti2 = rstHoja!Canti2
                    ZZZCanti3 = rstHoja!Canti3
                    ZZZReal = rstHoja!Real
                    rstHoja.Close
                End If
                
                ZZVector(dada, 22) = ZZZReal
                ZZVector(dada, 23) = ZZZLote1
                ZZVector(dada, 24) = ZZZLote2
                ZZVector(dada, 25) = ZZZLote3
                ZZVector(dada, 26) = ZZZCanti1
                ZZVector(dada, 27) = ZZZCanti2
                ZZVector(dada, 28) = ZZZCanti3
                
                ZZVector(dada, 31) = ""
                ZZVector(dada, 32) = ""
                ZZVector(dada, 33) = ""
                ZZVector(dada, 34) = ""
                ZZVector(dada, 35) = ""
                ZZVector(dada, 36) = ""
                ZZVector(dada, 37) = ""
                
                If Val(ZZVector(dada, 9)) <> 0 Then
                    ZZZEmpresa = "SI"
                    Call Calcula_StockPt
                    ZZVector(dada, 31) = Str$(ZZZSaldo)
                End If
                            
                If Val(ZZVector(dada, 10)) <> 0 Then
                    ZZZEmpresa = "SII"
                    Call Calcula_StockPt
                    ZZVector(dada, 32) = Str$(ZZZSaldo)
                End If
                            
                If Val(ZZVector(dada, 11)) <> 0 Then
                    ZZZEmpresa = "SIII"
                    Call Calcula_StockPt
                    ZZVector(dada, 33) = Str$(ZZZSaldo)
                End If
                            
                If Val(ZZVector(dada, 12)) <> 0 Then
                    ZZZEmpresa = "SIV"
                    Call Calcula_StockPt
                    ZZVector(dada, 34) = Str$(ZZZSaldo)
                End If
                            
                If Val(ZZVector(dada, 13)) <> 0 Then
                    ZZZEmpresa = "SV"
                    Call Calcula_StockPt
                    ZZVector(dada, 35) = Str$(ZZZSaldo)
                End If
                            
                If Val(ZZVector(dada, 14)) <> 0 Then
                    ZZZEmpresa = "SVI"
                    Call Calcula_StockPt
                    ZZVector(dada, 36) = Str$(ZZZSaldo)
                End If
                            
                If Val(ZZVector(dada, 15)) <> 0 Then
                    ZZZEmpresa = "SVII"
                    Call Calcula_StockPt
                    ZZVector(dada, 37) = Str$(ZZZSaldo)
                End If
                
                ZZZSaldo = Val(ZZVector(dada, 31)) + Val(ZZVector(dada, 32)) + Val(ZZVector(dada, 33)) + Val(ZZVector(dada, 34)) + Val(ZZVector(dada, 35)) + Val(ZZVector(dada, 36)) + Val(ZZVector(dada, 37))
            
                ZZZEntra = "N"
                Select Case TipoDatos.ListIndex
                    Case 0
                        If ZZZSaldo > 0 Then
                            ZZZEntra = "S"
                        End If
                    Case 1
                        If ZZZSaldo <= 0 Then
                            ZZZEntra = "S"
                        End If
                    Case Else
                        ZZZEntra = "S"
                End Select
            
            
            
        
        End Select
        
        
        
        
        If ZZZEntra = "S" Then
    
                
            ZZLugar = ZZLugar + 1
            
            Muestra.TextMatrix(ZZLugar, 1) = ZZVector(dada, 1)
            Muestra.TextMatrix(ZZLugar, 2) = ZZVector(dada, 2)
            Muestra.TextMatrix(ZZLugar, 3) = ZZVector(dada, 3)
            Muestra.TextMatrix(ZZLugar, 4) = ZZVector(dada, 4)
            Muestra.TextMatrix(ZZLugar, 5) = ZZVector(dada, 5)
            Muestra.TextMatrix(ZZLugar, 6) = ZZVector(dada, 6)
            Muestra.TextMatrix(ZZLugar, 7) = ZZVector(dada, 7)
            Muestra.TextMatrix(ZZLugar, 8) = ZZVector(dada, 8)
            Muestra.TextMatrix(ZZLugar, 9) = ZZVector(dada, 9)
            Muestra.TextMatrix(ZZLugar, 10) = ZZVector(dada, 10)
            Muestra.TextMatrix(ZZLugar, 11) = ZZVector(dada, 11)
            Muestra.TextMatrix(ZZLugar, 12) = ZZVector(dada, 12)
            Muestra.TextMatrix(ZZLugar, 13) = ZZVector(dada, 13)
            Muestra.TextMatrix(ZZLugar, 14) = ZZVector(dada, 14)
            Muestra.TextMatrix(ZZLugar, 15) = ZZVector(dada, 15)
    
            ZZZDias = Val(ZZVector(dada, 21))
            ZZZReal = Val(ZZVector(dada, 22))
            ZZZLote1 = Val(ZZVector(dada, 23))
            ZZZLote2 = Val(ZZVector(dada, 24))
            ZZZLote3 = Val(ZZVector(dada, 25))
            ZZZCanti1 = Val(ZZVector(dada, 26))
            ZZZCanti2 = Val(ZZVector(dada, 27))
            ZZZCanti3 = Val(ZZVector(dada, 28))
    
            If ZZZDias < 0 Then
                Muestra.Row = ZZLugar
                Muestra.Col = 3
                Muestra.CellBackColor = &H8080FF
                Muestra.Col = 4
                Muestra.CellBackColor = &H8080FF
                    Else
                Muestra.Row = ZZLugar
                Muestra.Col = 3
                Muestra.CellBackColor = &HC000&
                Muestra.Col = 4
                Muestra.CellBackColor = &HC000&
            End If
        
            If ZZZReal = 0 Then
                Muestra.Row = ZZLugar
                Muestra.Col = 5
                Muestra.CellBackColor = &H80FFFF
                Muestra.Col = 6
                Muestra.CellBackColor = &H80FFFF
                    Else
                If ZZZLote1 <> Val(zzzLote) And ZZZLote2 <> Val(zzzLote) And ZZZLote3 <> Val(zzzLote) Then
                    Muestra.Row = ZZLugar
                    Muestra.Col = 5
                    Muestra.CellBackColor = &H8080FF
                    Muestra.Col = 6
                    Muestra.CellBackColor = &H8080FF
                        Else
                    Muestra.Row = ZZLugar
                    Muestra.Col = 5
                    Muestra.CellBackColor = &HC000&
                    Muestra.Col = 6
                    Muestra.CellBackColor = &HC000&
                End If
            End If
    
    
    
            If Val(ZZVector(dada, 9)) <> 0 Then
                ZZZSaldoII = Val(ZZVector(dada, 9))
                ZZZSaldo = Val(ZZVector(dada, 31))
                
                Call Redondeo(ZZZSaldo)
                Call Redondeo(ZZZSaldoII)
                    
                If ZZZSaldo <= 0 Then
                    Muestra.Row = ZZLugar
                    Muestra.Col = 9
                    Muestra.CellBackColor = &H80FF80
                        Else
                    If ZZZSaldo <= ZZZSaldoII Then
                        Muestra.Row = ZZLugar
                        Muestra.Col = 9
                        Muestra.CellBackColor = &H8080FF
                            Else
                        Muestra.Row = ZZLugar
                        Muestra.Col = 9
                        Muestra.CellBackColor = &H80FFFF
                    End If
                End If
            End If
    
            If Val(ZZVector(dada, 10)) <> 0 Then
                ZZZSaldoII = Val(ZZVector(dada, 10))
                ZZZSaldo = Val(ZZVector(dada, 32))
                
                Call Redondeo(ZZZSaldo)
                Call Redondeo(ZZZSaldoII)
                    
                If ZZZSaldo <= 0 Then
                    Muestra.Row = ZZLugar
                    Muestra.Col = 10
                    Muestra.CellBackColor = &H80FF80
                        Else
                    If ZZZSaldo <= ZZZSaldoII Then
                        Muestra.Row = ZZLugar
                        Muestra.Col = 10
                        Muestra.CellBackColor = &H8080FF
                            Else
                        Muestra.Row = ZZLugar
                        Muestra.Col = 10
                        Muestra.CellBackColor = &H80FFFF
                    End If
                End If
            End If
    
            If Val(ZZVector(dada, 11)) <> 0 Then
                ZZZSaldoII = Val(ZZVector(dada, 11))
                ZZZSaldo = Val(ZZVector(dada, 33))
                
                Call Redondeo(ZZZSaldo)
                Call Redondeo(ZZZSaldoII)
                    
                If ZZZSaldo <= 0 Then
                    Muestra.Row = ZZLugar
                    Muestra.Col = 11
                    Muestra.CellBackColor = &H80FF80
                        Else
                    If ZZZSaldo <= ZZZSaldoII Then
                        Muestra.Row = ZZLugar
                        Muestra.Col = 11
                        Muestra.CellBackColor = &H8080FF
                            Else
                        Muestra.Row = ZZLugar
                        Muestra.Col = 11
                        Muestra.CellBackColor = &H80FFFF
                    End If
                End If
            End If
            
    
            If Val(ZZVector(dada, 12)) <> 0 Then
                ZZZSaldoII = Val(ZZVector(dada, 12))
                ZZZSaldo = Val(ZZVector(dada, 34))
                
                Call Redondeo(ZZZSaldo)
                Call Redondeo(ZZZSaldoII)
                    
                If ZZZSaldo <= 0 Then
                    Muestra.Row = ZZLugar
                    Muestra.Col = 12
                    Muestra.CellBackColor = &H80FF80
                        Else
                    If ZZZSaldo <= ZZZSaldoII Then
                        Muestra.Row = ZZLugar
                        Muestra.Col = 12
                        Muestra.CellBackColor = &H8080FF
                            Else
                        Muestra.Row = ZZLugar
                        Muestra.Col = 12
                        Muestra.CellBackColor = &H80FFFF
                    End If
                End If
            End If
    
            If Val(ZZVector(dada, 13)) <> 0 Then
                ZZZSaldoII = Val(ZZVector(dada, 13))
                ZZZSaldo = Val(ZZVector(dada, 35))
                
                Call Redondeo(ZZZSaldo)
                Call Redondeo(ZZZSaldoII)
                    
                If ZZZSaldo <= 0 Then
                    Muestra.Row = ZZLugar
                    Muestra.Col = 13
                    Muestra.CellBackColor = &H80FF80
                        Else
                    If ZZZSaldo <= ZZZSaldoII Then
                        Muestra.Row = ZZLugar
                        Muestra.Col = 13
                        Muestra.CellBackColor = &H8080FF
                            Else
                        Muestra.Row = ZZLugar
                        Muestra.Col = 13
                        Muestra.CellBackColor = &H80FFFF
                    End If
                End If
            End If
    
            If Val(ZZVector(dada, 14)) <> 0 Then
                ZZZSaldoII = Val(ZZVector(dada, 14))
                ZZZSaldo = Val(ZZVector(dada, 36))
                
                Call Redondeo(ZZZSaldo)
                Call Redondeo(ZZZSaldoII)
                    
                If ZZZSaldo <= 0 Then
                    Muestra.Row = ZZLugar
                    Muestra.Col = 14
                    Muestra.CellBackColor = &H80FF80
                        Else
                    If ZZZSaldo <= ZZZSaldoII Then
                        Muestra.Row = ZZLugar
                        Muestra.Col = 14
                        Muestra.CellBackColor = &H8080FF
                            Else
                        Muestra.Row = ZZLugar
                        Muestra.Col = 14
                        Muestra.CellBackColor = &H80FFFF
                    End If
                End If
            End If
    
            If Val(ZZVector(dada, 15)) <> 0 Then
                ZZZSaldoII = Val(ZZVector(dada, 15))
                ZZZSaldo = Val(ZZVector(dada, 37))
                
                Call Redondeo(ZZZSaldo)
                Call Redondeo(ZZZSaldoII)
                    
                If ZZZSaldo <= 0 Then
                    Muestra.Row = ZZLugar
                    Muestra.Col = 15
                    Muestra.CellBackColor = &H80FF80
                        Else
                    If ZZZSaldo <= ZZZSaldoII Then
                        Muestra.Row = ZZLugar
                        Muestra.Col = 15
                        Muestra.CellBackColor = &H8080FF
                            Else
                        Muestra.Row = ZZLugar
                        Muestra.Col = 15
                        Muestra.CellBackColor = &H80FFFF
                    End If
                End If
            End If
            
        End If
    
        Call Conecta_Empresa
        
        
        
    Next dada
    
    ZZZLugares = Lugar
    
    Muestra.Col = 1
    Muestra.Row = 1
    Muestra.TopRow = 1

End Sub

Private Sub Limpia_Vector()
    
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "M.Prima"
    
    Muestra.Col = 2
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 3
    Muestra.Text = "Planta"
    
    Muestra.Col = 4
    Muestra.Text = "Partida"
    
    Muestra.Col = 5
    Muestra.Text = "Planta"
    
    Muestra.Col = 6
    Muestra.Text = "Hoja"
    
    Muestra.Col = 7
    Muestra.Text = "Fecha"
    
    Muestra.Col = 8
    Muestra.Text = "Stock"
    
    Muestra.Col = 9
    Muestra.Text = "SI"
    
    Muestra.Col = 10
    Muestra.Text = "SII"
    
    Muestra.Col = 11
    Muestra.Text = "SIII"
    
    Muestra.Col = 12
    Muestra.Text = "SIV"
    
    Muestra.Col = 13
    Muestra.Text = "SV"
    
    Muestra.Col = 14
    Muestra.Text = "SVI"
    
    Muestra.Col = 15
    Muestra.Text = "SVII"
    
End Sub

Private Sub Muestra_DblClick()

    Select Case Muestra.Col
        Case 4
            Muestra.Col = 1
            ZZZArticulo = Muestra.Text
            
            Muestra.Col = 2
            ZZZDesArticulo = Muestra.Text
            
            Muestra.Col = 3
            ZZZEmpresaPartida = Muestra.Text
            
            Muestra.Col = 4
            ZZZPartida = Muestra.Text
            
            WEmpresaRevalida = ""
            Select Case ZZZEmpresaPartida
                Case "SI"
                    WEmpresaRevalida = "0001"
                    
                Case "SII"
                    WEmpresaRevalida = "0003"
                    
                Case "SIII"
                    WEmpresaRevalida = "0005"
                    
                Case "SV"
                    WEmpresaRevalida = "0007"
                    
                Case "SVI"
                    WEmpresaRevalida = "0010"
                    
                Case "SVII"
                    WEmpresaRevalida = "0011"
                    
                Case Else
            End Select
            
            If Val(ZZZPartida) <> 0 Then
                If Len(ZZZArticulo) = 10 Then
                    ZProgramaOrigen = 1
                    ZLoteRevalida = ZZZPartida
                    ZFechaRevalida = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                    ZArticuloRevalida = ZZZArticulo
                    ZDesArticuloRevalida = ZZZDesArticulo
                    PrgRevalida.Show
                        Else
                    ZProgramaOrigen = 1
                    ZFechaHoja = "  /  /    "
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Hoja"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZZPartida + "'"
                    ZSql = ZSql + " and Producto = " + "'" + ZZZArticulo + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHoja.RecordCount > 0 Then
                        ZFechaHoja = rstHoja!Fecha
                        rstHoja.Close
                    End If
                    ZLoteRevalida = ZZZPartida
                    ZFechaRevalida = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                    ZArticuloRevalida = ZZZArticulo
                    ZDesArticuloRevalida = ZZZDesArticulo
                    PrgRevalidaPt.Show
                End If
            End If
            
        Case 10
        
    End Select
    
    
    Rem PrgModifTerminado.Hide
    Rem Unload Me
    Rem PrgModPedTerminado.Show
    
End Sub

Private Sub Muestra_Click()
    If Muestra.TextMatrix(Muestra.Row, 0) = "" Then
        Muestra.TextMatrix(Muestra.Row, 0) = "X"
            Else
        Muestra.TextMatrix(Muestra.Row, 0) = ""
    End If
End Sub



Private Sub Form_Activate()
    XEmpresa = WEmpresaVerifica
    Call Conecta_Empresa
    If ZZZLee = "" Then
        Call Proceso_Click
    End If
    ZZZLee = ""
End Sub

Private Sub Calcula_Stock()

    Select Case ZZZEmpresa
        Case "SI"
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SII"
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SIII"
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SV"
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SVI"
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SVII"
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case Else
    End Select


    ZZZSaldo = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + ZZZArticulo + "'"
    ZSql = ZSql + " and Laudo.Laudo = " + "'" + ZZZPartida + "'"
    ZSql = ZSql + " and Laudo.Saldo <> 0"
    ZSql = ZSql + " Order by Laudo.Laudo"
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
                   
                    ZZZSaldo = ZZZSaldo + rstLaudo!Saldo
                   
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
    ZSql = ZSql + " FROM Guia"
    ZSql = ZSql + " Where Guia.Articulo = " + "'" + ZZZArticulo + "'"
    ZSql = ZSql + " and Guia.Lote = " + "'" + ZZZPartida + "'"
    ZSql = ZSql + " and Guia.Saldo <> 0"
    ZSql = ZSql + " Order by Guia.Codigo"
    spMovguia = ZSql
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
                   
        With rstMovguia
                   
            .MoveFirst
                   
            If .NoMatch = False Then
                Do
                   
                    If .EOF = True Then
                        Exit Do
                    End If
                
                    ZZZSaldo = ZZZSaldo + rstMovguia!Saldo
                    
                    .MoveNext
                   
                    If .EOF = True Then
                        Exit Do
                    End If
                   
                Loop
            End If
        End With
        rstMovguia.Close
    End If
    
    Call Redondeo(ZZZSaldo)

End Sub




Private Sub Calcula_StockPt()

    Select Case ZZZEmpresa
        Case "SI"
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SII"
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SIII"
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SV"
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SVI"
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SVII"
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case Else
    End Select


    ZZZSaldo = 0
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Where Hoja.Producto = " + "'" + ZZZArticulo + "'"
    ZSql = ZSql + " and Hoja.Hoja = " + "'" + ZZZPartida + "'"
    ZSql = ZSql + " and Hoja.Renglon = 1"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        ZZZSaldo = ZZZSaldo + rstHoja!Saldo
        rstHoja.Close
    End If
            
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Guia"
    ZSql = ZSql + " Where Guia.Terminado = " + "'" + ZZZArticulo + "'"
    ZSql = ZSql + " and Guia.Partida = " + "'" + ZZZPartida + "'"
    ZSql = ZSql + " and Guia.Saldo <> 0"
    ZSql = ZSql + " Order by Guia.Codigo"
    spMovguia = ZSql
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
                   
        With rstMovguia
                   
            .MoveFirst
                   
            If .NoMatch = False Then
                Do
                   
                    If .EOF = True Then
                        Exit Do
                    End If
                
                    ZZZSaldo = ZZZSaldo + rstMovguia!Saldo
                    
                    .MoveNext
                   
                    If .EOF = True Then
                        Exit Do
                    End If
                   
                Loop
            End If
        End With
        rstMovguia.Close
    End If
    
    Call Redondeo(ZZZSaldo)

End Sub


Private Sub Tipo_Click()
    Call Proceso_Click
    ZZZLee = "N"
End Sub






Private Sub Autorizo_Click()
    RowIni = Muestra.Row
    Rowfin = Muestra.RowSel
    WLugar = 0
    
    For Ciclo = RowIni To Rowfin
        Muestra.Row = Ciclo
        Muestra.Col = 8
        Muestra.Text = "Ajuste"
    Next Ciclo
    
    Muestra.Col = 1
End Sub

Private Sub AutorizoClave_Click()

    If WGraba <> "S" Then
        Call Ingresa_clave
            Else
            
        WGraba = ""

        For Ciclo = 1 To ZZZLugares
            
            If Muestra.TextMatrix(Ciclo, 0) = "X" Then
                
                ZZZArticulo = Muestra.TextMatrix(Ciclo, 1)
                ZZZPartida = Muestra.TextMatrix(Ciclo, 4)
                ZZZSI = Muestra.TextMatrix(Ciclo, 9)
                ZZZSII = Muestra.TextMatrix(Ciclo, 10)
                ZZZSIII = Muestra.TextMatrix(Ciclo, 11)
                ZZZSIV = Muestra.TextMatrix(Ciclo, 12)
                ZZZSV = Muestra.TextMatrix(Ciclo, 13)
                ZZZSVI = Muestra.TextMatrix(Ciclo, 14)
                ZZZSVII = Muestra.TextMatrix(Ciclo, 15)
                
                ZZZLargo = Len(ZZZArticulo)
                
                If Val(ZZZSI) <> 0 Then
                    ZZZEmpresa = "SI"
                    ZZSaldo = Val(ZZZSI)
                    If ZZZLargo = 10 Then
                        Call Graba_AjusteMP
                            Else
                        Call Graba_AjustePT
                    End If
                End If
                
                If Val(ZZZSII) <> 0 Then
                    ZZZEmpresa = "SII"
                    ZZZSaldo = Val(ZZZSII)
                    If ZZZLargo = 10 Then
                        Call Graba_AjusteMP
                            Else
                        Call Graba_AjustePT
                    End If
                End If
                
                If Val(ZZZSIII) <> 0 Then
                    ZZZEmpresa = "SIII"
                    ZZZSaldo = Val(ZZZSIII)
                    If ZZZLargo = 10 Then
                        Call Graba_AjusteMP
                            Else
                        Call Graba_AjustePT
                    End If
                End If
                
                If Val(ZZZSIV) <> 0 Then
                    ZZZEmpresa = "SIV"
                    ZZZSaldo = Val(ZZZSIV)
                    If ZZZLargo = 10 Then
                        Call Graba_AjusteMP
                            Else
                        Call Graba_AjustePT
                    End If
                End If
                
                If Val(ZZZSV) <> 0 Then
                    ZZZEmpresa = "SV"
                    ZZZSaldo = Val(ZZZSV)
                    If ZZZLargo = 10 Then
                        Call Graba_AjusteMP
                            Else
                        Call Graba_AjustePT
                    End If
                End If
                            
                If Val(ZZZSVI) <> 0 Then
                    ZZZEmpresa = "SVI"
                    ZZZSaldo = Val(ZZZSVI)
                    If ZZZLargo = 10 Then
                        Call Graba_AjusteMP
                            Else
                        Call Graba_AjustePT
                    End If
                End If
                
                If Val(ZZZSVII) <> 0 Then
                    ZZZEmpresa = "SVII"
                    ZZZSaldo = Val(ZZZSVII)
                    If ZZZLargo = 10 Then
                        Call Graba_AjusteMP
                            Else
                        Call Graba_AjustePT
                    End If
                End If
                
            End If

        Next Ciclo
        
        Call Proceso_Click
        
        Rem Muestra.Col = 1
            
    End If

End Sub

Sub Ingresa_clave()
    WClave.Text = ""
    Clave1.Visible = True
    WClave.SetFocus
End Sub

Private Sub CancelaGraba_Click()
    Clave1.Visible = False
End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WGraba = "N"
        WClave.Text = UCase(WClave.Text)
        If WClave.Text = "INSUMOS" Then
            WGraba = "S"
            Clave1.Visible = False
            Call AutorizoClave_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Archivo de Materias Primas")
            WClave.SetFocus
        End If
    End If
End Sub




Private Sub Graba_AjusteMP()

    XEmpresa = WEmpresa

    Select Case ZZZEmpresa
        Case "SI"
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SII"
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SIII"
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SV"
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SVI"
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SVII"
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case Else
    End Select




    WNroAjuste = 0
    
    spMovvar = "ListaMovvarNumero"
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
        With rstMovvar
            .MoveLast
            WNroAjuste = rstMovvar!Codigo + 1
        End With
        rstMovvar.Close
            Else
        WNroAjuste = 1
    End If

    Tipo = "M"
    Terminado = "  -     -   "
    Articulo = ZZZArticulo
    Cantidad = Str$(ZZZSaldo)
    Movi = "S"
    Lote = ZZZPartida
                    
    Renglon = 1
    Auxi = Str$(Renglon)
    Call Ceros(Auxi, 2)
                        
    Auxi1 = Str$(WNroAjuste)
    Call Ceros(Auxi1, 6)
                
    WCodigo = Str$(WNroAjuste)
    WRenglon = Str$(Renglon)
    ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    WFecha = ZFecha
    WFechaord = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
    WTipo = Tipo
    WArticulo = Articulo
    WTerminado = Terminado
    WCantidad = Cantidad
    WMovi = Movi
    WTipomov = "1"
    WObservaciones = "Ajuste de saldos de Materia Prima"
    WClave = Auxi1 + Auxi
    WDate = Date$
    WMarca = ""
    WLote = Lote
                
    XParam = "'" + WClave + "','" _
                 + WCodigo + "','" _
                 + WRenglon + "','" _
                 + WFecha + "','" _
                 + WTipo + "','" _
                 + WArticulo + "','" _
                 + WTerminado + "','" _
                 + WCantidad + "','" _
                 + WFechaord + "','" _
                 + WMovi + "','" _
                 + WTipomov + "','" _
                 + WObservaciones + "','" _
                 + WDate + "','" _
                 + WMarca + "','" _
                 + WLote + "'"
                         
    spMovvar = "AltaMovvar " + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
                        
    WControla = 0
    spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        
        WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
        WCodigo = Articulo
        WSalidas = Str$(rstArticulo!Salidas + Val(Cantidad))
        WEntradas = Str$(rstArticulo!Entradas)
        WDate = Date$
                
        XParam = "'" + WCodigo + "','" _
                     + WEntradas + "','" _
                     + WSalidas + "','" _
                     + WDate + "'"
                                           
        spArticulo = "ModificaArticuloMovimientos " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
        If WControla = 0 And Val(Lote) <> 0 Then
            XParam = "'" + Lote + "','" _
                         + Articulo + "'"
            spLaudo = "ListaLaudoArticulo " + XParam
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                WClave = rstLaudo!Clave
                WSaldo = Str$(rstLaudo!Saldo - Val(Cantidad))
                WDate = Date$
                rstLaudo.Close
                            
                XParam = "'" + WClave + "','" _
                             + WDate + "','" _
                             + WSaldo + "'"
                spLaudo = "ModificaLaudoSaldo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                    Else
                                
                XParam = "'" + Articulo + "','" _
                             + Lote + "'"
                spMovguia = "ListaMovguiaLote " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    WClave = rstMovguia!Clave
                    WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad))
                    WDate = Date$
                    rstMovguia.Close
                            
                    XParam = "'" + WClave + "','" _
                                 + WDate + "','" _
                                 + WSaldo + "'"
                    spMovguia = "ModificaMovguiaSaldo " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                End If
                            
            End If
        End If
    End If
    
    Call Conecta_Empresa
        
End Sub




Private Sub Graba_AjustePT()

    XEmpresa = WEmpresa

    Select Case ZZZEmpresa
        Case "SI"
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SII"
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SIII"
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SV"
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SVI"
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case "SVII"
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case Else
    End Select


    WNroAjuste = 0
    
    spMovvar = "ListaMovvarNumero"
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
        With rstMovvar
            .MoveLast
            WNroAjuste = rstMovvar!Codigo + 1
        End With
        rstMovvar.Close
            Else
        WNroAjuste = 1
    End If

    Tipo = "T"
    Articulo = "  -   -   "
    Terminado = ZZZArticulo
    Cantidad = Str$(ZZZSaldo)
    Movi = "S"
    Lote = ZZZPartida
                    
    Renglon = 1
    Auxi = Str$(Renglon)
    Call Ceros(Auxi, 2)
                        
    Auxi1 = Str$(WNroAjuste)
    Call Ceros(Auxi1, 6)
                
    WCodigo = Str$(WNroAjuste)
    WRenglon = Str$(Renglon)
    ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    WFecha = ZFecha
    WFechaord = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
    WTipo = Tipo
    WArticulo = Articulo
    WTerminado = Terminado
    WCantidad = Cantidad
    WMovi = Movi
    WTipomov = "1"
    WObservaciones = "Ajuste de saldos de Materia Prima"
    WClave = Auxi1 + Auxi
    WDate = Date$
    WMarca = ""
    WLote = Lote
                
    XParam = "'" + WClave + "','" _
                 + WCodigo + "','" _
                 + WRenglon + "','" _
                 + WFecha + "','" _
                 + WTipo + "','" _
                 + WArticulo + "','" _
                 + WTerminado + "','" _
                 + WCantidad + "','" _
                 + WFechaord + "','" _
                 + WMovi + "','" _
                 + WTipomov + "','" _
                 + WObservaciones + "','" _
                 + WDate + "','" _
                 + WMarca + "','" _
                 + WLote + "'"
                         
    spMovvar = "AltaMovvar " + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    
    WControla = 0
    spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        
        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
        WCodigo = Terminado
        WSalidas = Str$(rstTerminado!Salidas - Val(Cantidad))
        WEntradas = Str$(rstTerminado!Entradas)
        WDate = Date$
                
        XParam = "'" + WCodigo + "','" _
                     + WEntradas + "','" _
                     + WSalidas + "','" _
                     + WDate + "'"
                                           
        spTerminado = "ModificaTerminadoMovimientos " + XParam
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
        If WControla = 0 And Val(Lote) <> 0 Then
            XParam = "'" + Lote + "','" _
                         + Terminado + "'"
            spHoja = "ListaHojaProducto " + XParam
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                WClave = rstHoja!Clave
                WSaldo = Str$(rstHoja!Saldo - Val(Cantidad))
                WDate = Date$
                rstHoja.Close
                        
                XParam = "'" + WClave + "','" _
                             + WDate + "','" _
                             + WSaldo + "'"
                spHoja = "ModificaHojaSaldo " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            
                    Else
                                
                XParam = "'" + Terminado + "','" _
                             + Lote + "'"
                spMovguia = "ListaMovguiaLote1 " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    WClave = rstMovguia!Clave
                    WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad))
                    WDate = Date$
                    rstMovguia.Close
                            
                    XParam = "'" + WClave + "','" _
                                 + WDate + "','" _
                                 + WSaldo + "'"
                    spMovguia = "ModificaMovguiaSaldo " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                End If
                            
            End If
        End If
                    
    End If
    
    Call Conecta_Empresa
        
End Sub




