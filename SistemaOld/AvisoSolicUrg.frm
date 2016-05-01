VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgAvisoSolicUrgente 
   AutoRedraw      =   -1  'True
   Caption         =   "Verificacion de Vencimiento de M.P."
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Aviso 
      BackColor       =   &H00FF8080&
      Height          =   4455
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Impre 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "HAY SOLICITUDES DE MATERIA PRIMA URGENTES QUE TRAMITAR"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.CommandButton Acepta 
      Caption         =   "Aceptar"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   1200
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WPedPen.rpt"
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
Attribute VB_Name = "PrgAvisoSolicUrgente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim ZArti(10000, 10) As String
Dim Empe(10, 10) As String
Dim WSaldo As Double
Dim ZFecha As String
Dim ZFechaVto As String
Dim XMes As String
Dim XAno As String

Dim ZDias As String
Dim ZComparaI As Date
Dim ZComparaII As Date

Dim ZLaudo As String
Dim ZOrdFecha As String
Dim ZArticulo As String
Dim ZCantidad As String
Dim ZSaldo As String
Dim ZVto As String
Dim ZDesEmpresa As String
Dim ZTitulo As String
Dim ZEmpresa As String
Dim XEmpresa As String

Private Sub Acepta_Click()

    Dim Auxiliar(100, 15) As String
    Dim DiaFeriado(100) As String
    Dim XFec1 As String
    Dim XFec2 As String
    Dim SumaDia As Integer

    Erase DiaFeriado
    TotalFeriado = 0
    
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    spFeriado = "ListaFeriado"
    Set rstFeriado = db.OpenRecordset(spFeriado, dbOpenSnapshot, dbSQLPassThrough)
    If rstFeriado.RecordCount > 0 Then
        With rstFeriado
            .MoveFirst
            Do
                If .EOF = False Then
                    TotalFeriado = TotalFeriado + 1
                    DiaFeriado(TotalFeriado) = rstFeriado!Fecha
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstFeriado.Close
    End If

    ZSalida = "N"


    Erase Auxiliar
    WLugar = 0
    
    For Cicla = 1 To 11
    
        Select Case Cicla
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
        
        ZFechaEmision = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Solic"
        ZSql = ZSql + " Where Solic.Fecha = " + "'" + ZFechaEmision + "'"
        ZSql = ZSql + " and Solic.Marca <> " + "'" + "X" + "'"
        ZSql = ZSql + " Order by Solic.Clave"
        spSolic = ZSql
        Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
        If rstSolic.RecordCount > 0 Then
        
            With rstSolic
    
                .MoveFirst
                If .NoMatch = False Then
                    Do
                
                        ZSaldo = rstSolic!Cantidad - rstSolic!Entregado
                        
                        ZFecha = rstSolic!Fecha
                        ZFechaEntrega = rstSolic!Entrega
                        
                        ZFechaOrd = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
                        ZFechaEntregaOrd = Right$(ZFechaEntrega, 4) + Mid$(ZFechaEntrega, 4, 2) + Left$(ZFechaEntrega, 2)
        
                        WDias = 0
                        WSuma2 = "0"
                        WFechaHastaOrd = ZFechaEntregaOrd
                        WFechaDesdeOrd = ZFechaOrd
                        WFechaHasta = ZFechaEntrega
                        WFechaDesde = ZFecha
                                
                        If WFechaHastaOrd > WFechaDesdeOrd Then
                                
                            WSuma2 = "1"
                                
                            Do
                            
                                Feriado = "N"
                                For ZZCicla = 1 To TotalFeriado
                                    If DiaFeriado(ZZCicla) = WFechaDesde Then
                                        Feriado = "S"
                                        Exit For
                                    End If
                                Next ZZCicla
                                        
                                Rem 1 - DOMINGO
                                Rem 2 - LUNES
                                Rem 3 - MARTES
                                Rem 4 - MIERCOLES
                                Rem 5 - JUEVES
                                Rem 6 - VIERNES
                                Rem 7 - SABADO
                                XFec1 = WFechaDesde
                                strDia = Format$(XFec1, "dddd")
                                BDia = Format(XFec1, "w")
                                If BDia = 1 Or BDia = 7 Then
                                    Feriado = "S"
                                End If
                                
                                If Feriado = "N" Then
                                    WDias = WDias + 1
                                End If
                                SumaDia = 2
                                Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                                WFechaDesde = XFec2
                                            
                                If WFechaDesde = WFechaHasta Then
                                    Exit Do
                                End If
                            
                            Loop
                            
                        End If
    
                        If WDias <= 2 Then
                    
                            ZSalida = "S"
                            
                            Corte = rstSolic!Solicitud
                            Fecha = rstSolic!Fecha
                            Solicitante = rstSolic!Solicitante
                            Planta = rstSolic!Planta
                            Observaciones = rstSolic!Observaciones
                            
                            WLugar = WLugar + 1
                            Auxiliar(WLugar, 1) = Pusing("######", Str$(rstSolic!Solicitud))
                            Auxiliar(WLugar, 2) = rstSolic!Fecha
                            Auxiliar(WLugar, 3) = rstSolic!Solicitante
                            Auxiliar(WLugar, 4) = rstSolic!Planta
                            Auxiliar(WLugar, 5) = rstSolic!Articulo
                            Auxiliar(WLugar, 6) = ""
                            Auxiliar(WLugar, 7) = rstSolic!Cantidad - rstSolic!Entregado
                            Auxiliar(WLugar, 8) = rstSolic!Entrega
                            Auxiliar(WLugar, 9) = rstSolic!Obser
                            Auxiliar(WLugar, 10) = WEmpresa
                            Auxiliar(WLugar, 11) = rstSolic!Clave
                    
                        End If
                    
                        .MoveNext
                
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                End If
        
            End With
            rstSolic.Close
        End If
        
    Next Cicla
    
        
    If ZSalida = "S" Then
        PrgAvisoSolicUrgente.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For a = 1 To 10
            Beep
        Next a
        PrgAvisoSolicUrgente.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Cancela_click()
    PrgAvisoSolicUrgente.Hide
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
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

Private Sub Impre_Click()
    Call Cancela_click
End Sub
