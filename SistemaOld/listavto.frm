VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaVto 
   Caption         =   "Verificacion de Vencimientos de Materia Prima"
   ClientHeight    =   6555
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6555
   ScaleWidth      =   8145
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
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
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   6015
      Begin VB.TextBox Dias 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   10
         TabIndex        =   16
         Text            =   " "
         Top             =   1800
         Width           =   855
      End
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
         Height          =   420
         Left            =   4560
         MaskColor       =   &H00000000&
         TabIndex        =   12
         Top             =   1680
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2280
         TabIndex        =   11
         Top             =   1320
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
         TabIndex        =   1
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
         Left            =   2760
         TabIndex        =   10
         Top             =   2400
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
         Left            =   1200
         TabIndex        =   9
         Top             =   2400
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
         Left            =   4560
         MaskColor       =   &H00000000&
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
         Height          =   375
         Left            =   4560
         MaskColor       =   &H00000000&
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   300
         Left            =   2280
         TabIndex        =   0
         Top             =   360
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Dias"
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
         Left            =   600
         TabIndex        =   15
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
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
         Left            =   600
         TabIndex        =   14
         Top             =   360
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
         Left            =   600
         TabIndex        =   6
         Top             =   1320
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
         Left            =   600
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6840
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      ItemData        =   "listavto.frx":0000
      Left            =   120
      List            =   "listavto.frx":0007
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "PrgListaVto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim ZArti(10000, 10) As String
Dim Empe(12, 10) As String
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
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer

            
Private Sub Acepta_Click()

    ZSql = "DELETE ListaVencimiento"
    spListaVencimiento = ZSql
    Set rstListaVencimiento = db.OpenRecordset(spListaVencimiento, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            ZDesEmpresa = !Nombre
        End If
    End With
    
    ZTitulo = "del " + Desde.Text + " al " + Hasta.Text + " Dias : " + Dias.Text


    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    WFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    
    Erase ZArti
    ZLugar = 0
    
                
    Rem PROCESA LOS LAUDOS
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Articulo >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and Articulo <= " + "'" + Hasta.Text + "'"
    ZSql = ZSql + " and Saldo <> 0"
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
                
                If rstLaudo!Marca = "X" And rstLaudo!Saldo = 0 Then
                
                        Else
                    
                    WArticulo = rstLaudo!Articulo
                    WCantidad = rstLaudo!Liberada
                    WFecha = rstLaudo!Fecha
                    WLaudo = rstLaudo!Laudo
                    WPartiOri = rstLaudo!partiori
                    WOrden = rstLaudo!Orden
                    WDevuelta = IIf(IsNull(rstLaudo!Devuelta), "0", rstLaudo!Devuelta)
                    WRechazo = IIf(IsNull(rstLaudo!Rechazo), "0", rstLaudo!Rechazo)
                    WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                    Call Redondeo(WSaldo)
                    WVencimiento = IIf(IsNull(rstLaudo!fechavencimiento), "", rstLaudo!fechavencimiento)
                    WOrdVencimiento = IIf(IsNull(rstLaudo!OrdFechaVencimiento), "", rstLaudo!OrdFechaVencimiento)
                    
                    If WSaldo <> 0 Then
                        
                        ZLugar = ZLugar + 1
                        ZArti(ZLugar, 1) = WLaudo
                        ZArti(ZLugar, 2) = WArticulo
                        ZArti(ZLugar, 3) = Str$(WCantidad)
                        ZArti(ZLugar, 4) = Str$(WSaldo)
                        Select Case Val(WEmpresa)
                            Case 1
                                ZArti(ZLugar, 5) = "Pta SI"
                            Case 2
                                ZArti(ZLugar, 5) = "Pta PI"
                            Case 3
                                ZArti(ZLugar, 5) = "Pta SII"
                            Case 4
                                ZArti(ZLugar, 5) = "Pta PII"
                            Case 5
                                ZArti(ZLugar, 5) = "Pta SIII"
                            Case 6
                                ZArti(ZLugar, 5) = "Pta SVI"
                            Case 7
                                ZArti(ZLugar, 5) = "Pta SV"
                            Case 8
                                ZArti(ZLugar, 5) = "Pta PIII"
                            Case 9
                                ZArti(ZLugar, 5) = "Pta PIV"
                            Case 10
                                ZArti(ZLugar, 5) = "Pta SVI"
                            Case 11
                                ZArti(ZLugar, 5) = "Pta SVII"
                            Case Else
                        End Select
                        ZArti(ZLugar, 6) = WVencimiento
                        ZArti(ZLugar, 7) = WOrdVencimiento
                        ZArti(ZLugar, 8) = WFecha
                    
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
    
    
    Rem PROCESA LAS GUIAS DE TRASLADO INTERNOS
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Guia"
    ZSql = ZSql + " Where Articulo >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and Articulo <= " + "'" + Hasta.Text + "'"
    ZSql = ZSql + " and Saldo <> 0"
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
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 And rstMovguia!Codigo > 900000 Then
                
                        Else
                        
                    If rstMovguia!Tipo = "M" Then
                    
                        WArticulo = rstMovguia!Articulo
                        WCantidad = rstMovguia!Cantidad
                        WFecha = rstMovguia!Fecha
                        WCodigo = rstMovguia!Codigo
                        WMovi = rstMovguia!Movi
                        WDestino = rstMovguia!Destino
                        WTipomov = rstMovguia!Tipomov
                        WLaudo = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(WSaldo)
                        
                        If WSaldo <> 0 Then
                        
                            ZLugar = ZLugar + 1
                            ZArti(ZLugar, 1) = WLaudo
                            ZArti(ZLugar, 2) = WArticulo
                            ZArti(ZLugar, 3) = Str$(WCantidad)
                            ZArti(ZLugar, 4) = Str$(WSaldo)
                            ZArti(ZLugar, 5) = ""
                            ZArti(ZLugar, 6) = ""
                            ZArti(ZLugar, 7) = ""
                            ZArti(ZLugar, 8) = ""
                    
                        End If
                        
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
    
    
    For Ciclo = 1 To ZLugar
    
        ZVto = ""
        
        ZLaudo = ZArti(Ciclo, 1)
        ZArticulo = ZArti(Ciclo, 2)
        ZCantidad = ZArti(Ciclo, 3)
        ZSaldo = ZArti(Ciclo, 4)
        ZEmpresa = ZArti(Ciclo, 5)
        ZFechaVto = ZArti(Ciclo, 6)
        ZFecha = ZArti(Ciclo, 8)
        
        If Trim(ZEmpresa) = "" Then
        
            XEmpresa = WEmpresa
    
            Empe(1, 1) = "0001"
            Empe(1, 2) = "Empresa01"
            Empe(2, 1) = "0002"
            Empe(2, 2) = "Empresa02"
            Empe(3, 1) = "0003"
            Empe(3, 2) = "Empresa03"
            Empe(4, 1) = "0004"
            Empe(4, 2) = "Empresa04"
            Empe(5, 1) = "0005"
            Empe(5, 2) = "Empresa05"
            Empe(6, 1) = "0006"
            Empe(6, 2) = "Empresa06"
            Empe(7, 1) = "0007"
            Empe(7, 2) = "Empresa07"
            Empe(8, 1) = "0008"
            Empe(8, 2) = "Empresa08"
            Empe(9, 1) = "0009"
            Empe(9, 2) = "Empresa09"
            Empe(10, 1) = "0010"
            Empe(10, 2) = "Empresa10"
            Empe(11, 1) = "0011"
            Empe(11, 2) = "Empresa11"
            
            XHasta = 11
                
            For Ciclo2 = 1 To XHasta
    
                WEmpresa = Empe(Ciclo2, 1)
                txtOdbc = Empe(Ciclo2, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Laudo"
                ZSql = ZSql + " Where Laudo = " + "'" + ZLaudo + "'"
                ZSql = ZSql + " and Articulo = " + "'" + ZArticulo + "'"
                spLaudo = ZSql
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                        
                    ZFecha = rstLaudo!Fecha
                    ZFechaVto = IIf(IsNull(rstLaudo!fechavencimiento), "", rstLaudo!fechavencimiento)
                    
                    Select Case Val(WEmpresa)
                        Case 1
                            ZEmpresa = "Pta SI"
                        Case 2
                            ZEmpresa = "Pta PI"
                        Case 3
                            ZEmpresa = "Pta SII"
                        Case 4
                            ZEmpresa = "Pta PII"
                        Case 5
                            ZEmpresa = "Pta SIII"
                        Case 6
                            ZEmpresa = "Pta SIV"
                        Case 7
                            ZEmpresa = "Pta SV"
                        Case 8
                            ZEmpresa = "Pta PIII"
                        Case 9
                            ZEmpresa = "Pta PIV"
                        Case 10
                            ZEmpresa = "Pta SVI"
                        Case 11
                            ZEmpresa = "Pta SVII"
                        Case Else
                    End Select
                        
                    rstLaudo.Close
                    Exit For
        
                End If
            
            Next Ciclo2
            
            Call Conecta_Empresa
        
        End If
                    
                    
        ZVto = ""
        ZOrdFecha = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
        If ZFechaVto <> "" And ZFechaVto <> "  /  /    " And ZFechaVto <> "00/00/0000" Then
            Call Valida_fecha(ZFechaVto, Auxi)
            If Auxi = "S" Then
                ZVto = ZFechaVto
            End If
        End If
        
        If ZVto = "" Then
        
            ZMeses = 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Codigo = " + "'" + ZArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZMeses = rstArticulo!Meses
                rstArticulo.Close
            End If
        
            WMes = Val(Mid$(ZFecha, 4, 2))
            WAno = Val(Right$(ZFecha, 4))
            For ZCiclo = 1 To ZMeses
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
            If Val(Left$(ZFecha, 2)) <= 30 Then
                If Val(XMes) = 2 And Val(Left$(ZFecha, 2)) > 28 Then
                    ZVto = "28/" + XMes + "/" + XAno
                        Else
                    ZVto = Left$(ZFecha, 3) + XMes + "/" + XAno
                End If
                    Else
                If Val(XMes) = 2 Then
                    ZVto = "28/" + XMes + "/" + XAno
                        Else
                    ZVto = "30/" + XMes + "/" + XAno
                End If
            End If
            
        End If
        
        If ZFecha <> "" Then
        
        Do
            Call Valida_fecha(ZVto, Auxi)
            If Auxi = "S" Then
                Exit Do
                    Else
                XFec1 = ZVto
                SumaDia = 1
                Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                ZVto = XFec2
            End If
        Loop
        
        Rem WFechaActual = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        Rem WFechaVto = Right$(ZVto, 4) + Mid$(ZVto, 4, 2) + Left$(ZVto, 2)
        
        Select Case Val(Mid$(ZVto, 4, 2))
            Case 2
                If Val(Mid$(ZVto, 1, 2)) > 28 Then
                    ZVto = "28" + Mid$(ZVto, 3, 6)
                End If
            Case Else
                If Val(Mid$(ZVto, 1, 2)) > 31 Then
                    ZVto = "30" + Mid$(ZVto, 3, 6)
                End If
        End Select
            
        ZComparaI = Fecha.Text
        ZComparaII = ZVto
        ZDias = DateDiff("d", ZComparaI, ZComparaII)
        
        If Val(ZDias) < Val(Dias.Text) Then
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ListaVencimiento ("
            ZSql = ZSql + "Laudo ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Liberada ,"
            ZSql = ZSql + "Saldo ,"
            ZSql = ZSql + "FechaVencimiento ,"
            ZSql = ZSql + "Dias ,"
            ZSql = ZSql + "DesEmpresa ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "Origen )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZLaudo + "',"
            ZSql = ZSql + "'" + ZFecha + "',"
            ZSql = ZSql + "'" + ZOrdFecha + "',"
            ZSql = ZSql + "'" + ZArticulo + "',"
            ZSql = ZSql + "'" + ZCantidad + "',"
            ZSql = ZSql + "'" + ZSaldo + "',"
            ZSql = ZSql + "'" + ZVto + "',"
            ZSql = ZSql + "'" + ZDias + "',"
            ZSql = ZSql + "'" + ZDesEmpresa + "',"
            ZSql = ZSql + "'" + ZTitulo + "',"
            ZSql = ZSql + "'" + ZEmpresa + "')"
        
            spListaVencimiento = ZSql
            Set rstListaVencimiento = db.OpenRecordset(spListaVencimiento, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        End If
        
    Next Ciclo
        
            
    
    
    
    
    
    

    Listado.WindowTitle = "Verificacion de Vencimientos de Materia Prima"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT ListaVencimiento.Laudo, ListaVencimiento.Fecha, ListaVencimiento.Articulo, ListaVencimiento.Liberada, ListaVencimiento.Saldo, ListaVencimiento.FechaVencimiento, ListaVencimiento.Dias, ListaVencimiento.DesEmpresa, ListaVencimiento.Titulo, ListaVencimiento.Origen, " _
            + "Articulo.Descripcion " _
            + "From " _
            + DSQ + ".dbo.ListaVencimiento ListaVencimiento, " _
            + DSQ + ".dbo.Articulo Articulo " _
            + "Where " _
            + "ListaVencimiento.Articulo = Articulo.Codigo AND " _
            + "ListaVencimiento.Laudo >= 0 AND " _
            + "ListaVencimiento.Laudo <= 999999"
    Listado.Connect = Connect()
    
    Rem Listado.GroupSelectionFormula = "{Articulo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.ReportFileName = "WListaVencimiento.rpt"
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    PrgListaVto.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
        End If
    End If
End Sub

Private Sub Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Dias.SetFocus
    End If
End Sub

Private Sub Dias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
End Sub

Sub Form_Load()
    
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    Fecha.Text = "  /  /    "
    Dias.Text = ""
    
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    XIndice = 0
    
    Select Case XIndice
        Case 0
            spArticulo = "ListaArticulo"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Left$(rstArticulo!Codigo, 2) = "DY" Or Left$(rstArticulo!Codigo, 2) = "DW" Or Left$(rstArticulo!Codigo, 2) = "DS" Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstArticulo.Close
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus

End Sub


Private Sub pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    
    Select Case XIndice
        Case 0
            With rstArticulo
            
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                    Desde.Text = rstArticulo!Codigo
                    Hasta.Text = rstArticulo!Codigo
                End If
            End With
            Desde.SetFocus
        Case Else
    End Select
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spArticulo = "ListaArticuloConsulta"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
    
    With rstArticulo
        .MoveFirst
        Do
            If .EOF = False Then
            
                If Left$(rstArticulo!Codigo, 2) = "DY" Or Left$(rstArticulo!Codigo, 2) = "DW" Or Left$(rstArticulo!Codigo, 2) = "DS" Then
                    Da = Len(rstArticulo!Descripcion) - WEspacios
                    For Aaa = 1 To Da
                        If Left$(Ayuda.Text, WEspacios) = Mid$(rstArticulo!Descripcion, Aaa, WEspacios) Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                            Exit For
                        End If
                    Next Aaa
                End If
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











