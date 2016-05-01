VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPasaPtDy 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pasa PT a DY"
   ClientHeight    =   5205
   ClientLeft      =   1875
   ClientTop       =   1065
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Envase 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   12
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox LoteDY 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   9
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox LotePT 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Cantidad 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin MSMask.MaskEdBox Terminado 
         Height          =   300
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
         Left            =   3600
         TabIndex        =   4
         Top             =   480
         Width           =   975
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
         Left            =   3600
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Envase"
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
         Left            =   240
         TabIndex        =   11
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Partida DY"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Partida PT"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Cantidad"
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
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Terminado"
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "PrgPasaPtDy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WInicial As Double
Private WOrden As String
Private WClave As String
Private WDescripcion As String
Private WSaldo As Double

Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovvar As Recordset
Dim spMovvar As String

Dim XParam As String
Dim ZVector(10000) As String
Dim ZVectorII(10000, 5) As String
Dim Empe(100, 2) As String
Dim ZLugar As Integer
Dim ZPago As String
Dim ZZPrecios As Double
Dim ZZCantidad As Double
Dim ZCodigo As String
Dim ZRenglon As String


Private Sub Acepta_Click()

    If Val(Envase.Text) = 0 Then
        m$ = "Se debe informar tipo de envase"
        G% = MsgBox(m$, 0, "Actualizaion de Informe de Recepcion de Materia Prima")
        Exit Sub
    End If
    
    Terminado.Text = UCase(Trim(Terminado.Text))
    
    If Terminado.Text = "  -     -   " Then
        Exit Sub
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Terminado"
    ZSql = ZSql + " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        rstTerminado.Close
            Else
        Exit Sub
    End If

    
    
    
    
    ZZArticuloDy = "DY-" + Right$(Terminado.Text, 7)
    Rem ZZArticuloDy = "DY-305-510"
    
    
    ZPasa = "N"
Rem BY NAN
 XParam = "'" + Terminado.Text + "','" _
          + LotePT.Text + "'"
 Rem END BY NAN
    XParam = "'" + LotePT.Text + "','" _
               + Terminado.Text + "'"
    spHoja = "ListaHojaProducto " + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        WSaldo = rstHoja!Saldo
        ZPasa = "S"
        rstHoja.Close
        
            Else
            
        XParam = "'" + Terminado.Text + "','" _
                + LotePT.Text + "'"
        spMovguia = "ListaMovguiaLote1 " + XParam
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
            WSaldo = rstMovguia!Saldo
            ZPasa = "S"
            rstMovguia.Close
        End If
    End If
    
    If ZPasa = "N" Then
        m$ = "Partida de Producto Terminado inexistente"
        G% = MsgBox(m$, 0, "Actualizaion de Informe de Recepcion de Materia Prima")
        Exit Sub
    End If
    
    If Val(Cantidad.Text) > WSaldo Then
        m$ = "no existe cantidad suficiente para realizarla transferencia"
        G% = MsgBox(m$, 0, "Actualizaion de Informe de Recepcion de Materia Prima")
        Exit Sub
    End If
    
    
    spMovvar = "ListaMovvarNumero"
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
        With rstMovvar
            .MoveLast
            XCodigo = Str$(rstMovvar!Codigo + 1)
        End With
        rstMovvar.Close
            Else
        XCodigo = "1"
    End If


    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then

        WSalidas = Str$(rstTerminado!Salidas + Val(Cantidad.Text))
        WEntradas = Str$(rstTerminado!Entradas)
        WDate = Date$
    
        XParam = "'" + Terminado.Text + "','" _
                + WEntradas + "','" _
                + WSalidas + "','" _
                + WDate + "'"
                               
        spTerminado = "ModificaTerminadoMovimientos " + XParam
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
        XParam = "'" + LotePT.Text + "','" _
                    + Terminado.Text + "'"
        spHoja = "ListaHojaProducto " + XParam
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            WClave = rstHoja!Clave
            WSaldo = rstHoja!Saldo - Val(Cantidad.Text)
            WDate = Date$
            rstHoja.Close
            
            XParam = "'" + WClave + "','" _
                + WDate + "','" _
                + Str$(WSaldo) + "'"
            spHoja = "ModificaHojaSaldo " + XParam
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            XParam = "'" + Terminado.Text + "','" _
                    + LotePT.Text + "'"
            spMovguia = "ListaMovguiaLote1 " + XParam
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                WClave = rstMovguia!Clave
                WSaldo = rstMovguia!Saldo - Val(Cantidad.Text)
                WDate = Date$
                rstMovguia.Close
            
                XParam = "'" + WClave + "','" _
                    + WDate + "','" _
                    + Str$(WSaldo) + "'"
                spMovguia = "ModificaMovguiaSaldo " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
        End If
        
    End If
    
    
    ZZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)

    ZCodigo = XCodigo
    Call Ceros(ZCodigo, 6)
    ZRenglon = "1"
    Call Ceros(ZRenglon, 2)
    ZFecha = ZZFecha
    ZFechaOrd = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
    ZTipo = "T"
    ZArticulo = "  -   -   "
    ZTerminado = Terminado.Text
    ZCantidad = Cantidad.Text
    ZMovi = "S"
    ZTipomov = "1"
    ZObservaciones = "Traspaso de DY Lote :" + LoteDY.Text
    ZClave = ZCodigo + ZRenglon
    ZDate = Date$
    ZMarca = ""
    ZLote = LotePT.Text
    
    XParam = "'" + ZClave + "','" _
                + ZCodigo + "','" _
                + ZRenglon + "','" _
                + ZFecha + "','" _
                + ZTipo + "','" _
                + ZArticulo + "','" _
                + ZTerminado + "','" _
                + ZCantidad + "','" _
                + ZFechaOrd + "','" _
                + ZMovi + "','" _
                + ZTipomov + "','" _
                + ZObservaciones + "','" _
                + ZDate + "','" _
                + ZMarca + "','" _
                + ZLote + "'"
               
    spMovvar = "AltaMovvar " + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)











    spArticulo = "ConsultaArticulo " + "'" + ZZArticuloDy + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        WCodigo = ZZArticuloDy
        WLaboratorio = Str$(rstArticulo!Laboratorio)
        WEntradas = Str$(rstArticulo!Entradas + Val(Cantidad.Text))
        WCosto1 = Str$(rstArticulo!Costo1)
        WCosto3 = Str$(IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3))
        WDate = Date$
        rstArticulo.Close
    
        XParam = "'" + WCodigo + "','" _
                + WLaboratorio + "','" _
                + WEntradas + "','" _
                + WDate + "','" _
                + WCosto1 + "','" _
                + WCosto3 + "'"
                               
        spArticulo = "ModificaArticuloLaudo " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        
        ZEntra = "N"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + ZZArticuloDy + "'"
        ZSql = ZSql + " and Laudo.PartiOri = " + "'" + LoteDY.Text + "'"
        ZSql = ZSql + " Order by Laudo.FechaOrd, Laudo.Laudo"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            WClave = rstLaudo!Clave
            ZEntra = "S"
            WSaldo = rstLaudo!Saldo + Val(Cantidad.Text)
            WDate = Date$
            WLote = rstLaudo!Laudo
            rstLaudo.Close
            
            XParam = "'" + WClave + "','" _
                + WDate + "','" _
                + Str$(WSaldo) + "'"
            spLaudo = "ModificaLaudoSaldo " + XParam
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Guia"
            ZSql = ZSql + " Where Guia.Articulo = " + "'" + ZZArticuloDy + "'"
            ZSql = ZSql + " and Guia.PartiOri = " + "'" + LoteDY.Text + "'"
            ZSql = ZSql + " Order by Guia.FechaOrd, Guia.Codigo"
            spMovguia = ZSql
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                WLote = rstMovguia!Lote
                ZEntra = "S"
                WClave = rstMovguia!Clave
                WSaldo = rstMovguia!Saldo - Val(Cantidad)
                WDate = Date$
                rstMovguia.Close
            
                XParam = "'" + WClave + "','" _
                    + WDate + "','" _
                    + Str$(WSaldo) + "'"
                spMovguia = "ModificaMovguiaSaldo " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
                        
        End If
            
        If ZEntra = "S" Then
                    
            ZZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        
            ZCodigo = XCodigo
            Call Ceros(ZCodigo, 6)
            ZRenglon = "2"
            Call Ceros(ZRenglon, 2)
            ZFecha = ZZFecha
            ZFechaOrd = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
            ZTipo = "M"
            ZArticulo = ZZArticuloDy
            ZTerminado = "  -     -   "
            ZCantidad = Cantidad.Text
            ZMovi = "E"
            ZTipomov = "0"
            ZObservaciones = "Traspaso desde PT Lote " + LotePT.Text
            ZClave = ZCodigo + ZRenglon
            ZDate = Date$
            ZMarca = ""
            ZLote = Str$(WLote)
            
            XParam = "'" + ZClave + "','" _
                         + ZCodigo + "','" _
                         + ZRenglon + "','" _
                         + ZFecha + "','" _
                         + ZTipo + "','" _
                         + ZArticulo + "','" _
                         + ZTerminado + "','" _
                         + ZCantidad + "','" _
                         + ZFechaOrd + "','" _
                         + ZMovi + "','" _
                         + ZTipomov + "','" _
                         + ZObservaciones + "','" _
                         + ZDate + "','" _
                         + ZMarca + "','" _
                         + ZLote + "'"
                        
            spMovvar = "AltaMovvar " + XParam
            Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                    
            WLaudo = "950000"
            spLaudo = "ListaLaudoDy"
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                With rstLaudo
                    .MoveLast
                    WLaudo = Str$(rstLaudo!Laudo + 1)
                End With
                rstLaudo.Close
                    Else
                WLaudo = "950000"
            End If
                    
            WPartida = LoteDY.Text
            WCantidad = Cantidad.Text
    
            WRenglon = "1"
            WFecha = ZZFecha
            WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
            WOrden = "0"
            WArticulo = ZZArticuloDy
            WLiberada = Cantidad.Text
            WDevuelta = "0"
            WLote = WLaudo
            WRechazo = ""
            WActualiza = "N"
            WMarca = ""
            WInforme = "0"
            WOrigenOri = ""
            WPartiOri = WPartida
            WEnvase = Envase.Text
            WTransito = ""
            WSaldoTransito = ""
    
            Auxi1 = Str$(WLaudo)
            Call Ceros(Auxi1, 6)
            Auxi2 = Str$(WRenglon)
            Call Ceros(Auxi2, 2)
    
            WClave = Auxi1 + Auxi2
            WDate = Date$
            
            Sql1 = "INSERT INTO Laudo ("
            Sql2 = "Clave ,"
            Sql3 = "Laudo ,"
            Sql4 = "Renglon ,"
            Sql5 = "Fecha ,"
            Sql6 = "FechaOrd ,"
            Sql7 = "Articulo ,"
            Sql8 = "Liberada ,"
            Sql9 = "Devuelta ,"
            Sql10 = "Orden ,"
            Sql11 = "Marca ,"
            Sql12 = "Lote ,"
            Sql13 = "Rechazo ,"
            Sql14 = "Informe ,"
            Sql15 = "Actualiza ,"
            Sql16 = "WDate ,"
            Sql17 = "Saldo ,"
            Sql18 = "Origen ,"
            Sql19 = "PartiOri ,"
            Sql20 = "Envase ,"
            Sql21 = "Transito ,"
            Sql22 = "SaldoTransito )"
            Sql23 = "Values ("
            Sql24 = "'" + WClave + "',"
            Sql25 = "'" + WLaudo + "',"
            Sql26 = "'" + WRenglon + "',"
            Sql27 = "'" + WFecha + "',"
            Sql28 = "'" + WFechaord + "',"
            Sql29 = "'" + WArticulo + "',"
            Sql30 = "'" + WLiberada + "',"
            Sql31 = "'" + WDevuelta + "',"
            Sql32 = "'" + WOrden + "',"
            Sql33 = "'" + WMarca + "',"
            Sql34 = "'" + WLote + "',"
            Sql35 = "'" + WRechazo + "',"
            Sql36 = "'" + WInforme + "',"
            Sql37 = "'" + WActualiza + "',"
            Sql38 = "'" + WDate + "',"
            Sql39 = "'" + Cantidad.Text + "',"
            Sql40 = "'" + WOrigenOri + "',"
            Sql41 = "'" + WPartiOri + "',"
            Sql42 = "'" + WEnvase + "',"
            Sql43 = "'" + WTransito + "',"
            Sql44 = "'" + WSaldoTransito + "')"
    
            spLaudo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                      Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                      Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                      Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 + _
                      Sql41 + Sql42 + Sql43 + Sql44
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Laudo SET "
            ZSql = ZSql + "NroDespacho = " + "'" + "" + "',"
            ZSql = ZSql + "Procedencia = " + "'" + "" + "'"
            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            
        End If

    End If
    
    Terminado.Text = "  -     -   "
    Cantidad.Text = ""
    LotePT.Text = ""
    LoteDY.Text = ""
    Envase.Text = ""
    
    Terminado.SetFocus
    
End Sub

Private Sub Cancela_click()
    PrgPasaPtDy.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Terminado.Text = UCase(Terminado.Text)
        Cantidad.SetFocus
    End If
End Sub

Private Sub Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LotePT.SetFocus
    End If
End Sub

Private Sub LotePT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LoteDY.SetFocus
    End If
End Sub

Private Sub LoteDY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase.SetFocus
    End If
End Sub

Private Sub Envase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Terminado.SetFocus
    End If
End Sub

Sub Form_Load()
    Terminado.Text = "  -     -   "
    Cantidad.Text = ""
    LotePT.Text = ""
    LoteDY.Text = ""
    Envase.Text = ""
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


