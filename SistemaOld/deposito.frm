VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDeposito 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingresos de Depositos"
   ClientHeight    =   7830
   ClientLeft      =   30
   ClientTop       =   435
   ClientWidth     =   11880
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7830
   ScaleWidth      =   11880
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4440
      TabIndex        =   40
      Top             =   4320
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   12582912
      ForeColor       =   -2147483643
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
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   6
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   5
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   4
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   3
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3600
      TabIndex        =   32
      Top             =   3840
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   4080
      TabIndex        =   31
      Top             =   3360
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto1 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4320
      TabIndex        =   30
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   855
      Left            =   8520
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton LeeLectora 
      Caption         =   "Lectora"
      Height          =   300
      Left            =   3600
      TabIndex        =   27
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Lectora 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   26
      Top             =   2520
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "alta duplicado"
      Height          =   735
      Left            =   8640
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "baja duplicado"
      Height          =   615
      Left            =   8760
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   8640
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impredep.rpt"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Impresion"
      Height          =   300
      Left            =   2520
      TabIndex        =   23
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton LimpiaLinea 
      Caption         =   "Limpia Linea"
      Height          =   300
      Left            =   3600
      TabIndex        =   22
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ListBox WVector 
      Height          =   255
      Left            =   6960
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Deposito 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   855
   End
   Begin MSMask.MaskEdBox Acredita 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Importe 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      MaxLength       =   15
      TabIndex        =   4
      Text            =   " "
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Banco 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   2
      Text            =   " "
      Top             =   360
      Width           =   735
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   5880
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6300
      ItemData        =   "deposito.frx":0000
      Left            =   6000
      List            =   "deposito.frx":0007
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   2520
      TabIndex        =   11
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   300
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   1440
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin RichTextLib.RichTextBox Busqueda 
      Height          =   255
      Left            =   4920
      TabIndex        =   28
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   327680
      TextRTF         =   $"deposito.frx":0015
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   5295
      Left            =   0
      TabIndex        =   39
      Top             =   2040
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9340
      _Version        =   327680
      BackColor       =   12582912
      ForeColor       =   -2147483643
   End
   Begin VB.Label Label3 
      Caption         =   "Fec.Acreditacion"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Importe"
      Height          =   375
      Left            =   3000
      TabIndex        =   19
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Creditos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3720
      TabIndex        =   18
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Doc. : 1) Ef.    2) U$S   3) Ch. Terc."
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   7440
      Width           =   3255
   End
   Begin VB.Label DesBanco 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   2520
      TabIndex        =   16
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Banco"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nro. Deposito"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "PrgDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Auxi As String
Private dada As String
Private Vector(100, 6) As String

Private Numero As String
Private Imprelin As Single
Dim rstDepositos As Recordset
Dim spDepositos As String
Dim rstCuenta As Recordset
Dim spCuenta As String
Dim rstBanco As Recordset
Dim spBanco As String
Dim rstRecibos As Recordset
Dim spRecibos As String
Dim cParam As String
Dim XParam As String
Private WSalto As String
Dim Mira(100) As String
Dim Variable As String
Dim coderr As Integer
Dim ZZLugar As Integer

Dim ZEntraI(5000, 3) As String
Dim ZEntraII(5000, 5) As String


Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Private Sub Suma_Datos()

    Creditos.Caption = ""
    
    For iRow = 1 To 99
        Auxi = WVector1.TextMatrix(iRow, 5)
        Call Conver(Auxi, dada)
        If Val(Auxi) <> 0 Then
            Creditos.Caption = Str$(Val(Creditos.Caption) + Val(Auxi))
        End If
    Next iRow
    Creditos.Caption = Pusing("###,###,###.##", Creditos.Caption)
    
    Rem WVector1.Col = 1
    Rem WVector1.Row = 1
    
End Sub

Private Sub Lee_Datos()

    Renglon = 0
    Debito = 0
    Credito = 0
    
    Do
    
        Renglon = Renglon + 1
        Auxi1 = Str$(Renglon)
        Call Ceros(Auxi1, 2)
        ClaveDeposito = Deposito.Text + Auxi1
        
        spDepositos = "ConsultaDepositosClave " + " '" + ClaveDeposito + "'"
        Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
        If rstDepositos.RecordCount > 0 Then
        
            Credito = Credito + 1
            
            WVector1.Row = Credito
            
            WVector1.Col = 1
            WVector1.Text = rstDepositos!Tipo2
            WVector1.Col = 2
            WVector1.Text = rstDepositos!Numero2
            WVector1.Col = 3
            WVector1.Text = rstDepositos!Fecha2
            WVector1.Col = 4
            If rstDepositos!Observaciones2 <> "" Then
                WVector1.Text = rstDepositos!Observaciones2
            End If
            WVector1.Col = 5
            WVector1.Text = Str$(rstDepositos!Importe2)
            WVector1.Text = Pusing("###,###.##", WVector1.Text)
            WVector1.Col = 6
            WVector1.Text = ""
            
            rstDepositos.Close
            
                Else
            Exit Do
        End If
    Loop
    
End Sub

Sub Verifica_datos()
    If Importe.Text = 0 Then
        Importe.Text = "0"
    End If
End Sub

Sub Format_datos()
    Importe.Text = Pusing("###,###,###.##", Importe.Text)
End Sub

Sub Imprime_Datos()
    spBanco = "ConsultaBanco " + " '" + Banco.Text + "'"
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        Banco.Text = rstBanco!Banco
        DesBanco.Caption = rstBanco!Nombre
        rstBanco.Close
        Call Format_datos
    End If
End Sub

Private Sub cmdAdd_Click()

    Sql1 = "Select *"
    Sql2 = " FROM Depositos"
    Sql3 = " Where Depositos.Deposito = " + "'" + Deposito.Text + "'"
    spDepositos = Sql1 + Sql2 + Sql3
    Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
    If rstDepositos.RecordCount > 0 Then
        Existe = "S"
        rstDepositos.Close
    End If
    
    For iRow = 1 To 99
    
        ZVerifica = WVector1.TextMatrix(iRow, 2)
        If Val(ZVerifica) <> 0 Then
            For IRowII = 1 To 99
                ZVerificaII = WVector1.TextMatrix(IRowII, 2)
                If Val(ZVerificaII) <> 0 And iRow <> IRowII Then
                    If Val(ZVerifica) = Val(ZVerificaII) Then
                        Exit Sub
                    End If
                End If
            Next IRowII
        End If
        
    Next iRow

    If Deposito.Text <> "" And Fecha.Text <> "" And Banco.Text <> "" Then
    
    If Existe <> "S" Then
    
        Call Suma_Datos
        
        Debito = 0
        Credito = 0
        
        If Val(Importe.Text) <> 0 Then
            Debito = Val(Importe.Text)
        End If
        
        If Val(Creditos.Caption) <> 0 Then
            Credito = Val(Creditos.Caption)
        End If
        
        Erase Mira
        Counter = 0
        
        If Debito = Credito Then
    
            Renglon = 0
            
            For iRow = 1 To 99
            
                WRow = iRow
                
                WVector1.Col = 5
                WVector1.Row = iRow
                Auxi = WVector1.Text
                Call Conver(Auxi, dada)
                If Val(Auxi) <> 0 Then
                
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    Auxi2 = Str$(Val(Deposito.Text))
                    Call Ceros(Auxi2, 6)
                    
                    XClave = Auxi2 + Auxi1
                    XDeposito = Auxi2
                    XRenglon = Auxi1
                    XBanco = Banco.Text
                    XFecha = Fecha.Text
                    XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XImporte = Importe.Text
                    XAcredita = Acredita.Text
                    XAcreditaOrd = Right$(Acredita.Text, 4) + Mid$(Acredita.Text, 4, 2) + Left$(Acredita.Text, 2)
                    
                    WVector1.Col = 1
                    XTipo2 = WVector1.Text
                    WVector1.Col = 2
                    XNumero2 = WVector1.Text
                    WVector1.Col = 3
                    XFecha2 = WVector1.Text
                    WVector1.Col = 5
                    XImporte2 = WVector1.Text
                    WVector1.Col = 4
                    XObservaciones2 = WVector1.Text
                    WVector1.Col = 6
                    XBaja = WVector1.Text
                    
                    XEmpresa = "1"
                    XImpolis = ""
                    
                    XParam = "'" + XClave + "','" _
                            + XDeposito + "','" _
                            + XRenglon + "','" _
                            + XBanco + "','" _
                            + XFecha + "','" _
                            + XFechaOrd + "','" _
                            + XImporte + "','" _
                            + XAcredita + "','" _
                            + XAcreditaOrd + "','" _
                            + XTipo2 + "','" _
                            + XNumero2 + "','" _
                            + XFecha2 + "','" _
                            + XImporte2 + "','" _
                            + XObservaciones2 + "','" _
                            + XWEmpresa + "','" _
                            + XImpolis + "'"
                    spDepositos = "AltaDepositos "
                    Set rstDepositos = db.OpenRecordset(spDepositos + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    
                    TipoRecibo = Left$(XBaja, 1)
                    ClaveRecibo = Mid$(XBaja, 2, 10)
                    
                    Rem dada
                    Rem dada
                    Rem dada
                    Rem dada
                    Rem dada
                    
                    If Trim(XBaja) <> "" Then
                    
                        If TipoRecibo = "1" Then
                    
                            spRecibos = "ConsultaRecibosClave " + " '" + ClaveRecibo + "'"
                            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                            If rstRecibos.RecordCount > 0 Then
                                XObservaciones = "Deposito Nro : " + Str$(Deposito.Text) + " Banco : " + Left$(DesBanco.Caption, 20)
                                XParam = "'" + ClaveRecibo + "','" _
                                    + "X" + "','" _
                                    + XObservaciones + "'"
                                rstRecibos.Close
                                spRecibos = "ActualizaRecibos " + XParam
                                Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                                Counter = Counter + 1
                                Mira(Counter) = Left$(ClaveRecibo, 6)
                            End If
                            
                                Else
                                
                            Sql1 = "Select *"
                            Sql2 = " FROM RecibosProvi"
                            Sql3 = " Where RecibosProvi.Clave = " + "'" + ClaveRecibo + "'"
                            spRecibosProvi = Sql1 + Sql2 + Sql3
                            Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
                            If rstRecibosProvi.RecordCount > 0 Then
                                ZZReciboDefinitivo = IIf(IsNull(rstRecibosProvi!ReciboDefinitivo), "0", rstRecibosProvi!ReciboDefinitivo)
                                XObservaciones = "Deposito Nro : " + Str$(Deposito.Text) + " Banco : " + Left$(DesBanco.Caption, 20)
                                rstRecibosProvi.Close
                                ZSql = ""
                                ZSql = ZSql + "UPDATE RecibosProvi SET "
                                ZSql = ZSql + "Estado2 = " + "'" + "X" + "',"
                                ZSql = ZSql + "Destino = " + "'" + XObservaciones + "'"
                                ZSql = ZSql + " Where Clave = " + "'" + ClaveRecibo + "'"
                                spRecibosProvi = ZSql
                                Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
                                
                                If ZZReciboDefinitivo <> 0 Then
                                    Auxi3 = ZZReciboDefinitivo
                                    Call Ceros(Auxi3, 6)
                                    ZSql = ""
                                    ZSql = ZSql + "UPDATE Recibos SET "
                                    ZSql = ZSql + "Estado2 = " + "'" + "X" + "',"
                                    ZSql = ZSql + "Destino = " + "'" + XObservaciones + "'"
                                    ZSql = ZSql + " Where Recibo = " + "'" + Auxi3 + "'"
                                    ZSql = ZSql + " and Numero2 = " + "'" + XNumero2 + "'"
                                    spRecibos = ZSql
                                    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                                
                                Rem Counter = Counter + 1
                                Rem Mira(Counter) = "2" + Left$(ClaveRecibo, 6)
                            End If
                            
                        End If
                        
                    End If
                    
                End If
                
            Next iRow
            
            
            For Cicla = 1 To Counter
            
                Graba = "S"
                WRecibo = Mira(Cicla)
                
                spRecibos = "ConsultaRecibos " + "'" + WRecibo + "'"
                Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                If rstRecibos.RecordCount > 0 Then
                    With rstRecibos
                        .MoveFirst
                        Do
                            If .EOF = True Then
                                Exit Do
                            End If
                            If rstRecibos!Tiporeg = 2 Then
                                If rstRecibos!Estado2 <> "X" Then
                                    Graba = "N"
                                End If
                            End If
                            .MoveNext
                            If .EOF = True Then
                                Exit Do
                            End If
                        Loop
                    End With
                    rstRecibos.Close
                End If
            
                If Graba = "S" Then
                    XParam = "'" + WRecibo + "','" _
                                + XFecha + "','" _
                                + XFechaOrd + "','" _
                                + "X" + " '"
                    spRecibos = "ActualizaRecibosMarca " + XParam
                    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
            Next Cicla
        
            With rstEmpresa
                .Index = "Empresa"
                Claveven$ = "1"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    WCtaEfectivo = !CtaEfectivo
                    WCtaCheques = !CtaCheque
                End If
            End With
        
            WSalto = "N"
            Call ImpreDeposito

            Call CmdLimpiar_Click
            Deposito.SetFocus
        
        End If
        
    End If
    
    End If
End Sub

Private Sub cmdDelete_Click()
    If Deposito.Text <> "" Then
            Rem Borro los datos anteriores
            Rem For iRow = 0 To 20
            Rem     Auxi1 = Str$(iRow)
            Rem     Call Ceros(Auxi1, 2)
            Rem     .Seek "=", Deposito.text + Auxi1
            Rem     If .NoMatch = False Then
            Rem         .Delete
            Rem     End If
            Rem Next iRow
            
    End If
    Deposito.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Pantalla.Visible = False
    Call Limpia_Vector

    Deposito.Text = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Importe.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Acredita.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Creditos.Caption = ""
    Deposito.SetFocus
    
    spDepositos = "ConsultaUltimoDeposito "
    Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
    If rstDepositos.RecordCount > 0 Then
        Deposito.Text = rstDepositos!Deposito + 1
        rstDepositos.Close
            Else
        Deposito.Text = ""
    End If

End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    With rstEmpresa
        .Close
    End With
    Deposito.SetFocus
    PrgDeposito.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Command1_Click()
        WSalto = "S"
        Call ImpreDeposito
End Sub

Private Sub Command2_Click()

    Auxi1 = Deposito.Text
    Call Ceros(Auxi1, 6)
    Deposito.Text = Auxi1
    
    Sql1 = "DELETE Depositos"
    Sql2 = " Where Deposito = " + "'" + Deposito.Text + "'"
    spDepositos = Sql1 + Sql2
    Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)

End Sub




Private Sub Command3_Click()

            Renglon = 0
            For iRow = 0 To 19
                WRow = iRow
                WVector1.Col = 5
                WVector1.Row = iRow
                Auxi = WVector1.Text
                Call Conver(Auxi, dada)
                If Val(Auxi) <> 0 Then
                
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    Auxi2 = Str$(Val(Deposito.Text))
                    Call Ceros(Auxi2, 6)
                    
                    XClave = Auxi2 + Auxi1
                    XDeposito = Auxi2
                    XRenglon = Auxi1
                    XBanco = Banco.Text
                    XFecha = Fecha.Text
                    XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XImporte = Importe.Text
                    XAcredita = Acredita.Text
                    XAcreditaOrd = Right$(Acredita.Text, 4) + Mid$(Acredita.Text, 4, 2) + Left$(Acredita.Text, 2)
                    WVector1.Col = 1
                    XTipo2 = WVector1.Text
                    WVector1.Col = 2
                    XNumero2 = WVector1.Text
                    WVector1.Col = 3
                    XFecha2 = WVector1.Text
                    WVector1.Col = 5
                    XImporte2 = WVector1.Text
                    WVector1.Col = 4
                    XObservaciones2 = WVector1.Text
                    XEmpresa = "1"
                    XImpolis = ""
                    
                    XParam = "'" + XClave + "','" _
                            + XDeposito + "','" _
                            + XRenglon + "','" _
                            + XBanco + "','" _
                            + XFecha + "','" _
                            + XFechaOrd + "','" _
                            + XImporte + "','" _
                            + XAcredita + "','" _
                            + XAcreditaOrd + "','" _
                            + XTipo2 + "','" _
                            + XNumero2 + "','" _
                            + XFecha2 + "','" _
                            + XImporte2 + "','" _
                            + XObservaciones2 + "','" _
                            + XWEmpresa + "','" _
                            + XImpolis + "'"
                    spDepositos = "AltaDepositos "
                    Set rstDepositos = db.OpenRecordset(spDepositos + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
            Next iRow
            


End Sub

Private Sub Command4_Click()

    Erase ZEntraI
    ZLugar = 0

    Sql1 = "Select *"
    Sql2 = " FROM Depositos"
    Sql3 = " Where Depositos.FechaOrd >= " + "'" + "20080201" + "'"
    spDepositos = Sql1 + Sql2 + Sql3
    Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
    If rstDepositos.RecordCount > 0 Then
        With rstDepositos
            .MoveFirst
            Do
                If .EOF = False Then
                    WTipo2 = rstDepositos!Tipo2
                    If Val(WTipo2) = 3 Then
                        ZLugar = ZLugar + 1
                        ZEntraI(ZLugar, 1) = rstDepositos!Numero2
                        ZEntraI(ZLugar, 2) = rstDepositos!Deposito
                        ZEntraI(ZLugar, 3) = rstDepositos!Fecha2
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstDepositos.Close
    End If

    For iRow = 1 To ZLugar
    
        XNumero2 = ZEntraI(iRow, 1)
        XDeposito = ZEntraI(iRow, 2)
        XFecha2 = ZEntraI(iRow, 3)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM recibos"
        ZSql = ZSql + " Where recibos.Numero2 = " + "'" + XNumero2 + "'"
        ZSql = ZSql + " and recibos.fecha2 = " + "'" + XFecha2 + "'"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
            XObservaciones = "Deposito Nro : " + XDeposito
            rstRecibos.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE recibos SET "
            ZSql = ZSql + "Estado2 = " + "'" + "X" + "',"
            ZSql = ZSql + "Destino = " + "'" + XObservaciones + "'"
            ZSql = ZSql + " Where Numero2 = " + "'" + XNumero2 + "'"
            ZSql = ZSql + " and fecha2 = " + "'" + XFecha2 + "'"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        End If
                            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM RecibosProvi"
        ZSql = ZSql + " Where RecibosProvi.Numero2 = " + "'" + XNumero2 + "'"
        ZSql = ZSql + " and recibosprovi.fecha2 = " + "'" + XFecha2 + "'"
        spRecibosProvi = ZSql
        Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibosProvi.RecordCount > 0 Then
            XObservaciones = "Deposito Nro : " + XDeposito
            rstRecibosProvi.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE RecibosProvi SET "
            ZSql = ZSql + "Estado2 = " + "'" + "X" + "',"
            ZSql = ZSql + "Destino = " + "'" + XObservaciones + "'"
            ZSql = ZSql + " Where Numero2 = " + "'" + XNumero2 + "'"
            ZSql = ZSql + " and fecha2 = " + "'" + XFecha2 + "'"
            spRecibosProvi = ZSql
            Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
        End If
                            
    Next iRow
            
    Call CmdLimpiar_Click
    Deposito.SetFocus

End Sub

Private Sub Deposito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi1 = Deposito.Text
        Call Ceros(Auxi1, 6)
        Deposito.Text = Auxi1
        
        Existe = "N"
        ClaveDeposito = Deposito.Text + "01"
        
        spDepositos = "ConsultaDepositosClave " + " '" + ClaveDeposito + "'"
        Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
        If rstDepositos.RecordCount > 0 Then
            Existe = "S"
            If rstDepositos!Banco <> "" Then
                Banco.Text = rstDepositos!Banco
            End If
            If rstDepositos!Importe <> "" Then
                Importe.Text = rstDepositos!Importe
            End If
            Fecha.Text = rstDepositos!Fecha
            Acredita.Text = rstDepositos!Acredita
            rstDepositos.Close
        End If
        If Existe = "S" Then
            Call Imprime_Datos
            Call Lee_Datos
            Call Suma_Datos
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                Else
            Fecha.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Banco.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Banco.Text) <> 0 Then
            spBanco = "ConsultaBanco " + " '" + Banco.Text + "'"
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                Banco.Text = rstBanco!Banco
                DesBanco.Caption = rstBanco!Nombre
                rstBanco.Close
                Rem Call Imprime_Datos
                Acredita.SetFocus
                    Else
                Banco.SetFocus
            End If
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Acredita_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Acredita.Text, Auxi)
        If Auxi = "S" Then
            Importe.SetFocus
                Else
            Acredita.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_ImpreDep
End Sub

Private Sub Importe_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Importe.Text = Pusing("###,###,###.##", Importe.Text)
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub Consulta_Click()

     XRow = WVector1.Row
     XCol = WVector1.Col

     Opcion.Clear

     Opcion.AddItem "Bancos"
     Opcion.AddItem "Cheques terceros"

     Opcion.Visible = True
     
End Sub

Private Sub LeeLectora_Click()
    Lectora.Visible = True
    Lectora.Text = ""
    Lectora.SetFocus
End Sub

Private Sub LimpiaLinea_Click()
    
    NoToma = WVector1.Row
    Erase Vector
    Lugar = 0
    
    For iRow = 1 To 99
        If NoToma <> iRow Then
            Lugar = Lugar + 1
            WVector1.Row = iRow
            WVector1.Col = 1
            Vector(Lugar, 1) = WVector1.Text
            WVector1.Col = 2
            Vector(Lugar, 2) = WVector1.Text
            WVector1.Col = 3
            Vector(Lugar, 3) = WVector1.Text
            WVector1.Col = 4
            Vector(Lugar, 4) = WVector1.Text
            WVector1.Col = 5
            Vector(Lugar, 5) = WVector1.Text
            WVector1.Col = 6
            Vector(Lugar, 6) = WVector1.Text
        End If
        
        WVector1.Col = 1
        WVector1.Text = ""
        WVector1.Col = 2
        WVector1.Text = ""
        WVector1.Col = 3
        WVector1.Text = ""
        WVector1.Col = 4
        WVector1.Text = ""
        WVector1.Col = 5
        WVector1.Text = ""
        WVector1.Col = 6
        WVector1.Text = ""
        
    Next iRow
    
    For da = 1 To Lugar
        
        WVector1.Row = da
        WVector1.Col = 1
        WVector1.Text = Vector(da, 1)
        WVector1.Col = 2
        WVector1.Text = Vector(da, 2)
        WVector1.Col = 3
        WVector1.Text = Vector(da, 3)
        WVector1.Col = 4
        WVector1.Text = Vector(da, 4)
        WVector1.Col = 5
        WVector1.Text = Vector(da, 5)
        WVector1.Col = 6
        WVector1.Text = Vector(da, 6)
    
    Next da
    
    Call Suma_Datos

End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False

    Dim IngresaItem As String
    
    
    Pantalla.Clear
    WIndice.Clear
    WVector.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            spBanco = "ListaBancos"
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstBanco
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstBanco!Banco) + " " + rstBanco!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstBanco!Banco
                        WIndice.AddItem IngresaItem
                        WVector.AddItem ""
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstBanco.Close
                
        Case 1
        
            ZSql = ""
            ZSql = ZSql + "UPDATE RecibosProvi SET "
            ZSql = ZSql + " ReciboDefinitivo = 0"
            ZSql = ZSql + " Where ReciboDefinitivo is null"
            spRecibosProvi = ZSql
            Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
        
        
            Erase ZEntraI
            Erase ZEntraII
            
            ZLugarI = 0
            ZLugarII = 0
        
            ZSql = ""
            ZSql = ZSql + "Select Recibos.Tiporeg, Recibos.Estado2, Recibos.Tipo2, Recibos.TipoReg, Recibos.Importe2, Recibos.Numero2, Recibos.Fecha2, Recibos.Banco2, Recibos.Clave, Recibos.FechaOrd2"
            ZSql = ZSql + " FROM Recibos"
            ZSql = ZSql + " Where Recibos.TipoReg = '2'"
            ZSql = ZSql + " and Recibos.Estado2 <> 'X'"
            ZSql = ZSql + " and Recibos.Tipo2 = '02'"
            ZSql = ZSql + " Order by Recibos.FechaOrd2, Recibos.Numero2"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
            Rem spRecibos = "ListaRecibosNroCheque"
            Rem Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstRecibos.RecordCount Then
            
                With rstRecibos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Val(rstRecibos!Tiporeg) = 2 Then
                                If Val(rstRecibos!Tipo2) = 2 And rstRecibos!Estado2 <> "X" Then
                                
                                    ZLugarI = ZLugarI + 1
                                    Auxi$ = Str$(rstRecibos!Importe2)
                                    Auxi$ = Mascara("#,###,###.##", Auxi$)
                                    Numero = Str$(Val(rstRecibos!Numero2))
                                    Call Ceros(Numero, 6)
                                    IngresaItem = Numero + "  " + rstRecibos!Fecha2 + "  " + Auxi$ + "  " + rstRecibos!Banco2
                                    
                                    WOrdFecha2 = IIf(IsNull(rstRecibos!FechaOrd2), "", rstRecibos!FechaOrd2)
                                
                                    ZEntraI(ZLugarI, 1) = IngresaItem
                                    ZEntraI(ZLugarI, 2) = "1" + rstRecibos!Clave
                                    ZEntraI(ZLugarI, 3) = WOrdFecha2
                                    
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstRecibos.Close
                
            End If
            
            Sql1 = "Select RecibosProvi.Tiporeg, RecibosProvi.Estado2, RecibosProvi.Tipo2, RecibosProvi.TipoReg, RecibosProvi.Importe2, RecibosProvi.Numero2, RecibosProvi.Fecha2, RecibosProvi.Banco2, RecibosProvi.Clave, RecibosProvi.FechaOrd2, RecibosProvi.ReciboDefinitivo"
            Sql2 = " FROM RecibosProvi"
            Sql3 = " Where RecibosProvi.TipoReg = " + "'" + "2" + "'"
            Sql4 = " and RecibosProvi.Estado2 = " + "'" + "P" + "'"
            Sql5 = " and RecibosProvi.ReciboDefinitivo = " + "'" + "0" + "'"
            Sql6 = " Order by FechaOrd2, Numero2"
            spRecibosProvi = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6
            Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibosProvi.RecordCount > 0 Then
            
                With rstRecibosProvi
                    .MoveFirst
                    Do
                        If .EOF = False Then
                    
                            WTiporeg = IIf(IsNull(rstRecibosProvi!Tiporeg), "", rstRecibosProvi!Tiporeg)
                            WTipo2 = IIf(IsNull(rstRecibosProvi!Tipo2), "", rstRecibosProvi!Tipo2)
                            WEstado2 = IIf(IsNull(rstRecibosProvi!Estado2), "", rstRecibosProvi!Estado2)
                            WDefinitivo = IIf(IsNull(rstRecibosProvi!ReciboDefinitivo), "0", rstRecibosProvi!ReciboDefinitivo)
                        
                            If Val(WTiporeg) = 2 Then
                                If Val(WTipo2) = 2 And WEstado2 <> "X" And Val(WDefinitivo) = 0 Then
                            
                                    ZLugarII = ZLugarII + 1
                                    Auxi$ = Str$(rstRecibosProvi!Importe2)
                                    Auxi$ = Mascara("#,###,###.##", Auxi$)
                                    Numero = Str$(Val(rstRecibosProvi!Numero2))
                                    WFecha2 = IIf(IsNull(rstRecibosProvi!Fecha2), "", rstRecibosProvi!Fecha2)
                                    Call Ceros(Numero, 6)
                                    IngresaItem = Numero + "  " + rstRecibosProvi!Fecha2 + "  " + Auxi$ + "  " + rstRecibosProvi!Banco2
                                    
                                    WOrdFecha2 = IIf(IsNull(rstRecibosProvi!FechaOrd2), "", rstRecibosProvi!FechaOrd2)
                                
                                    ZEntraII(ZLugarII, 1) = IngresaItem
                                    ZEntraII(ZLugarII, 2) = "2" + rstRecibosProvi!Clave
                                    ZEntraII(ZLugarII, 3) = WOrdFecha2
                                    ZEntraII(ZLugarII, 4) = rstRecibosProvi!Numero2
                                    ZEntraII(ZLugarII, 5) = WFecha2
                                
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstRecibosProvi.Close
            
            End If
            
            ZZTotal = ZLugarI + ZLugarII
            ZLugarI = 0
            ZLugarII = 0
            
            For ZCicla = 1 To ZZTotal
            
                If ZEntraI(ZLugarI + 1, 1) <> "" And ZEntraII(ZLugarII + 1, 1) <> "" Then
                
                    If ZEntraI(ZLugarI + 1, 3) < ZEntraII(ZLugarII + 1, 3) Then
                
                        ZLugarI = ZLugarI + 1
                        IngresaItem = ZEntraI(ZLugarI, 1)
                        Pantalla.AddItem IngresaItem
                        IngresaItem = ZEntraI(ZLugarI, 2)
                        WIndice.AddItem IngresaItem
                        WVector.AddItem ""
                        
                            Else
                
                        ZLugarII = ZLugarII + 1
                        
                        ZZNumero2 = ZEntraII(ZLugarII, 4)
                        ZZFecha2 = ZEntraII(ZLugarII, 5)
                        
                        ZSql = ""
                        ZSql = ZSql + "Select Recibos.Numero2, Recibos.Fecha2"
                        ZSql = ZSql + " FROM Recibos"
                        ZSql = ZSql + " Where Recibos.Numero2 = " + "'" + ZZNumero2 + "'"
                        ZSql = ZSql + " and Recibos.Fecha2 = " + "'" + ZZFecha2 + "'"
                        spRecibos = ZSql
                        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                        If rstRecibos.RecordCount > 0 Then
                            rstRecibos.Close
                                Else
                            IngresaItem = ZEntraII(ZLugarII, 1)
                            Pantalla.AddItem IngresaItem
                            IngresaItem = ZEntraII(ZLugarII, 2)
                            WIndice.AddItem IngresaItem
                            WVector.AddItem ""
                        End If
                            
                    End If
                    
                        Else
                
                    If ZEntraI(ZLugarI + 1, 1) <> "" Then
                        ZLugarI = ZLugarI + 1
                        IngresaItem = ZEntraI(ZLugarI, 1)
                        Pantalla.AddItem IngresaItem
                        IngresaItem = ZEntraI(ZLugarI, 2)
                        WIndice.AddItem IngresaItem
                        WVector.AddItem ""
                    End If
                
                    If ZEntraII(ZLugarII + 1, 1) <> "" Then
                        ZLugarII = ZLugarII + 1
                        
                        ZZNumero2 = ZEntraII(ZLugarII, 4)
                        ZZFecha2 = ZEntraII(ZLugarII, 5)
                  Rem by nan
                        ZSql = ""
                        ZSql = ZSql + "Select Recibos.Numero2, Recibos.Fecha2"
                        ZSql = ZSql + " FROM Recibos"
                        ZSql = ZSql + " Where Recibos.Numero2 = " + "'" + ZZNumero2 + "'"
                        ZSql = ZSql + " and Recibos.Fecha2 = " + "'" + ZZFecha2 + "'"
                        spRecibos = ZSql
                        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                        If rstRecibos.RecordCount > 0 Then
                            rstRecibos.Close
                                Else
                            IngresaItem = ZEntraII(ZLugarII, 1)
                            Pantalla.AddItem IngresaItem
                            IngresaItem = ZEntraII(ZLugarII, 2)
                            WIndice.AddItem IngresaItem
                            WVector.AddItem ""
                        End If
                        
                    End If
                    
                End If
                
            Next ZCicla
        
     
        Case Else
    End Select
    
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_DblClick()

    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WBanco = WIndice.List(Indice)
            spBanco = "ConsultaBanco " + "'" + Str$(WBanco) + "'"
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                Banco.Text = WBanco
                DesBanco.Caption = rstBanco!Nombre
                rstBanco.Close
                        Else
                Banco.Text = ""
            End If
            Banco.SetFocus
        
        Case 1
            Indice = Pantalla.ListIndex
            Auxi = WVector.List(Indice)
            
            If Auxi <> "X" Then
            
                For iRow = 1 To 99
                    WVector1.Col = 5
                    WVector1.Row = iRow
                    Auxi = WVector1.Text
                    Call Conver(Auxi, dada)
                    If Val(Auxi) = 0 Then
                        Exit For
                    End If
                Next iRow
                
                If Mid$(WIndice.List(Indice), 1, 1) = "1" Then
    
                    Indice = Pantalla.ListIndex
                    ClaveRecibo = Mid$(WIndice.List(Indice), 2, 10)
                    spRecibos = "ConsultaRecibosClave " + "'" + ClaveRecibo + "'"
                    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                    If rstRecibos.RecordCount > 0 Then
                    
                        ZZEntra = "S"
                        For ZZCiclo = 1 To 100
                            If WVector1.TextMatrix(ZZCiclo, 2) = rstRecibos!Numero2 Then
                                ZZEntra = "N"
                                Exit For
                            End If
                        Next ZZCiclo
                        
                        If ZZEntra = "S" Then
                    
                            WVector1.Col = 1
                            WVector1.Text = "3"
                            
                            WVector1.Col = 2
                            WVector1.Text = rstRecibos!Numero2
                        
                            WVector1.Col = 3
                            WVector1.Text = rstRecibos!Fecha2
                        
                            WVector1.Col = 4
                            WVector1.Text = rstRecibos!Banco2
                        
                            WVector1.Col = 5
                            WVector1.Text = Str$(rstRecibos!Importe2)
                           Rem by nan 03-02-2014
                            WVector1.Text = Pusing("###,###,###.##", WVector1.Text)
                            
                            WVector1.Col = 6
                            WVector1.Text = WIndice.List(Indice)
                            
                            Call Suma_Datos
                            
                            WVector1.Row = XRow
                            WVector1.Col = 1
                            
                            Auxi = "X"
                            WVector.List(Indice) = Auxi
                            Pantalla.List(Indice) = ""
                            
                        End If
                        
                        rstRecibos.Close
                        
                    End If
                        
                            Else
                            
                    Indice = Pantalla.ListIndex
                    ClaveRecibo = Mid$(WIndice.List(Indice), 2, 10)
                    
                    Sql1 = "Select *"
                    Sql2 = " FROM RecibosProvi"
                    Sql3 = " Where RecibosProvi.Clave = " + "'" + ClaveRecibo + "'"
                    spRecibosProvi = Sql1 + Sql2 + Sql3
                    Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
                    If rstRecibosProvi.RecordCount > 0 Then
                    
                        ZZEntra = "S"
                        For ZZCiclo = 1 To 100
                            If WVector1.TextMatrix(ZZCiclo, 2) = rstRecibosProvi!Numero2 Then
                                ZZEntra = "N"
                                Exit For
                            End If
                        Next ZZCiclo
                        
                        If ZZEntra = "S" Then
                        
                            WVector1.Col = 1
                            WVector1.Text = "3"
                            
                            WVector1.Col = 2
                            WVector1.Text = rstRecibosProvi!Numero2
                        
                            WVector1.Col = 3
                            WVector1.Text = rstRecibosProvi!Fecha2
                        
                            WVector1.Col = 4
                            WVector1.Text = rstRecibosProvi!Banco2
                        
                            WVector1.Col = 5
                            WVector1.Text = Str$(rstRecibosProvi!Importe2)
                            
                            Rem by nan 03-2-2014
                            WVector1.Text = Pusing("###,###,###.##", WVector1.Text)
                            
                            WVector1.Col = 6
                            WVector1.Text = WIndice.List(Indice)
                            
                            Call Suma_Datos
                            
                            WVector1.Row = XRow
                            WVector1.Col = 1
                            
                            Auxi = "X"
                            WVector.List(Indice) = Auxi
                            Pantalla.List(Indice) = ""
                            
                        End If
                        
                        rstRecibosProvi.Close
                        
                    End If
                            
                End If
                        
                If WVector1.Row < 99 Then
                    WVector1.Row = WVector1.Row + 1
                    WVector1.Col = 1
                            Else
                    WVector1.Col = 1
                End If
            
            End If
                
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    Pantalla.Visible = False
    
    Deposito.Text = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Importe.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Acredita.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Creditos.Caption = ""
    
    spDepositos = "ConsultaUltimoDeposito "
    Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
    If rstDepositos.RecordCount > 0 Then
        Deposito.Text = rstDepositos!Deposito + 1
        rstDepositos.Close
            Else
        Deposito.Text = ""
    End If
     
End Sub

Private Sub ImpreDeposito()

    On Error GoTo WError
        
    da = 0
    With rstImpreDep
        .Index = "Deposito"
        .Seek ">=", 0
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

    With rstEmpresa
        .Index = "Empresa"
        Claveven$ = WEmpresa
        .Seek "=", Claveven$
        If .NoMatch = False Then
            Impretit = !Nombre
                Else
            Impretit = ""
        End If
    End With
        
    WRenglon = 0
        
    For iRow = 1 To 99
        WVector1.Col = 5
        WVector1.Row = iRow
        Auxi = WVector1.Text
        Call Conver(Auxi, dada)
        If Val(Auxi) <> 0 Then
            WRenglon = WRenglon + 1
            With rstImpreDep
                .AddNew
                !Deposito = Val(Deposito.Text)
                !Renglon = WRenglon
                !Fecha = Fecha.Text
                !Banco = Val(Banco.Text)
                !Nombre = DesBanco.Caption
                !Total = Val(Importe.Text)
                !Titulo = Impretit
                    
                WVector1.Col = 1
                Select Case Val(WVector1.Text)
                    Case 1
                        !Tipo = "Efectivo"
                        !Numero = ""
                        !Banco = ""
                        WVector1.Col = 5
                        !Importe = Val(WVector1.Text)
                    Case 2
                        !Tipo = "Dolares"
                        !Numero = ""
                        !Banco = ""
                        WVector1.Col = 5
                        !Importe = Val(WVector1.Text)
                    Case Else
                        !Tipo = "Cheque"
                        
                        WVector1.Col = 2
                        !Numero = WVector1.Text
                        
                        WVector1.Col = 4
                        Busqueda.Text = WVector1.Text
                        Foundpos = Busqueda.Find("/")
                        If Foundpos > 0 Then
                            !Descripcion = Left$(WVector1.Text, Foundpos)
                                Else
                            !Descripcion = WVector1.Text
                        End If
                        
                        WVector1.Col = 5
                        !Importe = Val(WVector1.Text)
                        
                End Select
                    
                .Update
            End With
            
        End If
    Next iRow
    
    listado.ReportFileName = "Impredep.rpt"
    
    listado.Destination = 1
    listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    listado.Action = 1
    
    listado.ReportFileName = "ImpredepII.rpt"
    
    listado.Destination = 1
    listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    listado.Action = 1
        
    Exit Sub
        
WError:
    Resume Next

End Sub

Private Sub Lectora_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Lectora.Text) = 31 Then
        
            ZZClaveCheque = ""
            Entra = "S"
        
            Sql1 = "Select *"
            Sql2 = " FROM Recibos"
            Sql3 = " Where Recibos.ClaveCheque = " + "'" + Lectora.Text + "'"
            spRecibos = Sql1 + Sql2 + Sql3
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
                Entra = "N"
                ZZClaveCheque = "1" + rstRecibos!Clave
                rstRecibos.Close
            End If
            
            If Entra = "S" Then
                Sql1 = "Select *"
                Sql2 = " FROM RecibosProvi"
                Sql3 = " Where RecibosProvi.ClaveCheque = " + "'" + Lectora.Text + "'"
                spRecibosProvi = Sql1 + Sql2 + Sql3
                Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
                If rstRecibosProvi.RecordCount > 0 Then
                    ZZClaveCheque = "2" + rstRecibosProvi!Clave
                    rstRecibosProvi.Close
                End If
            End If
            
            For iRow = 1 To 99
                WRow = iRow
                WVector1.Col = 6
                WVector1.Row = iRow
                If ZZClaveCheque = WVector1.Text Then
                    WVector1.Row = ZZLugar
                    WVector1.Col = 1
                    WVector1.Text = ""
                    Exit Sub
                End If
                WVector1.Col = 5
                WVector1.Row = iRow
                Auxi = WVector1.Text
                Call Conver(Auxi, dada)
                If Val(Auxi) = 0 Then
                    Exit For
                End If
            Next iRow
            
            Entra = "S"
        
            Sql1 = "Select *"
            Sql2 = " FROM Recibos"
            Sql3 = " Where Recibos.ClaveCheque = " + "'" + Lectora.Text + "'"
            Rem Sql4 = " and Recibos.Estado2 = " + "'" + "P" + "'"
            spRecibos = Sql1 + Sql2 + Sql3
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
            
                Entra = "N"
                
                If rstRecibos!Estado2 = "P" Then
                    
                    ZZEntra = "S"
                    For ZZCiclo = 1 To 100
                        If WVector1.TextMatrix(ZZCiclo, 2) = rstRecibos!Numero2 Then
                            ZZEntra = "N"
                            Exit For
                        End If
                    Next ZZCiclo
                    
                    If ZZEntra = "S" Then
                
                        WVector1.Col = 1
                        WVector1.Text = "3"
                        
                        WVector1.Col = 2
                        WVector1.Text = rstRecibos!Numero2
                    
                        WVector1.Col = 3
                        WVector1.Text = rstRecibos!Fecha2
                    
                        WVector1.Col = 4
                        WVector1.Text = rstRecibos!Banco2
                    
                        WVector1.Col = 5
                        WVector1.Text = Str$(rstRecibos!Importe2)
                        WVector1.Text = Pusing("###,###,###.##", WVector1.Text)
                        
                        WVector1.Col = 6
                        WVector1.Text = "1" + rstRecibos!Clave
                    
                        If WVector1.Row < 99 Then
                            WVector1.Row = WVector1.Row + 1
                            WVector1.Col = 1
                                Else
                            WVector1.Col = 1
                        End If
                        
                        Call Suma_Datos
                        
                        Rem wvector1.Row = XRow
                        Rem wvector1.Col = 0
                        
                    End If
                
                End If
                    
                rstRecibos.Close
                    
            End If
            
            If Entra = "S" Then
            
                Sql1 = "Select *"
                Sql2 = " FROM RecibosProvi"
                Sql3 = " Where RecibosProvi.ClaveCheque = " + "'" + Lectora.Text + "'"
                Sql4 = " and RecibosProvi.Estado2 = " + "'" + "P" + "'"
                spRecibosProvi = Sql1 + Sql2 + Sql3 + Sql4
                Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
                If rstRecibosProvi.RecordCount > 0 Then
                    
                    ZZEntra = "S"
                    For ZZCiclo = 1 To 100
                        If WVector1.TextMatrix(ZZCiclo, 2) = rstRecibosProvi!Numero2 Then
                            ZZEntra = "N"
                            Exit For
                        End If
                    Next ZZCiclo
                    
                    If ZZEntra = "S" Then
                        WVector1.Col = 1
                        WVector1.Text = "3"
                        
                        WVector1.Col = 2
                        WVector1.Text = rstRecibosProvi!Numero2
                    
                        WVector1.Col = 3
                        WVector1.Text = rstRecibosProvi!Fecha2
                    
                        WVector1.Col = 4
                        WVector1.Text = rstRecibosProvi!Banco2
                    
                        WVector1.Col = 5
                        WVector1.Text = Str$(rstRecibosProvi!Importe2)
                        WVector1.Text = Pusing("###,###,###.##", WVector1.Text)
                        
                        WVector1.Col = 6
                        WVector1.Text = "2" + rstRecibosProvi!Clave
                    
                        If WVector1.Row < 99 Then
                            WVector1.Row = WVector1.Row + 1
                            WVector1.Col = 1
                                Else
                            WVector1.Col = 1
                        End If
                        
                        Call Suma_Datos
                        
                        Rem wvector1.Row = XRow
                        Rem wvector1.Col = 0
                        
                    End If
                    
                    rstRecibosProvi.Close
                    
                        Else
                        
                    WVector1.Col = 1
                    WVector1.Text = ""
                    
                    WVector1.Col = 2
                    WVector1.Text = ""
                
                    WVector1.Col = 3
                    WVector1.Text = ""
                
                    WVector1.Col = 4
                    WVector1.Text = ""
                
                    WVector1.Col = 5
                    WVector1.Text = ""
                    
                    WVector1.Col = 6
                    WVector1.Text = ""
                
                    WVector1.Col = 1
                    
                    Call Suma_Datos
                    
                End If
            
            End If
            
        End If
        Lectora.Visible = False
    End If
    If KeyAscii = 27 Then
        Lectora.Visible = False
    End If
End Sub



Rem
Rem Controles de la grilla
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.Visible = True
            WTexto1.SetFocus
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.Visible = True
            WTexto2.SetFocus
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            WTexto3.Visible = True
            WTexto3.SetFocus
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            If Val(WVector1.Text) > 0 Then
                WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
            End If
        End If
        Rem Call Calcula_Click
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub


Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
            
        Case 34
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_Grilla()
    Select Case WVector1.Col
        Case 5
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            ZZTipo = WVector1.Text
            If Len(WVector1.Text) = 31 Then
                ZZLugar = WVector1.Row
                Lectora.Text = WVector1.Text
                Call Lectora_Keypress(13)
                ZZTipo = "99"
                Rem Exit Sub
            End If
            
            Rem Auxi$ = Str$(Val(WVector1.Text))
            Rem Call Ceros(Auxi$, 2)
            Rem WVector1.Text = Auxi$
            
            Select Case Val(WVector1.Text)
                Case 1, 2
                    WVector1.Col = 2
                    WVector1.Text = ""
                    WVector1.Col = 3
                    WVector1.Text = ""
                    WVector1.Col = 4
                    WVector1.Text = ""
                    WVector1.Col = 6
                    WVector1.Text = ""
                    WVector1.Col = 5
                    WVector1.Text = Importe.Text
                    WVector1.Text = Pusing("###,###,###.##", WVector1.Text)
                    Call Suma_Datos
                    WVector1.Col = 4
                    
                Case Else
                    WControl = "N"
                    
            End Select
            
        Case 5
            WVector1.Col = 5
            WVector1.Text = Pusing("###,###,###.##", WVector1.Text)
            Call Suma_Datos
            WVector1.Col = 5
           
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la grilla en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la Grilla
    
    WVector1.FixedCols = 1
    WVector1.Cols = 7
    WVector1.FixedRows = 1
    WVector1.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 400
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1150
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Nombre"
                WVector1.ColWidth(Ciclo) = 1600
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###,###.##"
            Case 6
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTituloVector(Ciclo).Text = WVector1.Text
        WTituloVector(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTituloVector(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTituloVector(Ciclo).Width = WVector1.CellWidth
        WTituloVector(Ciclo).Height = WVector1.CellHeight
        WTituloVector(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector1.Width = 11400
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub

