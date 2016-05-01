VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMovban 
   Caption         =   "Listado de Movimientos de Bancos"
   ClientHeight    =   5190
   ClientLeft      =   2430
   ClientTop       =   1155
   ClientWidth     =   6660
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5190
   ScaleWidth      =   6660
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6120
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   4335
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
         Height          =   540
         Left            =   3120
         TabIndex        =   16
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox HastaBanco 
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
         Left            =   2040
         TabIndex        =   14
         Text            =   " "
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox DesdeBanco 
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
         Left            =   2040
         TabIndex        =   13
         Text            =   " "
         Top             =   1200
         Width           =   855
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   240
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
         Left            =   2160
         TabIndex        =   8
         Top             =   2280
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
         Left            =   600
         TabIndex        =   7
         Top             =   2280
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
         Height          =   495
         Left            =   3120
         TabIndex        =   6
         Top             =   1560
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
         Height          =   495
         Left            =   3120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Banco"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Banco"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
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
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1095
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
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6360
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wmovban.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Movimietos de Bancos"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6240
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgMovban"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WInicial() As Variant ' Matriz de 2 dimensiones que contiene registros
Dim rstPagos As Recordset
Dim spPagos As String
Dim rstDepositos As Recordset
Dim spDepositos As String
Dim rstBanco As Recordset
Dim spBanco As String
Dim rstRecibos As Recordset
Dim spRecibos As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim XParam As String

Private Sub Acepta_Click()

    For XDa = 1 To 100
        WInicial(XDa) = 0
    Next XDa
    
    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia

    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
            .Update
        End If
    End With

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With

    da = 0
    With rstMovban
        .Index = "Clave"
        .Seek "=", da
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
    
    
    XParam = "'" + "19991017" + "','" _
                + DesdeBanco.Text + "','" _
                + HastaBanco.Text + "'"
    spPagos = "ListaPagosMovban " + XParam
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
            
    With rstPagos
            .MoveFirst
            Do

                If !FechaOrd > "19991017" Then
                
                If WDesde <= !FechaOrd And !FechaOrd <= WHasta Then
                    If Val(!Tipo2) = 2 Then
                        If Val(!Banco2) >= Val(DesdeBanco.Text) And Val(!Banco2) <= Val(HastaBanco.Text) Then
                            WBanco = !Banco2
                            WOrden = !Orden
                            WFecha = !Fecha
                            WFechaord = !FechaOrd
                            WAcredita = !Fecha2
                            WAcreditaOrd = !FechaOrd2
                            WObservaciones = ""
                            WObservaciones = !Observaciones
                            Rem If Val(!Proveedor) = 0 Then
                            Rem     WObservaciones = !Observaciones
                            Rem         Else
                            Rem     With rstProveedor
                            Rem         .Index = "Proveedor"
                            Rem         .Seek "=", !Proveedor
                            Rem        If .NoMatch = False Then
                            Rem             WObservaciones = !Nombre
                            Rem         End If
                            Rem     End With
                            Rem End If
                            WNumero = !Numero2
                            WImporte = !Importe2
                            WOrden = !Orden
                            WProveedor = !Proveedor
                
                            With rstMovban
                                .AddNew
                                !da = 0
                                !Banco = WBanco
                                !Fecha = WFecha
                                !FechaOrd = WFechaord
                                !Acredita = WAcredita
                                !AcreditaOrd = WAcreditaOrd
                                !Observaciones = WObservaciones
                                !Numero = WNumero
                                !Debito = 0
                                !Credito = WImporte
                                !Comprobante = WOrden
                                !Empresa = 1
                                !Titulo = WTitulo
                                !Titulo1 = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                !Proveedor = WProveedor
                                .Update
                            End With
                        End If
                    End If
                    
                    If Val(!Tiporeg) = 1 And Val(!Banco2) <> 0 Then
                        If Val(!Banco2) >= Val(DesdeBanco.Text) And Val(!Banco2) <= Val(HastaBanco.Text) Then
                            WBanco = !Banco2
                            WOrden = !Orden
                            WFecha = !Fecha
                            WFechaord = !FechaOrd
                            WAcredita = !Fecha
                            WAcreditaOrd = !FechaOrd
                            WObservaciones = ""
                            WObservaciones = !Observaciones
                            Rem If Val(!Proveedor) = 0 Then
                            Rem     WObservaciones = !Observaciones
                            Rem         Else
                            Rem     With rstProveedor
                            Rem         .Index = "Proveedor"
                            Rem         .Seek "=", !Proveedor
                            Rem        If .NoMatch = False Then
                            Rem             WObservaciones = !Nombre
                            Rem         End If
                            Rem     End With
                            Rem End If
                            WNumero = ""
                            WImporte = !Importe1
                            WOrden = !Orden
                            WProveedor = !Proveedor
                
                            With rstMovban
                                .AddNew
                                !da = 0
                                !Banco = WBanco
                                !Fecha = WFecha
                                !FechaOrd = WFechaord
                                !Acredita = WAcredita
                                !AcreditaOrd = WAcreditaOrd
                                !Observaciones = WObservaciones
                                !Numero = WNumero
                                !Debito = WImporte
                                !Credito = 0
                                !Comprobante = WOrden
                                !Empresa = 1
                                !Titulo = WTitulo
                                !Titulo1 = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                !Proveedor = WProveedor
                                .Update
                            End With
                        End If
                    End If
                    
                End If
                
                If WDesde > !FechaOrd Then
                    If Val(!Tipo2) = 2 Then
                        If Val(!Banco2) >= Val(DesdeBanco.Text) And Val(!Banco2) <= Val(HastaBanco.Text) Then
                            WInicial(!Banco2) = WInicial(!Banco2) - !Importe2
                        End If
                    End If
                    If Val(!Tiporeg) = 1 And !Banco2 <> 0 Then
                        If Val(!Banco2) >= Val(DesdeBanco.Text) And Val(!Banco2) <= Val(HastaBanco.Text) Then
                            WInicial(!Banco2) = WInicial(!Banco2) + !Importe1
                        End If
                    End If
                End If
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    rstPagos.Close
    
    End If

    XParam = "'" + "19991017" + "','" _
                + DesdeBanco.Text + "','" _
                + HastaBanco.Text + "','" _
                + "01" + "'"
    spDepositos = "ListaDepositosMovban " + XParam
    Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
    If rstDepositos.RecordCount > 0 Then
            
    With rstDepositos
            .MoveFirst
            Do
                If !FechaOrd > "19991017" Then
                If WDesde <= !FechaOrd And !FechaOrd <= WHasta Then
                        If Val(!Banco) >= Val(DesdeBanco.Text) And Val(!Banco) <= Val(HastaBanco.Text) Then
                            If Val(!Renglon) = 1 Then
                                WBanco = !Banco
                                WFecha = !Fecha
                                WFechaord = !FechaOrd
                                WAcredita = !Acredita
                                WAcreditaOrd = !AcreditaOrd
                                WObservaciones = "Deposito"
                                WNumero = !Deposito
                                WImporte = !Importe
                                WDeposito = !Deposito
                    
                                With rstMovban
                                    .AddNew
                                    !Banco = WBanco
                                    !Fecha = WFecha
                                    !FechaOrd = WFechaord
                                    !Acredita = WAcredita
                                    !AcreditaOrd = WAcreditaOrd
                                    !Observaciones = WObservaciones
                                    !Numero = WNumero
                                    !Credito = 0
                                    !Debito = WImporte
                                    !Comprobante = WDeposito
                                    !Empresa = 1
                                    !Titulo = WTitulo
                                    !Titulo1 = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                    !Proveedor = 0
                                    .Update
                                End With
                            End If
                        End If
                End If
                
                If WDesde > !FechaOrd Then
                        If Val(!Banco) >= Val(DesdeBanco.Text) And Val(!Banco) <= Val(HastaBanco.Text) Then
                                If Val(!Renglon) = 1 Then
                                    WInicial(!Banco) = WInicial(!Banco) + !Importe
                                End If
                        End If
                End If
                End If
                
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    rstDepositos.Close
    
    End If
    

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.Cuenta = " + "'" + "21" + "'"
    ZSql = ZSql + " or Recibos.Cuenta = " + "'" + "22" + "'"
    ZSql = ZSql + " or Recibos.Cuenta = " + "'" + "25" + "'"
    ZSql = ZSql + " or Recibos.Cuenta = " + "'" + "26" + "'"
    ZSql = ZSql + " or Recibos.Cuenta = " + "'" + "27" + "'"
    
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
            
    With rstRecibos
            .MoveFirst
            Do
                If !FechaOrd > "19991017" Then
                
                    If WDesde <= !FechaOrd And !FechaOrd <= WHasta Then
                    
                        WTipo = IIf(IsNull(!Tipo2), "0", !Tipo2)
                        If Val(WTipo) = 4 Then
                        
                            If Val(WEmpresa) = 1 Then
                            
                                Select Case Val(!Cuenta)
                                    Case 21
                                        WBanco = 3
                                        WFecha = !Fecha
                                        WFechaord = !FechaOrd
                                        WAcredita = !Fecha
                                        WAcreditaOrd = !FechaOrd
                                        WObservaciones = "Transferencia"
                                        WNumero = !Recibo
                                        WImporte = !Importe2
                                        WDeposito = !Recibo
                        
                                        With rstMovban
                                            .AddNew
                                            !Banco = WBanco
                                            !Fecha = WFecha
                                            !FechaOrd = WFechaord
                                            !Acredita = WAcredita
                                            !AcreditaOrd = WAcreditaOrd
                                            !Observaciones = WObservaciones
                                            !Numero = WNumero
                                            !Credito = 0
                                            !Debito = WImporte
                                            !Comprobante = WDeposito
                                            !Empresa = 1
                                            !Titulo = WTitulo
                                            !Titulo1 = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                            !Proveedor = 0
                                            .Update
                                        End With
                                    
                                    Case 22
                                        WBanco = 8
                                        WFecha = !Fecha
                                        WFechaord = !FechaOrd
                                        WAcredita = !Fecha
                                        WAcreditaOrd = !FechaOrd
                                        WObservaciones = "Transferencia"
                                        WNumero = !Recibo
                                        WImporte = !Importe2
                                        WDeposito = !Recibo
                            
                                        With rstMovban
                                            .AddNew
                                            !Banco = WBanco
                                            !Fecha = WFecha
                                            !FechaOrd = WFechaord
                                            !Acredita = WAcredita
                                            !AcreditaOrd = WAcreditaOrd
                                            !Observaciones = WObservaciones
                                            !Numero = WNumero
                                            !Credito = 0
                                            !Debito = WImporte
                                            !Comprobante = WDeposito
                                            !Empresa = 1
                                            !Titulo = WTitulo
                                            !Titulo1 = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                            !Proveedor = 0
                                            .Update
                                        End With
                                        
                                    Case 26
                                        WBanco = 12
                                        WFecha = !Fecha
                                        WFechaord = !FechaOrd
                                        WAcredita = !Fecha
                                        WAcreditaOrd = !FechaOrd
                                        WObservaciones = "Transferencia"
                                        WNumero = !Recibo
                                        WImporte = !Importe2
                                        WDeposito = !Recibo
                            
                                        With rstMovban
                                            .AddNew
                                            !Banco = WBanco
                                            !Fecha = WFecha
                                            !FechaOrd = WFechaord
                                            !Acredita = WAcredita
                                            !AcreditaOrd = WAcreditaOrd
                                            !Observaciones = WObservaciones
                                            !Numero = WNumero
                                            !Credito = 0
                                            !Debito = WImporte
                                            !Comprobante = WDeposito
                                            !Empresa = 1
                                            !Titulo = WTitulo
                                            !Titulo1 = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                            !Proveedor = 0
                                            .Update
                                        End With
                                        
                                    Case 27
                                        WBanco = 16
                                        WFecha = !Fecha
                                        WFechaord = !FechaOrd
                                        WAcredita = !Fecha
                                        WAcreditaOrd = !FechaOrd
                                        WObservaciones = "Transferencia"
                                        WNumero = !Recibo
                                        WImporte = !Importe2
                                        WDeposito = !Recibo
                            
                                        With rstMovban
                                            .AddNew
                                            !Banco = WBanco
                                            !Fecha = WFecha
                                            !FechaOrd = WFechaord
                                            !Acredita = WAcredita
                                            !AcreditaOrd = WAcreditaOrd
                                            !Observaciones = WObservaciones
                                            !Numero = WNumero
                                            !Credito = 0
                                            !Debito = WImporte
                                            !Comprobante = WDeposito
                                            !Empresa = 1
                                            !Titulo = WTitulo
                                            !Titulo1 = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                            !Proveedor = 0
                                            .Update
                                        End With
                                        
                                    Case Else
                                End Select
                                
                                    Else
                            
                                Select Case Val(!Cuenta)
                                    Case 22
                                        WBanco = 5
                                        WFecha = !Fecha
                                        WFechaord = !FechaOrd
                                        WAcredita = !Fecha
                                        WAcreditaOrd = !FechaOrd
                                        WObservaciones = "Transferencia"
                                        WNumero = !Recibo
                                        WImporte = !Importe2
                                        WDeposito = !Recibo
                        
                                        With rstMovban
                                            .AddNew
                                            !Banco = WBanco
                                            !Fecha = WFecha
                                            !FechaOrd = WFechaord
                                            !Acredita = WAcredita
                                            !AcreditaOrd = WAcreditaOrd
                                            !Observaciones = WObservaciones
                                            !Numero = WNumero
                                            !Credito = 0
                                            !Debito = WImporte
                                            !Comprobante = WDeposito
                                            !Empresa = 1
                                            !Titulo = WTitulo
                                            !Titulo1 = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                            !Proveedor = 0
                                            .Update
                                        End With
                                    
                                    Case 25
                                        WBanco = 12
                                        WFecha = !Fecha
                                        WFechaord = !FechaOrd
                                        WAcredita = !Fecha
                                        WAcreditaOrd = !FechaOrd
                                        WObservaciones = "Transferencia"
                                        WNumero = !Recibo
                                        WImporte = !Importe2
                                        WDeposito = !Recibo
                            
                                        With rstMovban
                                            .AddNew
                                            !Banco = WBanco
                                            !Fecha = WFecha
                                            !FechaOrd = WFechaord
                                            !Acredita = WAcredita
                                            !AcreditaOrd = WAcreditaOrd
                                            !Observaciones = WObservaciones
                                            !Numero = WNumero
                                            !Credito = 0
                                            !Debito = WImporte
                                            !Comprobante = WDeposito
                                            !Empresa = 1
                                            !Titulo = WTitulo
                                            !Titulo1 = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                            !Proveedor = 0
                                            .Update
                                        End With
                                        
                                    Case 26
                                        WBanco = 11
                                        WFecha = !Fecha
                                        WFechaord = !FechaOrd
                                        WAcredita = !Fecha
                                        WAcreditaOrd = !FechaOrd
                                        WObservaciones = "Transferencia"
                                        WNumero = !Recibo
                                        WImporte = !Importe2
                                        WDeposito = !Recibo
                            
                                        With rstMovban
                                            .AddNew
                                            !Banco = WBanco
                                            !Fecha = WFecha
                                            !FechaOrd = WFechaord
                                            !Acredita = WAcredita
                                            !AcreditaOrd = WAcreditaOrd
                                            !Observaciones = WObservaciones
                                            !Numero = WNumero
                                            !Credito = 0
                                            !Debito = WImporte
                                            !Comprobante = WDeposito
                                            !Empresa = 1
                                            !Titulo = WTitulo
                                            !Titulo1 = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                            !Proveedor = 0
                                            .Update
                                        End With
                                        
                                        
                                        
                                    Case Else
                                    
                                End Select
                                
                            End If
                            
                        End If
                        
                    End If
                    
                    If WDesde > !FechaOrd Then
                        WTipo = IIf(IsNull(!Tipo2), "0", !Tipo2)
                        If Val(WTipo) = 4 Then
                            WImporte = !Importe2
                            If Val(WEmpresa) = 1 Then
                                Select Case Val(!Cuenta)
                                    Case 21
                                        WInicial(3) = WInicial(3) + WImporte
                                    Case 22
                                        WInicial(8) = WInicial(8) + WImporte
                                    Case 25
                                        WInicial(12) = WInicial(12) + WImporte
                                    Case Else
                                End Select
                                    Else
                                Select Case Val(!Cuenta)
                                    Case 22
                                        WInicial(5) = WInicial(5) + WImporte
                                    Case 25
                                        WInicial(12) = WInicial(12) + WImporte
                                    Case Else
                                End Select
                            End If
                        End If
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
    
    
    If Val(WEmpresa) = 1 Then
        WInicial(3) = WInicial(3) - 4624.79 + 82277.33 - 21644.52
        Rem WInicial(3) = WInicial(3) - 4624.79 - 69613.86
        WInicial(8) = WInicial(8) + 65799.41 - 112141.1 + 11998.39 + 15008.45 + 10000.94 - 46211.58 + 46128.29 + 4135.52 - 434355.52 + 284428
        WInicial(9) = WInicial(9) - 982.73
        WInicial(11) = WInicial(11) + 34749.08
            Else
        WInicial(5) = WInicial(5) - 319209.66
    End If
    
    For XDa = 1 To 100
    
        If WInicial(XDa) <> 0 Then
    
            With rstMovban
                .AddNew
                !Banco = XDa
                !Fecha = "00/00/0000"
                !FechaOrd = "00000000"
                !Acredita = "00/00/0000"
                !AcreditaOrd = "00000000"
                !Observaciones = "Saldo Inicial"
                !Numero = 0
                If WInicial(XDa) > 0 Then
                    !Credito = 0
                    !Debito = WInicial(XDa)
                        Else
                    !Credito = Abs(WInicial(XDa))
                    !Debito = 0
                End If
                !Comprobante = "000000"
                !Empresa = 1
                !Titulo = WTitulo
                !Titulo1 = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                !Proveedor = 0
                .Update
            End With
        End If
        
    Next XDa
    
    da = 0
    With rstMovban
        .Index = "Clave"
        .Seek "=", da
        If .NoMatch = False Then
            Do
                .Edit
                
                WBanco = !Banco
                WNombre = ""
                
                spBanco = "ConsultaBancos " + "'" + Str$(WBanco) + "'"
                Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                If rstBanco.RecordCount > 0 Then
                    WNombre = rstBanco!Nombre
                    rstBanco.Close
                End If
                
                WProveedor = IIf(IsNull(!Proveedor), "0", !Proveedor)
                If Val(WProveedor) <> 0 Then
                    spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    If RstProveedor.RecordCount > 0 Then
                        WObservaciones = RstProveedor!Nombre
                        RstProveedor.Close
                    End If
                    !Observaciones = WObservaciones
                End If
                
                !Nombre = WNombre
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    listado.GroupSelectionFormula = "{Movban.banco} in " + DesdeBanco + " to " + HastaBanco
    Rem Listado.GroupSelectionFormula = "{Movban.banco} in 0 to 9999"
    
    listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    
    If Impresora.Value = True Then
        listado.Destination = 1
            Else
        listado.Destination = 0
    End If
    listado.Action = 1
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    With rstMovban
        .Close
    End With
    Desde.SetFocus
    PrgMovban.Hide
    Unload Me
    Menu.SetFocus
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

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Movban
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            DesdeBanco.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Private Sub DesdeBanco_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaBanco.Text = DesdeBanco.Text
        HastaBanco.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub HastaBanco_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()

    ReDim WInicial(1 To 100)

    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    DesdeBanco.Text = 1
    HastaBanco.Text = 9999
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spBanco = "ListaBancos"
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        With rstBanco
            .MoveFirst
            Do
                If .EOF = False Then
                    Auxi = Str$(rstBanco!Banco)
                    Call Ceros(Auxi, 4)
                    IngresaItem = Auxi + " " + rstBanco!Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = rstBanco!Banco
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstBanco.Close
    End If
            
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    WBanco = WIndice.List(Indice)
    spBanco = "ConsultaBanco " + "'" + Str$(WBanco) + "'"
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        DesdeBanco.Text = rstBanco!Banco
        HastaBanco.Text = rstBanco!Banco
        rstBanco.Close
                Else
        DesdeBanco.Text = WBanco
        HastaBanco.Text = WBanco
    End If
    Desde.SetFocus
    
End Sub

