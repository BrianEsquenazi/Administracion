VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Prgpago 
   Caption         =   "Ingresos de Pagos a Proveedores"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   540
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6780
   ScaleWidth      =   9480
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   21
      Text            =   " "
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   19
      Text            =   " "
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox Concepto 
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Text            =   " "
      Top             =   360
      Width           =   735
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   6840
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3240
      TabIndex        =   12
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSDBGrid.DBGrid DbGrid1 
      Height          =   3375
      Left            =   120
      OleObjectBlob   =   "pago2.frx":0000
      TabIndex        =   10
      Top             =   2400
      Width           =   9255
   End
   Begin VB.TextBox Orden 
      Height          =   285
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   735
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "pago2.frx":09C2
      Left            =   5520
      List            =   "pago2.frx":09C9
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   7320
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   7320
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6240
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   8400
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   6240
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Importe"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label sfd 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Creditos 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Doc. : 1) Ef.   2) Bco.  3) Ch. Terc."
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Label DesConcepto 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   2520
      TabIndex        =   14
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Concepto"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nro. Orden de Pago"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "Prgpago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 10 ' Número máximo de campos del conjunto de registros.

Private Sub Suma_Datos()
    Debitos.Caption = ""
    Creditos.Caption = ""
    
    For iRow = 0 To 9
        DbGrid1.Col = 4
        DbGrid1.Row = iRow
        If Val(DbGrid1.Text) <> 0 Then
            Debitos.Caption = Str$(Val(Debitos.Caption) + CDbl(DbGrid1.Text))
        End If
        DbGrid1.Col = 10
        DbGrid1.Row = iRow
        If Val(DbGrid1.Text) <> 0 Then
            Creditos.Caption = Str$(Val(Creditos.Caption) + CDbl(DbGrid1.Text))
        End If
    Next iRow
    Debitos.Caption = PUsing("###,###.##", Debitos.Caption)
    Creditos.Caption = PUsing("###,###.##", Creditos.Caption)
    DbGrid1.Col = 0
    DbGrid1.Row = 0
End Sub

Private Sub Lee_Datos()
    Renglon = 0
    Debito = 0
    Credito = 0
    Do
        With rstPagos
            .Index = "Clave"
            Renglon = Renglon + 1
            Auxi1 = Str$(Renglon)
            Call Ceros(Auxi1, 2)
            .Seek "=", Orden.Text + Auxi1
            If .NoMatch = False Then
                Select Case Val(!TipoReg)
                    Case 1
                        Debito = Debito + 1
                        DbGrid1.Row = Debito - 1
                        DbGrid1.Col = 0
                        DbGrid1.Text = !Tipo1
                        DbGrid1.Col = 1
                        DbGrid1.Text = !Letra1
                        DbGrid1.Col = 2
                        DbGrid1.Text = !Punto1
                        DbGrid1.Col = 3
                        DbGrid1.Text = !Numero1
                        DbGrid1.Col = 4
                        DbGrid1.Text = !Importe1
                        DbGrid1.Text = PUsing("###,###.##", DbGrid1.Text)
                    Case 2
                        Credito = Credito + 1
                        DbGrid1.Row = Credito - 1
                        DbGrid1.Col = 5
                        DbGrid1.Text = !Tipo2
                        DbGrid1.Col = 6
                        DbGrid1.Text = !Numero2
                        DbGrid1.Col = 7
                        DbGrid1.Text = !Fecha2
                        DbGrid1.Col = 8
                        DbGrid1.Text = !Banco2
                        DbGrid1.Col = 9
                        If !Observaciones <> "" Then
                            DbGrid1.Text = !Observaciones
                        End If
                        DbGrid1.Col = 10
                        DbGrid1.Text = !Importe2
                        DbGrid1.Text = PUsing("###,###.##", DbGrid1.Text)
                    Case Else
                End Select
                    Else
                Exit Do
            End If
        End With
    Loop
End Sub
Sub Verifica_datos()
End Sub
Sub Format_datos()
    Rem Retganancias.text = PUsing("###,###.##", Retganancias.text)
End Sub

Sub Imprime_Datos()
    With rstProveedor
        .Index = "Proveedor"
        .Seek "=", Proveedor.Text
        If .NoMatch = False Then
            Proveedor.Text = !Proveedor
            DesProveedor.Caption = !Nombre
            Call Format_datos
        End If
    End With
End Sub

Private Sub cmdAdd_Click()

    If Orden.Text <> "" And Fecha.Text <> "" And Proveedor.Text <> "" Then
    
    If Existe <> "S" Then
    
        Call Suma_Datos
        
        Debito = 0
        Credito = 0
        If Val(Debitos.Caption) <> 0 Then
            Debito = CDbl(Debitos.Caption)
        End If
        
        If Val(Creditos.Caption) <> 0 Then
            Credito = CDbl(Creditos.Caption)
        End If
        
        If Debito = Credito Then
    
        With rstPagos
            Renglon = 0
            .Index = "Clave"
            For iRow = 0 To 9
                WRow = iRow
                DbGrid1.Col = 4
                DbGrid1.Row = iRow
                If Val(DbGrid1.Text) <> 0 Then
                    .AddNew
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    !Orden = Orden.Text
                    !Renglon = Auxi1
                    !Proveedor = Proveedor.Text
                    !Fecha = Fecha.Text
                    If Tipo1.Value = True Then
                        !TipoOrd = "1"
                    End If
                    If Tipo2.Value = True Then
                        !TipoOrd = "2"
                    End If
                    !TipoReg = "1"
                    DbGrid1.Col = 0
                    !Tipo1 = DbGrid1.Text
                    DbGrid1.Col = 1
                    !Letra1 = DbGrid1.Text
                    DbGrid1.Col = 2
                    !Punto1 = DbGrid1.Text
                    DbGrid1.Col = 3
                    !Numero1 = DbGrid1.Text
                    DbGrid1.Col = 4
                    !Importe1 = DbGrid1.Text
                    !Tipo2 = ""
                    !Numero2 = ""
                    !Fecha2 = ""
                    !Banco2 = 0
                    !Importe2 = 0
                    !Observaciones = " "
                    !Empresa = 1
                    !Clave = !Orden + !Renglon
                    !Concepto = 0
                    .Update
                    .Bookmark = .LastModified
                    
                    WLetra = !Letra1
                    WTipo = !Tipo1
                    WPunto = !Punto1
                    WNumero = !Numero1
                    WImporte = !Importe1
                    
                    With rstCtaCtePrv
                        .Index = "CtaCte"
                        Auxi$ = Proveedor.Text
                        Call Ceros(Auxi$, 6)
                        Claveven$ = Auxi$
                        Claveven$ = Claveven$ + WLetra + WTipo + WPunto + WNumero
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            .Edit
                            !Saldo = !Saldo - WImporte
                            .Update
                            .Bookmark = .LastModified
                        End If
                    End With
                    
                End If
                
                DbGrid1.Col = 10
                DbGrid1.Row = iRow
                If Val(DbGrid1.Text) <> 0 Then
                    .AddNew
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    !Orden = Orden.Text
                    !Renglon = Auxi1
                    !Proveedor = Proveedor.Text
                    !Fecha = Fecha.Text
                    If Tipo1.Value = True Then
                        !TipoOrd = "1"
                    End If
                    If Tipo2.Value = True Then
                        !TipoOrd = "2"
                    End If
                    !TipoReg = "2"
                    !Tipo1 = ""
                    !Letra1 = ""
                    !Punto1 = ""
                    !Numero1 = ""
                    !Importe1 = 0
                    DbGrid1.Col = 5
                    !Tipo2 = DbGrid1.Text
                    DbGrid1.Col = 6
                    !Numero2 = DbGrid1.Text
                    DbGrid1.Col = 7
                    !Fecha2 = DbGrid1.Text
                    DbGrid1.Col = 8
                    !Banco2 = Val(DbGrid1.Text)
                    DbGrid1.Col = 9
                    !Observaciones = DbGrid1.Text
                    DbGrid1.Col = 10
                    !Importe2 = DbGrid1.Text
                    !Empresa = 1
                    !Clave = !Orden + !Renglon
                    !Concepto = 0
                    .Update
                    .Bookmark = .LastModified
                    
                    With rstRecibos
                        .Index = "Clave"
                        DbGrid1.Col = 11
                        Claveven$ = DbGrid1.Text
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            .Edit
                            !Estado2 = "X"
                            .Update
                        End If
                    End With
                    
                End If
                
            Next iRow
        End With
        
        With rstEmpresa
            .Index = "Empresa"
            Claveven$ = "1"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                WCtaProveedor = !CtaProveedores
                WCtaEfectivo = !CtaEfectivo
                WCtaCheques = !CtaCheque
            End If
        End With
        
        With rstImputac
            Renglon = 0
            .Index = "Clave"
            
            For iRow = 0 To 9
                WRow = iRow
                DbGrid1.Col = 4
                DbGrid1.Row = iRow
                If Val(DbGrid1.Text) <> 0 Then
                    .AddNew
                    !TipoMovi = "3"
                    !Proveedor = Proveedor.Text
                    !TipoComp = "01"
                    !LetraComp = "A"
                    !PuntoComp = "0000"
                    !NroComp = Orden.Text
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    !Renglon = Auxi1$
                    !Fecha = Fecha.Text
                    !Observaciones = DesProveedor.Caption
                    !Cuenta = WCtaProveedore
                    !Debito = 0
                    !Credito = CDbl(DbGrid1.Text)
                    !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !Titulo = "Pagos"
                    !Empresa = 1
                    !Clave = !TipoMovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
                    .Update
                End If
                
                DbGrid1.Col = 10
                DbGrid1.Row = iRow
                If Val(DbGrid1.Text) <> 0 Then
                    .AddNew
                    !TipoMovi = "3"
                    !Proveedor = Proveedor.Text
                    !TipoComp = "01"
                    !LetraComp = "A"
                    !PuntoComp = "0000"
                    !NroComp = Orden.Text
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    !Renglon = Auxi1$
                    !Fecha = Fecha.Text
                    !Observaciones = DesProveedor.Caption
                    !Debito = CDbl(DbGrid1.Text)
                    !Credito = 0
                    !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !Titulo = "Pagos"
                    !Empresa = 1
                    DbGrid1.Col = 5
                    Select Case Val(DbGrid1.Text)
                        Case 2
                            !Cuenta = "999999"
                            With rstBanco
                                DbGrid1.Col = 8
                                .Index = "Banco"
                                .Seek "=", Val(DbGrid1.Text)
                                If .NoMatch = False Then
                                    WCuenta = !Cuenta
                                End If
                            End With
                            !Cuenta = WCtaCheques
                        Case 3
                            !Cuenta = WCtaCheques
                        Case Else
                            !Cuenta = WCtaEfectivo
                    End Select
                    !Clave = !TipoMovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
                    .Update
                End If
                
            Next iRow
        End With
        
        WLetra = " "
        WTipo = "04"
        WPunto = "0000"
        WNumero = Orden.Text
        WProveedor = Proveedor.Text
        
        Call Ceros(WNumero, 8)
        Call Ceros(WProveedor, 6)
        
        With rstCtaCtePrv
            .Index = "CtaCte"
            .Seek "=", WProveedor + WLetra + WTipo + WPunto + WNumero
            If .NoMatch Then
                .AddNew
                !Proveedor = Proveedor.Text
                !Letra = WLetra
                !Tipo = WTipo
                !Punto = WPunto
                !Numero = WNumero
                !Fecha = Fecha.Text
                !Estado = "1"
                !Vencimiento = Fecha.Text
                !Total = Debito * -1
                !Saldo = 0
                !Clave = WProveedor + WLetra + WTipo + WPunto + WNumero
                !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                !OrdVencimiento = !OrdFecha
                !Impre = "OP"
                !Empresa = 1
                .Update
                .Bookmark = .LastModified
                    Else
                .Edit
                !Proveedor = Proveedor.Text
                !Letra = WLetra
                !Tipo = WTipo
                !Punto = WPunto
                !Numero = WNumero
                !Fecha = Fecha.Text
                !Estado = "1"
                !Vencimiento = Fecha.Text
                !Total = Debito * -1
                !Saldo = 0
                !Clave = WProveedor + WLetra + WTipo + WPunto + WNumero
                !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                !OrdVencimiento = !OrdFecha
                !Impre = "OP"
                !Empresa = 1
                .Update
                .Bookmark = .LastModified
            End If
        End With

        Call CmdLimpiar_Click
        Orden.SetFocus
        
        End If
        
        End If
        
    End If
End Sub

Private Sub cmdDelete_Click()
    If Orden.Text <> "" Then
                
            Rem Borro los datos anteriores
            
            Rem For iRow = 0 To 20
            Rem     Auxi1 = Str$(iRow)
            Rem     Call Ceros(Auxi1, 2)
            Rem     .Seek "=", Orden.text + Auxi1
            Rem     If .NoMatch = False Then
            Rem         .Delete
            Rem     End If
            Rem Next iRow

    End If
    Proveedor.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    For iCol = 0 To 10
        For iRow = 0 To 9
            DbGrid1.Col = iCol
            DbGrid1.Row = iRow
            DbGrid1.Text = ""
        Next iRow
    Next iCol
    Orden.Text = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Fecha.Text = "  /  /    "
    Tipo1.Value = True
    Tipo2.Value = False
    Debitos.Caption = ""
    Creditos.Caption = ""
    Orden.SetFocus
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Orden.SetFocus
    Prgpago.Hide
    Menu.SetFocus
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub
Private Sub Orden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi1 = Orden.Text
        Call Ceros(Auxi1, 6)
        Orden.Text = Auxi1
        
        With rstPagos
            Existe = "N"
            .Index = "Clave"
            Claveven$ = Orden.Text + "01"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Existe = "S"
                Proveedor.Text = !Proveedor
                Fecha.Text = !Fecha
            End If
        End With
        If Existe = "S" Then
            Call Imprime_Datos
            Call Lee_Datos
            Call Suma_Datos
            DbGrid1.Col = 0
            DbGrid1.Row = 0
            DbGrid1.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Proveedor.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Proveedor.Text) <> 0 Then
            With rstProveedor
                .Index = "Proveedor"
                Claveven$ = Proveedor.Text
                .Seek "=", Proveedor.Text
                If .NoMatch Then
                    Proveedor.Text = Claveven$
                    Proveedor.SetFocus
                        Else
                    Proveedor.Text = !Proveedor
                    DesProveedor.Caption = !Nombre
                    Rem Call Imprime_Datos
                    DbGrid1.Col = 0
                    DbGrid1.Row = 0
                    DbGrid1.SetFocus
                End If
            End With
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     XRow = DbGrid1.Row
     XCol = DbGrid1.Col

     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Cuenta Corrientes"
     Opcion.AddItem "Cheques terceros"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            With rstProveedor
                .Index = "Proveedor"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Proveedor + " " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case 1
            With rstCtaCtePrv
                .Index = "ClaveImpre"
                .Seek ">", Proveedor.Text + Space$(100)
                If .NoMatch = False Then
                Do
                    If .EOF = False Then
                        If Val(Proveedor.Text) = Val(!Proveedor) Then
                            If !Saldo <> 0 Then
                                Auxi$ = Str$(!Saldo)
                                Auxi$ = PUsing("###,###.##", Auxi$)
                                IngresaItem = !Impre + " " + !Letra + " " + !Punto + " " + !Numero + " " + !Fecha + " " + Auxi$
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Clave
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                                Else
                        Exit Do
                    End If
                Loop
                End If
            End With
            
        Case 2
            With rstRecibos
                .Index = "Clave"
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Val(!TipoReg) = 2 Then
                            If Val(!Tipo2) = 2 And !Estado2 <> "X" Then
                                Auxi$ = Str$(!Importe2)
                                Auxi$ = PUsing("###,###.##", Auxi$)
                                IngresaItem = !Numero2 + " " + !Banco2 + " " + !Fecha2 + " " + Auxi$
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Clave
                                WIndice.AddItem IngresaItem
                            End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
     
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            With rstProveedor
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                Proveedor.Text = Claveven$
                .Index = "Proveedor"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    DesProveedor.Caption = !Nombre
                            Else
                    Proveedor.Text = ""
                End If
            End With
                
            Proveedor.SetFocus
            
        Case 1
            With rstCtaCtePrv

                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "CtaCte"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 0
                    DbGrid1.Text = !Tipo
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 1
                    DbGrid1.Text = !Letra
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 2
                    DbGrid1.Text = !Punto
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 3
                    DbGrid1.Text = !Numero
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 4
                    DbGrid1.Text = !Saldo
                    DbGrid1.Text = PUsing("###,###.##", DbGrid1.Text)
                    
                    Call Suma_Datos
                    
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 4
                    
                End If
            End With
                
            DbGrid1.Row = XRow
            DbGrid1.Col = 0
            DbGrid1.SetFocus
            
        Case 2
            With rstRecibos
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Clave"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    DbGrid1.Col = 6
                    DbGrid1.Text = !Numero2
                
                    DbGrid1.Col = 8
                    DbGrid1.Text = ""
                
                    DbGrid1.Col = 7
                    DbGrid1.Text = !Fecha2
                
                    DbGrid1.Col = 9
                    DbGrid1.Text = !Banco2
                
                    DbGrid1.Col = 10
                    DbGrid1.Text = !Importe2
                    DbGrid1.Text = PUsing("###,###.##", DbGrid1.Text)
                    
                    DbGrid1.Col = 11
                    DbGrid1.Text = Claveven$
                    
                    Call Suma_Datos
                    
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 4
                    
                End If
                If DbGrid1.Row < 10 Then
                    DbGrid1.Row = DbGrid1.Row + 1
                    DbGrid1.Col = 5
                    KeyCode = 0
                            Else
                    DbGrid1.Col = 5
                    KeyCode = 0
                End If
            End With
                
        Case Else
    End Select
    
End Sub
Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DbGrid1.Col
    
            Case 0
                If KeyCode = 13 Then
                    If Val(DbGrid1.Text) = 1 Or Val(DbGrid1.Text) = 2 Or Val(DbGrid1.Text) = 3 Then
                        Auxi$ = Str$(Val(DbGrid1.Text))
                        Call Ceros(Auxi$, 2)
                        DbGrid1.Text = Auxi$
                        DbGrid1.Col = 1
                        KeyCode = 0
                            Else
                        DbGrid1.Col = 0
                        KeyCode = 0
                    End If
                End If
                
            Case 1
                If KeyCode = 13 Then
                    DbGrid1.Text = Left$(DbGrid1.Text, 1)
                    If DbGrid1.Text = "A" Or DbGrid1.Text = "C" Then
                        DbGrid1.Col = 2
                        KeyCode = 0
                        Rem no hago anda
                            Else
                        DbGrid1.Col = 1
                        KeyCode = 0
                    End If
                End If
                
            Case 2
                If KeyCode = 13 Then
                    Auxi$ = Str$(Val(DbGrid1.Text))
                    Call Ceros(Auxi$, 4)
                    DbGrid1.Text = Auxi$
                    DbGrid1.Col = 3
                    KeyCode = 0
                End If
                
            Case 3
                If KeyCode = 13 Then
                
                    Auxi$ = Str$(Val(DbGrid1.Text))
                    Call Ceros(Auxi$, 8)
                    DbGrid1.Text = Auxi$
                
                    With rstCtaCtePrv
                        .Index = "CtaCte"
                        Auxi$ = Proveedor.Text
                        Call Ceros(Auxi$, 6)
                        Claveven$ = Auxi$
                        DbGrid1.Col = 1
                        Claveven$ = Claveven$ + DbGrid1.Text
                        DbGrid1.Col = 0
                        Claveven$ = Claveven$ + DbGrid1.Text
                        DbGrid1.Col = 2
                        Claveven$ = Claveven$ + DbGrid1.Text
                        DbGrid1.Col = 3
                        Claveven$ = Claveven$ + DbGrid1.Text
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            DbGrid1.Col = 4
                            XRow = DbGrid1.Row
                            If Val(DbGrid1.Text) = 0 Then
                                DbGrid1.Text = !Saldo
                                Call Suma_Datos
                                DbGrid1.Col = 4
                                DbGrid1.Row = XRow
                            End If
                            DbGrid1.Col = 4
                            KeyCode = 0
                                Else
                            DbGrid1.Col = 0
                            KeyCode = 0
                        End If
                    End With
                End If
                
            Case 4
                If KeyCode = 13 Then
                
                    With rstCtaCtePrv
                        .Index = "CtaCte"
                        Auxi$ = Proveedor.Text
                        Call Ceros(Auxi$, 6)
                        Claveven$ = Auxi$
                        DbGrid1.Col = 1
                        Claveven$ = Claveven$ + DbGrid1.Text
                        DbGrid1.Col = 0
                        Claveven$ = Claveven$ + DbGrid1.Text
                        DbGrid1.Col = 2
                        Claveven$ = Claveven$ + DbGrid1.Text
                        DbGrid1.Col = 3
                        Claveven$ = Claveven$ + DbGrid1.Text
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            Saldo = !Saldo
                                Else
                            Saldo = 0
                        End If
                    End With
                
                    DbGrid1.Col = 4
                    If Val(DbGrid1.Text) > Saldo Then
                        DbGrid1.Text = ""
                        DbGrid1.Col = 4
                        KeyCode = 0
                            Else
                        DbGrid1.Text = PUsing("###,###.##", DbGrid1.Text)
                        Call Suma_Datos
                        If DbGrid1.Row < 10 Then
                            DbGrid1.Row = DbGrid1.Row + 1
                            DbGrid1.Col = 0
                            KeyCode = 0
                                Else
                            DbGrid1.Col = 0
                            KeyCode = 0
                        End If
                    End If
                End If
                
            Case 5
                If KeyCode = 13 Then
                    If Val(DbGrid1.Text) = 1 Or Val(DbGrid1.Text) = 2 Or Val(DbGrid1.Text) = 3 Then
                        Auxi$ = Str$(Val(DbGrid1.Text))
                        Call Ceros(Auxi$, 2)
                        DbGrid1.Text = Auxi$
                        
                        Select Case Val(DbGrid1.Text)
                        
                            Case 1
                                DbGrid1.Col = 6
                                DbGrid1.Text = ""
                                DbGrid1.Col = 7
                                DbGrid1.Text = ""
                                DbGrid1.Col = 8
                                DbGrid1.Text = ""
                                DbGrid1.Col = 9
                                DbGrid1.Text = ""
                                DbGrid1.Col = 10
                                KeyCode = 0
                                
                            Case 3
                                Call Consulta_Click
                                
                            Case Else
                                DbGrid1.Col = 6
                                KeyCode = 0
                                
                        End Select
                        
                            Else
                            
                        DbGrid1.Col = 5
                        KeyCode = 0
                        
                    End If
                End If
                
            Case 6
                If KeyCode = 13 Then
                    DbGrid1.Col = 5
                    If Val(DbGrid1.Text) = 3 Then
                        DbGrid1.Col = 6
                        KeyCode = 0
                            Else
                        Auxi$ = Str$(Val(DbGrid1.Text))
                        Call Ceros(Auxi$, 8)
                        DbGrid1.Text = Auxi$
                        DbGrid1.Col = 7
                        KeyCode = 0
                    End If
                End If
                
            Case 7
                If KeyCode = 13 Then
                    DbGrid1.Col = 7
                    
                    Call Valida_fecha(DbGrid1.Text, Auxi)
                    If Auxi <> "S" Then
                        DbGrid1.Col = 7
                        KeyCode = 0
                                Else
                        DbGrid1.Col = 8
                        KeyCode = 0
                    End If
                End If
                
            Case 8
                If KeyCode = 13 Then
                    With rstBanco
                        .Index = "Banco"
                        DbGrid1.Col = 8
                        Claveven$ = DbGrid1.Text
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            DbGrid1.Col = 9
                            DbGrid1.Text = !Nombre
                            DbGrid1.Col = 10
                            KeyCode = 0
                                Else
                            DbGrid1.Col = 8
                            KeyCode = 0
                        End If
                    End With
                End If

            Case 10
                If KeyCode = 13 Then
                    iRow = DbGrid1.Row
                    DbGrid1.Col = 10
                    DbGrid1.Text = PUsing("###,###.##", DbGrid1.Text)
                    Call Suma_Datos
                    DbGrid1.Row = iRow
                    If DbGrid1.Row < 10 Then
                        DbGrid1.Row = DbGrid1.Row + 1
                        DbGrid1.Col = 5
                        KeyCode = 0
                            Else
                        DbGrid1.Col = 5
                        KeyCode = 0
                    End If
                End If

            Case Else
                
    End Select
    
End Sub

' Cuando el usuario hace clic en el icono Agregar, esta subrutina agrega una
' nueva fila a la variable RowBuf y un marcador a la variable NewRowBookmark
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows = mTotalRows + 1
ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
NewRowBookmark = mTotalRows - 1 'Establece el marcador a la última fila.

' El bucle siguiente agrega un nuevo registro a la base de datos.
For iCol = 0 To UBound(UserData, 1)
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
    Else
        ' Si no se establece ningún valor para la columna, usa DefaultValue
        UserData(iCol, mTotalRows - 1) = DbGrid1.Columns(iCol).DefaultValue
    End If
Next iCol

End Sub

' Esta subrutina elimina una fila basándose en su marcador.
Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

' Mueve todas las filas encima de la fila eliminada de
' la matriz.

For iRow = Bookmark + 1 To mTotalRows - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows = mTotalRows - 1

End Sub

' Se llama a esta subrutina cada vez que DBGrid quiere mostrar
' datos nuevos.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid está solicitando filas, así que se las damos

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' Si StartLocation es Null, empieza a leer por el final
' o por el principio del conjunto de datos.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Busca la posición para empezar a leer, basándose en el marcador
    ' StartLocation y en la variable iIncr
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfiere datos de nuestra matriz de conjunto de datos al objeto RowBuf
' que DBGrid utiliza para presentar los datos
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
    For iCol = 0 To UBound(UserData, 1)
        RowBuf.Value(iRow, iCol) = UserData(iCol, CurRow&)
    Next iCol
    ' Establece el marcador mediante CurRow&, que es también
    ' nuestro índice de matriz
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

' Esta subrutina actualiza los datos de la matriz después de
' haberse modificado.

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Se están actualizando los datos

' Actualiza cada columna de la matriz de conjuntos de datos
For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub


Private Sub Form_Load()

ReDim UserData(0 To 9, 0 To 11)

mTotalRows& = 10

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DbGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DbGrid1.Columns.Count - 1 To 0 Step -1
     DbGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 11
    DbGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DbGrid1.Columns(newcnt).Caption = "Tipo"
             DbGrid1.Columns(newcnt).Width = 400
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 1
             DbGrid1.Columns(newcnt).Caption = "Letra"
             DbGrid1.Columns(newcnt).Width = 450
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 2
             DbGrid1.Columns(newcnt).Caption = "Punto"
             DbGrid1.Columns(newcnt).Width = 600
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 3
             DbGrid1.Columns(newcnt).Caption = "Numero"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 4
             DbGrid1.Columns(newcnt).Caption = "Importe"
             DbGrid1.Columns(newcnt).Width = 1200
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 5
             DbGrid1.Columns(newcnt).Caption = "Tipo"
             DbGrid1.Columns(newcnt).Width = 400
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 6
             DbGrid1.Columns(newcnt).Caption = "Numero"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 7
             DbGrid1.Columns(newcnt).Caption = "Fecha"
             DbGrid1.Columns(newcnt).Width = 1150
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 8
             DbGrid1.Columns(newcnt).Caption = "Banco"
             DbGrid1.Columns(newcnt).Width = 500
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 9
             DbGrid1.Columns(newcnt).Caption = "Nombre"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 10
             DbGrid1.Columns(newcnt).Caption = "Importe"
             DbGrid1.Columns(newcnt).Width = 1200
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 11
             DbGrid1.Columns(newcnt).Caption = ""
             DbGrid1.Columns(newcnt).Width = 1
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case Else

     End Select
     DbGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
     
    Tipo1.Value = True
    Tipo2.Value = False
    
End Sub
