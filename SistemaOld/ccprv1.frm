VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PrgCcprv1 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Cuenta Corriente de Proveedores"
   ClientHeight    =   7905
   ClientLeft      =   450
   ClientTop       =   585
   ClientWidth     =   11085
   LinkTopic       =   "Form2"
   ScaleHeight     =   7905
   ScaleWidth      =   11085
   Begin VB.TextBox Ayuda 
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   120
      Width           =   8295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Datos"
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1815
      Begin VB.OptionButton Total 
         Caption         =   "Total"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Pendiente 
         Caption         =   "Pendiente"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox Proveedor 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   0
      Text            =   " "
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Lee datos"
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   2595
      ItemData        =   "ccprv1.frx":0000
      Left            =   2520
      List            =   "ccprv1.frx":0007
      TabIndex        =   3
      Top             =   600
      Width           =   8295
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   3855
      Left            =   240
      TabIndex        =   14
      Top             =   3960
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6800
      _Version        =   327680
      Rows            =   10000
      Cols            =   9
   End
   Begin VB.Label Saldo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   375
      Left            =   8640
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label DesProveedor 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proveedor"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "PrgCcprv1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private Importe1 As Double
Private Importe2 As Double
Private Importe3 As Double
Private WTipo As Integer
Private WSalida As String
Private WSaldo As Double
Dim RstCtaPrv As Recordset
Dim spCtaprv As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim cParam As String
Dim XParam As String

Dim ZZNroInterno(10000) As String

Private Sub cmdClose_Click()
    PrgCcprv1.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Command1_Click()
    Sql1 = "UPDATE CtaCtePrv SET "
    Sql2 = " Saldo  = 0"
    Sql3 = " Where Proveedor = " + "'" + Proveedor.Text + "'"
    spCtaprv = Sql1 + Sql2 + Sql3
    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spProveedor = "ListaProveedoresOrdConsulta"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        With RstProveedor
            .MoveFirst
            Do
                If .EOF = False Then
                    Auxi = !Proveedor
                    Call Ceros(Auxi, 11)
                    IngresaItem = Auxi + "      " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Proveedor
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        RstProveedor.Close
    End If
            
    Rem Pantalla.Visible = True

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Muestra_DblClick()

    Muestra.Col = 1
    Tipo = Muestra.Text
    
    If Tipo = "OP" Or Tipo = "AN" Then
        Muestra.Col = 2
        WOPago = Trim(Str$(Val(Muestra.Text)))
        PrgpagoConsulta.Show
    End If
    
    If Tipo = "FC" Then
        Muestra.Col = 2
        WPasaNroInterno = ZZNroInterno(Muestra.Row)
      Rem by nan 5-07-2013
      
   Rem   PrgComprasConsulta.Show
        PrgConsultaCompras.Show
    
    End If
    
End Sub

Private Sub Pantalla_Click()

    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spProveedor = "ConsultaProveedores " + "'" + Claveven$ + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        Proveedor.Text = RstProveedor!Proveedor
        DesProveedor.Caption = RstProveedor!Nombre
        RstProveedor.Close
        Call Proceso_Click
            Else
        Proveedor.Text = Claveven$
    End If
    Proveedor.SetFocus
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    Muestra.ColWidth(0) = 150
    Muestra.ColWidth(1) = 500
    Muestra.ColWidth(2) = 1000
    Muestra.ColWidth(3) = 1300
    Muestra.ColWidth(4) = 1000
    Muestra.ColWidth(5) = 1100
    Muestra.ColWidth(6) = 1000
    Muestra.ColWidth(7) = 1300
    Muestra.ColWidth(8) = 1300
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Tipo"
    
    Muestra.Col = 2
    Muestra.Text = "Numero"
    
    Muestra.Col = 3
    Muestra.Text = "Fecha"
    
    Muestra.Col = 4
    Muestra.Text = "Debito"
    
    Muestra.Col = 5
    Muestra.Text = "Credito"
    
    Muestra.Col = 6
    Muestra.Text = "Saldo"
    
    Muestra.Col = 7
    Muestra.Text = "Vencimiento"
    
    Muestra.Col = 8
    Muestra.Text = "Vencimiento"
    
    Muestra.Col = 1
    Muestra.Row = 1
 
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Pendiente.Value = True
    
    Rem Proveedor.SetFocus
    
End Sub

Private Sub Proceso_Click()

    Erase ZZNroInterno

    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Tipo"
    
    Muestra.Col = 2
    Muestra.Text = "Numero"
    
    Muestra.Col = 3
    Muestra.Text = "Fecha"
    
    Muestra.Col = 4
    Muestra.Text = "Debito"
    
    Muestra.Col = 5
    Muestra.Text = "Credito"
    
    Muestra.Col = 6
    Muestra.Text = "Saldo"
    
    Muestra.Col = 7
    Muestra.Text = "Vencimiento"
    
    Muestra.Col = 8
    Muestra.Text = "Vencimiento"
    
    Muestra.Col = 1
    Muestra.Row = 1
    
    Renglon = 0
    WSaldo = 0
    
    XParam = "'" + Proveedor.Text + "','" _
                 + Proveedor.Text + "'"
    spCtaprv = "ListaCtaprvDesdeHasta " + XParam
    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
    If RstCtaPrv.RecordCount > 0 Then
    
    With RstCtaPrv
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                If Proveedor.Text <> !Proveedor Then
                    Exit Do
                End If
                                                                        
                If !Total >= 0 Then
                    Importe1 = !Total
                    Importe2 = 0
                    Importe3 = !Saldo
                        Else
                    Importe1 = 0
                    Importe2 = Abs(!Total)
                    Importe3 = !Saldo
                End If
                
                Call Redondeo(Importe3)
                
                If Importe3 <> 0 Or Total.Value = True Then
                
                        Renglon = Renglon + 1
                        
                        Muestra.Row = Renglon
                                                        
                        Muestra.Col = 1
                        Muestra.Text = !Impre
                        
                        Muestra.Col = 2
                        Muestra.Text = !Numero
                
                        Muestra.Col = 3
                        Muestra.Text = !Fecha
                
                        If Importe1 <> 0 Then
                            Muestra.Col = 4
                            Muestra.Text = Alinea("###,###.##", Str$(Importe1))
                                Else
                            Muestra.Col = 4
                            Muestra.Text = ""
                        End If
                
                        If Importe2 <> 0 Then
                            Muestra.Col = 5
                            Muestra.Text = Alinea("###,###.##", Str$(Importe2))
                                Else
                            Muestra.Col = 5
                            Muestra.Text = ""
                        End If
                
                        If Importe3 <> 0 Then
                            Muestra.Col = 6
                            Muestra.Text = Alinea("###,###.##", Str$(Importe3))
                                Else
                            Muestra.Col = 6
                            Muestra.Text = ""
                        End If
                        
                        WSaldo = WSaldo + Importe3
                
                        Muestra.Col = 7
                        Muestra.Text = !Vencimiento
                        
                        If !Vencimiento1 <> "" Then
                            Muestra.Col = 8
                            Muestra.Text = !Vencimiento1
                                Else
                            Muestra.Col = 8
                            Muestra.Text = ""
                        End If
                        
                        ZZNroInterno(Renglon) = !NroInterno
                    
                End If
                    
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
                If Proveedor.Text <> !Proveedor Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    RstCtaPrv.Close
    
    End If
    
    Saldo.Caption = Alinea("###,###.##", Str$(WSaldo))
    Muestra.Col = 1
    Muestra.Row = 1

End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WProveedor = Proveedor.Text
        Proveedor.Text = WProveedor
        spProveedor = "ConsultaProveedores "
        XParam = "'" & Proveedor.Text & "'"
        Set RstProveedor = db.OpenRecordset(spProveedor + XParam, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            DesProveedor.Caption = RstProveedor!Nombre
            RstProveedor.Close
            Call Proceso_Click
            Muestra.SetFocus
                Else
            Proveedor.SetFocus
        End If
    End If
End Sub

Private Sub Limpia_Vector()

    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Tipo"
    
    Muestra.Col = 2
    Muestra.Text = "Numero"
    
    Muestra.Col = 3
    Muestra.Text = "Fecha"
    
    Muestra.Col = 4
    Muestra.Text = "Debito"
    
    Muestra.Col = 5
    Muestra.Text = "Credito"
    
    Muestra.Col = 6
    Muestra.Text = "Saldo"
    
    Muestra.Col = 7
    Muestra.Text = "Vencimiento"
    
    Muestra.Col = 8
    Muestra.Text = "Vencimiento"
    
    Muestra.Col = 1
    Muestra.Row = 1

     Call Consulta_Click

End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    If Ayuda.Text <> "" Then
        spProveedor = "ListaProveedoresOrdConsultaII " + "'" + Ayuda.Text + "'"
            Else
        spProveedor = "ListaProveedoresOrdConsulta"
    End If
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
    
    With RstProveedor
        .MoveFirst
        Do
            If .EOF = False Then
            
                da = Len(!Nombre) - WEspacios
                
                For aa = 1 To da
                    If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(!Nombre), aa, WEspacios) Then
                    
                    
                        Auxi = Str$(!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + "    " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
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
    
    RstProveedor.Close
    
    End If
    
    End If

End Sub


