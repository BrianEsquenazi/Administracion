VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PrgCC 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Cuenta Corriente de Clientes"
   ClientHeight    =   8175
   ClientLeft      =   435
   ClientTop       =   435
   ClientWidth     =   11040
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   11040
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   5895
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   10398
      _Version        =   327680
      Rows            =   2000
      Cols            =   9
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8040
      TabIndex        =   10
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Todos 
         Caption         =   "Total"
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
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Pendiente 
         Caption         =   "Pendiente"
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
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6240
      TabIndex        =   9
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton Total 
         Caption         =   "Total"
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
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Documentos 
         Caption         =   "Documentos"
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
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton CtaCte 
         Caption         =   "Cta.Cte."
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4560
      TabIndex        =   8
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Dolares 
         Caption         =   "Dolares"
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
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Pesos 
         Caption         =   "Pesos"
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
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Cliente 
      Height          =   375
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   6
      Text            =   " "
      Top             =   1680
      Width           =   1695
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
      Height          =   300
      Left            =   3480
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
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
      Height          =   1425
      ItemData        =   "cc.frx":0000
      Left            =   120
      List            =   "cc.frx":0007
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
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
      Height          =   300
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Saldo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   8520
      TabIndex        =   19
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
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
      Left            =   7560
      TabIndex        =   18
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   3240
      TabIndex        =   7
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "PrgCC"
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
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstCtacte As Recordset
Dim spCtecte As String
Dim XParam As String

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub cmdClose_Click()
        
    Cliente.Text = ""
    DesCliente.Caption = ""

    Pesos.Value = True
    CtaCte.Value = True
    Pendiente.Value = True
    
    Cliente.Text = ""
    Saldo.Caption = ""
    
    Cliente.SetFocus
                
    PrgCC.Hide
    Unload Me
    PrgPed.Show
    
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstCliente!Cliente + "     " + rstCliente!Razon
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstCliente!Cliente
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
        End With
    End If
    
    Pantalla.Visible = True

End Sub

Private Sub Muestra_Click()

    Muestra.Col = 1
    Tipo = Muestra.Text
    
    If Tipo = "Rec" Then
        Muestra.Col = 2
        WRecibo = Muestra.Text
        PrgRecPed.Show
    End If

End Sub

Private Sub pantalla_Click()
    Rem Pantalla.Visible = False
       
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    Cliente.Text = Claveven$
    
    spCliente = "ConsultaCliente " + "'" + Claveven$ + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Cliente.Text = rstCliente!Cliente
        DesCliente.Caption = rstCliente!Razon
        rstCliente.Close
        Call Proceso_Click
            Else
        Cliente.Text = Claveven$
    End If
    Cliente.SetFocus
    
End Sub

Private Sub Form_Load()

    Muestra.ColWidth(0) = 50
    Muestra.ColWidth(1) = 600
    Muestra.ColWidth(2) = 1000
    Muestra.ColWidth(3) = 1200
    Muestra.ColWidth(4) = 1500
    Muestra.ColWidth(5) = 1500
    Muestra.ColWidth(6) = 1500
    Muestra.ColWidth(7) = 1200
    Muestra.ColWidth(8) = 1200
    
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

    Cliente.Text = ""
    DesCliente.Caption = ""

    Pesos.Value = True
    CtaCte.Value = True
    Pendiente.Value = True
    
    Cliente.Text = PCliente
    Call lee
   
End Sub

Private Sub Proceso_Click()

    Cliente.Text = UCase(Cliente.Text)
    
    WSalida = "N"
    
    Muestra.Clear
        
    Renglon = 0
    WSaldo = 0
    Cambia = "N"
    
    XParam = "'" + Cliente.Text + "','" _
                 + Cliente.Text + "'"
    spCtacte = "ListaCtacteDesdeHasta" + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
    
    With rstCtacte
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                WPasa = "N"
                
                If CtaCte.Value = True Then
                    If !Tipo < 50 Then
                        WPasa = "S"
                    End If
                End If
                
                If Documentos.Value = True Then
                    If !Tipo >= 50 Then
                        WPasa = "S"
                    End If
                End If
                
                If Total.Value = True Then
                    WPasa = "S"
                End If
                    
                If WPasa = "S" Then
                    If Pesos.Value = True Then
                        If !Total > 0 Then
                            Importe1 = !Total
                            Importe2 = 0
                                Else
                            Importe1 = 0
                            Importe2 = !Total
                        End If
                        Importe3 = !Saldo
                            Else
                        If !Totalus > 0 Then
                            Importe1 = !Totalus
                            Importe2 = 0
                                Else
                            Importe1 = 0
                            Importe2 = !Totalus
                        End If
                        Importe3 = !Saldous
                    End If
                    
                    Call Redondeo(Importe3)
                
                    If Importe3 <> 0 Or Todos.Value = True Then
                
                        Renglon = Renglon + 1
                        
                        Muestra.Row = Renglon
                        
                        Select Case !Tipo
                            Case 1
                                Muestra.Col = 1
                                Muestra.Text = "Fac"
                            Case 2
                                Muestra.Col = 1
                                Muestra.Text = "Dev"
                            Case 3
                                Muestra.Col = 1
                                Muestra.Text = "Fac"
                            Case 4
                                Muestra.Col = 1
                                Select Case Left$(!Impre, 2)
                                    Case "DC"
                                        Muestra.Text = "D.C"
                                    Case "CH"
                                        Muestra.Text = "CHR"
                                    Case Else
                                        Muestra.Text = "N/D"
                                End Select
                            Case 5
                                Muestra.Col = 1
                                Select Case Left$(!Impre, 2)
                                    Case "DC"
                                        Muestra.Text = "D.C"
                                    Case "CH"
                                        Muestra.Text = "CHR"
                                    Case Else
                                        Muestra.Text = "N/C"
                                End Select
                            Case 6
                                Muestra.Col = 1
                                Muestra.Text = "Rec"
                            Case 7
                                Muestra.Col = 1
                                Muestra.Text = "Ant"
                            Case 50
                                Muestra.Col = 1
                                Muestra.Text = "Doc"
                            Case Else
                        End Select
                        
                        Muestra.Col = 2
                        Muestra.Text = Pusing("######", Str$(!Numero))
                
                        Muestra.Col = 3
                        Muestra.Text = !Fecha

                        If Importe1 <> 0 Then
                            Muestra.Col = 4
                            Muestra.Text = Pusing("###,###,###.##", Str$(Importe1))
                                Else
                            Muestra.Col = 4
                            Muestra.Text = ""
                        End If
                
                        If Importe2 <> 0 Then
                            Muestra.Col = 5
                            Muestra.Text = Pusing("###,###,###.##", Str$(Importe2))
                                Else
                            Muestra.Col = 5
                            Muestra.Text = ""
                        End If
                
                        If Importe3 <> 0 Then
                            Muestra.Col = 6
                            Muestra.Text = Pusing("###,###,###.##", Str$(Importe3))
                                Else
                            Muestra.Col = 6
                            Muestra.Text = ""
                        End If
                        
                        WSaldo = WSaldo + Importe3
                
                        Muestra.Col = 7
                        Muestra.Text = !Vencimiento
                        
                        Muestra.Col = 8
                        Muestra.Text = !Vencimiento1
                    
                    End If
                    
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    
    rstCtacte.Close
    
    End If
    
    Saldo.Caption = Pusing("###,###,###.##", Str$(WSaldo))
    
    Muestra.Row = 1

End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        WCliente = Cliente.Text
        Cliente.Text = WCliente
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
                DesCliente.Caption = rstCliente!Razon
                rstCliente.Close
                Call Proceso_Click
                    Else
                Cliente.SetFocus
        End If
    End If
End Sub

Private Sub lee()
        Cliente.Text = UCase(Cliente.Text)
        WCliente = Cliente.Text
        Cliente.Text = WCliente
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
                DesCliente.Caption = rstCliente!Razon
                rstCliente.Close
                    Else
                Cliente.SetFocus
        End If
        Call Proceso_Click
End Sub

