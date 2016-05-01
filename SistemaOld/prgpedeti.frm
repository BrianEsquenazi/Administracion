VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPedeti 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresion de Etiquetas de Exportacion"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin VB.Frame Impresion 
      Caption         =   "Impresion de Etiquetas"
      Height          =   2175
      Left            =   2760
      TabIndex        =   29
      Top             =   2400
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton WConfirma 
         Caption         =   "Confirma"
         Height          =   375
         Left            =   960
         TabIndex        =   35
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox WHasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2760
         TabIndex        =   34
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox WDesde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   33
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox WTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2520
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Imprimer desde"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Total Etiquetas"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox Marca3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8520
      MaxLength       =   20
      TabIndex        =   28
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Destino 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   27
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox Marca2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5160
      MaxLength       =   20
      TabIndex        =   25
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   22
      Text            =   " "
      Top             =   7080
      Visible         =   0   'False
      Width           =   5175
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
      Height          =   1740
      ItemData        =   "prgpedeti.frx":0000
      Left            =   3480
      List            =   "prgpedeti.frx":0007
      TabIndex        =   1
      Top             =   6480
      Visible         =   0   'False
      Width           =   8175
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wpedeti.rpt"
      Destination     =   1
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   500
      Left            =   10440
      TabIndex        =   21
      Top             =   0
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   3840
      TabIndex        =   20
      Top             =   6600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Marca1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   18
      Text            =   " "
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Cliente 
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   15
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   13
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Pedido 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   500
      Left            =   8520
      TabIndex        =   10
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   5640
      Width           =   11535
      Begin VB.TextBox WPalet 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10080
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox WPeso 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8760
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Text            =   " "
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox WCantidad 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   6
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label WCantienv 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7560
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label WEnvase 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6360
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label WDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton Impre 
      Caption         =   "Impresion"
      Height          =   500
      Left            =   9480
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3975
      Left            =   240
      OleObjectBlob   =   "prgpedeti.frx":0015
      TabIndex        =   3
      Top             =   1560
      Width           =   11535
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6960
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Destino"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "Marca"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de pedido"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgPedeti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 7 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Auxi As String
Private XLinea As Single
Private WDireccion As String
Private WLocalidad As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim Auxiliar(100) As String
Dim ImprePalet(1000, 10) As String

Private Sub BorraConsulta_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstAuxiliar
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    PrgPedeti.Hide
    Menu.Show
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Pedeti
End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 12 Then
        WLinea.Text = DBGrid1.Row + 1
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -     -   "
        WLinea.Text = ""
    End If
    
    DBGrid1.Col = 1
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 2
    If Val(DBGrid1.Text) <> 0 Then
        WCantidad.Text = Pusing("###,###.##", DBGrid1.Text)
            Else
        WCantidad.Text = ""
    End If
    
    DBGrid1.Col = 3
    WEnvase.Caption = Pusing("###,###.##", DBGrid1.Text)
    
    DBGrid1.Col = 4
    WCantienv.Caption = Pusing("###,###.##", DBGrid1.Text)
    
    DBGrid1.Col = 5
    If Val(DBGrid1.Text) <> 0 Then
        WPeso.Text = Pusing("###,###.##", DBGrid1.Text)
            Else
        WPeso.Text = ""
    End If
    
    DBGrid1.Col = 6
    WPalet.Text = DBGrid1.Text
    
    WPeso.SetFocus

End Sub

Private Sub Impre_Click()

    Erase ImprePalet
    LugarPalet = 0

    Da = 0
    With rstPedeti
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

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""

    Renglon = 0
    DBGrid1.Refresh
        
    With rstPedeti
        
        Renglon = 0
        .Index = "Clave"
                                        
        For a = 0 To 9
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Articulo = UCase(DBGrid1.Text)
                    
                DBGrid1.Col = 1
                Nombre = DBGrid1.Text
                    
                aa = Len(Nombre)
                    
                DBGrid1.Col = 2
                Cantidad = Val(DBGrid1.Text)
                    
                DBGrid1.Col = 3
                Envase = Val(DBGrid1.Text)
                    
                DBGrid1.Col = 4
                Cantienv = Val(DBGrid1.Text)
                    
                DBGrid1.Col = 5
                Peso = Val(DBGrid1.Text)
                
                DBGrid1.Col = 6
                Palet = DBGrid1.Text
                    
                spEnvase = "ConsultaEnvases " + "'" + Str$(Envase) + "'"
                Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvase.RecordCount > 0 Then
                    WPeso = rstEnvase!Kilos
                    rstEnvase.Close
                End If
                    
                If Cantidad <> 0 Then
                
                    If Val(Palet) = 0 Then
                        
                        For aa = 1 To Cantienv
                        
                            Renglon = Renglon + 1
                    
                            .AddNew
                        
                            !Clave = Renglon
                            !Impre1 = Marca1.Text
                            !Impre2 = Marca2.Text
                            !Impre3 = Destino.Text
                            !Impre4 = Left$(Nombre, 16)
                            !Impre5 = Peso + WPeso
                            !Impre6 = WPeso
                            !Impre7 = Renglon
                            !Impre8 = Marca3.Text
                            !Impre9 = Mid$(Nombre, 17, 16)
                            .Update
                        
                        Next aa
                        
                            Else
                            
                        ImprePalet(Val(Palet), 1) = Palet
                        ImprePalet(Val(Palet), 2) = Str$(Val(ImprePalet(Val(Palet), 2)) + (WPeso * Cantienv))
                        If Val(Peso) <> 0 Then
                            ImprePalet(Val(Palet), 3) = Str$(Peso)
                        End If
                        ImprePalet(Val(Palet), 4) = "PALET " + Str$(Palet)
                        
                    
                    End If
                        
                End If
                                        
            Next iRow
            
        Next a
            
    End With
    
    For Ciclo = 1 To 100
    
        If Val(ImprePalet(Ciclo, 1)) <> 0 Then
            Renglon = Renglon + 1
        End If
        
    Next Ciclo

    Impresion.Visible = True
    
    WTotal.Text = Str$(Renglon)
    WDesde.Text = "1"
    WHasta.Text = Str$(Renglon)
    
    WDesde.SetFocus
    
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
            WPeso.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WConfirma_Click()

    Impresion.Visible = False
    Call Graba_Click
    
End Sub

Private Sub WPeso_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPeso.Text = Pusing("###,###.##", WPeso.Text)
        WPalet.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WPalet_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Alta_Vector
        Call Ingresa_Click
        WPeso.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Graba_Click()

    Erase ImprePalet
    LugarPalet = 0

    Da = 0
    With rstPedeti
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

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""

        Renglon = 0
        DBGrid1.Refresh
        
        With rstPedeti
        
            Renglon = 0
            .Index = "Clave"
                                        
            For a = 0 To 9
        
                Suma = a * 10
                DBGrid1.FirstRow = Suma
            
                For iRow = 0 To 9
                
                    WRow = iRow
                    DBGrid1.Row = WRow
                    
                    DBGrid1.Col = 0
                    Articulo = UCase(DBGrid1.Text)
                    
                    DBGrid1.Col = 1
                    Nombre = DBGrid1.Text
                    
                    aa = Len(Nombre)
                    
                    DBGrid1.Col = 2
                    Cantidad = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 3
                    Envase = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 4
                    Cantienv = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 5
                    Peso = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 6
                    Palet = DBGrid1.Text
                    
                    spEnvase = "ConsultaEnvases " + "'" + Str$(Envase) + "'"
                    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvase.RecordCount > 0 Then
                        WPeso = rstEnvase!Kilos
                        rstEnvase.Close
                    End If
                    
                    If Cantidad <> 0 Then
                        
                        If Val(Palet) = 0 Then
                        
                            For aa = 1 To Cantienv
                        
                                Renglon = Renglon + 1
                        
                                If Renglon >= Val(WDesde.Text) And Renglon <= Val(WHasta.Text) Then
                    
                                    .AddNew
                        
                                    !Clave = Renglon
                                    !Impre1 = Marca1.Text
                                    !Impre2 = Marca2.Text
                                    !Impre3 = Destino.Text
                                    !Impre4 = Left$(Nombre, 16)
                                    !Impre5 = Peso + WPeso
                                    !Impre6 = WPeso
                                    !Impre7 = Renglon
                                    !Impre8 = Marca3.Text
                                    !Impre9 = Mid$(Nombre, 17, 16)
                                    .Update
                            
                                End If
                        
                            Next aa
                            
                                Else
                            
                            ImprePalet(Val(Palet), 1) = Palet
                            Da = ImprePalet(Val(Palet), 2)
                            ImprePalet(Val(Palet), 2) = Str$(Val(ImprePalet(Val(Palet), 2)) + (WPeso * Cantienv))
                            Da = ImprePalet(Val(Palet), 2)
                            If Val(Peso) <> 0 Then
                                ImprePalet(Val(Palet), 3) = Str$(Peso)
                            End If
                            ImprePalet(Val(Palet), 4) = "PALLET Nro.:" + Str$(Palet)
                            
                        End If
                        
                    End If
                                        
                Next iRow
            
            Next a
            
        End With
        
        For Ciclo = 1 To 100
    
            If Val(ImprePalet(Ciclo, 1)) <> 0 Then
            
                Nombre = ImprePalet(Ciclo, 4)
                Peso = Val(ImprePalet(Ciclo, 2))
                WPeso = Val(ImprePalet(Ciclo, 3))
            
                With rstPedeti
        
                    .Index = "Clave"
                                        
                    Renglon = Renglon + 1
                        
                    If Renglon >= Val(WDesde.Text) And Renglon <= Val(WHasta.Text) Then
                    
                        .AddNew
                        !Clave = Renglon
                        !Impre1 = Marca1.Text
                        !Impre2 = Marca2.Text
                        !Impre3 = Destino.Text
                        !Impre4 = Left$(Nombre, 16)
                        !Impre5 = Peso + WPeso
                        !Impre6 = WPeso
                        !Impre7 = Renglon
                        !Impre8 = Marca3.Text
                        !Impre9 = Mid$(Nombre, 17, 16)
                        .Update
                    
                    End If
                    
                End With
                
            End If
        
        Next Ciclo
   Rem by nan
        Listado.Destination = 1
        Rem Listado.Destination = 0
        Listado.Action = 1
        
        Rem Call Impresion
        
        Call Limpia_Click

        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
    
        Pedido.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WEnvase.Caption = ""
    WCantienv.Caption = ""
    WPeso.Text = ""
    WPalet.Text = ""
    
    WPeso.SetFocus
    
End Sub

Private Sub Limpia_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WEnvase.Caption = ""
    WCantienv.Caption = ""
    WPeso.Text = ""
    WPalet.Text = ""
    
    Marca1.Text = ""
    Marca2.Text = ""
    Marca3.Text = ""
    Destino.Text = ""
    
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Rem     Observaciones.Text = ""
    
    For a = 0 To 9
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 6
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    Pedido.Text = ""
    
    DBGrid1.FirstRow = 0
    Renglon = 0

    Pedido.SetFocus

End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 40 Then
                            DBGrid1.Row = DBGrid1.Row + 1
                            DBGrid1.Col = 0
                            KeyCode = 0
                        End If
                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
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
        UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
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

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 6, 0 To 100)

mTotalRows& = 100

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DBGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DBGrid1.Columns.Count - 1 To 0 Step -1
      DBGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 6
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
             
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Envase"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
             
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Cant.Env."
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
             
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Tara"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
             
         Case 6
             DBGrid1.Columns(newcnt).Caption = "Nro. Palet"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgPedeti.Caption = "Impresion de Etiquetas de Exportacion :  " + !Nombre
        End If
    End With
 
 
    Pedido.Text = ""
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Pedido.SetFocus
    
End Sub

Private Sub Proceso_Click()

    On Error GoTo WError
    
    Marca1.Text = ""
    Marca2.Text = ""
    Marca3.Text = ""
    Destino.Text = ""
    
    For a = 0 To 9
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 6
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    
    spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    If rstPedido.RecordCount > 0 Then
            With rstPedido
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Marca1.Text = rstPedido!Marca1
                        Marca2.Text = rstPedido!Marca2
                        Marca3.Text = rstPedido!Marca3
                        Destino.Text = rstPedido!Destino
        
                        If rstPedido!Envase1 <> 0 And rstPedido!Canti1 <> 0 Then
        
                            Renglon = Renglon + 1
            
                            Lugar1 = Int((Renglon - 1) / 10) * 10
                            Lugar2 = Renglon - Lugar1
                
                            DBGrid1.FirstRow = Lugar1
                            DBGrid1.Row = Lugar2 - 1
                
                            DBGrid1.Col = 0
                            DBGrid1.Text = rstPedido!Terminado
                            Auxi1 = rstPedido!Terminado
                
                            DBGrid1.Col = 2
                            DBGrid1.Text = Pusing("###,###.##", rstPedido!Cantidad - rstPedido!Facturado)
                   
                            DBGrid1.Col = 3
                            DBGrid1.Text = Pusing("###,###", rstPedido!Envase1)
                    
                            DBGrid1.Col = 4
                            DBGrid1.Text = Pusing("###,###", rstPedido!Canti1)
                            
                            Auxiliar(Renglon) = Auxi1
                
                        End If
                
                        If rstPedido!Envase2 <> 0 And rstPedido!Canti2 <> 0 Then
        
                            Renglon = Renglon + 1
                
                            Lugar1 = Int((Renglon - 1) / 10) * 10
                            Lugar2 = Renglon - Lugar1
                
                            DBGrid1.FirstRow = Lugar1
                            DBGrid1.Row = Lugar2 - 1
                
                            DBGrid1.Col = 0
                            DBGrid1.Text = rstPedido!Terminado
                            Auxi1 = rstPedido!Terminado
                
                            DBGrid1.Col = 2
                            DBGrid1.Text = Pusing("###,###.##", rstPedido!Cantidad - rstPedido!Facturado)
                    
                            DBGrid1.Col = 3
                            DBGrid1.Text = Pusing("###,###", rstPedido!Envase2)
                    
                            DBGrid1.Col = 4
                            DBGrid1.Text = Pusing("###,###", rstPedido!Canti2)
                            
                            Auxiliar(Renglon) = Auxi1
                    
                        End If
                
                        If rstPedido!Envase3 <> 0 And rstPedido!Canti3 <> 0 Then
        
                            Renglon = Renglon + 1
                    
                            Lugar1 = Int((Renglon - 1) / 10) * 10
                            Lugar2 = Renglon - Lugar1
                        
                            DBGrid1.FirstRow = Lugar1
                            DBGrid1.Row = Lugar2 - 1
                    
                            DBGrid1.Col = 0
                            DBGrid1.Text = rstPedido!Terminado
                            Auxi1 = rstPedido!Terminado
                
                            DBGrid1.Col = 2
                            DBGrid1.Text = Pusing("###,###.##", rstPedido!Cantidad - rstPedido!Facturado)
                    
                            DBGrid1.Col = 3
                            DBGrid1.Text = Pusing("###,###", rstPedido!Envase3)
                    
                            DBGrid1.Col = 4
                            DBGrid1.Text = Pusing("###,###", rstPedido!Canti3)
                            
                            Auxiliar(Renglon) = Auxi1

                        End If
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedido.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Renglon = 1 To WRenglon
    
        Auxi1 = Auxiliar(Renglon)
        
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                        
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
        
        spPrecios = "ConsultaPrecios " + "'" + Cliente.Text + Auxi1 + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            DBGrid1.Col = 1
            DBGrid1.Text = rstPrecios!Descripcion
            rstPrecios.Close
        End If
        
    Next Renglon
    
    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    DBGrid1.FirstRow = 0
    
    WPeso.SetFocus
    
    Exit Sub

WError:

    Resume Next

End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) <> 0 Then
                
            DBGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
            
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
            
            DBGrid1.Col = 3
            DBGrid1.Text = Pusing("###,###.##", WEnvase.Caption)
            
            DBGrid1.Col = 4
            DBGrid1.Text = Pusing("###,###.##", WCantienv.Caption)
            
            DBGrid1.Col = 5
            DBGrid1.Text = Pusing("###,###.##", WPeso.Text)
            
            DBGrid1.Col = 6
            DBGrid1.Text = WPalet.Text
                
            Rem DbGrid1.Row = Renglon
            DBGrid1.Row = Lugar2 - 1
            DBGrid1.Col = 0
            
            DBGrid1.Row = DBGrid1.Row + 1
            
    End If

End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            Fecha.Text = rstPedido!Fecha
            Cliente.Text = rstPedido!Cliente
            rstPedido.Close
            
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                rstCliente.Close
            End If
            Call Proceso_Click
                Else
            WPedido = Pedido.Text
            Call Limpia_Click
            Pedido.Text = WPedido
            Fecha.SetFocus
        End If
    End If
End Sub
