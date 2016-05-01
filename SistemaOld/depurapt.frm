VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDepuraPt 
   AutoRedraw      =   -1  'True
   Caption         =   "       "
   ClientHeight    =   7320
   ClientLeft      =   150
   ClientTop       =   690
   ClientWidth     =   11550
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11550
   Begin VB.TextBox Cantidad 
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
      Left            =   6360
      MaxLength       =   11
      TabIndex        =   13
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Porce 
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
      Left            =   6360
      MaxLength       =   11
      TabIndex        =   12
      Text            =   " "
      Top             =   600
      Width           =   1095
   End
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
      Left            =   2640
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Cancelagraba 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   960
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
   Begin VB.CommandButton Graba 
      Caption         =   "&Graba"
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
      Left            =   9720
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Autorizo 
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
      Left            =   8640
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   6135
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10821
      _Version        =   393216
      Rows            =   4000
      Cols            =   9
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11400
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
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
      Left            =   7560
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin MSMask.MaskEdBox Desde 
      Height          =   300
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSMask.MaskEdBox Hasta 
      Height          =   300
      Left            =   2280
      TabIndex        =   10
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   8880
      TabIndex        =   16
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   300
      Left            =   7560
      TabIndex        =   17
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad de Resto"
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
      Left            =   3960
      TabIndex        =   15
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "% de Resto Sobre Total"
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
      Left            =   3960
      TabIndex        =   14
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta P.terminado"
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
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde P.Terminado"
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
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "PrgDepuraPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim XParam As String
Dim WTotal As Integer
Dim WGraba As String
Dim Vector(1000, 10) As String
Dim WTerminado As String
Dim WNumero As String
Dim WLote As String
Dim WSaldo As String
Dim ZSaldo As Double

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

Private Sub cmdClose_Click()
    With rstEmpresa
        .Close
    End With
    PrgDepuraPt.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    Muestra.Font.Bold = True
    
    Muestra.ColWidth(0) = 50
    Muestra.ColWidth(1) = 1400
    Muestra.ColWidth(2) = 2000
    Muestra.ColWidth(3) = 1200
    Muestra.ColWidth(4) = 1200
    Muestra.ColWidth(5) = 1200
    Muestra.ColWidth(6) = 1200
    Muestra.ColWidth(7) = 1200
    Muestra.ColWidth(8) = 1200
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Codigo"
    Muestra.ColAlignment(1) = flexAlignLeftCenter
    
    Muestra.Col = 2
    Muestra.Text = "Descripcion"
    Muestra.ColAlignment(2) = flexAlignLeftCenter
    
    Muestra.Col = 3
    Muestra.Text = "Numero"
    Muestra.ColAlignment(3) = flexAlignRightCenter
    
    Muestra.Col = 4
    Muestra.Text = "Fecha"
    Muestra.ColAlignment(4) = flexAlignRightCenter
    
    Muestra.Col = 5
    Muestra.Text = "Lote"
    Muestra.ColAlignment(5) = flexAlignRightCenter
    
    Muestra.Col = 6
    Muestra.Text = "Cantidad"
    Muestra.ColAlignment(6) = flexAlignRightCenter
    
    Muestra.Col = 7
    Muestra.Text = "Saldo"
    Muestra.ColAlignment(7) = flexAlignRightCenter
    
    Muestra.Col = 8
    Muestra.Text = "Estado"
    Muestra.ColAlignment(8) = flexAlignLeftCenter
   
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    Porce.Text = ""
    Cantidad.Text = ""
    
    Rem DesdeFecha.SetFocus
    
End Sub

Private Sub Graba_Click()

    If WGraba <> "S" Then
        Call Ingresa_clave
            Else
            
        WGraba = ""

        For Ciclo = 1 To WTotal
    
            Muestra.Row = Ciclo
            Muestra.Col = 8
            If Muestra.Text = "Ajuste" Then
        
                Muestra.Col = 1
                WTerminado = Muestra.Text
                Muestra.Col = 3
                WNumero = Muestra.Text
                Muestra.Col = 5
                WLote = Muestra.Text
                Muestra.Col = 7
                WSaldo = Muestra.Text
                
                Call Graba_Ajuste
                
            
            End If
    
        Next Ciclo
    
        Call cmdClose_Click
    End If

End Sub

Private Sub Proceso_Click()

    WSalida = "N"
        
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Codigo"
    Muestra.ColAlignment(1) = flexAlignLeftCenter
    
    Muestra.Col = 2
    Muestra.Text = "Descripcion"
    Muestra.ColAlignment(2) = flexAlignLeftCenter
    
    Muestra.Col = 3
    Muestra.Text = "Numero"
    Muestra.ColAlignment(3) = flexAlignRightCenter
    
    Muestra.Col = 4
    Muestra.Text = "Fecha"
    Muestra.ColAlignment(4) = flexAlignRightCenter
    
    Muestra.Col = 5
    Muestra.Text = "Lote"
    Muestra.ColAlignment(5) = flexAlignRightCenter
    
    Muestra.Col = 6
    Muestra.Text = "Cantidad"
    Muestra.ColAlignment(6) = flexAlignRightCenter
    
    Muestra.Col = 7
    Muestra.Text = "Saldo"
    Muestra.ColAlignment(7) = flexAlignRightCenter
    
    Muestra.Col = 8
    Muestra.Text = "Estado"
    Muestra.ColAlignment(8) = flexAlignLeftCenter
    
    Renglon = 0
    WSaldo = 0
    
    Uno = "Select * FROM Hoja Where Producto >= " + "'" + Desde.Text + "'"
    Dos = " and Producto <= " + "'" + Hasta.Text + "'"
    Tres = " and Saldo <> 0 and Renglon = 1"
    Cuatro = " ORDER BY Producto"
    spHoja = Uno + Dos + Tres + Cuatro
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    With rstHoja
    
        .MoveFirst
        If .NoMatch = False Then
            Do
                
                Entra = "N"
                
                WReal = IIf(IsNull(rstHoja!realant), "0", rstHoja!realant)
                If WReal = 0 Then
                    WReal = !Real
                End If
                
                If Val(Cantidad.Text) <> 0 And !Saldo <= Val(Cantidad.Text) Then
                    Entra = "S"
                End If
                
                If Val(Porce.Text) <> 0 Then
                    If WReal <> 0 Then
                        WPorce = (!Saldo / WReal) * 100
                            Else
                        WPorce = 0
                    End If
                    If WPorce <= Val(Porce.Text) Then
                        Entra = "S"
                    End If
                End If
                
                If Entra = "S" Then
                    
                    Renglon = Renglon + 1
            
                    Vector(Renglon, 1) = !Producto
                    Vector(Renglon, 3) = !Hoja
                    Vector(Renglon, 4) = !Fecha
                    Vector(Renglon, 5) = !Hoja
                    Vector(Renglon, 6) = WReal
                    Vector(Renglon, 7) = Str$(!Saldo)
                    Vector(Renglon, 8) = ""
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    End If
    
    
    Uno = "Select * FROM Guia Where Terminado >= " + "'" + Desde.Text + "'"
    Dos = " and Terminado <= " + "'" + Hasta.Text + "'"
    Tres = " and Saldo <> 0 "
    Cuatro = "ORDER BY Terminado"
    spMovguia = Uno + Dos + Tres + Cuatro
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    With rstMovguia
    
        .MoveFirst
        If .NoMatch = False Then
            Do
                
                Entra = "N"
                
                If Val(Cantidad.Text) <> 0 And !Saldo <= Val(Cantidad.Text) Then
                    Entra = "S"
                End If
                
                WCantidad = IIf(IsNull(rstMovguia!Cantidadant), "0", rstMovguia!Cantidadant)
                If WCantidad = 0 Then
                    WCantidad = !Cantidad
                End If
                
                If Val(Porce.Text) <> 0 Then
                    If WCantidad <> 0 Then
                        WPorce = (!Saldo / WCantidad) * 100
                            Else
                        WPorce = 0
                    End If
                    If WPorce <= Val(Porce.Text) Then
                        Entra = "S"
                    End If
                End If
                
                If Entra = "S" Then
                    
                    Renglon = Renglon + 1
                    
                    Vector(Renglon, 1) = !Terminado
                    Vector(Renglon, 3) = !Codigo
                    Vector(Renglon, 4) = !Fecha
                    Vector(Renglon, 5) = !Lote
                    Vector(Renglon, 6) = WCantidad
                    Vector(Renglon, 7) = Str$(!Saldo)
                    Vector(Renglon, 8) = ""
                    
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    End If
    
    
    For dada = 1 To Renglon
    
        WTerminado = Vector(dada, 1)
        spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            Vector(dada, 2) = rstTerminado!Descripcion
            rstTerminado.Close
        End If
        
    Next dada
    
    For Ciclo = 1 To Renglon

        For dada = Ciclo + 1 To Renglon

            If Vector(Ciclo, 1) > Vector(dada, 1) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8

            End If

        Next dada

    Next Ciclo
    
    
    For dada = 1 To Renglon
    
        WTerminado = Vector(dada, 1)
        spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            Vector(dada, 2) = rstTerminado!Descripcion
            rstTerminado.Close
        End If
        
    Next dada
    
    For Ciclo = 1 To Renglon

        For dada = Ciclo + 1 To Renglon

            If Vector(Ciclo, 1) > Vector(dada, 1) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8

            End If

        Next dada

    Next Ciclo
    
    For Ciclo = 1 To Renglon
        
        Muestra.Row = Ciclo
        
        Muestra.Col = 1
        Muestra.Text = Vector(Ciclo, 1)
        
        Muestra.Col = 2
        Muestra.Text = Vector(Ciclo, 2)
        
        Muestra.Col = 3
        Muestra.Text = Vector(Ciclo, 3)
        
        Muestra.Col = 4
        Muestra.Text = Vector(Ciclo, 4)
        
        Muestra.Col = 5
        Muestra.Text = Vector(Ciclo, 5)
        
        Muestra.Col = 6
        Muestra.Text = Vector(Ciclo, 6)
        
        ZSaldo = Val(Vector(Ciclo, 7))
        Call Redondeo(ZSaldo)
        Vector(Ciclo, 7) = Str$(ZSaldo)
        Muestra.Col = 7
        Muestra.Text = Vector(Ciclo, 7)
        Muestra.Text = Pusing("###,###.##", Muestra.Text)
        
        Muestra.Col = 8
        Muestra.Text = Vector(Ciclo, 8)
    
    Next Ciclo
    
    WTotal = Renglon
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    Muestra.TopRow = 1
    
    Muestra.SetFocus

End Sub

Private Sub Limpia_Vector()

    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Codigo"
    Muestra.ColAlignment(1) = flexAlignLeftCenter
    
    Muestra.Col = 2
    Muestra.Text = "Descripcion"
    Muestra.ColAlignment(2) = flexAlignLeftCenter
    
    Muestra.Col = 3
    Muestra.Text = "Numero"
    Muestra.ColAlignment(3) = flexAlignRightCenter
    
    Muestra.Col = 4
    Muestra.Text = "Fecha"
    Muestra.ColAlignment(4) = flexAlignRightCenter
    
    Muestra.Col = 5
    Muestra.Text = "Lote"
    Muestra.ColAlignment(5) = flexAlignRightCenter
    
    Muestra.Col = 6
    Muestra.Text = "Cantidad"
    Muestra.ColAlignment(6) = flexAlignRightCenter
    
    Muestra.Col = 7
    Muestra.Text = "Saldo"
    Muestra.ColAlignment(7) = flexAlignRightCenter
    
    Muestra.Col = 8
    Muestra.Text = "Estado"
    Muestra.ColAlignment(8) = flexAlignLeftCenter
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Cantidad.SetFocus
    End If
End Sub

Private Sub Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cantidad.Text = Pusing("###,###.##", Cantidad.Text)
        Porce.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Porce_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Porce.Text = Pusing("###,###.##", Porce.Text)
        Call Proceso_Click
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
        If WClave.Text = "SALDO" Then
            WGraba = "S"
            Clave1.Visible = False
            Call Graba_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Archivo de Materias Primas")
            WClave.SetFocus
        End If
    End If
End Sub

Private Sub Graba_Ajuste()

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
    Terminado = WTerminado
    Cantidad = WSaldo
    Movi = "S"
    Lote = WLote
                    
    Renglon = 1
    Auxi = Str$(Renglon)
    Call Ceros(Auxi, 2)
                        
    Auxi1 = Str$(WNroAjuste)
    Call Ceros(Auxi1, 6)
                
    WCodigo = Str$(WNroAjuste)
    WRenglon = Str$(Renglon)
    WFecha = Fecha.Text
    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
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
        
End Sub


