VERSION 5.00
Begin VB.Form PrgPedConsulta 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verirficacion de Costo "
   ClientHeight    =   4785
   ClientLeft      =   135
   ClientTop       =   1755
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   11850
   Visible         =   0   'False
   Begin VB.Frame MuestraCosto 
      Height          =   2535
      Left            =   6240
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox FechaCotiza 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2280
         TabIndex        =   36
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton CerrarPanta 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   1320
         TabIndex        =   34
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox CostoReposicion 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2280
         TabIndex        =   32
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox CostoStd 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2280
         TabIndex        =   30
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox CostoUltCpa 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2280
         TabIndex        =   28
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Cot."
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reposicion"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Costo Std."
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Costo Ult. Cpa"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox FactorPT 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3840
      TabIndex        =   25
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox CostoPT 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3840
      TabIndex        =   23
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox FechaPrecio 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3840
      TabIndex        =   22
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Termi 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3840
      TabIndex        =   13
      Top             =   600
      Width           =   1695
   End
   Begin VB.Frame Datos 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2535
      Left            =   8880
      TabIndex        =   1
      Top             =   240
      Width           =   2775
      Begin VB.CommandButton AvisoError 
         Caption         =   "Sistema sin Conexion"
         Height          =   1215
         Left            =   600
         Picture         =   "Prgpedconsulta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label WStock5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1200
         TabIndex        =   19
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Stock5 
         Caption         =   "Stock"
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
         TabIndex        =   18
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label WStock4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1200
         TabIndex        =   17
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Stock4 
         Caption         =   "Stock"
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
         TabIndex        =   16
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label WStock3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1200
         TabIndex        =   15
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Stock3 
         Caption         =   "Stock"
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
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Stock2 
         Caption         =   "Stock"
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
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Stock1 
         Caption         =   "Stock"
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
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label WStock2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1200
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label WStock1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1200
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Disponible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label StkPedido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Stock 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Disponible"
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
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Pedido"
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
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Stock"
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
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Factor"
      Height          =   375
      Left            =   3840
      TabIndex        =   26
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Costo"
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Precio"
      Height          =   375
      Left            =   3840
      TabIndex        =   21
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRODUCTO"
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "PrgPedConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Auxi As String
Private WImpre(10) As String
Private WEnvase(10) As String
Private WVector(6, 3) As String
Private XEnvase(40, 6) As String
Private XLinea As Single
Private WDirentrega As String
Private WInicio As Integer
Private Auxiliar(100, 3) As String
Private XSaldo As Double
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim XParam As String
Dim ClavePedido(100)
Dim Producto As String
Dim Costo As Double
Dim ZTipoCosto As Integer

Private Sub CerrarPanta_Click()
    MuestraCosto.Visible = False
End Sub

Private Sub cmdClose_Click()

    With rstEmpresa
        .Close
    End With
    PrgPedConsulta.Hide
    Unload Me
    PrgPedido.Show
    
End Sub

Private Sub CostoPT_dblclick()

    If Left$(Termi.Text, 2) = "PT" Then
    
        ZTipoCosto = 2
        Producto = Termi.Text
        Call Calcula_Costo(Producto, Costo)
        CostoUltCpa.Text = Str$(Costo)
        CostoUltCpa.Text = Pusing("###,###.##", CostoUltCpa.Text)
    
        ZTipoCosto = 1
        Producto = Termi.Text
        Call Calcula_Costo(Producto, Costo)
        CostoStd.Text = Str$(Costo)
        CostoStd.Text = Pusing("###,###.##", CostoStd.Text)
    
        ZTipoCosto = 3
        Producto = Termi.Text
        Call Calcula_Costo(Producto, Costo)
        CostoReposicion.Text = Str$(Costo)
        CostoReposicion.Text = Pusing("###,###.##", CostoReposicion.Text)
        
        FechaCotiza.Text = ""
        
        MuestraCosto.Visible = True
    
            Else

        ZZArti = Left$(Termi.Text, 3) + Right$(Termi.Text, 7)
        spArticulo = "ConsultaArticulo " + "'" + ZZArti + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            CostoUltCpa.Text = Str$(rstArticulo!Costo1)
            CostoUltCpa.Text = Pusing("###,###.##", CostoUltCpa.Text)
            CostoStd.Text = Str$(rstArticulo!Costo2)
            CostoStd.Text = Pusing("###,###.##", CostoStd.Text)
            ZCosto4 = IIf(IsNull(rstArticulo!Costo4), "0", rstArticulo!Costo4)
            CostoReposicion.Text = Str$(ZCosto4)
            CostoReposicion.Text = Pusing("###,###.##", CostoReposicion.Text)
            MuestraCosto.Visible = True
            rstArticulo.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cotiza"
        ZSql = ZSql + " Where Cotiza.Articulo = " + "'" + ZZArti + "'"
        ZSql = ZSql + " Order by Cotiza"
        spCotiza = ZSql
        Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
        If rstCotiza.RecordCount > 0 Then
            With rstCotiza
                .MoveLast
                FechaCotiza.Text = rstCotiza!Fecha
            End With
            rstCotiza.Close
        End If
        
    End If

End Sub

Private Sub Form_Load()

    Muestra1.ColWidth(0) = 150
    Muestra1.ColWidth(1) = 600
    Muestra1.ColWidth(2) = 900
    Muestra1.ColWidth(3) = 750
    
    Muestra1.Row = 0
    
    Muestra1.Col = 1
    Muestra1.Text = "Tipo"
    
    Muestra1.Col = 2
    Muestra1.Text = "Partida"
    
    Muestra1.Col = 3
    Muestra1.Text = "Stock"
    
    Muestra2.ColWidth(0) = 100
    Muestra2.ColWidth(1) = 900
    Muestra2.ColWidth(2) = 800
    Muestra2.ColWidth(3) = 1300
    
    Muestra2.Row = 0
    
    Muestra2.Col = 1
    Muestra2.Text = "Cliente"
    
    Muestra2.Col = 3
    Muestra2.Text = "Fecha"
    
    Muestra2.Col = 2
    Muestra2.Text = "Canti."

    Pedido.Text = WXPed
    
    Call Muestra
    
End Sub

Private Sub Muestra()

    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 10 Then
        Stock1.Caption = "Stock SI"
        Stock2.Caption = "Stock SII"
        Stock3.Caption = "Stock SIII"
        Stock4.Caption = "Stock SIV"
        Stock5.Caption = "Stock SV"
            Else
        Stock1.Caption = "Stock PI"
        Stock2.Caption = "Stock PII"
        Stock3.Caption = "Stock PV"
        Stock4.Caption = "Stock PVI"
        Stock5.Caption = ""
    End If

    WStock1.Caption = ""
    WStock2.Caption = ""
    WStock3.Caption = ""
    WStock4.Caption = ""
    WStock5.Caption = ""

    Muestra.Col = 1
    Termi.Text = Muestra.Text
    XProducto = Termi.Text
    
    Muestra1.Clear
    Muestra1.Row = 0
    
    Muestra1.Col = 1
    Muestra1.Text = "Tipo"
    
    Muestra1.Col = 2
    Muestra1.Text = "Partida"
    
    Muestra1.Col = 3
    Muestra1.Text = "Stock"
    
    Muestra2.Clear
    Muestra2.Row = 0
    
    Muestra2.Col = 1
    Muestra2.Text = "Cliente"
    
    Muestra2.Col = 3
    Muestra2.Text = "Fecha"
    
    Muestra2.Col = 2
    Muestra2.Text = "Canti."

    Renglon = 0
    XStock = 0
    XPedido = 0
    
    If Left$(XProducto, 2) <> "PT" Then
        WTipopro = "M"
            Else
        WTipopro = "T"
    End If
        
    Select Case WTipopro
        Case "M"
            WArti = Left$(XProducto, 3) + Right$(XProducto, 7)
            
            XParam = "'" + WArti + "','" _
                 + WArti + "'"
            spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
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
                    
                            If rstLaudo!Articulo = WArti Then
                            
                                XSaldo = rstLaudo!Saldo
                                Call Redondeo(XSaldo)
                                
                                If XSaldo <> 0 Then
                            
                                    Renglon = Renglon + 1
                                    Muestra1.Row = Renglon
                            
                                    Muestra1.Col = 1
                                    Muestra1.Text = Left$(WArti, 2)
                        
                                    Muestra1.Col = 2
                                    Muestra1.Text = rstLaudo!Laudo
                            
                                    Muestra1.Col = 3
                                    Muestra1.Text = Pusing("###,###", Str$(XSaldo))
                        
                                    XStock = XStock + XSaldo
                                    
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
                rstLaudo.Close
            End If
            
            
            Rem PROCESA LAS GUIAS DE TRASLADO INTERNOS
    
            XParam = "'" + WArti + "','" _
                        + WArti + "'"
            spMovguia = "ListaMovguiaArticuloDesdeHasta" + XParam
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then

                With rstMovguia
    
                    .MoveFirst
            
                    If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                                Else
                        
                            If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArti Then
                    
                                WArticulo = rstMovguia!Articulo
                                WCantidad = rstMovguia!Cantidad
                                WFecha = rstMovguia!Fecha
                                WCodigo = rstMovguia!Codigo
                                WMovi = rstMovguia!Movi
                                WDestino = rstMovguia!Destino
                                WTipomov = rstMovguia!Tipomov
                                WSaldo = rstMovguia!Saldo
                                
                                Renglon = Renglon + 1
                                Muestra1.Row = Renglon
                            
                                Muestra1.Col = 1
                                Muestra1.Text = Left$(WArti, 2)
                        
                                WPartiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                                If Trim(WPartiOri) <> "" Then
                                    WParti = WPartiOri
                                        Else
                                    WParti = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                                End If
                                
                                Muestra1.Col = 2
                                Muestra1.Text = WParti
                            
                                Muestra1.Col = 3
                                Muestra1.Text = Pusing("###,###", Str$(rstMovguia!Saldo))
                        
                                XStock = XStock + rstMovguia!Saldo
                                
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
            
            
            
            
            Renglon = 0
    
            spPedido = "ListaPedidoTerminado " + "'" + Termi.Text + "'"
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                With rstPedido
                    .MoveFirst
                    If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        XPed = rstPedido!Cantidad - rstPedido!Facturado
                        If XPed <> 0 Then
                        If Pedido.Text <> rstPedido!Pedido Then
                            Renglon = Renglon + 1
                            Muestra2.Row = Renglon
                    
                            Muestra2.Col = 1
                            Muestra2.Text = rstPedido!Cliente
                
                            Muestra2.Col = 3
                            Muestra2.Text = rstPedido!FecEntrega
                            
                            Muestra2.Col = 2
                            Muestra2.Text = Pusing("###,###", Str$(XPed))
                        
                            XPedido = XPedido + XPed
                        End If
                        End If
                        
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                    End If
                End With
            End If
            
            
            
            
            Cliente.Text = UCase(Cliente.Text)
            Termi.Text = UCase(Termi.Text)
            ZZArti = Left$(Termi.Text, 3) + Right$(Termi.Text, 7)
    
            WClave = Cliente.Text + ZZArti
    
            spPreciosMp = "ConsultaPreciosMp " + "'" + WClave + "'"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
                FechaPrecio.Text = IIf(IsNull(rstPreciosMp!Fecha), "", rstPreciosMp!Fecha)
                rstPreciosMp.Close
            End If
            
            CostoPT.Text = ""
            FactorPT.Text = ""
            
            spArticulo = "ConsultaArticulo " + "'" + ZZArti + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                CostoPT.Text = Str$(rstArticulo!Costo2)
                CostoPT.Text = Pusing("###,###.##", CostoPT.Text)
                If Val(CostoPT.Text) <> 0 Then
                    ZZPrecioVenta = Muestra.TextMatrix(Muestra.Row, 5)
                    FactorPT.Text = Str$(Val(ZZPrecioVenta) / Val(CostoPT.Text))
                    FactorPT.Text = Pusing("###.##", FactorPT.Text)
                End If
                rstArticulo.Close
            End If
            
            If Left$(XProducto, 2) = "DW" And Val(CostoPT.Text) = 0 Then
            
                ZTipoCosto = 1
                Producto = XProducto
                Call Calcula_Costo(Producto, Costo)
                CostoPT.Text = Str$(Costo)
                CostoPT.Text = Pusing("###,###.##", CostoPT.Text)
                
                If Val(CostoPT.Text) <> 0 Then
                    ZZPrecioVenta = Muestra.TextMatrix(Muestra.Row, 5)
                    FactorPT.Text = Str$(Val(ZZPrecioVenta) / Val(CostoPT.Text))
                    FactorPT.Text = Pusing("###.##", FactorPT.Text)
                End If
                
            End If
            
            WStock1.Caption = Pusing("###,###.##", Str$(XStock))
            WStock2.Caption = Pusing("###,###.##", WStock2.Caption)
            WStock3.Caption = Pusing("###,###.##", WStock3.Caption)
            WStock4.Caption = Pusing("###,###.##", WStock4.Caption)
            WStock5.Caption = Pusing("###,###.##", WStock5.Caption)
            
            Stock.Caption = Pusing("###,###.##", Str$(Val(WStock1.Caption) + Val(WStock2.Caption) + Val(WStock3.Caption) + Val(WStock4.Caption) + Val(WStock5.Caption)))
            StkPedido.Caption = Pusing("###,###.##", Str$(XPedido))
            Disponible.Caption = Pusing("###,###.##", Str$(Val(Stock.Caption) - XPedido))
            
    
        Case Else
            Rem lee pt
    
            XParam = "'" + XProducto + "','" _
                         + XProducto + "'"
            spHoja = "ListaHojaProductoDesdeHasta" + XParam
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                With rstHoja
                    .MoveFirst
                    If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
            
                        If rstHoja!Marca = "X" And rstHoja!Saldo = 0 Then
                            Else
                        If Val(rstHoja!Renglon) = 1 Then
                 
                            XHoja = rstHoja!Hoja
                            XSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                            Call Redondeo(XSaldo)
                        
                            If XSaldo <> 0 Then
                        
                                Renglon = Renglon + 1
                                Muestra1.Row = Renglon
                            
                                Muestra1.Col = 1
                                Muestra1.Text = "PT"
                        
                                Muestra1.Col = 2
                                Muestra1.Text = XHoja
                            
                                Muestra1.Col = 3
                                Muestra1.Text = Pusing("###,###", Str$(XSaldo))
                        
                                XStock = XStock + XSaldo
                            
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
            End If
    
            XParam = "'" + XProducto + "','" _
                         + XProducto + "'"
            spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                With rstMovguia
                    .MoveFirst
                    If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        If rstMovguia!Marca = "X" Then
                                Else
                        If rstMovguia!Tipo = "T" Then
                
                            XLote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                            XSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                            Call Redondeo(XSaldo)
                    
                            If XSaldo <> 0 Then
                                Renglon = Renglon + 1
                                Muestra1.Row = Renglon
                        
                                Muestra1.Col = 1
                                Muestra1.Text = "PT"
                
                                Muestra1.Col = 2
                                Muestra1.Text = XLote
                        
                                Muestra1.Col = 3
                                Muestra1.Text = Pusing("###,###", Str$(XSaldo))
                        
                                XStock = XStock + XSaldo
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
            End If
    
    
            Renglon = 0
    
            spPedido = "ListaPedidoTerminado " + "'" + Termi.Text + "'"
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                With rstPedido
                    .MoveFirst
                    If .NoMatch = False Then
            
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        XPed = rstPedido!Cantidad - rstPedido!Facturado
                        If XPed <> 0 Then
                        If Pedido.Text <> rstPedido!Pedido Then
                            Renglon = Renglon + 1
                            Muestra2.Row = Renglon
                    
                            Muestra2.Col = 1
                            Muestra2.Text = rstPedido!Cliente
                
                            Muestra2.Col = 3
                            Muestra2.Text = rstPedido!FecEntrega
                            
                            Muestra2.Col = 2
                            Muestra2.Text = Pusing("###,###", Str$(XPed))
                        
                            XPedido = XPedido + XPed
                        End If
                        End If
                
                        .MoveNext
                
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                    End If
                End With
            End If
            
            WSalidaError = ""
            On Error GoTo Control_error
            
            If Val(WEmpresa) = 1 Then
            
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock1.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                         WStock2.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
            
                    WEmpresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock3.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    WEmpresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock4.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    WEmpresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock5.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                        Else
                        
                    WEmpresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock1.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                         WStock2.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
            
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock3.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    
                    WEmpresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock4.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
            End If
            
            On Error GoTo 0
            
            
            
            Cliente.Text = UCase(Cliente.Text)
            Termi.Text = UCase(Termi.Text)
            WClave = Cliente.Text + Termi.Text
    
            spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                FechaPrecio.Text = IIf(IsNull(rstPrecios!Fecha), "", rstPrecios!Fecha)
                rstPrecios.Close
            End If
            
            
            If Left$(Termi.Text, 2) = "PT" Then
            
                ZTipoCosto = 1
                Producto = Termi.Text
                Call Calcula_Costo(Producto, Costo)
                CostoPT.Text = Str$(Costo)
                CostoPT.Text = Pusing("###,###.##", CostoPT.Text)
                
                If Val(CostoPT.Text) <> 0 Then
                    ZZPrecioVenta = Muestra.TextMatrix(Muestra.Row, 5)
                    FactorPT.Text = Str$(Val(ZZPrecioVenta) / Val(CostoPT.Text))
                    FactorPT.Text = Pusing("###.##", FactorPT.Text)
                End If
                
            End If
            
            
            WStock1.Caption = Pusing("###,###.##", WStock1.Caption)
            WStock2.Caption = Pusing("###,###.##", WStock2.Caption)
            WStock3.Caption = Pusing("###,###.##", WStock3.Caption)
            WStock4.Caption = Pusing("###,###.##", WStock4.Caption)
            WStock5.Caption = Pusing("###,###.##", WStock5.Caption)
            
            Stock.Caption = Pusing("###,###.##", Str$(Val(WStock1.Caption) + Val(WStock2.Caption) + Val(WStock3.Caption) + Val(WStock4.Caption) + Val(WStock5.Caption)))
            StkPedido.Caption = Pusing("###,###.##", Str$(XPedido))
            Disponible.Caption = Pusing("###,###.##", Str$(Val(Stock.Caption) - XPedido))
        
    End Select
    
    Exit Sub
    
Control_error:
    Rem MsgBox Err.Description
    Beep
    WSalidaError = "N"
    AvisoError.Visible = True
    Stock1.Visible = False
    Stock2.Visible = False
    Stock3.Visible = False
    Stock4.Visible = False
    Stock5.Visible = False
    WStock1.Visible = False
    WStock2.Visible = False
    WStock3.Visible = False
    WStock4.Visible = False
    WStock5.Visible = False
    Label4.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Disponible.Visible = False
    Stock.Visible = False
    StkPedido.Visible = False
    Resume Next
    
End Sub
    
Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim ZZVector(100, 2) As String
    Dim ZZAuxiliar(100, 3) As String
    
    Erase ZZAuxiliar
    ZZRenglon = 0
    
    ZZVector(1, 1) = Producto
    ZZVector(1, 2) = "1"
    ZZLugar = 1
    ZZCicla = 0
    
    Costo = 0
    
    Do
        ZZCicla = ZZCicla + 1
        If ZZVector(ZZCicla, 1) <> "" Then
    
            ZZEntra = "S"
            
            spComposicion = "ConsultaComposicionProducto " + "'" + ZZVector(ZZCicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            If rstComposicion.RecordCount > 0 Then
                With rstComposicion
                    .MoveFirst
                    Do
                        If .EOF = False Then
                    
                            ZZEntra = "N"
                        
                            ZZTipo = rstComposicion!Tipo
                            ZZArticulo1 = rstComposicion!Articulo1
                            ZZArticulo2 = rstComposicion!Articulo2
                            ZZCantidad = rstComposicion!Cantidad
                            
                            If Left$(ZZArticulo1, 2) = "DW" Then
                                ZZTipo = "T"
                                ZZArticulo2 = Left$(ZZArticulo1, 3) + "00" + Right$(ZZArticulo1, 7)
                            End If
                            
                            Select Case ZZTipo
                                Case "T"
                                    If Producto <> ZZArticulo2 Then
                                        ZZLugar = ZZLugar + 1
                                        ZZVector(ZZLugar, 1) = ZZArticulo2
                                        ZZVector(ZZLugar, 2) = Str$(ZZCantidad * Val(ZZVector(ZZCicla, 2)))
                                    End If
                                Case "M"
                                    ZZRenglon = ZZRenglon + 1
                                    ZZAuxiliar(ZZRenglon, 1) = ZZArticulo1
                                    ZZAuxiliar(ZZRenglon, 2) = ZZCantidad
                                    ZZAuxiliar(ZZRenglon, 3) = ZZVector(ZZCicla, 2)
                                Case Else
                            End Select
                            
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstComposicion.Close
            End If
            
            If ZZEntra = "S" And Left$(ZZVector(ZZCicla, 1), 2) = "DW" Then
                ZZRenglon = ZZRenglon + 1
                ZZAuxiliar(ZZRenglon, 1) = Left$(ZZVector(ZZCicla, 1), 3) + Right$(ZZVector(ZZCicla, 1), 7)
                ZZAuxiliar(ZZRenglon, 2) = 1
                ZZAuxiliar(ZZRenglon, 3) = ZZVector(ZZCicla, 2)
            End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For DA = 1 To ZZRenglon
        ZZArticulo = ZZAuxiliar(DA, 1)
        ZZCantidad = ZZAuxiliar(DA, 2)
        ZZCantidadII = ZZAuxiliar(DA, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Select Case ZTipoCosto
                Case 1
                    WCosto = (ZZCantidad * rstArticulo!Costo2 * Val(ZZCantidadII))
                Case 2
                    WCosto = (ZZCantidad * rstArticulo!Costo1 * Val(ZZCantidadII))
                Case 3
                    Costo4 = IIf(IsNull(rstArticulo!Costo4), "0", rstArticulo!Costo4)
                    If Costo4 = 0 Then
                        Costo4 = IIf(IsNull(rstArticulo!Costo2), "0", rstArticulo!Costo2)
                    End If
                    WCosto = (ZZCantidad * Costo4 * Val(ZZCantidadII))
                Case Else
                    WCosto = (ZZCantidad * rstArticulo!Costo2 * Val(ZZCantidadII))
            End Select
            Costo = Costo + WCosto
            rstArticulo.Close
        End If
    Next DA
    
    
End Sub


