VERSION 5.00
Begin VB.Form PrgHojaDesvio 
   Caption         =   "Detalle de Partidas con desvio a Utilizar"
   ClientHeight    =   4890
   ClientLeft      =   3585
   ClientTop       =   2115
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   7410
   Begin VB.TextBox Usa5 
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
      Left            =   4680
      TabIndex        =   26
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Usa4 
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
      Left            =   4680
      TabIndex        =   25
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox Usa3 
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
      Left            =   4680
      TabIndex        =   24
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Usa2 
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
      Left            =   4680
      TabIndex        =   23
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Usa1 
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
      Left            =   4680
      TabIndex        =   22
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Cantidad1 
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
      Left            =   2880
      TabIndex        =   20
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Cantidad2 
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
      Left            =   2880
      TabIndex        =   19
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Cantidad3 
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
      Left            =   2880
      TabIndex        =   18
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Cantidad4 
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
      Left            =   2880
      TabIndex        =   17
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox Cantidad5 
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
      Left            =   2880
      TabIndex        =   16
      Top             =   3480
      Width           =   1575
   End
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
      Left            =   1800
      TabIndex        =   15
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox MateriaPrima 
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
      Left            =   1800
      TabIndex        =   14
      Top             =   200
      Width           =   1575
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "ACEPTAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   11
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CheckBox Confirma5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   9
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox Partida5 
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
      Left            =   1080
      TabIndex        =   8
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CheckBox Confirma4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   7
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Partida4 
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
      Left            =   1080
      TabIndex        =   6
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CheckBox Confirma3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   5
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Partida3 
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
      Left            =   1080
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CheckBox Confirma2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   3
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Partida2 
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
      Left            =   1080
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CheckBox Confirma1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   1
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Partida1 
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
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "A UTILIZAR"
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
      Left            =   4680
      TabIndex        =   27
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "CANTIDAD"
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
      Left            =   2880
      TabIndex        =   21
      Top             =   1200
      Width           =   1575
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
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Materia Prima"
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
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PARTIDA"
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
      Left            =   1080
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "PrgHojaDesvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ZZSaldo As Double
Dim ZLugar As Integer

Private Sub Aceptar_Click()

    ZCargaDesvio(ZLugar, 3) = "N"
    ZColumna = 3
    
    If Confirma1.Value = 1 And Val(Partida1.Text) <> 0 Then
        ZColumna = ZColumna + 1
        ZCargaDesvio(ZLugar, ZColumna) = Partida1.Text
        ZCargaDesvio(ZLugar, ZColumna + 5) = Usa1.Text
    End If
    If Confirma2.Value = 1 And Val(Partida2.Text) <> 0 Then
        ZColumna = ZColumna + 1
        ZCargaDesvio(ZLugar, ZColumna) = Partida2.Text
        ZCargaDesvio(ZLugar, ZColumna + 5) = Usa2.Text
    End If
    If Confirma3.Value = 1 And Val(Partida3.Text) <> 0 Then
        ZColumna = ZColumna + 1
        ZCargaDesvio(ZLugar, ZColumna) = Partida3.Text
        ZCargaDesvio(ZLugar, ZColumna + 5) = Usa3.Text
    End If
    If Confirma4.Value = 1 And Val(Partida4.Text) <> 0 Then
        ZColumna = ZColumna + 1
        ZCargaDesvio(ZLugar, ZColumna) = Partida4.Text
        ZCargaDesvio(ZLugar, ZColumna + 5) = Usa4.Text
    End If
    If Confirma5.Value = 1 And Val(Partida5.Text) <> 0 Then
        ZColumna = ZColumna + 1
        ZCargaDesvio(ZLugar, ZColumna) = Partida5.Text
        ZCargaDesvio(ZLugar, ZColumna + 5) = Usa5.Text
    End If
    
    For Ciclo = 1 To 100
        If Val(ZCargaDesvio(Ciclo, 2)) <> 0 And ZCargaDesvio(Ciclo, 3) = "S" Then
            Call Proceso
            Exit Sub
        End If
    Next Ciclo
    
    PrgHojaDesvio.Hide
    Unload Me
    PrgHoja.Show
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    Call Proceso
End Sub

Private Sub Proceso()

    MateriaPrima.Text = ""
    Cantidad.Text = ""
    
    Partida1.Text = ""
    Partida2.Text = ""
    Partida3.Text = ""
    Partida4.Text = ""
    Partida5.Text = ""
    
    Cantidad1.Text = ""
    Cantidad2.Text = ""
    Cantidad3.Text = ""
    Cantidad4.Text = ""
    Cantidad5.Text = ""
    
    Confirma1.Value = 0
    Confirma2.Value = 0
    Confirma3.Value = 0
    Confirma4.Value = 0
    Confirma5.Value = 0
    
    Usa1.Text = ""
    Usa2.Text = ""
    Usa3.Text = ""
    Usa4.Text = ""
    Usa5.Text = ""
    
    ZLugarDesvio = 0
    For Ciclo = 1 To 100
        If Val(ZCargaDesvio(Ciclo, 2)) <> 0 And ZCargaDesvio(Ciclo, 3) = "S" Then
        
            ZZArticulo = ZCargaDesvio(Ciclo, 1)
            ZLugar = Ciclo
            LugarDesvio = 0
            
            MateriaPrima.Text = ZCargaDesvio(Ciclo, 1)
            Cantidad.Text = ZCargaDesvio(Ciclo, 2)
    
            XParam = "'" + ZZArticulo + "','" _
                     + ZZArticulo + "'"
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
                        
                        WArticulo = rstLaudo!Articulo
                        WCantidad = rstLaudo!Liberada
                        WLaudo = rstLaudo!Laudo
                        ZZSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                        Call Redondeo(ZZSaldo)
                
                        If rstLaudo!Marca = "X" And rstLaudo!Saldo = 0 Then
                
                                Else
                        
                            If rstLaudo!Articulo = ZZArticulo And ZZSaldo <> 0 Then
                
                                Entra = "N"
                                    
                                If WLaudo >= 190000 And WLaudo <= 194999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 990000 And WLaudo <= 994999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 290000 And WLaudo <= 294999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 390000 And WLaudo <= 394999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 490000 And WLaudo <= 494999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 590000 And WLaudo <= 594999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 690000 And WLaudo <= 694999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 790000 And WLaudo <= 794999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 890000 And WLaudo <= 894999 Then
                                    Entra = "S"
                                End If
                                If Entra = "S" Then
                                    LugarDesvio = LugarDesvio + 1
                                    Select Case LugarDesvio
                                        Case 1
                                            Partida1.Text = WLaudo
                                            Cantidad1.Text = Str$(ZZSaldo)
                                        Case 2
                                            Partida2.Text = WLaudo
                                            Cantidad2.Text = Str$(ZZSaldo)
                                        Case 3
                                            Partida3.Text = WLaudo
                                            Cantidad3.Text = Str$(ZZSaldo)
                                        Case 4
                                            Partida4.Text = WLaudo
                                            Cantidad4.Text = Str$(ZZSaldo)
                                        Case 5
                                            Partida5.Text = WLaudo
                                            Cantidad5.Text = Str$(ZZSaldo)
                                        Case Else
                                    End Select
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
    
            XParam = "'" + ZZArticulo + "','" _
                        + ZZArticulo + "'"
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
                        
                        WArticulo = rstMovguia!Articulo
                        WCantidad = rstMovguia!Cantidad
                        WFecha = rstMovguia!Fecha
                        WCodigo = rstMovguia!Codigo
                        WMovi = rstMovguia!Movi
                        WDestino = rstMovguia!Destino
                        WTipomov = rstMovguia!Tipomov
                        ZZSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(ZZSaldo)
                        WLaudo = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                
                        If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                                Else
                        
                            If rstMovguia!Tipo = "M" And rstMovguia!Articulo = ZZArticulo And ZZSaldo <> 0 Then
                    
                        
                                Entra = "N"
                                    
                                If WLaudo >= 190000 And WLaudo <= 194999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 990000 And WLaudo <= 994999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 290000 And WLaudo <= 294999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 390000 And WLaudo <= 394999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 490000 And WLaudo <= 494999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 590000 And WLaudo <= 594999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 690000 And WLaudo <= 694999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 790000 And WLaudo <= 794999 Then
                                    Entra = "S"
                                End If
                                If WLaudo >= 890000 And WLaudo <= 894999 Then
                                    Entra = "S"
                                End If
                                If Entra = "S" Then
                                    LugarDesvio = LugarDesvio + 1
                                    Select Case LugarDesvio
                                        Case 1
                                            Partida1.Text = WLaudo
                                            Cantidad1.Text = Str$(ZZSaldo)
                                        Case 2
                                            Partida2.Text = WLaudo
                                            Cantidad2.Text = Str$(ZZSaldo)
                                        Case 3
                                            Partida3.Text = WLaudo
                                            Cantidad3.Text = Str$(ZZSaldo)
                                        Case 4
                                            Partida4.Text = WLaudo
                                            Cantidad4.Text = Str$(ZZSaldo)
                                        Case 5
                                            Partida5.Text = WLaudo
                                            Cantidad5.Text = Str$(ZZSaldo)
                                        Case Else
                                    End Select
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
            
            Exit Sub
        
        End If
    Next Ciclo


End Sub


