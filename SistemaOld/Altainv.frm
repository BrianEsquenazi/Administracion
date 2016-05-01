VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAltainv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Talones de Inventario"
   ClientHeight    =   6525
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   11805
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   6525
   ScaleWidth      =   11805
   Visible         =   0   'False
   Begin Crystal.CrystalReport Listado 
      Left            =   9960
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impreord.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   500
      Left            =   4560
      TabIndex        =   12
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   11
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   500
      Left            =   240
      TabIndex        =   9
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      Height          =   500
      Left            =   2400
      TabIndex        =   8
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   500
      Left            =   1320
      TabIndex        =   6
      Top             =   5880
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   11535
      Begin VB.TextBox WUbicacion 
         Height          =   285
         Left            =   7800
         MaxLength       =   10
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox WTalon 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         MaxLength       =   6
         TabIndex        =   25
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox WObservaciones 
         Height          =   285
         Left            =   8880
         MaxLength       =   20
         TabIndex        =   23
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox WLote 
         Height          =   285
         Left            =   5880
         MaxLength       =   20
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox WCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         MaxLength       =   10
         TabIndex        =   15
         Text            =   " "
         Top             =   600
         Width           =   975
      End
      Begin MSMask.MaskEdBox WTerminado 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox WTipo 
         Height          =   285
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   13
         Text            =   " "
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   7
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   3000
         TabIndex        =   5
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ubicacion"
         Height          =   255
         Left            =   7800
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Talon"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   8880
         TabIndex        =   24
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote"
         Height          =   255
         Left            =   5880
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   6840
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   4080
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Materia Prima"
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P. Terminado"
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M/T"
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.Label WDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   4080
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   500
      Left            =   3480
      TabIndex        =   2
      Top             =   5880
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4185
      Left            =   120
      OleObjectBlob   =   "Altainv.frx":0000
      TabIndex        =   1
      Top             =   480
      Width           =   11565
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10680
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nro Movimiento"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgAltainv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 9 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Tipo As String
Private Articulo As String
Private Terminado As String
Private Auxiliar(100, 8) As String
Dim rstInventario As Recordset
Dim spInventario As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim XParam As String

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    DBGrid1.Col = 3
    DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
    DBGrid1.Col = 5
    DBGrid1.Text = ""
    
    DBGrid1.Col = 6
    DBGrid1.Text = ""
    
    DBGrid1.Col = 7
    DBGrid1.Text = ""
    
    DBGrid1.Col = 8
    DBGrid1.Text = ""
    
    WTalon.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WLote.Text = ""
    WCantidad.Text = ""
    WUbicacion.Text = ""
    WObservaciones.Text = ""
    
    WLinea.Text = ""
    WTalon.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    PrgAltainv.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) <> 0 Then
        WLinea.Text = DBGrid1.Row + 1
        WTalon.Text = DBGrid1.Text
            Else
        WTalon.Text = ""
        WLinea.Text = ""
    End If
    
    DBGrid1.Col = 1
    WTipo.Text = DBGrid1.Text

    DBGrid1.Col = 2
    If Len(DBGrid1.Text) = 12 Then
        WTerminado.Text = DBGrid1.Text
            Else
        WTerminado.Text = "  -     -   "
    End If

    DBGrid1.Col = 3
    If Len(DBGrid1.Text) = 10 Then
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -   -   "
    End If
    
    DBGrid1.Col = 4
    WDescripcion.Caption = DBGrid1.Text
    
    DBGrid1.Col = 5
    WLote.Text = DBGrid1.Text
    
    DBGrid1.Col = 6
    WCantidad.Text = DBGrid1.Text
    
    DBGrid1.Col = 7
    WUbicacion.Text = DBGrid1.Text
    
    DBGrid1.Col = 8
    WObservaciones.Text = DBGrid1.Text
    
    WTalon.SetFocus

End Sub

Private Sub Graba_Click()

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 20) * 20
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    spInventario = "BorrarInventario " + "'" + Codigo.Text + "'"
    Set rstInventario = db.OpenRecordset(spInventario, dbOpenDynaset, dbSQLPassThrough)
    
    Renglon = 0
    DBGrid1.Refresh
                
    For A = 0 To 4
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Talon = DBGrid1.Text
            
            DBGrid1.Col = 1
            Tipo = DBGrid1.Text
                                       
            DBGrid1.Col = 2
            Terminado = DBGrid1.Text
                    
            DBGrid1.Col = 3
            Articulo = DBGrid1.Text
            
            DBGrid1.Col = 5
            Lote = DBGrid1.Text
                    
            DBGrid1.Col = 6
            Cantidad = DBGrid1.Text
            
            DBGrid1.Col = 7
            Ubicacion = DBGrid1.Text
            
            DBGrid1.Col = 8
            Observaciones = DBGrid1.Text
                    
            If Talon <> "" Then
            
                WPartida = ""
                If (Left$(Articulo, 2) = "DY" Or Left$(Articulo, 2) = "DW" Or Left$(Articulo, 2) = "DS") And WEmpresa <> "0008" Then
                
                    WEntra = "N"
                    WPartida = Lote
                    
                    XParam = "'" + Lote + "','" _
                                 + Articulo + "'"
                    spLaudo = "ListaLaudoArticuloPartiOri " + XParam
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        Lote = rstLaudo!laudo
                        WEntra = "S"
                        rstLaudo.Close
                    End If
                    
                    If WEntra = "N" Then
                
                        Sql1 = "Select *"
                        Sql2 = " FROM Guia"
                        Sql3 = " Where Guia.Articulo = " + "'" + Articulo + "'"
                        Sql4 = " and Guia.PartiOri = " + "'" + Lote + "'"
                        spMovguia = Sql1 + Sql2 + Sql3 + Sql4
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            Lote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                            WEntra = "S"
                            rstMovguia.Close
                        End If
                    
                    End If
                    
                End If
                    
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Codigo.Text)
                Call Ceros(Auxi1, 6)
                
                WCodigo = Codigo.Text
                WRenglon = Str$(Renglon)
                WTipo = Tipo
                WArticulo = Articulo
                WTerminado = Terminado
                WTalon = Talon
                WCantidad = Cantidad
                WLote = Lote
                WObservaciones = Observaciones
                WUbicacion = Ubicacion
                WClave = Auxi1 + Auxi
                
                Sql1 = "INSERT INTO Inventario ("
                Sql2 = "Clave ,"
                Sql3 = "Numero ,"
                Sql4 = "Renglon ,"
                Sql5 = "Tipo ,"
                Sql6 = "Articulo ,"
                Sql7 = "Terminado ,"
                Sql8 = "Talon ,"
                Sql9 = "Cantidad ,"
                Sql10 = "Lote ,"
                Sql11 = "Ubicacion ,"
                Sql12 = "Observaciones ,"
                Sql13 = "Partida )"
                Sql14 = "Values ("
                Sql15 = "'" + WClave + "',"
                Sql16 = "'" + WCodigo + "',"
                Sql17 = "'" + WRenglon + "',"
                Sql18 = "'" + WTipo + "',"
                Sql19 = "'" + WArticulo + "',"
                Sql20 = "'" + WTerminado + "',"
                Sql21 = "'" + WTalon + "',"
                Sql22 = "'" + WCantidad + "',"
                Sql23 = "'" + WLote + "',"
                Sql24 = "'" + WUbicacion + "',"
                Sql25 = "'" + WObservaciones + "',"
                Sql26 = "'" + WPartida + "')"
            
                spInventario = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                           Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                           Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26
                Set rstInventario = db.OpenRecordset(spInventario, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                
        Next iRow
            
    Next A
        
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Codigo.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    
    WTalon.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WLote.Text = ""
    WCantidad.Text = ""
    WObservaciones.Text = ""
    WUbicacion.Text = ""
    
    WTalon.SetFocus
    
End Sub

Private Sub Limpia_Click()

    WLinea.Text = ""
    
    WTalon.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WLote.Text = ""
    WCantidad.Text = ""
    WObservaciones.Text = ""
    WUbicacion.Text = ""

    Codigo.Text = ""
    
    For A = 0 To 4
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 8
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next A
    
    Codigo.Text = "1"
    
    spInventario = "ListaInventarioNumero"
    Set rstInventario = db.OpenRecordset(spInventario, dbOpenSnapshot, dbSQLPassThrough)
    If rstInventario.RecordCount > 0 Then
        With rstInventario
            .MoveLast
            Codigo.Text = rstInventario!Numero + 1
        End With
        rstInventario.Close
    End If
    
    DBGrid1.FirstRow = 0
    Renglon = 0

    Codigo.SetFocus

End Sub

Private Sub WTalon_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(WTalon.Text) <> 0 Then
            WTipo.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WTipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WTipo.Text = "M" Or WTipo.Text = "T" Then
            If WTipo.Text = "M" Then
                WArticulo.SetFocus
                    Else
                WTerminado.SetFocus
            End If
                Else
            WTipo.SetFocus
        End If
    End If
End Sub

Private Sub WTerminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem If Left$(WTerminado.Text, 2) <> "NI" Then
            spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WDescripcion.Caption = Left$(rstTerminado!Descripcion, 15)
                rstTerminado.Close
                WLote.SetFocus
                    Else
                WTerminado.SetFocus
            End If
        Rem         Else
        Rem     WDescripcion.Caption = ""
        Rem     WLote.Text = ""
        Rem     WCantidad.SetFocus
        Rem End If
    End If
End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WDescripcion.Caption = Left$(rstArticulo!Descripcion, 15)
            rstArticulo.Close
            WLote.SetFocus
               Else
            WArticulo.SetFocus
        End If
    End If
End Sub

Private Sub WLote_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        If WTipo.Text = "M" Then
        
            WEntra = "N"
                
            If (Left$(WArticulo.Text, 2) = "DY" Or Left$(WArticulo.Text, 2) = "DW" Or Left$(WArticulo.Text, 2) = "DS") And WEmpresa <> "0008" Then
            
                XParam = "'" + WLote.Text + "','" _
                             + WArticulo.Text + "'"
                Rem spLaudo = "ListaLaudoArticulo " + XParam
                spLaudo = "ListaLaudoArticuloPartiOri " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                
                    Sql1 = "Select *"
                    Sql2 = " FROM Guia"
                    Sql3 = " Where Guia.Articulo = " + "'" + WArticulo.Text + "'"
                    Sql4 = " and Guia.PartiOri = " + "'" + WLote.Text + "'"
                    spMovguia = Sql1 + Sql2 + Sql3 + Sql4
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                
                End If
                    
                    Else
        
                WControla = 0
                WArticulo.Text = UCase(WArticulo.Text)
                spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                    rstArticulo.Close
                End If
            
                If WControla = 0 Then
            
                    XParam = "'" + WLote.Text + "','" _
                            + WArticulo.Text + "'"
                    spLaudo = "ListaLaudoArticulo " + XParam
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        WCanti = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                        WEntra = "S"
                        rstLaudo.Close
                    End If
                
                    If WEntra = "N" Then
                        XParam = "'" + WArticulo.Text + "','" _
                                + WLote.Text + "'"
                        spMovguia = "ListaMovguiaLote " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            WCanti = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                            WEntra = "S"
                            rstMovguia.Close
                        End If
                    End If
                
                        Else
                    
                    WEntra = "S"
                
                End If
                
            End If
            
            If WEntra = "N" Then
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + WLote.Text + " inexistente"
                G% = MsgBox(m$, 0, "Movimientos Varios de Stock")
                WLote.SetFocus
                    Else
                WCantidad.SetFocus
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote.Text + "','" _
                        + WTerminado.Text + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WCanti = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WTerminado.Text + "','" _
                            + WLote.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WCanti = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra = "N" Then
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
                G% = MsgBox(m$, 0, "Movimientos Varios de Stock")
                WLote.SetFocus
                    Else
                WCantidad.SetFocus
            End If
            
        End If
    
    End If
    
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        WUbicacion.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WUbicacion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WObservaciones.SetFocus
    End If
End Sub

Private Sub WObservaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Alta_Vector
        Call Ingresa_Click
        WTalon.SetFocus
    End If
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 50 Then
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
ReDim UserData(0 To 8, 0 To 50)

mTotalRows& = 50

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
For i = 0 To 8
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Talon"
             DBGrid1.Columns(newcnt).Width = 900
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Tipo"
             DBGrid1.Columns(newcnt).Width = 400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Prod.Terminado"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Materia Prima"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 2000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Lote"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 6
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 900
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 7
             DBGrid1.Columns(newcnt).Caption = "Ubicacion"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 8
             DBGrid1.Columns(newcnt).Caption = "Observaciones"
             DBGrid1.Columns(newcnt).Width = 2000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    Codigo.Text = "1"
    
    spInventario = "ListaInventarioNumero"
    Set rstInventario = db.OpenRecordset(spInventario, dbOpenSnapshot, dbSQLPassThrough)
    If rstInventario.RecordCount > 0 Then
        With rstInventario
            .MoveLast
            Codigo.Text = rstInventario!Numero + 1
        End With
        rstInventario.Close
    End If
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgAltainv.Caption = "Ingreso de Talones de Inventario :  " + !Nombre
        End If
    End With
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Codigo.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For A = 0 To 4
    Suma = A * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 8
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
    Erase Auxiliar
    Renglon = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM Inventario"
    Sql3 = " Where Numero = " + Codigo.Text
    Sql4 = " Order by Clave"
    spInventario = Sql1 + Sql2 + Sql3 + Sql4
    Set rstInventario = db.OpenRecordset(spInventario, dbOpenSnapshot, dbSQLPassThrough)
    If rstInventario.RecordCount > 0 Then
    
        With rstInventario
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 20) * 20
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstInventario!Talon
                
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstInventario!Tipo
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = rstInventario!Terminado
                    Auxi1 = rstInventario!Terminado
                
                    DBGrid1.Col = 3
                    DBGrid1.Text = rstInventario!Articulo
                    Auxi2 = rstInventario!Articulo
                
                    If (Left$(rstInventario!Articulo, 2) = "DY" Or Left$(rstInventario!Articulo, 2) = "DW" Or Left$(rstInventario!Articulo, 2) = "DS") And WEmpresa <> "0008" Then
                        DBGrid1.Col = 5
                        DBGrid1.Text = rstInventario!Partida
                            Else
                        DBGrid1.Col = 5
                        DBGrid1.Text = rstInventario!Lote
                    End If
                    
                    DBGrid1.Col = 6
                    DBGrid1.Text = Pusing("###,###.##", rstInventario!Cantidad)
                
                    DBGrid1.Col = 7
                    DBGrid1.Text = rstInventario!Ubicacion
                    
                    DBGrid1.Col = 8
                    DBGrid1.Text = rstInventario!Observaciones
                    
                    Auxiliar(Renglon, 1) = Auxi1
                    Auxiliar(Renglon, 2) = Auxi2
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstInventario.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0

    For Da = 1 To WRenglon
    
        Auxi1 = Auxiliar(Da, 1)
        Auxi2 = Auxiliar(Da, 2)
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 20) * 20
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
    
        spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DBGrid1.Col = 4
            DBGrid1.Text = rstTerminado!Descripcion
            rstTerminado.Close
        End If
        
        spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DBGrid1.Col = 4
            DBGrid1.Text = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
    Next Da

    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 20) * 20
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 20) * 20
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    WTalon.SetFocus

End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) = 0 Then

            If Renglon < 15 Then
    
                Renglon = Renglon + 1
            
                Lugar1 = Int((Renglon - 1) / 20) * 20
                Lugar2 = Renglon - Lugar1
                
                DBGrid1.FirstRow = Lugar1
                DBGrid1.Row = Lugar2 - 1
                    
                WAnterior = DBGrid1.Row
                
                DBGrid1.Col = 0
                DBGrid1.Text = WTalon.Text
                
                DBGrid1.Col = 1
                DBGrid1.Text = WTipo.Text
                
                DBGrid1.Col = 2
                DBGrid1.Text = WTerminado.Text
                
                DBGrid1.Col = 3
                DBGrid1.Text = WArticulo.Text
                
                DBGrid1.Col = 4
                DBGrid1.Text = WDescripcion.Caption
                
                DBGrid1.Col = 5
                DBGrid1.Text = WLote.Text
                    
                DBGrid1.Col = 6
                DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
            
                DBGrid1.Col = 7
                DBGrid1.Text = WUbicacion.Text
            
                DBGrid1.Col = 8
                DBGrid1.Text = WObservaciones.Text
            
                DBGrid1.Row = Renglon
                DBGrid1.Col = 0
                
                    Else
                    
                m$ = "No se puede ingresar mas talones en este movimiento"
                A% = MsgBox(m$, 0, "Archivo de Articulos")
                
            End If
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WTalon.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WTipo.Text
            
            DBGrid1.Col = 2
            DBGrid1.Text = WTerminado.Text
            
            DBGrid1.Col = 3
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 4
            DBGrid1.Text = WDescripcion.Caption
            
            DBGrid1.Col = 5
            DBGrid1.Text = WLote.Text
                
            DBGrid1.Col = 6
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
            
            DBGrid1.Col = 7
            DBGrid1.Text = WUbicacion.Text
            
            DBGrid1.Col = 8
            DBGrid1.Text = WObservaciones.Text
            
            DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spInventario = "ListaInventario " + "'" + Codigo.Text + "'"
        Set rstInventario = db.OpenRecordset(spInventario, dbOpenSnapshot, dbSQLPassThrough)
        If rstInventario.RecordCount > 0 Then
            rstInventario.Close
            Call Proceso_Click
            WTalon.SetFocus
                Else
            WCodigo = Codigo.Text
            Call Limpia_Click
            Codigo.Text = WCodigo
            WTalon.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

