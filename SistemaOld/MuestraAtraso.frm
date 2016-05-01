VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgMuestraAtraso 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Aviso de Atraso en Entrega"
   ClientHeight    =   8325
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11790
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   8325
   ScaleWidth      =   11790
   Begin Crystal.CrystalReport ListaGRilla 
      Left            =   10080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "muestra.rpt"
   End
   Begin VB.ListBox Lista 
      Height          =   645
      Left            =   2640
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox Pantalla 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5715
      ItemData        =   "MuestraAtraso.frx":0000
      Left            =   3480
      List            =   "MuestraAtraso.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   4815
   End
   Begin MSFlexGridLib.MSFlexGrid Vector 
      Height          =   7455
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13150
      _Version        =   327680
      BackColor       =   16777215
      ForeColor       =   4210752
      FocusRect       =   2
      GridLines       =   0
   End
   Begin VB.CommandButton Modifica 
      Caption         =   "  Modifica Aviso             (F3)"
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
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Baja 
      Caption         =   "  Borra Aviso             (F2)"
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
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Alta 
      Caption         =   "Alta (F1)"
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
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Imprexx 
      Caption         =   "Impresion (F9)"
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
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Fin (F10)"
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
      Left            =   7200
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "PrgMuestraAtraso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstAtraso As Recordset
Dim spAtraso As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim XParam As String
Dim Auxiliar(20000) As String
Dim XEmpresa As String
Dim WFecha As String
Dim WFecha2 As String
Dim WDia As String
Dim WMes As String
Dim WCod As String
Dim ColumnaOpcion As Integer
Dim Seleccion As String
Dim WPasa(10000) As String

Private Sub Alta_Click()
    WPosi1 = Vector.TopRow
    WPosi2 = Vector.Row
    WPosi3 = Vector.Col
    Vector.Visible = False
    WAtraso = "0"
    PrgAtrasoEntrega.Show
End Sub

Private Sub Modifica_Click()
    WPosi1 = Vector.TopRow
    WPosi2 = Vector.Row
    WPosi3 = Vector.Col
    Vector.Visible = False
    Fila = Vector.Row
    WAtraso = Auxiliar(Fila)
    If Val(WAtraso) <> 0 Then
        PrgAtrasoEntrega.Show
    End If
End Sub

Private Sub Baja_Click()
    WPosi1 = Vector.TopRow
    WPosi2 = Vector.Row
    WPosi3 = Vector.Col
    Fila = Vector.Row
    WAtraso = Auxiliar(Fila)
    If Val(WAtraso) <> 0 Then
        T$ = "Borrar Registro"
        m$ = "Desea Borrar el aviso de atraso"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            Sql1 = "DELETE Atraso"
            Sql2 = " Where Numero = " + "'" + WAtraso + "'"
            spAtraso = Sql1 + Sql2
            Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
            Call Proceso_Click
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    PrgMuestraAtraso.Hide
    Unload Me
    Close
    End
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
        
    Select Case ColumnaOpcion
        Case 0, 1
            Sql1 = "Select *"
            Sql2 = " FROM Atraso"
            Sql3 = " Order by Numero"
            spAtraso = Sql1 + Sql2 + Sql3
        Case 2
            Sql1 = "Select *"
            Sql2 = " FROM Atraso"
            Sql3 = " Where Atraso.Fecha = " + "'" + Seleccion + "'"
            Sql4 = " Order by Numero"
            spAtraso = Sql1 + Sql2 + Sql3 + Sql4
        Case 4
            Sql1 = "Select *"
            Sql2 = " FROM Atraso"
            Sql3 = " Where Atraso.DesCliente = " + "'" + Seleccion + "'"
            Sql4 = " Order by Numero"
            spAtraso = Sql1 + Sql2 + Sql3 + Sql4
        Case 5
            Sql1 = "Select *"
            Sql2 = " FROM Atraso"
            Sql3 = " Where Atraso.Terminado = " + "'" + Seleccion + "'"
            Sql4 = " Order by Numero"
            spAtraso = Sql1 + Sql2 + Sql3 + Sql4
        Case 6
            Sql1 = "Select *"
            Sql2 = " FROM Atraso"
            Sql3 = " Where Atraso.Problema = " + "'" + Seleccion + "'"
            Sql4 = " Order by Numero"
            spAtraso = Sql1 + Sql2 + Sql3 + Sql4
        Case 7
            Sql1 = "Select *"
            Sql2 = " FROM Atraso"
            Sql3 = " Where Atraso.Articulo = " + "'" + Seleccion + "'"
            Sql4 = " Order by Numero"
            spAtraso = Sql1 + Sql2 + Sql3 + Sql4
        Case Else
    End Select
            
    Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
    If rstAtraso.RecordCount > 0 Then
        With rstAtraso
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    WLugar = WLugar + 1
                    Auxiliar(WLugar) = Str$(!Numero)
                    
                    Vector.TextMatrix(WLugar, 1) = Str$(!Numero)
                    Vector.TextMatrix(WLugar, 2) = !Fecha
                    Vector.TextMatrix(WLugar, 3) = Str$(!Pedido)
                    Vector.TextMatrix(WLugar, 4) = !DesCliente
                    Vector.TextMatrix(WLugar, 5) = !Terminado
                    Vector.TextMatrix(WLugar, 6) = !Problema
                    Vector.TextMatrix(WLugar, 7) = !Articulo
                    Vector.TextMatrix(WLugar, 8) = !FechaEntrega
                    
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
        
        End With
        rstAtraso.Close
    End If
    
    Vector.Visible = True
    
    Renglon = Renglon + 1
    Vector.Row = Renglon
    
    Vector.Col = 0
    Vector.Text = ""
    
    If WPosi1 <> 0 And WPosi2 <> 0 And WPosi3 <> 0 Then
        Vector.TopRow = WPosi1
        Vector.Col = WPosi3
        Vector.Row = WPosi2
            Else
        If WLugar > 20 Then
            Vector.TopRow = WLugar - 20
                Else
            Vector.TopRow = 1
        End If
        Vector.Col = 1
        Vector.Row = WLugar
    End If
    
    Vector.SetFocus
    
End Sub

Private Sub Vector_DblClick()

    ColumnaOpcion = Vector.Col
    WPosi1 = 1
    WPosi2 = 1
    WPosi3 = 1
    
    pantalla.Clear
    pantalla.AddItem ""
    Select Case ColumnaOpcion
        Case 2
            Sql1 = "Select DISTINCT OrdFecha"
            Sql2 = " FROM Atraso"
            Sql3 = " Order by OrdFecha"
            spAtraso = Sql1 + Sql2 + Sql3
            Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
            If rstAtraso.RecordCount > 0 Then
                With rstAtraso
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            pantalla.AddItem Right$(!OrdFecha, 2) + "/" + Mid$(!OrdFecha, 5, 2) + "/" + Left$(!OrdFecha, 4)
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstAtraso.Close
            End If
            pantalla.Visible = True
            
        Case 4
            Sql1 = "Select DISTINCT DesCliente"
            Sql2 = " FROM Atraso"
            Sql3 = " Order by DesCliente"
            spAtraso = Sql1 + Sql2 + Sql3
            Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
            If rstAtraso.RecordCount > 0 Then
                With rstAtraso
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            pantalla.AddItem !DesCliente
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstAtraso.Close
            End If
            pantalla.Visible = True
            
        Case 5
            Sql1 = "Select DISTINCT Terminado"
            Sql2 = " FROM Atraso"
            Sql3 = " Order by Terminado"
            spAtraso = Sql1 + Sql2 + Sql3
            Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
            If rstAtraso.RecordCount > 0 Then
                With rstAtraso
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            pantalla.AddItem !Terminado
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstAtraso.Close
            End If
            pantalla.Visible = True
            
        Case 6
            Sql1 = "Select DISTINCT Problema"
            Sql2 = " FROM Atraso"
            Sql3 = " Order by Problema"
            spAtraso = Sql1 + Sql2 + Sql3
            Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
            If rstAtraso.RecordCount > 0 Then
                With rstAtraso
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            pantalla.AddItem !Problema
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstAtraso.Close
            End If
            pantalla.Visible = True
            
        Case 7
            Sql1 = "Select DISTINCT Articulo"
            Sql2 = " FROM Atraso"
            Sql3 = " Order by Articulo"
            spAtraso = Sql1 + Sql2 + Sql3
            Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
            If rstAtraso.RecordCount > 0 Then
                With rstAtraso
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            pantalla.AddItem !Articulo
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstAtraso.Close
            End If
            pantalla.Visible = True
            
        Case Else
        
    End Select
    
    Rem Vector.Col = 10
    Rem Vector.Col = 1
    Rem WXSol = Vector.Text
    Rem PrgSol.Show
End Sub

Private Sub Form_Activate()
    Call Proceso_Click
End Sub

Private Sub Vector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Alta_Click
        Case 113
            Call Baja_Click
        Case 114
            Call Modifica_Click
        Case 120
            Call Impre_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub

Private Sub Impre_Click()
  
    Rem RowIni = Muestra.Row
    Rem RowFin = Muestra.RowSel
    
    Rem DesdeNumero = Vector(RowIni, 1)
    Rem HastaNumero = Vector(RowFin, 1)
    
    Rem Listado.Destination = 1
    Rem DbConnect = db.Connect
    Rem DSQ = getDatabase(DbConnect)
    Rem Listado.SQLQuery = "SELECT MuestraImpre.Numero, MuestraImpre.Fecha, MuestraImpre.Codigo, MuestraImpre.Descripcion, MuestraImpre.Cantidad, MuestraImpre.DescriCLiente, MuestraImpre.Cliente, MuestraImpre.Observaciones, MuestraImpre.Fecha2, MuestraImpre.Codigo2, MuestraImpre.Descripcion2, MuestraImpre.Lote, MuestraImpre.Observaciones2, MuestraImpre.Cantidad2 " _
    rem                 + "From " _
    rem                 + DSQ + ".dbo.MuestraImpre MuestraImpre " _
    rem                 + "Where " _
    rem                 + "MuestraImpre.Numero >= 0 AND " _
    rem                 + "MuestraImpre.Numero <= 999999 " _
    rem                 + "Order By MuestraImpre.Numero ASC"
    Rem Listado.Connect = Connect()
    Rem Listado.Action = 1
    
End Sub

Private Sub Limpia_Vector()

    Vector.Clear

    Rem ponga la muestra en negritas
    Rem Vector.Font.Bold = True

    ' Establesco loa Valores de la muestra
    
    Vector.FixedCols = 1
    Vector.Cols = 9
    Vector.FixedRows = 1
    Vector.Rows = 20000
    
    Vector.ColWidth(0) = 200
    Vector.Row = 0
    
    For Ciclo = 1 To Vector.Cols - 1
        Vector.Col = Ciclo
        Select Case Ciclo
            Case 1
                Vector.Text = "Numero"
                Vector.ColWidth(Ciclo) = 900
                Vector.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                Vector.Text = "Fecha"
                Vector.ColWidth(Ciclo) = 1000
                Vector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                Vector.Text = "Pedido"
                Vector.ColWidth(Ciclo) = 900
                Vector.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                Vector.Text = "Cliente"
                Vector.ColWidth(Ciclo) = 1900
                Vector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                Vector.Text = "P.Terminado"
                Vector.ColWidth(Ciclo) = 1200
                Vector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 6
                Vector.Text = "Problema"
                Vector.ColWidth(Ciclo) = 3000
                Vector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 7
                Vector.Text = "M.Prima"
                Vector.ColWidth(Ciclo) = 1200
                Vector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                Vector.Text = "F.Entrega"
                Vector.ColWidth(Ciclo) = 1000
                Vector.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Vector.AllowUserResizing = flexResizeBoth
    
    Vector.Col = 1
    Vector.Row = 1
    
End Sub

Private Sub Pantalla_Click()
    If pantalla.ListIndex <> 0 Then
        Seleccion = pantalla.Text
            Else
        Seleccion = ""
        ColumnaOpcion = 0
    End If
    pantalla.Visible = False
    Call Proceso_Click
End Sub
