VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPasaDW 
   AutoRedraw      =   -1  'True
   Caption         =   "Pasa DW"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      Begin MSMask.MaskEdBox ArticuloDy 
         Height          =   300
         Left            =   1680
         TabIndex        =   6
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Terminado 
         Height          =   300
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   360
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
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "ArticuloDy"
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
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Terminado"
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "PrgPasaDW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WInicial As Double
Private WOrden As String
Private WClave As String
Private WDescripcion As String
Private WSaldo As Double

Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstCotiza As Recordset
Dim spCotiza As String
Dim rstEspecificaciones As Recordset
Dim spEspecificaciones As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstPrestamo As Recordset
Dim spPrestamo As String
Dim rstPrueart As Recordset
Dim spPrueart As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String

Dim XParam As String
Dim ZVector(10000) As String
Dim ZVectorII(10000, 5) As String
Dim Empe(100, 2) As String
Dim ZLugar As Integer
Dim ZPago As String
Dim ZZPrecios As Double
Dim ZZCantidad As Double

Private Sub Acepta_Click()

    Terminado.Text = UCase(Terminado.Text)
    ArticuloDy.Text = UCase(ArticuloDy.Text)
    
    If Terminado.Text = "  -     -   " Then
        Exit Sub
    End If
    If ArticuloDy.Text = "  -   -   " Then
        Exit Sub
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Terminado"
    ZSql = ZSql + " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        rstTerminado.Close
            Else
        Exit Sub
    End If

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Codigo = " + "'" + ArticuloDy.Text + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        rstArticulo.Close
            Else
        Exit Sub
    End If
    
    ZFecha = "20040101"
    
    Empe(1, 1) = "0001"
    Empe(1, 2) = "Empresa01"
    Empe(2, 1) = "0003"
    Empe(2, 2) = "Empresa03"
    Empe(3, 1) = "0007"
    Empe(3, 2) = "Empresa07"
    
    XEmpresa = WEmpresa
    
    For CiclaEmpresa = 1 To 3
    
    WEmpresa = Empe(CiclaEmpresa, 1)
    txtOdbc = Empe(CiclaEmpresa, 2)
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Erase ZVector
    ZLugar = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM Estadistica"
    Sql3 = " Where Estadistica.Articulo = " + "'" + Terminado.Text + "'"
    Sql4 = " and Estadistica.OrdFecha >= " + "'" + ZFecha + "'"
    Sql5 = " Order by Estadistica.Numero"
    spEstadistica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
            .MoveFirst
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar) = rstEstadistica!Clave
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
            
        End With
        
        rstEstadistica.Close
        
    End If
    
    
    For Ciclo = 1 To ZLugar
    
        ZClave = ZVector(Ciclo)
        
        ZCampo1 = Left$(ArticuloDy.Text, 3) + "00" + Right$(ArticuloDy.Text, 7)
        ZCampo2 = Left$(ZCampo1, 8)
        ZCampo3 = "DW"
        ZCampo4 = "M"
        ZCampo5 = ArticuloDy.Text
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Estadistica SET "
        ZSql = ZSql + " Estadistica.Articulo = " + "'" + ZCampo1 + "',"
        ZSql = ZSql + " Estadistica.WArticulo = " + "'" + ZCampo2 + "',"
        ZSql = ZSql + " Estadistica.TipoPro = " + "'" + ZCampo3 + "',"
        ZSql = ZSql + " Estadistica.TipoProDy = " + "'" + ZCampo4 + "',"
        ZSql = ZSql + " Estadistica.ArticuloDy = " + "'" + ZCampo5 + "'"
        ZSql = ZSql + " Where Estadistica.Clave = " + "'" + ZClave + "'"
        
        spEstadistica = ZSql
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
        
    
    
    
    
    
    
    
    
    
    
    
    Erase ZVector
    ZLugar = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM Movvar"
    Sql3 = " Where Movvar.Terminado = " + "'" + Terminado.Text + "'"
    Sql4 = " and Movvar.FechaOrd >= " + "'" + ZFecha + "'"
    Sql5 = " Order by Movvar.Clave"
    spMovvar = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
    
        With rstMovvar
            .MoveFirst
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar) = rstMovvar!Clave
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
            
        End With
        
        rstMovvar.Close
        
    End If
    
    
    For Ciclo = 1 To ZLugar
    
        ZClave = ZVector(Ciclo)
        
        ZCampo1 = "M"
        ZCampo2 = ArticuloDy.Text
        ZCampo3 = "  -     -   "
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Movvar SET "
        ZSql = ZSql + " Movvar.Tipo = " + "'" + ZCampo1 + "',"
        ZSql = ZSql + " Movvar.Articulo = " + "'" + ZCampo2 + "',"
        ZSql = ZSql + " Movvar.Terminado = " + "'" + ZCampo3 + "'"
        ZSql = ZSql + " Where Movvar.Clave = " + "'" + ZClave + "'"
        
        spMovvar = ZSql
        Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
        
    
    
    
    
    
    
    
    
    Erase ZVector
    ZLugar = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM MovLab"
    Sql3 = " Where MovLab.Terminado = " + "'" + Terminado.Text + "'"
    Sql4 = " and MovLab.FechaOrd >= " + "'" + ZFecha + "'"
    Sql5 = " Order by MovLab.Clave"
    spMovlab = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
            .MoveFirst
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar) = rstMovlab!Clave
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
            
        End With
        
        rstMovlab.Close
        
    End If
    
    
    For Ciclo = 1 To ZLugar
    
        ZClave = ZVector(Ciclo)
        
        ZCampo1 = "M"
        ZCampo2 = ArticuloDy.Text
        ZCampo3 = "  -     -   "
    
        ZSql = ""
        ZSql = ZSql + "UPDATE MovLab SET "
        ZSql = ZSql + " MovLab.Tipo = " + "'" + ZCampo1 + "',"
        ZSql = ZSql + " MovLab.Articulo = " + "'" + ZCampo2 + "',"
        ZSql = ZSql + " MovLab.Terminado = " + "'" + ZCampo3 + "'"
        ZSql = ZSql + " Where MovLab.Clave = " + "'" + ZClave + "'"
        
        spMovlab = ZSql
        Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    Erase ZVector
    ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Guia"
    ZSql = ZSql + " Where Guia.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " and Guia.FechaOrd >= " + "'" + ZFecha + "'"
    ZSql = ZSql + " Order by Guia.Clave"
    spGuia = ZSql
    Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
    If rstGuia.RecordCount > 0 Then
    
        With rstGuia
            .MoveFirst
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar) = rstGuia!Clave
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
            
        End With
        
        rstGuia.Close
        
    End If
    
    
    For Ciclo = 1 To ZLugar
    
        ZClave = ZVector(Ciclo)
        
        ZCampo1 = "M"
        ZCampo2 = ArticuloDy.Text
        ZCampo3 = "  -     -   "
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Guia SET "
        ZSql = ZSql + " Guia.Tipo = " + "'" + ZCampo1 + "',"
        ZSql = ZSql + " Guia.Articulo = " + "'" + ZCampo2 + "',"
        ZSql = ZSql + " Guia.Terminado = " + "'" + ZCampo3 + "',"
        ZSql = ZSql + " Guia.PartiOri = Guia.Lote "
        ZSql = ZSql + " Where Guia.Clave = " + "'" + ZClave + "'"
        
        spGuia = ZSql
        Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Erase ZVector
    ZLugar = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM Hoja"
    Sql3 = " Where Hoja.Terminado = " + "'" + Terminado.Text + "'"
    Sql4 = " and Hoja.FechaOrd >= " + "'" + ZFecha + "'"
    Sql5 = " Order by Hoja.Clave"
    spHoja = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
            .MoveFirst
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar) = rstHoja!Clave
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
            
        End With
        
        rstHoja.Close
        
    End If
    
    
    
    For Ciclo = 1 To ZLugar
    
        ZClave = ZVector(Ciclo)
        
        ZCampo1 = "M"
        ZCampo2 = ArticuloDy.Text
        ZCampo3 = "  -     -   "
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + " Hoja.Tipo = " + "'" + ZCampo1 + "',"
        ZSql = ZSql + " Hoja.Articulo = " + "'" + ZCampo2 + "',"
        ZSql = ZSql + " Hoja.Terminado = " + "'" + ZCampo3 + "'"
        ZSql = ZSql + " Where Hoja.Clave = " + "'" + ZClave + "'"
        
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Erase ZVector
    Erase ZVectorII
    ZLugar = 0
    ZRenglon = "1"
    
    Sql1 = "Select *"
    Sql2 = " FROM Hoja"
    Sql3 = " Where Hoja.Producto = " + "'" + Terminado.Text + "'"
    Sql4 = " and Hoja.FechaOrd >= " + "'" + ZFecha + "'"
    Sql5 = " and Hoja.Renglon = " + "'" + ZRenglon + "'"
    Sql6 = " Order by Hoja.Clave"
    spHoja = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
            .MoveFirst
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    If rstHoja!Real <> 0 Then
                    
                        ZLugar = ZLugar + 1
                        ZVector(ZLugar) = Str$(rstHoja!Hoja)
                        ZVectorII(ZLugar, 1) = Str$(rstHoja!Real)
                        ZVectorII(ZLugar, 2) = rstHoja!Fecha
                        
                    End If
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
            
        End With
        
        rstHoja.Close
        
    End If
    
    
    
    For Ciclo = 1 To ZLugar
    
        ZHoja = ZVector(Ciclo)
        ZReal = ZVectorII(Ciclo, 1)
        ZFecha = ZVectorII(Ciclo, 2)
        
        ZCampo1 = Left$(ArticuloDy.Text, 3) + "00" + Right$(ArticuloDy.Text, 7)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + " Hoja.Producto = " + "'" + ZCampo1 + "'"
        ZSql = ZSql + " Where Hoja.Hoja = " + "'" + ZHoja + "'"
        
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        WTipomov = "9"
        WDestino = "9"
    
        spMovguia = "ListaMovguiaNumero " + "'" + WTipomov + "'"
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
            With rstMovguia
                .MoveLast
                    Do
                    WCodigo = Str$(rstMovguia!Codigo + 1)
                    If Val(WCodigo) > 900000 Then
                        .MovePrevious
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstMovguia.Close
                Else
            WCodigo = "1"
        End If
    
    
        ZTipo = "T"
        ZTerminado = Left$(ArticuloDy.Text, 3) + "00" + Right$(ArticuloDy.Text, 7)
        ZArticulo = "  -   -   "
        ZCantidad = ZReal
        ZMovi = "S"
        ZLote = ZHoja
        ZTransito = ""
        ZOrden = ""
        ZDescontar = ""
    
        Auxi1 = WCodigo
        Call Ceros(Auxi1, 6)
        Auxi = "01"
    
        WTipomov = WTipomov
        WDestino = WDestino
        WCodigo = WCodigo
        WRenglon = "1"
        WFecha = ZFecha
        WFechaord = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
        WTipo = ZTipo
        WArticulo = ZArticulo
        WTerminado = ZTerminado
        WCantidad = ZCantidad
        WMovi = ZMovi
        WObservaciones = "Hoja de Produccion Nro. " + ZLote
        WClave = WTipomov + Auxi1 + Auxi
        WDate = Date$
        WMarca = ""
        WPartida = ZLote
        WLote = ""
        WWSaldo = "0"
        WPartiOri = ZLote
        WTransito = ZTransito
        WOrden = ZOrden
        WDescontar = ZDescontar
        
        Sql1 = "INSERT INTO Guia ("
        Sql2 = "Clave ,"
        Sql3 = "TipoMov ,"
        Sql4 = "Codigo ,"
        Sql5 = "Renglon ,"
        Sql6 = "Fecha ,"
        Sql7 = "Tipo ,"
        Sql8 = "Articulo ,"
        Sql9 = "Terminado ,"
        Sql10 = "Cantidad ,"
        Sql11 = "FechaOrd ,"
        Sql12 = "Movi,"
        Sql13 = "Observaciones,"
        Sql14 = "Marca,"
        Sql15 = "Destino,"
        Sql16 = "Lote,"
        Sql17 = "Saldo,"
        Sql18 = "Partida,"
        Sql19 = "PartiOri,"
        Sql20 = "Transito,"
        Sql21 = "Orden,"
        Sql22 = "Descontar )"
        Sql23 = "Values ("
        Sql24 = "'" + WClave + "',"
        Sql25 = "'" + WTipomov + "',"
        Sql26 = "'" + WCodigo + "',"
        Sql27 = "'" + WRenglon + "',"
        Sql28 = "'" + WFecha + "',"
        Sql29 = "'" + WTipo + "',"
        Sql30 = "'" + WArticulo + "',"
        Sql31 = "'" + WTerminado + "',"
        Sql32 = "'" + WCantidad + "',"
        Sql33 = "'" + WFechaord + "',"
        Sql34 = "'" + WMovi + "',"
        Sql35 = "'" + WObservaciones + "',"
        Sql36 = "'" + WMarca + "',"
        Sql37 = "'" + WDestino + "',"
        Sql38 = "'" + WLote + "',"
        Sql39 = "'" + WWSaldo + "',"
        Sql40 = "'" + WPartida + "',"
        Sql41 = "'" + WPartiOri + "',"
        Sql42 = "'" + WTransito + "',"
        Sql43 = "'" + WOrden + "',"
        Sql44 = "'" + WDescontar + "')"
        
        spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 + _
                Sql41 + Sql42 + Sql43 + Sql44
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                
                
                
                
        ZTipo = "M"
        ZTerminado = "  -     -   "
        ZArticulo = ArticuloDy.Text
        ZCantidad = ZReal
        ZMovi = "E"
        ZLote = ZHoja
        ZTransito = ""
        ZOrden = ""
        ZDescontar = ""
    
        Auxi1 = WCodigo
        Call Ceros(Auxi1, 6)
        Auxi = "02"
    
        WTipomov = WTipomov
        WDestino = WDestino
        WCodigo = WCodigo
        WRenglon = "2"
        WFecha = ZFecha
        WFechaord = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
        WTipo = ZTipo
        WArticulo = ZArticulo
        WTerminado = ZTerminado
        WCantidad = ZCantidad
        WMovi = ZMovi
        WObservaciones = "Hoja de Produccion Nro. " + ZLote
        WClave = WTipomov + Auxi1 + Auxi
        WDate = Date$
        WMarca = ""
        WPartida = ""
        WLote = ZLote
        WWSaldo = "0"
        WPartiOri = ZLote
        WTransito = ZTransito
        WOrden = ZOrden
        WDescontar = ZDescontar
    
        Sql1 = "INSERT INTO Guia ("
        Sql2 = "Clave ,"
        Sql3 = "TipoMov ,"
        Sql4 = "Codigo ,"
        Sql5 = "Renglon ,"
        Sql6 = "Fecha ,"
        Sql7 = "Tipo ,"
        Sql8 = "Articulo ,"
        Sql9 = "Terminado ,"
        Sql10 = "Cantidad ,"
        Sql11 = "FechaOrd ,"
        Sql12 = "Movi,"
        Sql13 = "Observaciones,"
        Sql14 = "Marca,"
        Sql15 = "Destino,"
        Sql16 = "Lote,"
        Sql17 = "Saldo,"
        Sql18 = "Partida,"
        Sql19 = "PartiOri,"
        Sql20 = "Transito,"
        Sql21 = "Orden,"
        Sql22 = "Descontar )"
        Sql23 = "Values ("
        Sql24 = "'" + WClave + "',"
        Sql25 = "'" + WTipomov + "',"
        Sql26 = "'" + WCodigo + "',"
        Sql27 = "'" + WRenglon + "',"
        Sql28 = "'" + WFecha + "',"
        Sql29 = "'" + WTipo + "',"
        Sql30 = "'" + WArticulo + "',"
        Sql31 = "'" + WTerminado + "',"
        Sql32 = "'" + WCantidad + "',"
        Sql33 = "'" + WFechaord + "',"
        Sql34 = "'" + WMovi + "',"
        Sql35 = "'" + WObservaciones + "',"
        Sql36 = "'" + WMarca + "',"
        Sql37 = "'" + WDestino + "',"
        Sql38 = "'" + WLote + "',"
        Sql39 = "'" + WWSaldo + "',"
        Sql40 = "'" + WPartida + "',"
        Sql41 = "'" + WPartiOri + "',"
        Sql42 = "'" + WTransito + "',"
        Sql43 = "'" + WOrden + "',"
        Sql44 = "'" + WDescontar + "')"
        
        spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 + _
                Sql41 + Sql42 + Sql43 + Sql44
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
                        
    Erase ZVector
    ZLugar = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM Precios"
    Sql3 = " Where Precios.Terminado = " + "'" + Terminado.Text + "'"
    Sql4 = " Order by Precios.Clave"
    spPrecios = Sql1 + Sql2 + Sql3 + Sql4
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
    
        With rstPrecios
            .MoveFirst
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar) = rstPrecios!Clave
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
            
        End With
        
        rstPrecios.Close
        
    End If
    
    
    For Ciclo = 1 To ZLugar
    
        ZClave = ZVector(Ciclo)
        
        spPrecios = "ConsultaPrecios " + "'" + ZClave + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
        
            ZCliente = rstPrecios!Cliente
            ZTerminado = rstPrecios!Terminado
            ZPrecio = Str$(rstPrecios!Precio)
            ZDescripcion = rstPrecios!Descripcion
            ZFecha = IIf(IsNull(rstPrecios!Fecha), "", rstPrecios!Fecha)
            ZPago = IIf(IsNull(rstPrecios!Pago), "0", rstPrecios!Pago)
    
            ZFecha1 = IIf(IsNull(rstPrecios!Fecha1), "", rstPrecios!Fecha1)
            ZFactura1 = IIf(IsNull(rstPrecios!Factura1), "", rstPrecios!Factura1)
            ZZPrecios = IIf(IsNull(rstPrecios!Precio1), "0", rstPrecios!Precio1)
            ZZCantidad = IIf(IsNull(rstPrecios!Cantidad1), "0", rstPrecios!Cantidad1)
            ZPrecios1 = Str$(ZZPrecios)
            ZCantidad1 = Str$(ZZCantidad)
    
            ZFecha2 = IIf(IsNull(rstPrecios!fecha2), "", rstPrecios!fecha2)
            ZFactura2 = IIf(IsNull(rstPrecios!Factura2), "", rstPrecios!Factura2)
            ZZPrecios = IIf(IsNull(rstPrecios!Precio2), "0", rstPrecios!Precio2)
            ZZCantidad = IIf(IsNull(rstPrecios!Cantidad2), "0", rstPrecios!Cantidad2)
            ZPrecios2 = Str$(ZZPrecios)
            ZCantidad2 = Str$(ZZCantidad)
    
            ZFecha3 = IIf(IsNull(rstPrecios!Fecha3), "", rstPrecios!Fecha3)
            ZFactura3 = IIf(IsNull(rstPrecios!Factura3), "", rstPrecios!Factura3)
            ZZPrecios = IIf(IsNull(rstPrecios!Precio3), "0", rstPrecios!Precio3)
            ZZCantidad = IIf(IsNull(rstPrecios!Cantidad3), "0", rstPrecios!Cantidad3)
            ZPrecios3 = Str$(ZZPrecios)
            ZCantidad3 = Str$(ZZCantidad)
    
            ZFecha4 = IIf(IsNull(rstPrecios!Fecha4), "", rstPrecios!Fecha4)
            ZFactura4 = IIf(IsNull(rstPrecios!Factura4), "", rstPrecios!Factura4)
            ZZPrecios = IIf(IsNull(rstPrecios!Precio4), "0", rstPrecios!Precio4)
            ZZCantidad = IIf(IsNull(rstPrecios!Cantidad4), "0", rstPrecios!Cantidad4)
            ZPrecios4 = Str$(ZZPrecios)
            ZCantidad4 = Str$(ZZCantidad)
    
            ZFecha5 = IIf(IsNull(rstPrecios!Fecha5), "", rstPrecios!Fecha5)
            ZFactura5 = IIf(IsNull(rstPrecios!Factura5), "", rstPrecios!Factura5)
            ZZPrecios = IIf(IsNull(rstPrecios!Precio5), "0", rstPrecios!Precio5)
            ZZCantidad = IIf(IsNull(rstPrecios!Cantidad5), "0", rstPrecios!Cantidad5)
            ZPrecios5 = Str$(ZZPrecios)
            ZCantidad5 = Str$(ZZCantidad)
            
            rstPrecios.Close
            
        End If
        
        ZClave = ZCliente + ArticuloDy.Text
        
        spPreciosMp = "ConsultaPreciosMp " + "'" + ZClave + "'"
        Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
        If rstPreciosMp.RecordCount > 0 Then
        
            XParam = "'" + ZClave + "','" + ZCliente + "','" + ArticuloDy.Text + "','" + ZPrecio + "','" _
                         + ZFecha1 + "','" + ZFactura1 + "','" + ZPrecios1 + "','" + ZCantidad1 + "','" _
                         + ZFecha2 + "','" + ZFactura2 + "','" + ZPrecios2 + "','" + ZCantidad2 + "','" _
                         + ZFecha3 + "','" + ZFactura3 + "','" + ZPrecios3 + "','" + ZCantidad3 + "','" _
                         + ZFecha4 + "','" + ZFactura4 + "','" + ZPrecios4 + "','" + ZCantidad4 + "','" _
                         + ZFecha5 + "','" + ZFactura5 + "','" + ZPrecios5 + "','" + ZCantidad5 + "','" _
                         + Date$ + "','" + ZFecha + "','" + ZPago + "'"
            Set rstPreciosMp = db.OpenRecordset("ModificaPreciosMp " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                
                    Else
        
            XParam = "'" + ZClave + "','" + ZCliente + "','" + ArticuloDy.Text + "','" + ZPrecio + "','" _
                         + ZFecha1 + "','" + ZFactura1 + "','" + ZPrecios1 + "','" + ZCantidad1 + "','" _
                         + ZFecha2 + "','" + ZFactura2 + "','" + ZPrecios2 + "','" + ZCantidad2 + "','" _
                         + ZFecha3 + "','" + ZFactura3 + "','" + ZPrecios3 + "','" + ZCantidad3 + "','" _
                         + ZFecha4 + "','" + ZFactura4 + "','" + ZPrecios4 + "','" + ZCantidad4 + "','" _
                         + ZFecha5 + "','" + ZFactura5 + "','" + ZPrecios5 + "','" + ZCantidad5 + "','" _
                         + Date$ + "','" + ZFecha + "','" + ZPago + "'"
            Set rstPreciosMp = db.OpenRecordset("AltaPreciosMp " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                
        End If
        
    
    Next Ciclo
    
    Next CiclaEmpresa
    
    
    
    
    
    Select Case Val(XEmpresa)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 11
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    
    Terminado.Text = "  -     -   "
    ArticuloDy.Text = "  -   -   "
    
    Terminado.SetFocus
    
End Sub

Private Sub Cancela_click()
    PrgPasaDW.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Terminado.Text = UCase(Terminado.Text)
        ArticuloDy.SetFocus
    End If
End Sub

Private Sub ArticuloDy_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ArticuloDy.Text = UCase(ArticuloDy.Text)
        Terminado.SetFocus
    End If
End Sub


Sub Form_Load()
    Terminado.Text = "  -     -   "
    ArticuloDy.Text = "  -   -   "
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub



