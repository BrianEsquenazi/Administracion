VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCierre 
   AutoRedraw      =   -1  'True
   Caption         =   "Actualizacion del Stock"
   ClientHeight    =   6405
   ClientLeft      =   1410
   ClientTop       =   1155
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   ScaleHeight     =   6405
   ScaleWidth      =   9585
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   3135
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Top             =   720
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
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "PrgCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstInventario As Recordset
Dim spInventario As String
Dim rstEntdev As Recordset
Dim spEntdev As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim XParam As String
Private Uno As String
Private Dos As String
Private Tres As String
Private Auxi As String
Private Auxi1 As String
Private Auxi2 As String
Private WArticulo As String
Private WTerminado As String
Private WLote As String
Private WCantidad As String
Private WLiberada As String
Private WMarca As String
Private WSaldo As String
Dim Vector(5000, 10) As String
Dim WVector(50000) As String
Dim WVectorII(50000) As String

Dim WDesde1 As String
Dim WHasta1 As String
Dim WDesde2 As String
Dim WHasta2 As String
Dim WDesde3 As String
Dim WHasta3 As String
Dim WDesde4 As String
Dim WHasta4 As String
Dim WDesde5 As String
Dim WHasta5 As String
Dim WDesde6 As String
Dim WHasta6 As String
Dim WDesde7 As String
Dim WHasta7 As String
Dim WDesde8 As String
Dim WHasta8 As String

Private Sub Cancelar_Click()

    PrgCierre.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Aceptar_Click()

    Rem HERNAN ACTUALIZACION
    Rem PARAMETROS DE MATERIA PRIMA

Rem  WDesde1 = "AA-000-000"
 Rem WHasta1 = "XX-030-999"
     
Rem    WDesde2 = "BL-000-000"
Rem    WHasta2 = "BS-999-999"
    
Rem     WDesde5 = "EC-000-100"
Rem     WHasta5 = "SR-999-100"
    
Rem     WDesde6 = "WA-067-100"
 Rem    WHasta6 = "XV-999-999"
    
 Rem    WDesde7 = "DA-005-100"
 Rem    WHasta7 = "DQ-410-100"
        
  Rem  WDesde2 = ""
  Rem  WHasta2 = ""
        
   Rem WDesde5 = ""
   Rem WHasta5 = ""
        
  Rem  WDesde6 = ""
  Rem  WHasta6 = ""
        
  Rem  WDesde7 = ""
  Rem  WHasta7 = ""
        
  Rem  WDesde8 = "CD-020-100"
  Rem  WHasta8 = "CM-000-100"

    Rem PARAMETROS DE PRODUCTO TERMINADO
    
    WDesde3 = "NK-25024-000"
    WHasta3 = "NK-25024-777"
    
  Rem   WDesde4 = "RE-05106-000"
  Rem  WHasta4 = "RE-25301-999"
 Rem   WDesde4 = "RE-00000-000"
 Rem   WHasta4 = "RE-99999-999"



Rem procesa el total de las materias

    Rem HERNAN ACTUALIZACION
    Rem PARAMETROS DE MATERIA PRIMA

  Rem  WDesde1 = "DY-000-000"
  Rem  WHasta1 = "DY-999-999"
    
  Rem  WDesde2 = "CO-000-000"
  Rem  WHasta2 = "CO-999-999"
    
 Rem   WDesde5 = "DS-000-000"
 Rem   WHasta5 = "DS-999-999"
    
  Rem  WDesde6 = "DK-000-000"
  Rem  WHasta6 = "DK-999-999"
    
  Rem  WDesde7 = "DQ-000-000"
  Rem  WHasta7 = "DQ-999-999"
        
  Rem  WDesde8 = "DW-000-000"
  Rem  WHasta8 = "DW-999-999"

    Rem PARAMETROS DE PRODUCTO TERMINADO
    
 Rem   WDesde3 = "NK-08164-100"
 Rem   WHasta3 = "NK-08164-100"
   
 Rem   WDesde4 = "NK-25056-777"
 Rem   WHasta4 = "NK-25056-777"

    Call Valida_fecha(Fecha.Text, Auxi)
    If Auxi = "N" Then
        m$ = "La fecha de cierre informada no es valida"
        ca% = MsgBox(m$, 0, "Actualizacion de Stock")
        Exit Sub
    End If
    
    T$ = "Actualizacion de Stock"
    m$ = "!!! ATENCION !!!   Se actualizara el stock con el recuento informado, Desea realizar el proceso      "
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% <> 6 Then
        Exit Sub
    End If

    WFecha = Fecha.Text
    WOrdFecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
    XParam = "'" + WFecha + "','" _
                 + WOrdFecha + "'"
                 
    Rem spArticulo = "ModificaArticuloFecha " + XParam
    Rem Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem spArticulo = "ModificaArticuloInicial0"
    Rem Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem spTerminado = "ModificaTerminadoFecha " + XParam
    Rem Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem spTerminado = "ModificaTerminadoInicial0"
    Rem Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    
    
    

    Rem Procesa las materias primas

    Erase WVector
    Lugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Order by Codigo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
    
        With rstArticulo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                
                WArticulo = IIf(IsNull(rstArticulo!Codigo), "", rstArticulo!Codigo)
                WArticulo = UCase(WArticulo)
                
                Pasa = "N"
                
                If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                    Pasa = "S"
                End If
                
                If Pasa = "S" Then
                    Lugar = Lugar + 1
                    WVector(Lugar) = rstArticulo!Codigo
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstArticulo.Close
    End If
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Articulo SET "
        ZSql = ZSql & "FechaCierre = '" + WFecha + "' ,"
        ZSql = ZSql & "OrdFechaCierre = '" + WOrdFecha + "' ,"
        ZSql = ZSql & "Inicial = 0 ,"
        ZSql = ZSql & "Laboratorio = 0 ,"
        ZSql = ZSql & "Pedido = 0 "
        ZSql = ZSql & " Where Codigo = " + "'" + WVector(Ciclo) + "'"
                
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    

    
    
    
    
    
    
    
    
    
    
    

    Rem Procesa las productos terminados

    Erase WVector
    Lugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Terminado"
    ZSql = ZSql + " Order by Codigo"
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
        With rstTerminado
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                
                WTerminado = IIf(IsNull(rstTerminado!Codigo), "", rstTerminado!Codigo)
                XCodigo = Val(Mid$(WTerminado, 4, 5))
                
                Pasa = "N"
                
                If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                    Pasa = "S"
                End If
                If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                    Pasa = "S"
                End If
                
                If Pasa = "S" Then
                    Lugar = Lugar + 1
                    WVector(Lugar) = rstTerminado!Codigo
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstTerminado.Close
    End If
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Terminado SET "
        ZSql = ZSql & "FechaCierre = '" + WFecha + "' ,"
        ZSql = ZSql & "OrdFechaCierre = '" + WOrdFecha + "' ,"
        ZSql = ZSql & "Inicial = 0 ,"
        ZSql = ZSql & "Proceso = 0 "
        ZSql = ZSql & " Where Codigo = " + "'" + WVector(Ciclo) + "'"
                
        spTerminado = ZSql
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
    
    
    
    
    

    
    
    
    
    
    
    
    
    
    

    Rem spLaudo = "ModificaLaudoMarca"
    Rem Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem spHoja = "ModificaHojaMarca"
    Rem Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem spMovguia = "ModificaMovguiaMarca"
    Rem Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem spMovvar = "ModificaMovvarMarca"
    Rem Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem spMovlab = "ModificaMovlabMarca"
    Rem Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem spEstadistica = "ModificaEstadisticaMarca"
    Rem Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem spEntdev = "ModificaEntdevMarca"
    Rem Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
        
    
    
    
    

    Rem Procesa los Laudos

    Erase WVector
    Lugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Order by Laudo"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
        With rstLaudo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                
                WArticulo = IIf(IsNull(rstLaudo!Articulo), "", rstLaudo!Articulo)
                ZSaldo = rstLaudo!Saldo
                
                Pasa = "N"
                
                If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                    Pasa = "S"
                End If
                
                If Pasa = "S" Then
                    WMarca = IIf(IsNull(rstLaudo!Marca), "", rstLaudo!Marca)
                    If Trim(WMarca) = "" Or ZSaldo <> 0 Then
                        Lugar = Lugar + 1
                        WVector(Lugar) = rstLaudo!Clave
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
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Laudo SET "
        ZSql = ZSql & "Marca = 'X' ,"
        ZSql = ZSql & "Saldo = 0 ,"
        ZSql = ZSql & "Liberada = 0 ,"
        ZSql = ZSql & "Devuelta = 0"
        ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo) + "'"
                
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
    
    
    
    
    Rem Procesa las Hojas
    
    Erase WVector
    Erase WVectorII
    Lugar = 0


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Order by Hoja"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                
                WTerminado = IIf(IsNull(rstHoja!Producto), "", rstHoja!Producto)
                XCodigo = Val(Mid$(WTerminado, 4, 5))
                ZSaldo = rstHoja!Saldo
                
                Pasa = "N"
                
                If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                    Pasa = "S"
                End If
                If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                    Pasa = "S"
                End If
                
                If Pasa = "S" Then
                    WMarca = IIf(IsNull(rstHoja!Marca), "", rstHoja!Marca)
                    If Trim(WMarca) = "" Or ZSaldo <> 0 Then
                        Lugar = Lugar + 1
                        WVector(Lugar) = rstHoja!Clave
                        If rstHoja!Tipo = "M" Then
                            WVectorII(Lugar) = rstHoja!Articulo
                                Else
                            WVectorII(Lugar) = rstHoja!Terminado
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
        rstHoja.Close
    End If
    
    For Ciclo = 1 To Lugar
    
    
        If Len(WVectorII(Ciclo)) = 10 Then
        
            WArticulo = WVectorII(Ciclo)
            
            Pasa = "N"
            
            If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                Pasa = "S"
            End If
            If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                Pasa = "S"
            End If
            If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                Pasa = "S"
            End If
            If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                Pasa = "S"
            End If
            If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                Pasa = "S"
            End If
            If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                Pasa = "S"
            End If
            
            
            
                Else
                
            WTerminado = WVectorII(Ciclo)
            
            Pasa = "N"
            
            If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                Pasa = "S"
            End If
            If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                Pasa = "S"
            End If
            
        End If
            
        If Pasa = "S" Then
    
            ZSql = ""
            ZSql = ZSql & "UPDATE Hoja SET "
            ZSql = ZSql & "Marca = 'X' ,"
            ZSql = ZSql & "Saldo = 0 ,"
            ZSql = ZSql & "Real = 0"
            ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo) + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZSql = ""
            ZSql = ZSql & "UPDATE Hoja SET "
            ZSql = ZSql & "Saldo = 0 ,"
            ZSql = ZSql & "Real = 0"
            ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo) + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                
        End If
    
    Next Ciclo
    
    
    
    
    
    
    
    
    Rem Procesa las guias
    
    
    Erase WVector
    Lugar = 0

    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Guia"
    ZSql = ZSql + " Order by Codigo"
    spMovguia = ZSql
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                
                WTipo = IIf(IsNull(rstMovguia!Tipo), "", rstMovguia!Tipo)
                WArticulo = IIf(IsNull(rstMovguia!Articulo), "  -   -   ", rstMovguia!Articulo)
                WTerminado = IIf(IsNull(rstMovguia!Terminado), "  -     -   ", rstMovguia!Terminado)
                ZSaldo = rstMovguia!Saldo
                
                If WTipo = "M" Then
                    
                    Pasa = "N"
                    
                    If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                        Pasa = "S"
                    End If
                    
                        Else
                
                    Pasa = "N"
                    
                    If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                        Pasa = "S"
                    End If
                    If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                        Pasa = "S"
                    End If
                    
                End If
                        
                If Pasa = "S" Then
                    WMarca = IIf(IsNull(rstMovguia!Marca), "", rstMovguia!Marca)
                    If Trim(WMarca) = "" Or ZSaldo <> 0 Then
                        Lugar = Lugar + 1
                        WVector(Lugar) = rstMovguia!Clave
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
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Guia SET "
        ZSql = ZSql & "Marca = 'X' ,"
        ZSql = ZSql & "Saldo = 0 ,"
        ZSql = ZSql & "Cantidad = 0"
        ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo) + "'"
                
        spMovguia = ZSql
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem Procesa las Movimientos Varios
    
    Erase WVector
    Lugar = 0

    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Movvar"
    ZSql = ZSql + " Order by Codigo"
    spMovvar = ZSql
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
    
        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                
                WTipo = IIf(IsNull(rstMovvar!Tipo), "", rstMovvar!Tipo)
                WArticulo = IIf(IsNull(rstMovvar!Articulo), "  -   -   ", rstMovvar!Articulo)
                WTerminado = IIf(IsNull(rstMovvar!Terminado), "  -     -   ", rstMovvar!Terminado)
                
                If WTipo = "M" Then
                
                    Pasa = "N"
                
                    If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                        Pasa = "S"
                    End If
                    
                        Else
                
                    Pasa = "N"
                    
                    If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                        Pasa = "S"
                    End If
                    If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                        Pasa = "S"
                    End If
                    
                End If
                        
                If Pasa = "S" Then
                    WMarca = IIf(IsNull(rstMovvar!Marca), "", rstMovvar!Marca)
                    If Trim(WMarca) = "" Then
                        Lugar = Lugar + 1
                        WVector(Lugar) = rstMovvar!Clave
                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstMovvar.Close
    End If
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Movvar SET "
        ZSql = ZSql & "Marca = 'X' "
        ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo) + "'"
                
        spMovvar = ZSql
        Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem Procesa las Movimientos de Laboratorio
    
    Erase WVector
    Lugar = 0

    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Movlab"
    ZSql = ZSql + " Order by Codigo"
    spMovlab = ZSql
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                
                WTipo = IIf(IsNull(rstMovlab!Tipo), "", rstMovlab!Tipo)
                WArticulo = IIf(IsNull(rstMovlab!Articulo), "  -   -   ", rstMovlab!Articulo)
                WTerminado = IIf(IsNull(rstMovlab!Terminado), "  -     -   ", rstMovlab!Terminado)
                
                If WTipo = "M" Then
                
                    Pasa = "N"
                        
                    If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                        Pasa = "S"
                    End If
                    
                        Else
                
                    Pasa = "N"
                    
                    If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                        Pasa = "S"
                    End If
                    If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                        Pasa = "S"
                    End If
                    
                End If
                        
                If Pasa = "S" Then
                    WMarca = IIf(IsNull(rstMovlab!Marca), "", rstMovlab!Marca)
                    If Trim(WMarca) = "" Then
                        Lugar = Lugar + 1
                        WVector(Lugar) = rstMovlab!Clave
                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstMovlab.Close
    End If
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Movlab SET "
        ZSql = ZSql & "Marca = 'X' "
        ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo) + "'"
                
        spMovlab = ZSql
        Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem Procesa las estadisticas
    
    Erase WVector
    Lugar = 0

    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Order by Numero"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                WTerminado = IIf(IsNull(rstEstadistica!Articulo), "", rstEstadistica!Articulo)
                WArticulo = Left$(WTerminado, 3) + Right$(WTerminado, 7)
                XCodigo = Val(Mid$(WTerminado, 4, 5))
                
                If Left$(WTerminado, 2) = "PT" Or Left$(WTerminado, 2) = "YQ" Or Left$(WTerminado, 2) = "YF" Or Left$(WTerminado, 2) = "YP" Then
                
                    Pasa = "N"
                    
                    If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                        Pasa = "S"
                    End If
                    If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                        Pasa = "S"
                    End If
                    
                        Else
                
                    Pasa = "N"
                    
                    If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                        Pasa = "S"
                    End If
                
                End If
                        
                If Pasa = "S" Then
                    WMarca = IIf(IsNull(rstEstadistica!Marca), "", rstEstadistica!Marca)
                    If Trim(WMarca) = "" Then
                        Lugar = Lugar + 1
                        WVector(Lugar) = IIf(IsNull(rstEstadistica!Clave), "", rstEstadistica!Clave)
                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstEstadistica.Close
    End If
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Estadistica SET "
        ZSql = ZSql & "Marca = 'X' "
        ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo) + "'"
                
        spEstadistica = ZSql
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem Procesa las entradas de devoluciones
    
    Erase WVector
    Lugar = 0

    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EntDev"
    ZSql = ZSql + " Order by Codigo"
    spEntdev = ZSql
    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    If rstEntdev.RecordCount > 0 Then
    
        With rstEntdev
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                
                WTerminado = rstEntdev!Terminado
                WArticulo = Left$(rstEntdev!Terminado, 3) + Right$(rstEntdev!Terminado, 7)
                XCodigo = Val(Mid$(WTerminado, 4, 5))
       Rem by nan 22-11 AGREGO  nk Y RE
                If Left$(WTerminado, 2) = "PT" Or Left$(WTerminado, 2) = "RE" Or Left$(WTerminado, 2) = "NK" Or Left$(WTerminado, 2) = "YQ" Or Left$(WTerminado, 2) = "YF" Or Left$(WTerminado, 2) = "YP" Then
                
                    Pasa = "N"
                    
                    If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                        Pasa = "S"
                    End If
                    If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                        Pasa = "S"
                    End If
                    
                        Else
                
                    Pasa = "N"
                    
                    If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                        Pasa = "S"
                    End If
                
                End If
                        
                If Pasa = "S" Then
                    WMarca = IIf(IsNull(rstEntdev!Marca), "", rstEntdev!Marca)
                    If Trim(WMarca) = "" Then
                        Lugar = Lugar + 1
                        WVector(Lugar) = rstEntdev!Clave
                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstEntdev.Close
    End If
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE EntDev SET "
        ZSql = ZSql & "Marca = 'X' "
        ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo) + "'"
                
        spEntdev = ZSql
        Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Erase Vector
    Renglon = 0

    spInventario = "ListaInventarioTotal"
    Set rstInventario = db.OpenRecordset(spInventario, dbOpenSnapshot, dbSQLPassThrough)
    If rstInventario.RecordCount > 0 Then
        
    With rstInventario
    
        .MoveFirst
        
        If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WTipo = rstInventario!Tipo
                WArticulo = rstInventario!Articulo
                WTerminado = rstInventario!Terminado
Rem BY NAN 23-11 AGREGO NK Y RE
                If Left$(WTerminado, 2) = "SE" Or Left$(WTerminado, 2) = "RE" Or Left$(WTerminado, 2) = "PT" Or Left$(WTerminado, 2) = "NK" Or Left$(WTerminado, 2) = "YQ" Or Left$(WTerminado, 2) = "YF" Or Left$(WTerminado, 2) = "YP" Then
                
                    Pasa = "N"
                    
                    If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                        Pasa = "S"
                    End If
                    If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                        Pasa = "S"
                    End If
                    
                        Else
                
                    Pasa = "N"
                    
                    If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                        Pasa = "S"
                    End If
                
                End If
                        
                If Pasa = "S" Then
                
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstInventario!Tipo
                    Vector(Renglon, 2) = rstInventario!Articulo
                    Vector(Renglon, 3) = rstInventario!Terminado
                    Vector(Renglon, 4) = Str$(rstInventario!Cantidad)
                    Vector(Renglon, 5) = rstInventario!Lote
                    
                End If
                
                .MoveNext
                        
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
    End With
            
    rstInventario.Close
            
    End If
    
    
    For Ciclo = 1 To Renglon
    
                WTipo = Vector(Ciclo, 1)
                WArticulo = Vector(Ciclo, 2)
                WTerminado = Vector(Ciclo, 3)
                WCantidad = Vector(Ciclo, 4)
                WLote = Vector(Ciclo, 5)
                
                If WTipo = "M" Then
    
                    Rem procesa la materia prima
                
                    WEntra = "N"
                    WControla = 0
                    
                    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                        rstArticulo.Close
                    End If
            
                    If WControla = 0 Then
            
                        XParam = "'" + WLote + "','" _
                                    + WArticulo + "'"
                        spLaudo = "ListaLaudoArticulo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                        
                            WClave = rstLaudo!Clave
                            WDate = Date$
                            WLiberada = Str$(rstLaudo!Liberada + Val(WCantidad))
                            WSaldo = Str$(rstLaudo!Saldo + Val(WCantidad))
                            WMarca = ""
                            
                            WEntra = "S"
                            rstLaudo.Close
                            
                            XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "','" _
                                    + WLiberada + "','" _
                                    + WMarca + "'"
                                    
                            spLaudo = "ModificaLaudoSaldoCierre " + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                        End If
                
                        If WEntra = "N" Then
                        
                            XParam = "'" + WArticulo + "','" _
                                    + WLote + "'"
                            spMovguia = "ListaMovguiaLote " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                            
                                WClave = rstMovguia!Clave
                                WCanti = Str$(rstMovguia!Cantidad + Val(WCantidad))
                                WSaldo = Str$(rstMovguia!Saldo + Val(WCantidad))
                                WMarca = ""
                            
                                WEntra = "S"
                                rstMovguia.Close
                                
                                XParam = "'" + WClave + "','" _
                                        + WSaldo + "','" _
                                        + WCanti + "','" _
                                        + WMarca + "'"
                                    
                                spMovguia = "ModificaMovguiaSaldoCierre " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                
                            End If
                        End If
                    
                    End If
                    
                                
                                Else
                                
    
                    Rem procesa el producto terminado
                    
                    WEntra = "N"
                    WControla = 0
                    
                    spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        rstTerminado.Close
                    End If
            
                    If WControla = 0 Then
                        XParam = "'" + WLote + "','" _
                                + WTerminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                        
                            WClave = rstHoja!Clave
                            WReal = Str$(rstHoja!Real + Val(WCantidad))
                            WSaldo = Str$(rstHoja!Saldo + Val(WCantidad))
                            WMarca = ""
                            
                            WEntra = "S"
                            rstHoja.Close
                            
                            XParam = "'" + WClave + "','" _
                                        + WSaldo + "','" _
                                        + WReal + "','" _
                                        + WMarca + "'"
                                        
                            spHoja = "ModificaHojaSaldoCierre " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            
                        End If
                        
                        If WEntra = "N" Then
                        
                            XParam = "'" + WTerminado + "','" _
                                    + WLote + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                            
                                WClave = rstMovguia!Clave
                                WCanti = Str$(rstMovguia!Cantidad + Val(WCantidad))
                                WSaldo = Str$(rstMovguia!Saldo + Val(WCantidad))
                                WMarca = ""
                            
                                WEntra = "S"
                                rstMovguia.Close
                                
                                XParam = "'" + WClave + "','" _
                                        + WSaldo + "','" _
                                        + WCanti + "','" _
                                        + WMarca + "'"
                                        
                                spMovguia = "ModificaMovguiaSaldoCierre " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                        End If
                    
                    End If
                    
                End If
                
    Next Ciclo
    
    Call Cancelar_Click

End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgCierre.Caption = "Actualizacion del Stock :  " + !Nombre
        End If
    End With

End Sub
