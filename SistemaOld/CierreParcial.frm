VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCierreParcial 
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
Attribute VB_Name = "PrgCierreParcial"
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

Private Sub Cancelar_Click()

    PrgCierreParcial.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Aceptar_Click()

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
                
                aa = !Clave
                
                Renglon = Renglon + 1
                Vector(Renglon, 1) = rstInventario!Tipo
                Vector(Renglon, 2) = rstInventario!Articulo
                Vector(Renglon, 3) = rstInventario!Terminado
                Vector(Renglon, 4) = Str$(rstInventario!Cantidad)
                Vector(Renglon, 5) = rstInventario!Lote
                
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
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgCierreParcial.Caption = "Actualizacion del Stock :  " + !Nombre
        End If
    End With

End Sub
