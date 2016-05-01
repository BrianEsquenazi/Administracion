VERSION 5.00
Begin VB.Form PrgGraba1 
   Caption         =   "Lectura de Prestamos entre Plantas"
   ClientHeight    =   4620
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4620
   ScaleWidth      =   6390
   Begin VB.Frame Frame2 
      Caption         =   "Control de Grabacion"
      Height          =   1815
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   4215
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cerrar"
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "PrgGraba1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPrestamo As Recordset
Dim spPrestamo As String
Dim XParam As String
Dim WClave As String
Dim WCodigo As String
Dim WRenglon As String
Dim WFecha As String
Dim OrdFecha As String
Dim WObservaciones As String
Dim WTipo As String
Dim WArticulo As String
Dim WTerminado As String
Dim WCantidad As String
Dim WCosto As String
Dim WDestino As String
Dim WTermino As String

Private Sub Acepta_Click()

    Call Proceso
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    PrgGraba1.Hide
    Unload Me
    End
End Sub

Private Sub Proceso()

    Rem On Error GoTo Error

    OPEN_FILE_Presta
    
    Pasa = 0

    With rstPresta
            .Index = "Codigo"
            .MoveFirst
            Do
            
                If Pasa = 0 Then
                    Pasa = 1
                    Corte = !Codigo
                    Renglon = 0
                End If
                If Corte <> !Codigo Then
                    Corte = !Codigo
                    Renglon = 0
                End If
                
                Renglon = Renglon + 1
                
                XFecha = !Fecha
                
                XLug = 0
                WTermino = ""
                
                For Ciclo = 1 To 12
                    If Mid$(XFecha, Ciclo, 1) = "/" Then
                        Call Ceros(WTermino, 2)
                        XLug = XLug + 1
                        Select Case XLug
                            Case 1
                                WFecha = WTermino + "/"
                            Case Else
                                WFecha = WFecha + WTermino + "/"
                                WAno = Mid$(XFecha, Ciclo + 1, 2)
                                If Val(WAno) > 60 Then
                                    WFecha = WFecha + "19" + WAno
                                        Else
                                    WFecha = WFecha + "20" + WAno
                                End If
                                Exit For
                        End Select
                        WTermino = ""
                    End If
                    If Mid$(XFecha, Ciclo, 1) <> "/" Then
                        WTermino = WTermino + Mid$(XFecha, Ciclo, 1)
                    End If
                Next Ciclo
            
                WCodigo = Str$(!Codigo)
                WRenglon = Str$(Renglon)
                WOrdFecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                WObservaciones = ""
                WTipo = !Tipo
                WArticulo = !Articulo
                WTerminado = !Terminado
                WCantidad = Str$(!Cantidad)
                WCosto = Str$(!Costo)
                WDestino = ""
                
                Call Ceros(WCodigo, 6)
                Call Ceros(WRenglon, 2)
                
                WClave = WCodigo + WRenglon
                
                
                
                Rem spPrestamo = "ConsultaPrestamo " + "'" + WCodigo + "'"
                Rem Set rstPrestamo = db.OpenRecordset(spPrestamo, dbOpenSnapshot, dbSQLPassThrough)
                Rem If rstPrestamo.RecordCount = 0 Then
                    XParam = "'" + WClave + "','" _
                         + WCodigo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WOrdFecha + "','" _
                         + WObservaciones + "','" _
                         + WTipo + "','" _
                         + WArticulo + "','" _
                         + WTerminado + "','" _
                         + WCantidad + "','" _
                         + WCosto + "','" _
                         + WDestino + "'"
                                         
                    Set rstPrestamo = db.OpenRecordset("AltaPrestamo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                Rem End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    Exit Sub
    
Error:
     coderr = Err
     Resume Next
     
End Sub




