VERSION 5.00
Begin VB.Form Prglee20 
   Caption         =   "Generacion de traspaso de datos"
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
Attribute VB_Name = "Prglee20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WLinea As String
Private WImporte As Double
Private WTeorico As Double
Private WReal As Double
Private WCantidad As Double
Private WCotiza As String * 6

Private Sub Acepta_Click()
 
    Call Proceso
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Prglee20.Hide
    Unload Me
    Menu.SetFocus
End Sub


Private Sub Proceso()

    On Error GoTo Error
    
    
    Open "a:" + WEmpresa + "cot.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
   
                XCotiza = Mid$(WLinea, 1, 6)
                Select Case Val(WEmpresa)
                    Case 3
                        WCotiza = Str$(Val(Mid$(WLinea, 1, 6)) + 6500)
                    Case 4
                        WCotiza = Str$(Val(Mid$(WLinea, 1, 6)) + 6800)
                    Case 5
                        WCotiza = Str$(Val(Mid$(WLinea, 1, 6)) + 7000)
                    Case Else
                End Select
                Call Ceros(WCotiza, 6)
                WRenglon = Mid$(WLinea, 7, 2)
                WFecha = Mid$(WLinea, 13, 2) + "/" + Mid$(WLinea, 11, 2) + "/19" + Mid$(WLinea, 9, 2)
                WProveedor = Mid$(WLinea, 15, 11)
                WArticulo = Mid$(WLinea, 26, 2) + "-" + Mid$(WLinea, 28, 3) + "-" + Mid$(WLinea, 31, 3)
                WPrecio = Val(Mid$(WLinea, 34, 8))
                WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                WCondicion = Mid$(WLinea, 42, 15)
                WObservaciones = Mid$(WLinea, 57, 20)
                WClave = WCotiza + WRenglon
                
                With rstCotiza
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Cotiza = WCotiza
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Proveedor = WProveedor
                            !Articulo = WArticulo
                            !Precio = WPrecio
                            !FechaOrd = WFechaord
                            !Condicion = WCondicion
                            !Observaciones = WObservaciones
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Cotiza = WCotiza
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Proveedor = WProveedor
                            !Articulo = WArticulo
                            !Precio = WPrecio
                            !FechaOrd = WFechaord
                            !Condicion = WCondicion
                            !Observaciones = WObservaciones
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1
                
    
   


















    Call Cancela_click


    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Exit Sub
    
    
Error:
Stop
     coderr = Err
     Resume Next
     
End Sub




