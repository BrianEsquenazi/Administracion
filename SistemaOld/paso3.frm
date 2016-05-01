VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPaso3 
   Caption         =   "Traspaso de Ordenes de Compra"
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
      Begin MSMask.MaskEdBox Fecha 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3120
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cerrar"
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
   End
End
Attribute VB_Name = "PrgPaso3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()

    WFectraspaso = Mid$(Fecha.Text, 4, 2) + "-" + Left$(Fecha.Text, 2) + "-" + Right$(Fecha.Text, 4)
    Call Proceso
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Prgtraspa.Hide
    Unload Me
    Menu.SetFocus
End Sub


Sub Form_Load()
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
End Sub


Private Sub Proceso()

    On Error GoTo Error
    
    
    'Orden de compra
    
    coderr = 0
    With rstWOrden
        .Index = "Clave"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
   Rem ORDENES DE COMPRA
        
    coderr = 0
    With rstOrden
            .Index = "Clave"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                
                    WOrden = !Orden
                    Wrenglon = !Renglon
                    WFecha = !Fecha
                    WFechaord = !FechaOrd
                    WProveedor = !Proveedor
                    WArticulo = !Articulo
                    WCantidad = !Cantidad
                    WPrecio = !Precio
                    WFecha1 = !Fecha1
                    WFecha2 = !Fecha2
                    WCondicion = !Condicion
                    WRecibida = !Recibida
                    WClave = !Clave
                
                    With rstWOrden
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Orden = WOrden
                            !Renglon = Wrenglon
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Proveedor = WProveedor
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Precio = WPrecio
                            !Fecha1 = WFecha1
                            !Fecha2 = WFecha2
                            !Condicion = WCondicion
                            !Recibida = WRecibida
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Orden = WOrden
                            !Renglon = Wrenglon
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Proveedor = WProveedor
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Precio = WPrecio
                            !Fecha1 = WFecha1
                            !Fecha2 = WFecha2
                            !Condicion = WCondicion
                            !Recibida = WRecibida
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                    End With
                End If
                
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




