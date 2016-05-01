VERSION 5.00
Begin VB.Form PrgPaso4 
   Caption         =   "Recepcion de Ordenes de Compra"
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
Attribute VB_Name = "PrgPaso4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()

    Call Proceso
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    PrgRecep.Hide
    Unload Me
    Menu.SetFocus
End Sub


Private Sub Proceso()

    On Error GoTo Error
    
    
   Rem ORDENES DE COMPRA
        
    coderr = 0
    With rstWOrden
            .Index = "Clave"
            .MoveFirst
            Do
                
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
                
                With rstOrden
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




