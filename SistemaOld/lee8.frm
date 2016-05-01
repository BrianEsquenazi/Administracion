VERSION 5.00
Begin VB.Form Prglee8 
   AutoRedraw      =   -1  'True
   Caption         =   "grabacion de cotizaciones"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.CommandButton Acepta 
      Caption         =   "Acepta"
      Height          =   300
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancela"
      Height          =   300
      Left            =   2880
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "Prglee8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Auxi1 As String
Private Auxi2 As String

Private Sub Acepta_Click()

    mas = 6400

    With rstCotiza1
            .Index = "Clave"
            .MoveFirst
            Do
                WClave = !Clave
                WCotiza = !Cotiza
                WRenglon = !Renglon
                WFecha = !Fecha
                WProveedor = !Proveedor
                WArticulo = !Articulo
                WPrecio = !Precio
                WCondicion = !Condicion
                WObservaciones = !Observaciones
                WFechaord = !FechaOrd
                WDate = !WDate
                    
                With rstCotiza
                        .Index = "Clave"
                        .AddNew
                        !Cotiza = WCotiza + mas
                        !Renglon = WRenglon
                        !Fecha = WFecha
                        !Proveedor = WProveedor
                        !Articulo = WArticulo
                        !Precio = WPrecio
                        !FechaOrd = WFechaord
                        !Condicion = WCondicion
                        !Observaciones = WObservaciones
                        Auxi1 = !Cotiza
                        Call Ceros(Auxi1, 6)
                        Auxi2 = !Renglon
                        Call Ceros(Auxi2, 2)
                        !Clave = Auxi1 + Auxi2
                        !WDate = WDate
                        .Update
                End With
                        
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    With rstCotiza1
        .Close
    End With
    With rstCotiza
        .Close
    End With
    DbsCotiza.Close
    Prglee8.Hide
    Menu.Show
End Sub

