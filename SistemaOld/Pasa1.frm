VERSION 5.00
Begin VB.Form PrgPasa1 
   Caption         =   "Trasposo de Clientes"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
End
Attribute VB_Name = "PrgPasa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancelar_Click()
    With rstClientes
        .Close
    End With
    DbsVentas.Close
    PrgPasa1.Hide
    Menu.Show
End Sub

Private Sub Aceptar_Click()

    Open "A:" + WEmpresa + "clie.txt" For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, Linea
        
        Cliente = Mid$(Linea, 1, 6)
        Razon = Mid$(Linea, 8, 40)
        Direccion = Mid$(Linea, 49, 40)
        Localidad = Mid$(Linea, 90, 40)
        Postal = Mid$(Linea, 131, 4)
        Telefono = Mid$(Linea, 138, 15)
        Contacto = ""
        Observaciones = ""
        Cuit = Mid$(Linea, 156, 15)
        Vendedor = Val(Mid$(Linea, 182, 4))
        email = ""
        fax = ""
        Rubro = Val(Mid$(Linea, 177, 4))
        Horario = Mid$(Linea, 253, 8)
        Pago1 = Val(Mid$(Linea, 172, 4))
        pago2 = Val(Mid$(Linea, 232, 4))
        Limite = 0
        Minimo = 0
        DirEntrega = Mid$(Linea, 191, 40)
        Provincia = "1"
        Iva = "1"
        
        With rstClientes
        
            .Index = "Cliente"
            .Seek "=", Cliente
            If .NoMatch Then
                .AddNew
                !Cliente = Cliente
                !Razon = Razon
                !Direccion = Direccion
                !Localidad = Localidad
                !Postal = Postal
                !Telefono = Telefono
                !Contacto = Contacto
                !Observaciones = Observaciones
                !Cuit = Cuit
                !Vendedor = Vendedor
                !email = email
                !fax = fax
                !Rubro = Rubro
                !Horario = Horario
                !Pago1 = Pago1
                !pago2 = pago2
                !Limite = Limite
                !Minimo = Minimo
                !DirEntrega = DirEntrega
                !Provincia = Provincia
                !Iva = "1"
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.

End Sub


