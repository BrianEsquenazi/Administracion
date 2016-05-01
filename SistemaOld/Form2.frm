VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form2"
   ScaleHeight     =   6375
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox IValor20 
      Height          =   285
      Left            =   0
      MaxLength       =   70
      TabIndex        =   9
      Text            =   " "
      Top             =   3240
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.TextBox IValor19 
      Height          =   285
      Left            =   0
      MaxLength       =   70
      TabIndex        =   8
      Text            =   " "
      Top             =   2880
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.TextBox IValor18 
      Height          =   285
      Left            =   0
      MaxLength       =   70
      TabIndex        =   7
      Text            =   " "
      Top             =   2520
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.TextBox IValor17 
      Height          =   285
      Left            =   0
      MaxLength       =   70
      TabIndex        =   6
      Text            =   " "
      Top             =   2160
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.TextBox IValor16 
      Height          =   285
      Left            =   0
      MaxLength       =   70
      TabIndex        =   5
      Text            =   " "
      Top             =   1800
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.TextBox IValor15 
      Height          =   285
      Left            =   0
      MaxLength       =   70
      TabIndex        =   4
      Text            =   " "
      Top             =   1440
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.TextBox IValor14 
      Height          =   285
      Left            =   0
      MaxLength       =   70
      TabIndex        =   3
      Text            =   " "
      Top             =   1080
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.TextBox IValor13 
      Height          =   285
      Left            =   0
      MaxLength       =   70
      TabIndex        =   2
      Text            =   " "
      Top             =   720
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.TextBox IValor12 
      Height          =   285
      Left            =   0
      MaxLength       =   70
      TabIndex        =   1
      Text            =   " "
      Top             =   360
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.TextBox IValor11 
      Height          =   285
      Left            =   0
      MaxLength       =   70
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   5040
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Rem dada
    Rem dada
    Rem dada
    
    If CertificadoNo.Value = 1 Then
        WCertificado1 = "0"
    End If
    If CertificadoSi.Value = 1 Then
        WCertificado1 = "1"
    End If
    
    If EstadoNo.Value = 1 Then
        WEstado1 = "0"
    End If
    If EstadoSi.Value = 1 Then
        WEstado1 = "1"
    End If
    
    ZVencimiento = Vencimiento.Text
    WClave = ""
       
    ZSql = ""
    ZSql = ZSql + "Select * "
    ZSql = ZSql + " FROM Informe"
    ZSql = ZSql + " Where Informe = " + "'" + Informe.Text + "'"
    ZSql = ZSql + " and Orden = " + "'" + Orden.Text + "'"
    ZSql = ZSql + " and Articulo = " + "'" + Producto.Text + "'"
    spInforme = ZSql
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        WClave = rstInforme!Clave
        rstInforme.Close
    End If
    
    ZVencimiento = Vencimiento.Text
    ZOrdVencimiento = Right$(ZVencimiento, 4) + Mid$(ZVencimiento, 4, 2) + Left$(ZVencimiento, 2)
        
    ZSql = ""
    ZSql = ZSql + "UPDATE Informe SET "
    ZSql = ZSql + "Certificado1 = " + "'" + WCertificado1 + "',"
    ZSql = ZSql + "Certificado2 = " + "'" + Certificado2.Text + "',"
    ZSql = ZSql + "Estado1 = " + "'" + WEstado1 + "',"
    ZSql = ZSql + "Estado2 = " + "'" + Estado2.Text + "',"
    ZSql = ZSql + "FechaVencimiento = " + "'" + ZVencimiento + "',"
    ZSql = ZSql + "FechaElaboracion = " + "'" + FechaElaboracion.Text + "',"
    ZSql = ZSql + "OrdFechaVencimiento = " + "'" + ZOrdVencimiento + "'"
    ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
    
    spInforme = ZSql
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select * "
    ZSql = ZSql + " FROM laudo"
    ZSql = ZSql + " Where Informe = " + "'" + Informe.Text + "'"
    ZSql = ZSql + " and Orden = " + "'" + Orden.Text + "'"
    ZSql = ZSql + " and Articulo = " + "'" + Producto.Text + "'"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        WClave = rstLaudo!Clave
        rstLaudo.Close
    End If
    ZFechaElaboracion = FechaElaboracion.Text



    ZSql = ""
    ZSql = ZSql + "UPDATE Laudo SET "
    ZSql = ZSql + "FechaVencimiento = " + "'" + ZVencimiento + "',"
    ZSql = ZSql + "FechaElaboracion = " + "'" + ZFechaElaboracion + "'"
    ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"

    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    
    












                    WProveedor = !Proveedor
                    WTipo = !Tipo
                    WLetra = !Letra
                    WPunto = !Punto
                    WNumero = !Numero
                    WFecha = !Fecha
                    Wvencimiento = !Vencimiento
                    WPeriodo = !Periodo
                    WNeto = !Neto
                    WIva21 = !Iva21
                    WIva5 = !Iva5
                    WIva27 = !Iva27
                    WIb = !Ib
            
                    ZIva105 = IIf(IsNull(!Iva105), "0", !Iva105)
                    WIva27 = WIva27 + ZIva105
                    
                    WExento = !Exento
                    WImpre = !Impre
                    WOrdFecha = !OrdFecha
                    WContado = !Contado
                    XEmpresa = !Empresa
                    WNroInterno = !NroInterno
                    ZSoloIva = IIf(IsNull(rstIvaComp!SoloIva), "0", rstIvaComp!SoloIva)
                    If ZSoloIva = 1 Then
                        WNeto = 0
                    End If
                
                
                WProveedor = !Proveedor
                WNombre = ""
                WCuit = ""
                
                spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    WNombre = RstProveedor!Nombre
                    WCuit = RstProveedor!Cuit
                    RstProveedor.Close
                End If
                
                !Nombre = WNombre
                !Cuit = WCuit
                
                
                
                    With rstIva
                        .AddNew
                        !NroInterno = WNroInterno
                        !Proveedor = WProveedor
                        !Tipo = WTipo
                        !Letra = WLetra
                        !Punto = WPunto
                        !Numero = WNumero
                        !Fecha = WFecha
                        !Vencimiento = Wvencimiento
                        !Periodo = WPeriodo
                        !Concepto = WConcepto
                        !Neto = WNeto
                        !Iva21 = WIva21
                        !Iva5 = WIva5
                        !Iva27 = WIva27
                        !Ib = WIb
                        !Exento = WExento
                        !Impre = WImpre
                        !OrdFecha = WOrdFecha
                        !Contado = WContado
                        !Empresa = XEmpresa
                        !Titulo = WTitulo
                        !TituloII = WTituloII
                        .Update
                    End With
                    
    
    
    
    
        
