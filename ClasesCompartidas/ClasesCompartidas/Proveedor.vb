Public Class Proveedor
    Public id As String
    Public razonSocial, direccion, codPostal, localidad, telefono, email, observaciones, cuit, nombreCheque, porceIBProvincia, porceIBCABA, cai, observacionCompleta, cufe1, cufe2, cufe3, diasPlazo, numeroIB, numeroSEDRONAR As String
    Public provincia, region, tipo, codIva, condicionIB1, condicionIB2, categoria, categoriaCalif, tipoInscripcionIB, certificados, estado, calificacion As Nullable(Of Integer)
    Public vtoSEDRONAR, vtoCategoria, vtoCAI, vtoCertificados, vtoCalificacion, vtoCUFE1, vtoCUFE2, vtoCUFE3 As String
    Public cuenta As CuentaContable
    Public rubro As RubroProveedor
    Public estaDefinidoCompleto As Boolean
    Public Sub New(ByVal codigo As String, ByVal nombre As String)
        id = codigo
        razonSocial = nombre
        estaDefinidoCompleto = False
    End Sub

    Public Sub New(ByVal codigo As String, ByVal nombre As String, ByVal dir As String, ByVal codigoPostal As String, ByVal loc As String,
                   ByVal tel As String, ByVal mail As String, ByVal obs1 As String, ByVal claveCUIT As String, ByVal cheque As String,
                   ByVal porceProv As String, ByVal porceCABA As String, ByVal claveCAI As String, ByVal obs2 As String,
                   ByVal cuf1 As String, ByVal cuf2 As String, ByVal cuf3 As String, ByVal prov As Integer, ByVal reg As Integer, ByVal dias As String, ByVal tipoProv As Integer,
                   ByVal iva As Integer, ByVal condicion1IB As Integer, ByVal condicion2IB As Integer, ByVal nroIB As String, ByVal SEDRONAR As String, ByVal cat As Integer,
                   ByVal calificacionCategoria As Integer, ByVal tipoIB As Integer, ByVal certificaciones As Integer, ByVal tipoEstado As Integer, ByVal calif As Integer,
                   ByVal SEDRONARVto As String, ByVal categoriaVto As String, ByVal CAIVto As String, ByVal certificadosVto As String, ByVal calificacionVto As String,
                   ByVal cufe1Vto As String, ByVal cufe2Vto As String, ByVal cufe3Vto As String, ByVal cuentaContable As CuentaContable, ByVal rubroProv As RubroProveedor)
        id = codigo
        razonSocial = nombre
        direccion = dir
        codPostal = codigoPostal
        localidad = loc
        telefono = tel
        email = mail
        observaciones = obs1
        cuit = claveCUIT
        nombreCheque = cheque
        porceIBProvincia = porceProv
        porceIBCABA = porceCABA
        cai = claveCAI
        observacionCompleta = obs2
        cufe1 = cuf1
        cufe2 = cuf2
        cufe3 = cuf3
        provincia = prov
        region = reg
        diasPlazo = dias
        tipo = tipoProv
        codIva = iva
        condicionIB1 = condicion1IB
        condicionIB2 = condicion2IB
        numeroIB = nroIB
        numeroSEDRONAR = SEDRONAR
        categoria = cat
        categoriaCalif = calificacionCategoria
        tipoInscripcionIB = tipoIB
        certificados = certificaciones
        estado = tipoEstado
        calificacion = calif
        vtoSEDRONAR = SEDRONARVto
        vtoCategoria = categoriaVto
        vtoCAI = CAIVto
        vtoCertificados = certificadosVto
        vtoCalificacion = calificacionVto
        vtoCUFE1 = cufe1Vto
        vtoCUFE2 = cufe2Vto
        vtoCUFE3 = cufe3Vto
        cuenta = cuentaContable
        rubro = rubroProv
        estaDefinidoCompleto = True
    End Sub
    
    Public Overrides Function ToString() As String
        Return razonSocial
    End Function
End Class
