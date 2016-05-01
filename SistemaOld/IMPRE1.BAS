Impre2784:

        IF Rete2784# = 0 THEN RETURN

        SELECT CASE RecordEmp.Impresora
                CASE "2"
                        LPRINT CHR$(27) + CHR$(38) + CHR$(107) + CHR$(50) + CHR$(83);
                CASE ELSE
                        LPRINT CHR$(15);
        END SELECT

        M# = Rete2784#
        Mes% = VAL(MID$(RecordOr4.Fecha, 3, 2))
        Cuatri$ = ""

        IF Mes% <= 4 THEN
           Cuatri$ = "Primer Cuatrimestre"
        ELSE
            IF Mes% >= 5 AND Mes% <= 8 THEN
                  Cuatri$ = "Segundo Cuatrimestre"
            ELSE
                   IF Mes% >= 9 THEN
                         Cuatri$ = "Tercer Cuatrimestre"
                   END IF
            END IF
        END IF

        EmpNombre$ = RecordEmp.Nombre
        EmpDireccion$ = RecordEmp.Direccion
        EmpLocalidad$ = RecordEmp.Localidad
        EmpCuit$ = RecordEmp.Cuit

        FOR Ciclo% = 1 TO 2

                IF Ciclo% = 1 THEN
                        Nombre.Empre$ = EmpNombre$
                                ELSE
                        Nombre.Empre$ = RecordPrv.Nombre
                END IF

            LPRINT
            LPRINT "                                                     COMPROBANTE DE RETENCION"
            LPRINT "                                                IMPUESTO A LAS GANACIAS RG 2784                            "
            LPRINT EmpNombre$
            LPRINT EmpDireccion$
            LPRINT EmpLocalidad$
            LPRINT "Clave Unica de Identificacion Tributaria : "; EmpCuit$
            LPRINT "----------------------------------------------------------------------------------------------------------------------------------"
            LPRINT "SUJETO RETENIDO                                                                               |"
            LPRINT "                                                                                              |"
            LPRINT USING "Nombre/Razon Social : \                              \                                        |"; RecordPrv.Nombre
            LPRINT USING "Domicilio           : \                              \                                        |"; RecordPrv.Direccion
            LPRINT USING "Clave Unica de Identificacion Tributaria : \               \                                  |"; RecordPrv.Cuit
            LPRINT "..............................................................................................|"
            LPRINT "                                                                                              |"
            LPRINT "DETALLE DE LA RETENCION                                                                       |"
            LPRINT "                                                                                              |"
            LPRINT USING "Concepto de la Retencion : \                              \                                   |"; Leyenda$(VAL(RecordPrv.pago))
            LPRINT USING "Importe Pagado           : ###,###,###,###.##                                                 |"; Total# - Rete2784# - Rete3125#
            LPRINT USING "Importe Retenido         : ###,###,###,###.##                                                 |"; Rete2784#
            LPRINT "                                                                                              |----------------------------------"
            LPRINT ".................................................................................................................................."
            LPRINT
            LPRINT USING "La Presente Retencion efectuada el \        \ se informara en la Declaracion Jurada del \                  \ ."; fnimpredate$(FNREVDATE$(RecordOr4.Fecha)); Cuatri$
            LPRINT ""
            LPRINT "=================================================================================================================================="
        NEXT Ciclo%
        LPRINT CHR$(12)
        RETURN

