Redondeo:
            B = x * 100
            B1 = x * 10000
            VALOR = INT(B)
            Redondeo = INT(B1)
            REDONDEO1 = INT(B) * 100
            DIFE = Redondeo - REDONDEO1
            IF DIFE >= 50 THEN VALOR = VALOR + 1
            x = VALOR / 100
            RETURN



