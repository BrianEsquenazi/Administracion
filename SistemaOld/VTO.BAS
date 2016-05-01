      DATA 31,28,31,30,31,30,31,31,30,31,30,31
      DATA "0101","0105","2505", , ,"0907", ,"1210", ,"2512", , , , , ,
Calcula.Fecha:

44170 Dg% = 0
44180    An$ = MID$(WFECHA$, 5, 2): An# = VAL(An$)
44190    Me$ = MID$(WFECHA$, 3, 2): Me% = VAL(Me$)
44200    Di$ = MID$(WFECHA$, 1, 2): Di% = VAL(Di$)
44210 'CANTIDAD DE DIAS HASTA LA FECHA
44220    Dg% = Di% + PLAZO% - 1
44230    FOR Aa = Me% TO 12
44240        IF (An# MOD 4 = 0) AND Me% = 2 THEN Ds(2) = 29 ELSE Ds(2) = 28
44250        IF Dg% <= Ds(Aa) THEN 44300
44260        Dg% = Dg% - Ds(Aa)
44270    NEXT Aa
44280    An# = An# + 1
44290    Me% = 1: GOTO 44230
44300 Di% = Dg%: Di$ = RIGHT$("0" + MID$(STR$(Di%), 2, LEN(STR$(Di%)) - 1), 2)
44310 Me% = Aa: Me$ = RIGHT$("0" + MID$(STR$(Me%), 2, LEN(STR$(Me%)) - 1), 2)
44320 An$ = RIGHT$("0" + MID$(STR$(An#), 2, LEN(STR$(An#)) - 1), 2)
44330 FVENC$ = Di$ + Me$ + An$
44340 RETURN

