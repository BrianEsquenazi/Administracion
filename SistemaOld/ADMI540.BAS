DECLARE SUB Ingreso (Campo$, Tipo$, SKIP$, Largo%, NumDec%, Fila%, columna%, COLFOR%, COLBAC%, Forma$, A$)
DECLARE SUB VidPop (pant$)

'靈컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴캠
'튞OMBRE DEL prvGRAMA            : admi540                                  �
'튦ITULO DEL prvGRAMA            : ORDENES DE PAGO                          �
'튔ECHA                          : 03.10.90                                 �
'聃컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴캭

	COMMON WMenu%, WChoice%, WEmpresa$

	REM $INCLUDE: 'create.dcl'
	REM $INCLUDE: 'FBan.dcl'
	REM $INCLUDE: 'FIvc.dcl'
	REM $INCLUDE: 'FOr3.dcl'
	REM $INCLUDE: 'FOr4.dcl'
	REM $INCLUDE: 'FCtc.dcl'
	REM $INCLUDE: 'FPrv.dcl'
	REM $INCLUDE: 'FPto.dcl'
	REM $INCLUDE: 'FRet.dcl'
	REM $INCLUDE: 'FEmp.dcl'
	REM $INCLUDE: 'FTip.dcl'
	REM $INCLUDE: 'FObr.dcl'
	REM $INCLUDE: 'FCon.dcl'

	XMenu% = WMenu%
	XChoice% = WChoice%

	''wempresa$ = "0002"

	DIM FR$(20)
	DIM DS(20)
	DIM Vector$(8, 13), CALCDAY%(12)
	DIM M$(2), M%(12)
	DIM TX$(20), T$(20), IP(20), Leyenda$(14), Porce$(3, 2)
	DIM Mes$(12)
	DIM Rete$(3, 10)
	DIM Desc$(20)
	DIM Elije$(100, 5)
	DIM XObra$(20, 5)
	DIM Dato$(10, 2)
	DIM Error$(100)

	REM $INCLUDE: 'PRNUSING.FN'
	REM $INCLUDE: 'REDONDEO.FN'
	REM $INCLUDE: 'IMPREDAT.FN'
	REM $INCLUDE: 'REVDATE.FN'

	DEF FNFECHA$ (X$) = RIGHT$(X$, 2) + MID$(X$, 3, 2) + LEFT$(X$, 2)
	DEF FNDFECHA$ (X$) = RIGHT$(X$, 2) + "/" + MID$(X$, 3, 2) + "/" + LEFT$(X$, 2)
	DEL$ = STRING$(75, 32)

	Leyenda$(1) = "Ejercicio Prof.Lib. c/Aj.Inf."
	Leyenda$(2) = "Alquileres y Arrendamientos"
	Leyenda$(3) = "Compra de Bienes"

	OpOr3% = 28
	Wst% = FOr3R%
	CALL Errores(Error$(), Wst%, "")
	CLOSE

	Lmsg% = 24
	Cmsg% = 5

	OPEN "R", #31, "TABGAN", 128
	FIELD #31, 8 AS T$(1), 8 AS T$(2), 8 AS T$(3), 8 AS T$(4), 8 AS T$(5), 8 AS T$(6), 8 AS T$(7), 8 AS T$(8), 8 AS T$(9), 8 AS T$(10), 8 AS T$(11), 8 AS T$(12), 8 AS T$(13), 8 AS T$(14)

	GET #31, 1
	DATA "10","13.5","17","20.5","24"

	FOR Ciclo% = 1 TO 5
		READ TX$(Ciclo%)
	NEXT Ciclo%

	FOR I% = 1 TO 4
	    IP(I%) = CVD(T$(I% + 3)) / 100
	NEXT I%

	'CLOSE #31

	CPO$ = "######,###.##"
	M% = VAL(LEFT$(DATE$, 2))

	FOR J% = 1 TO 11
		READ Desc$(J%)
	NEXT J%

	DATA "Fac.","N.D.","N.C.","Apl.","C.Re","Vari","Tran","Ant.","Rem","Fac","CHR"

	WIDTH "Lpt1:", 255

ABRO:

	EmpresaBan$ = WEmpresa$
	Wst% = FBanO%
	CALL Errores(Error$(), Wst%, BanMsg$)

	EmpresaOr3$ = WEmpresa$
	Wst% = FOr3O%
	CALL Errores(Error$(), Wst%, Or3Msg$)

	EmpresaOr4$ = WEmpresa$
	Wst% = FOr4O%
	CALL Errores(Error$(), Wst%, Or4Msg$)

	EmpresaCtc$ = WEmpresa$
	Wst% = FCtcO%
	CALL Errores(Error$(), Wst%, CtcMsg$)

	EmpresaPrv$ = WEmpresa$
	Wst% = FPrvO%
	CALL Errores(Error$(), Wst%, PrvMsg$)

	EmpresaPto$ = WEmpresa$
	Wst% = FPtoO%
	CALL Errores(Error$(), Wst%, PtoMsg$)

	EmpresaRet$ = WEmpresa$
	Wst% = FRetO%
	CALL Errores(Error$(), Wst%, RetMsg$)

	EmpresaIvc$ = WEmpresa$
	Wst% = FIvcO%
	CALL Errores(Error$(), Wst%, IvcMsg$)

	EmpresaTip$ = WEmpresa$
	Wst% = FTipO%
	CALL Errores(Error$(), Wst%, TipMsg$)

	EmpresaObr$ = WEmpresa$
	Wst% = FObrO%
	CALL Errores(Error$(), Wst%, ObrMsg$)

	EmpresaCon$ = WEmpresa$
	Wst% = FConO%
	CALL Errores(Error$(), Wst%, ConMsg$)

	Wst% = FEmpO%
	CALL Errores(Error$(), Wst%, EmpMsg$)

	OpEmp% = 5
	ClaveEmp$ = WEmpresa$
	Wst% = FEmpR
	IF Wst% = 0 THEN
		NomEmpresa$ = LEFT$(RecordEmp.Nombre, 20)
			ELSE
		NomEmpresa$ = WEmpresa$
	END IF

	OPEN "R", #99, WEmpresa$ + "Tabl.dat", 50
	FIELD #99, 2 AS Codtabla$, 40 AS Impretabla1$, 3 AS Impretabla2$, 1 AS Tipotabla$

cabecera:

	LOCATE 1, 1
	VidPop "admi540"
	LOCATE 1, 1: CALL QPRINT("Empresa : " + NomEmpresa$, 120)


	LOCATE 23, 2: CALL QPRINT("  F1 = Menu Principal   F2 = Cancela    F3 = Campo Anterior", 11)
	Dato% = 1
	Opcion$ = ""

	WHILE Dato% <> 6
105         ON Dato% GOSUB ORDEN.PAGO, Fecha.Orden, Tipo, Proveedor, Banco.Cabecera
107         IF VAL(A$) = 1 THEN GOTO Fin
		IF VAL(A$) = 2 THEN
			GOSUB Limpio
			GOTO cabecera
		END IF
		IF VAL(A$) = 3 AND Dato% <> 1 THEN Dato% = Dato% - 2
		IF Dato% <= 0 THEN Dato% = 4
	WEND

	GOSUB ACEPTO
	IF VAL(A$) = 0 THEN
		IF Opcion$ = "N" THEN
			Dato% = 1
			GOTO 105
		END IF
				ELSE
		Dato% = 10
		GOTO 107
	END IF

DETALLE:

	Linea% = 1
	columna% = 1

INGRESO.DATOS:

	A$ = ""
	WHILE VAL(A$) <> 1
		LOCATE 23, 2: CALL QPRINT(" ," + CHR$(25) + ",�컴,컴� = MOVIMIENTOS     F1 = FIN INGRESO    F10 = MENU PRINCIPAL", 11)
		SELECT CASE columna%

			CASE 1
				Vector$(Linea%, columna%) = ""
			CASE 2
				IF VAL(WTipo$) <> 5 THEN
					DO
						CALL Ingreso(Vector$(Linea%, columna%), "N", "", 6, 0, Linea% + 7, 14, 7, 0, "######", A$)
						IF VAL(A$) = 1 OR VAL(A$) = 10 OR VAL(A$) = 11 OR VAL(A$) = 12 THEN EXIT DO
					LOOP UNTIL Vector$(Linea%, columna%) <> ""
							 ELSE
					Vector$(Linea%, columna%) = WOrden$
					LOCATE Linea% + 7, 14: CALL QPRINT(Vector$(Linea%, columna%), 7)
				END IF

			CASE 3
				SELECT CASE VAL(WTipo$)
					CASE 1
						LOCATE 1, 1
						VidPop "FILL PAGE 1"

						LOCATE 1, 1
						VidPop "admi540B"

						DO
							CALL Ingreso(Vector$(Linea%, columna%), "N", "", 2, 0, Linea% + 7, 22, 7, 0, "##", A$)
							IF VAL(Vector$(Linea%, columna%)) = 10 THEN EXIT DO
						LOOP UNTIL ((VAL(Vector$(Linea%, columna%)) >= 1 AND VAL(Vector$(Linea%, columna%)) <= 7) OR VAL(Vector$(Linea%, columna%)) = 11) AND VAL(Vector$(Linea%, columna%)) <> 5

						LOCATE 1, 1
						VidPop "DISPLAY PAGE 1"

						LOCATE Linea% + 7, 22: CALL QPRINT(LEFT$(Desc$(VAL(Vector$(Linea%, columna%))), 4), 7)

					CASE 2
						Vector$(Linea%, columna%) = "07"
						LOCATE Linea% + 7, 22: CALL QPRINT(LEFT$(Desc$(VAL(Vector$(Linea%, columna%))), 4), 7)

					CASE 3, 9
						Vector$(Linea%, columna%) = "06"
						LOCATE Linea% + 7, 22: CALL QPRINT(LEFT$(Desc$(VAL(Vector$(Linea%, columna%))), 4), 7)

					CASE 4
						Vector$(Linea%, columna%) = "05"
						LOCATE Linea% + 7, 22: CALL QPRINT(LEFT$(Desc$(VAL(Vector$(Linea%, columna%))), 4), 7)

					CASE 5
						Vector$(Linea%, columna%) = "08"
						LOCATE Linea% + 7, 22: CALL QPRINT(LEFT$(Desc$(VAL(Vector$(Linea%, columna%))), 4), 7)

					CASE 6
						LOCATE 1, 1
						VidPop "FILL PAGE 1"

						LOCATE 1, 1
						VidPop "admi540B"

						DO
							CALL Ingreso(Vector$(Linea%, columna%), "N", "", 2, 0, Linea% + 7, 22, 7, 0, "##", A$)
							IF VAL(Vector$(Linea%, columna%)) = 10 THEN EXIT DO
						LOOP UNTIL (VAL(Vector$(Linea%, columna%)) = 9 OR VAL(Vector$(Linea%, columna%)) = 4)

						LOCATE 1, 1
						VidPop "DISPLAY PAGE 1"

						LOCATE Linea% + 7, 22: CALL QPRINT(LEFT$(Desc$(VAL(Vector$(Linea%, columna%))), 4), 7)

					CASE ELSE
				END SELECT

				IF (VAL(WTipo$) = 1 AND VAL(Vector$(Linea%, 3)) <> 6) OR VAL(WTipo$) = 6 THEN
					ClaveCtc$ = WProveedor$ + Vector$(Linea%, 3) + Vector$(Linea%, 2) + "00"
					OpCtc% = 5
					Wst% = FCtcR
					IF Wst% <> 0 THEN
						CALL Errores(Error$(), Wst%, CtcMsg$)
						columna% = columna% - 2
								 ELSE
						FOR Contador% = 1 TO 8
							IF Contador% <> Linea% AND VAL(Vector$(Contador%, 2)) <> 0 AND Vector$(Contador%, 2) = Vector$(Linea%, 2) AND Vector$(Contador%, 3) = Vector$(Linea%, 3) THEN
								columna% = columna% - 3
								LOCATE Lmsg%, Cmsg%: PRINT SPACE$(70);
								LOCATE Lmsg%, Cmsg%: PRINT "Factura ya ingresada a al orden de pago  ";
								TECLA$ = INPUT$(1)
								LOCATE Lmsg%, Cmsg%: PRINT SPACE$(70);
								Vector$(Linea%, 1) = ""
								Vector$(Linea%, 2) = ""
								Vector$(Linea%, 3) = ""
							END IF
						NEXT Contador%
						IF VAL(Vector$(Linea%, 5)) = 0 THEN
							IF VAL(Vector$(Linea%, 3)) = 10 THEN
								Vector$(Linea%, 5) = STR$(RecordCtc.Importe * 100)
								Vector$(Linea%, 11) = STR$(RecordCtc.Saldo * 100)
									ELSE
								Vector$(Linea%, 5) = STR$(RecordCtc.Saldo * 100)
								Vector$(Linea%, 11) = ""
							END IF
						END IF
						Vector$(Linea%, 12) = RecordCtc.Emision
						Vector$(Linea%, 13) = RecordCtc.Vencimiento
						RecordPrv.pago = RecordIvc.retencion
					END IF
				END IF

			CASE 4
				Vector$(Linea%, columna%) = LEFT$(Vector$(Linea%, columna%), 30)
				CALL Ingreso(Vector$(Linea%, columna%), "F", "", 30, 0, Linea% + 7, 27, 7, 0, "", A$)

			CASE 5
				WIngreso$ = STR$(ABS(VAL(Vector$(Linea%, 5))))
				CALL Ingreso(WIngreso$, "N", "", 11, 2, Linea% + 7, 67, 7, 0, "######,###.##", A$)
				Vector$(Linea%, 5) = WIngreso$

				IF (VAL(WTipo$) = 1 OR VAL(WTipo$) = 6) AND VAL(Vector$(Linea%, 3)) <> 6 THEN
					COMPARA1# = INT(RecordCtc.Saldo)
					IF VAL(Vector$(Linea%, 3)) = 10 THEN

						LOCATE 1, 1
						VidPop "Fill Page 1"
						LOCATE 1, 1
						VidPop "admi540g"

						XDada$ = STR$(ABS(VAL(Vector$(Linea%, 11))))
						CALL Ingreso(XDada$, "N", "", 11, 2, 11, 49, 7, 0, "######,###.##", A$)
						Vector$(Linea%, 11) = XDada$

						LOCATE 1, 1
						VidPop "Display Page 1"

						COMPARA2# = INT(VAL(XDada$) / 100)

							ELSE

						COMPARA2# = INT(VAL(WIngreso$) / 100)

					END IF

					IF COMPARA1# < COMPARA2# THEN
						LOCATE Lmsg%, Cmsg%: PRINT SPACE$(70);
						LOCATE Lmsg%, Cmsg%: PRINT "Importe mayor al saldo de la factura ";
						TECLA$ = INPUT$(1)
						LOCATE Lmsg%, Cmsg%: PRINT SPACE$(70);
						Vector$(Linea%, 5) = ""
						LOCATE Linea% + 7, 67: PRINT USING "######,###.##"; VAL(Vector$(Linea%, columna%))
						columna% = 0
					END IF
				END IF

				Totfac# = 0

				FOR I% = 1 TO 8
					IF VAL(Vector$(I%, 3)) = 3 OR VAL(Vector$(I%, 3)) = 4 THEN
						Totfac# = (Totfac# - VAL(Vector$(I%, 5)) / 100)
												  ELSE
						Totfac# = (Totfac# + VAL(Vector$(I%, 5)) / 100)
					END IF
				NEXT I%

			CASE 6
				IF VAL(WTipo$) = 3 THEN

					LOCATE 1, 1
					VidPop "Fill Page 1"
					LOCATE 1, 1
					VidPop "admi540C"
					DO
						CALL Ingreso(Vector$(Linea%, 10), "F", "", 13, 0, 11, 44, 13, 0, "", A$)

						'OpCue% = 5
						'ClaveCue$ = Vector$(Linea%, 10)
						'Wst% = FCueR
						Wst% = 0
						IF Wst% <> 0 THEN
							'CALL Errores(Error$(), Wst%, CueMsg$)
								ELSE
							'LOCATE 12, 44: PRINT RecordCue.Descripcion
							WConfirma$ = ""
							DO
								CALL Ingreso(WConfirma$, "F", "", 1, 0, 13, 44, 13, 0, "", A$)
							LOOP UNTIL WConfirma$ = "S" OR WConfirma$ = "N"
							IF WConfirma$ = "N" THEN
								Wst% = 99
							END IF
						END IF
					LOOP UNTIL Wst% = 0

					LOCATE 1, 1
					VidPop "Display Page 1"

				END IF

			CASE ELSE
		END SELECT

		SELECT CASE VAL(A$)
			CASE 12
				'CLS
				'PRINT LINEA%, columna%
				'tecla$ = INPUT$(1)
				IF Linea% <> 1 THEN
					IF columna% = 2 THEN
						Linea% = Linea% - 1
					END IF
				END IF
			CASE 11
				IF Linea% <> 8 THEN
					IF columna% = 2 THEN
						Linea% = Linea% + 1
					END IF
				END IF
			CASE 14, 0
				IF columna% <> 6 THEN
					columna% = columna% + 1
							  ELSE
					IF Linea% <> 8 THEN
						Linea% = Linea% + 1
						columna% = 1
								ELSE
						Linea% = 1
						columna% = 1
					END IF
				END IF
			CASE 13
				IF columna% <> 1 THEN
					columna% = columna% - 1
				END IF
			CASE 10
				GOSUB Fin
			CASE ELSE
		END SELECT
	WEND

	Dato% = 1: Opcion$ = ""

	WImporte$ = ""
	WObservaciones$ = ""

200   WHILE Dato% <> 3
		ON Dato% GOSUB Importe.Pago, Observaciones
		IF VAL(A$) = 1 THEN GOTO Fin
		IF VAL(A$) = 2 THEN
			GOSUB Limpio.Linea
			GOSUB Limpio
			Dato% = 1
			GOTO cabecera
		END IF
		IF VAL(A$) = 3 AND Dato% <> 1 THEN Dato% = Dato% - 2
		IF Dato% <= 0 THEN Dato% = 1
	 WEND

	 GOSUB ACEPTO

	 IF VAL(A$) = 0 THEN
		IF Opcion$ = "N" THEN
			Dato% = 1
			GOTO cabecera
		END IF
	 END IF

	GOSUB Calcula.retencion
	GOSUB GRABO.CARGA
	GOSUB IMPREORDEN
	GOSUB Impre2784
	GOSUB Impre3125
	GOSUB Limpio
	GOTO Fin

ACEPTO:

	Opcion$ = "S"
	DO
		LOCATE 24, 5: CALL QPRINT("Acepta Datos [ S/N ] ", 7)
		CALL Ingreso(Opcion$, "A", "", 1, 0, 24, 27, 7, 0, "", A$)
		LOCATE 24, 5: CALL QPRINT(STRING$(50, 32), 7)
	LOOP UNTIL ASC(Opcion$) = 83 OR ASC(Opcion$) = 78
	RETURN

ORDEN.PAGO:

	IF VAL(WOrden$) = 0 THEN
		OpOr4% = 13
		ClaveOr4$ = SPACE$(20)
		Wst% = FOr4R
		IF Wst% <> 0 THEN
			WOrden$ = "1"
					 ELSE
			WOrden$ = STR$(VAL(RecordOr4.ORDEN) + 1)
		END IF
	END IF
	CALL Ingreso(WOrden$, "N", "", 6, 0, 3, 9, 7, 0, "######", A$)
	IF VAL(A$) = 0 THEN
		OpOr4% = 5
		ClaveOr4$ = WOrden$
		Wst% = FOr4R
		IF Wst% = 0 THEN
			Wst% = 5
			CALL Errores(Error$(), Wst%, CueMsg$)
					ELSE
			Dato% = Dato% + 1
		END IF
	END IF
	RETURN

Fecha.Orden:
	DO
		IF VAL(WFecha$) = 0 THEN
			WFecha$ = MID$(DATE$, 4, 2) + MID$(DATE$, 1, 2) + RIGHT$(DATE$, 2)
		END IF

		DO
			Fecha$ = WFecha$
			CALL Ingreso(Fecha$, "Y", "", 6, 0, 3, 23, 7, 0, "", A$)
			CALL Validate(Fecha$, Mensaje$)
		LOOP UNTIL Mensaje$ = "Yes" OR VAL(A$) <> 0
		WFecha$ = Fecha$

		WFEC$ = RIGHT$(WFecha$, 4)

		Compara$ = FNREVDATE$(Fecha$)

	LOOP UNTIL Compara$ > XFECHA$

	Dato% = Dato% + 1
	RETURN

Tipo:

	LOCATE 1, 1
	VidPop "Fill Page 1"
	LOCATE 1, 1
	VidPop "admi540Z"
	DO
		CALL Ingreso(WTipo$, "N", "", 1, 0, 3, 38, 7, 0, "#", A$)
	LOOP UNTIL (VAL(WTipo$) > 0 AND VAL(WTipo$) < 6) OR VAL(A$) = 1 OR VAL(A$) = 2 OR VAL(A$) = 3
	LOCATE 1, 1
	VidPop "Display Page 1"
	LOCATE 3, 38: CALL QPRINT(WTipo$, 7)

	Dato% = Dato% + 1

Proveedor:

	IF WTipo$ <> "1" AND WTipo$ <> "5" AND WTipo$ <> "6" AND WTipo$ <> "9" THEN
		WProveedor$ = SPACE$(6)
		LOCATE 3, 51: PRINT SPACE$(20)
		LOCATE 4, 51: PRINT SPACE$(20)
		Dato% = Dato% + 1
		RETURN
	END IF

	DO
		CALL Ingreso(WProveedor$, "N", "", 6, 0, 3, 51, 7, 0, "######", A$)

		IF VAL(WProveedor$) = 0 THEN
			XDATO% = 1
			GOSUB Selecciona.Datos
			WProveedor$ = Codigo$
			A$ = ""
			COLOR 13
			LOCATE 3, 51: PRINT USING "######"; VAL(WProveedor$)
			COLOR 7
		END IF

		OpPrv% = 5
		KynPrv% = 0
		ClavePrv$ = WProveedor$
		IF VAL(A$) = 0 THEN
			Wst% = FPrvR
			IF Wst% <> 0 THEN
				CALL Errores(Error$(), Wst%, PrvMsg$)
					ELSE
				LOCATE 4, 51: PRINT LEFT$(RecordPrv.Nombre, 20)
				OpTip% = 5
				ClaveTip$ = RecordPrv.Tipo
				Wst% = FTipR
				IF Wst% <> 0 THEN
					RecordTip.Cuenta = SPACE$(13)
				END IF
			END IF
		END IF
	LOOP UNTIL Wst% = 0 OR VAL(A$) <> 0
	Dato% = Dato% + 1
	RETURN

Banco.Cabecera:

	IF VAL(WTipo$) <> 2 THEN
		Dato% = Dato% + 1
		RETURN
	END IF

	DO
		CALL Ingreso(XBanco$, "N", "", 2, 0, 4, 9, 7, 0, "##", A$)
		IF VAL(A$) <> 0 THEN EXIT DO
		OpBan% = 5
		ClaveBan$ = XBanco$
		Wst% = FBanR
		IF Wst% <> 0 THEN
			CALL Errores(Error$(), Wst%, BanMsg$)
		END IF
	LOOP UNTIL Wst% = 0 OR VAL(A$) <> 0
	Cuenta$(2) = RecordBan.Cuenta
	WIMpre$ = RecordBan.Nombre

	Dato% = Dato% + 1
	RETURN

Importe.Pago:

	A$ = ""
	WImporte$ = STR$(Totfac# * 100)
	LOCATE 17, 67: PRINT USING "######,###.##"; Totfac#
	Total# = Totfac#
	Dato% = Dato% + 1
	RETURN

Observaciones:

	CALL Ingreso(WObservaciones$, "F", "", 50, 0, 19, 3, 7, 0, "", A$)
	Dato% = Dato% + 1
	RETURN

CONFIRMA.DATOS:

	Comfirma$ = "S"
	CALL Ingreso(CONFIRMA$, "G", "", 1, 0, 16, 28, 7, 0, "", A$)
	RETURN

Limpio:

	WOrden$ = ""
	WFecha$ = ""
	WProveedor$ = ""
	WFechaVto$ = ""
	WConcepto$ = ""
	WConcepto1$ = ""
	XBanco$ = ""
	WTipo$ = ""
	WBanco$ = ""
	WNroCheque$ = ""
	WDespacho$ = ""
	FACTU# = 0
	FACTU$ = ""
	Neto# = 0
	Neto$ = ""
	pago# = 0
	ERASE Cuenta$
	ERASE Vector$
	RETURN

Limpio.Linea:

	WImporte$ = ""
	WObservaciones$ = ""
	Efectivo$ = ""
	LOCATE 17, 26: CALL QPRINT(SPACE$(13), 7)
	LOCATE 17, 51: CALL QPRINT(SPACE$(10), 7)
	LOCATE 17, 67: CALL QPRINT(SPACE$(14), 7)
	LOCATE 19, 3: CALL QPRINT(SPACE$(50), 7)
	IvaPago# = 0: pago# = 0
	RETURN

GRABO.CARGA:

	WGrabacion$ = "Yes"
	RecordOr4.ORDEN = WOrden$
	RecordOr4.Fecha = FNFECHA$(WFecha$)
	RecordOr4.Proveedor = WProveedor$
	RecordOr4.Vencimiento = FNFECHA$(WFechaVto$)
	RecordOr4.Observaciones = WObservaciones$
	RecordOr4.Banco2 = XBanco$
	RecordOr4.Tipo = WTipo$
	RecordOr4.pago = (VAL(WImporte$) / 100) - Rete2784# - Rete3125#
	RecordOr4.Rete2784 = Rete2784#
	RecordOr4.Rete3125 = Rete3125#
	RecordOr4.Marca = ""

	ClaveOr4$ = WOrden$
	OpOr4% = 2
	Wst% = FOr4W
	IF Wst% <> 0 THEN
		CALL Errores(Error$(), Wst%, Or4Msg$)
		WGrabacion$ = "No"
		RETURN
	END IF

	FOR J% = 1 TO 8

		IF VAL(Vector$(J%, 5)) <> 0 THEN

			' Grabacion de Orden de Pago

			J$ = STR$(J%)
			CALL Ceros(J$, 2)

			RecordOr3.ORDEN = WOrden$
			RecordOr3.Renglon = J$
			RecordOr3.Fecha = FNFECHA$(Vector$(J%, 1))
			RecordOr3.Tipo = Vector$(J%, 3)
			RecordOr3.Numero = Vector$(J%, 2)
			RecordOr3.Descripcion = Vector$(J%, 4)
			RecordOr3.Importe = VAL(Vector$(J%, 5))
			RecordOr3.Rete3125 = VAL(Vector$(J%, 9))
			RecordOr3.Importe2 = VAL(Vector$(J%, 11))
			RecordOr3.Cuenta = Vector$(J%, 10)

			ClaveOr3$ = RecordOr3.ORDEN + RecordOr3.Renglon
			OpOr3% = 2
			Wst% = FOr3W
			IF ST.009% <> 0 THEN
				CALL Errores(Error$(), Wst%, Or3Msg$)
				WGrabacion$ = "No"
				RETURN
			END IF

		END IF
	NEXT J%

	FOR Da% = 1 TO Rete%

		OpRet% = 5
		ClaveRet$ = WFEC$ + Rete$(Da%, 1)
		Wst% = FRetR%

		IF Wst% = 0 THEN

			RecordRet.Neto = VAL(Rete$(Da%, 2))
			RecordRet.Retenido = RecordRet.Retenido + VAL(Rete$(Da%, 3))
			RecordRet.Bruto = RecordRet.Bruto + VAL(Rete$(Da%, 4))
			RecordRet.iva = RecordRet.iva + VAL(Rete$(Da%, 5))

			OpRet% = 3
			Wst% = FRetW
			CALL Errores(Error$(), Wst%, RetMsg$)

		END IF

	NEXT Da%
	RETURN

Fin:
	OpOr3% = 28
	Wst% = FOr3R%
	CALL Errores(Error$(), Wst%, "")
	CLOSE

	WMenu% = XMenu%
	WChoice% = XChoice%

	CHAIN "Adminis"

8000 '
	LOCATE 23, 2: PRINT STRING$(79, 32)
	RETURN

Calcula.retencion:

	ERASE Rete$
	Rete% = 0

	Rete3125# = 0
	Rete2784# = 0

	VidPop "Fill Page 1"
	VidPop "Admi540i"

	CALL Ingreso(iva$, "N", "", 5, 2, 11, 53, 7, 0, "###.##", A$)

	VidPop "Display Page 1"
		    
	IF VAL(WTipo$) <> 1 AND VAL(WTipo$) <> 5 AND VAL(WTipo$) <> 9 THEN RETURN

	GOSUB Calcula.2784
	GOSUB Calcula.3125
	
	Rete2784# = FnRedondeo#(Rete2784#)
	Rete3125# = FnRedondeo#(Rete3125#)

	Rete1$ = STR$(Rete2784# * 100)
	Rete2$ = STR$(Rete3125# * 100)

	VidPop "Fill Page 1"
	VidPop "Admi540R"

	CALL Ingreso(Rete1$, "N", "", 8, 2, 11, 52, 7, 0, "###,###.##", A$)
	CALL Ingreso(Rete2$, "N", "", 8, 2, 12, 52, 7, 0, "###,###.##", A$)
	 
	VidPop "Display Page 1"

	Rete2784# = VAL(Rete1$) / 100
	Rete3125# = VAL(Rete2$) / 100

	RETURN

Calcula.2784:

	IF VAL(RecordPrv.pago) = 1 THEN
		Concepto$ = "8"
			ELSE
		IF VAL(RecordPrv.pago) = 2 THEN
			Concepto$ = "13"
				ELSE
			IF VAL(RecordPrv.pago) = 3 THEN
				Concepto$ = "6"
					ELSE
				RETURN
			END IF
		END IF
	END IF

	Porce$(1, 1) = RecordPrv.Responsable1
	Porce$(1, 2) = STR$(RecordPrv.Porce1)

	Porce$(2, 1) = RecordPrv.Responsable2
	Porce$(2, 2) = STR$(RecordPrv.Porce2)

	Porce$(3, 1) = RecordPrv.Responsable3
	Porce$(3, 2) = STR$(RecordPrv.Porce3)

	RESPONSABLE$ = ""
	RESPONSABLE% = 0

	Resp$ = "N"

	FOR J% = 1 TO 3
		IF VAL(Porce$(J%, 1)) <> 0 THEN
			RESPONSABLE% = RESPONSABLE% + 1
			RESPONSABLE$ = Porce$(J%, 1)
			Resp$ = "S"
		END IF
	NEXT J%

	IMPOR# = 0
	Neto# = 0
	
	Bruto# = Total#
	Neto# = (Total# / (1 + (VAL(iva$) / 10000)))
	iva# = Total# - Neto#
	TBASE# = Neto#

	IF Resp$ = "S" THEN
		ACU# = TBASE#
		ACUBRU# = Bruto#
		ACUIVA# = iva#
		GOSUB Varios.Responsables
		RETURN
			ELSE
		OpRet% = 5
		ClaveRet$ = WFEC$ + WProveedor$
		Wst% = FRetR
		IF Wst% <> 0 THEN
			GOSUB INICIALIZO.CAMPOS
		END IF
	END IF

	NETPAG# = RecordRet.Neto
	TOTRET# = RecordRet.Retenido

	GOSUB Calculo

	RETURN

Varios.Responsables:

	FOR J% = 1 TO RESPONSABLE%

		IF VAL(Porce$(J%, 1)) = 0 THEN EXIT FOR

		OpRet% = 5
		ClaveRet$ = WFEC$ + Porce$(J%, 1)
		Wst% = FRetR
		IF Wst% <> 0 THEN
			GOSUB INICIALIZO.CAMPOS
		END IF

		NETPAG# = RecordRet.Neto
		TOTRET# = RecordRet.Retenido
		TBASE# = (ACU# * (VAL(Porce$(J%, 2)) / 10000))
		Bruto# = ACUBRU# * (VAL(Porce$(J%, 2)) / 10000)
		iva# = ACUIVA# * (VAL(Porce$(J%, 2)) / 10000)

		OpPrv% = 5
		ClavePrv$ = Porce$(J%, 1)
		Ws% = FPrvR

		GOSUB Calculo
		Retenido# = Retenido# + Rete2784#

	NEXT J%
	ACU# = 0
	ACUBRU# = 0
	ACUIVA# = 0
	Rete2784# = Retenido#
	Retenido# = 0

	RETURN

INICIALIZO.CAMPOS:

	IF Resp$ = "N" THEN
		RecordRet.Proveedor = WProveedor$
				ELSE
		RecordRet.Proveedor = Porce$(J%, 1)
	END IF

	RecordRet.Fecha = WFEC$
	RecordRet.Neto = 0
	RecordRet.Anticipo = 0
	RecordRet.Bruto = 0
	RecordRet.iva = 0
	RecordRet.Retenido = 0

	ClaveRet$ = WFEC$ + RecordRet.Proveedor
	OpRet% = 2
	Wst% = FRetW
	CALL Errores(Error$(), Wst%, RetMsg$)
	RETURN

Calculo:

	IF VAL(Concepto$) >= 1 AND VAL(Concepto$) <= 6 THEN
		MINI# = CVD(T$(1)) / 100
				ELSE
		IF VAL(Concepto$) >= 7 AND VAL(Concepto$) <= 12 THEN
			MINI# = CVD(T$(2)) / 100
					ELSE
			MINI# = CVD(T$(3)) / 100
		END IF
	 END IF

	 Acupag# = NETPAG# + TBASE#
	 TSUB# = Acupag# - MINI#

	 IF INT(TSUB#) <= 0 THEN
		TSUB# = 0
		Rete2784# = 0
		GOTO Graba.Rete
	END IF

	IF VAL(RecordPrv.Ganancia) = 2 AND (VAL(Concepto$) >= 1 AND VAL(Concepto$) <= 7) THEN
		Tasa# = .02
			ELSE
		IF VAL(RecordPrv.Ganancia) = 2 AND VAL(Concepto$) >= 8 THEN
			Tasa# = .28
		END IF
	END IF

	IF VAL(RecordPrv.Ganancia) = 1 THEN
		Tasa# = .06
		IF VAL(Concepto$) >= 1 AND VAL(Concepto$) <= 7 THEN
			Tasa# = .01
		END IF
	END IF

	IF VAL(RecordPrv.Ganancia) = 1 AND VAL(Concepto$) = 8 THEN

		Rete2784# = 0
		TOPE# = 0
		TOPE1# = 0

		FOR JI% = 1 TO 4

			IF JI% = 1 AND TSUB# <= IP(JI%) THEN
				TOPE1# = TSUB#
				GOSUB Auxi
				EXIT FOR
			END IF

			IF JI% = 1 AND TSUB# > IP(JI%) THEN
				TOPE1# = IP(JI%)
				GOSUB Auxi

					ELSE

				IF TSUB# <= IP(JI%) THEN
					TOPE# = IP(JI% - 1)
					TOPE1# = TSUB#
					GOSUB Auxi
					EXIT FOR
				END IF

				IF TSUB# > IP(JI%) THEN
					TOPE# = IP(JI% - 1)
					TOPE1# = IP(JI%)
					GOSUB Auxi
				END IF

			END IF

		NEXT JI%

		IF JI% > 4 AND TSUB# > IP(4) THEN
			TOPE1# = TSUB#
			TOPE# = IP(4)
			JI% = 5
			GOSUB Auxi
		END IF

					ELSE

		Rete2784# = (TSUB# * Tasa#)

	END IF

	Rete2784# = Rete2784# - TOTRET#

	IF Rete2784# < CVD(T$(8)) / 100 THEN
		Rete2784# = 0
			ELSE
		IF Rete2784# > Neto# THEN
			NORET# = Rete2784#
			NORET$ = "N"
			Rete2784# = 0
		END IF
	END IF

Graba.Rete:

	IF NORET$ <> "N" THEN

		Rete% = Rete% + 1

		IF Resp$ = "N" THEN
			Rete$(Rete%, 1) = WProveedor$
				ELSE
			Rete$(Rete%, 1) = Porce$(J%, 1)
		END IF

		Rete$(Rete%, 2) = STR$(Acupag#)
		Rete$(Rete%, 3) = STR$(Rete2784#)
		Rete$(Rete%, 4) = STR$(Bruto#)
		Rete$(Rete%, 5) = STR$(iva#)

	END IF

	RETURN

Auxi:

	SUM# = TOPE1# - TOPE#
	SUM# = (SUM# * VAL(TX$(JI%))) / 100
	Rete2784# = Rete2784# + SUM#
	SUM# = 0
	TOPE# = 0
	TOPE1# = 0
	RETURN

Limpio.Variables:

	MINI# = 0
	Acupag# = 0
	pago# = 0
	ACUIMP# = 0
	TSUB# = 0
	Tasa# = 0
	Rete2784# = 0
	INSC$ = " "
	CON$ = " "
	NORET$ = " "
	NORET# = 0
	Bruto# = 0
	TBASE# = 0
	iva# = 0
	IMPOR# = 0
	ERASE Porce$
	RETURN


Calcula.3125:

	IF VAL(WTipo$) <> 1 THEN RETURN

	IF VAL(RecordPrv.Ganancia) = 1 AND VAL(RecordPrv.ReteIva) <> 0 THEN
	    GOSUB Retencion.3125
			ELSE
	    RETURN
	END IF
	RETURN

Retencion.3125:

	Rete3125# = 0
	Iva3125# = 0
	Factura# = 0
	Tasa# = VAL(RecordPrv.ReteIva) / 100

	FOR J% = 1 TO 8

		IF VAL(Vector$(J%, 5)) <> 0 THEN

			IF VAL(Vector$(J%, 3)) = 1 OR VAL(Vector$(J%, 3)) = 2 OR VAL(Vector$(J%, 3)) = 3 THEN

				ClaveIvc$ = WProveedor$ + Vector$(J%, 3) + Vector$(J%, 2)
				OpIvc% = 5
				Wst% = FIvcR
				IF Wst% = 0 THEN

					IF RecordIvc.Marca <> "X" THEN

						retencion# = 0

						Iva3125# = (RecordIvc.iva + RecordIvc.Iva4 + RecordIvc.Iva27)

						IF Iva3125# > CVD(T$(12)) / 100 THEN
							retencion# = Iva3125# * (Tasa# / 100)
						END IF

						Vector$(J%, 9) = STR$(retencion#)

						Rete3125# = Rete3125# + retencion#

					END IF

				END IF
			END IF
		END IF
	NEXT J%
	RETURN


IMPREORDEN:

	WIMpre$ = ""
	DO
		LOCATE 24, 5: CALL QPRINT("Imprime la O.P. [ S/N ] ", 7)
		CALL Ingreso(WIMpre$, "A", "", 1, 0, 24, 35, 7, 0, "", A$)
		LOCATE 24, 5: CALL QPRINT(STRING$(50, 32), 7)
	LOOP UNTIL ASC(WIMpre$) = 83 OR ASC(WIMpre$) = 78

	IF WIMpre$ = "N" THEN RETURN

	XCo% = 0
	ERASE XObra$

	FOR Da% = 1 TO 8

		IF VAL(Vector$(Da%, 5)) <> 0 THEN

			ClaveIvc$ = WProveedor$ + Vector$(Da%, 3) + Vector$(Da%, 2)
			OpIvc% = 5
			Wst% = FIvcR
			IF Wst% = 0 THEN

				Dato$(1, 1) = RecordIvc.Obra1
				Dato$(1, 2) = STR$(RecordIvc.Importe1)

				Dato$(2, 1) = RecordIvc.Obra2
				Dato$(2, 2) = STR$(RecordIvc.Importe2)

				Dato$(3, 1) = RecordIvc.Obra3
				Dato$(3, 2) = STR$(RecordIvc.Importe3)

				Dato$(4, 1) = RecordIvc.Obra4
				Dato$(4, 2) = STR$(RecordIvc.Importe4)

				Dato$(5, 1) = RecordIvc.Obra5
				Dato$(5, 2) = STR$(RecordIvc.Importe5)

				Dato$(6, 1) = RecordIvc.Obra6
				Dato$(6, 2) = STR$(RecordIvc.Importe6)

				Dato$(7, 1) = RecordIvc.Obra7
				Dato$(7, 2) = STR$(RecordIvc.Importe7)

				Dato$(8, 1) = RecordIvc.Obra8
				Dato$(8, 2) = STR$(RecordIvc.Importe8)

				Dato$(9, 1) = RecordIvc.Obra9
				Dato$(9, 2) = STR$(RecordIvc.Importe9)

				Dato$(10, 1) = RecordIvc.Obra10
				Dato$(10, 2) = STR$(RecordIvc.Importe10)

				FOR Xx% = 1 TO 10

					IF VAL(Dato$(Xx%, 1)) <> 0 AND VAL(Dato$(Xx%, 2)) <> 0 THEN

						Obra$ = Dato$(Xx%, 1)
						Concepto$ = RecordIvc.Concepto
						Importe# = VAL(Dato$(Xx%, 2))

						GOSUB Ingresa.Obra

					END IF

				NEXT Xx%

			END IF

		END IF
	NEXT Da%

	GOSUB Calcula.Obra

	RecordPrv.Nombre = RecordPrv.Nombre + SPACE$(40)

	SELECT CASE RecordEmp.Impresora
		CASE "2"
			LPRINT CHR$(27) + CHR$(38) + CHR$(107) + CHR$(50) + CHR$(83);
		CASE ELSE
			LPRINT CHR$(15);
	END SELECT


 FOR Ciclo% = 1 TO 1

			LPRINT "+--------------------------------------------------------------------------------------------------+"
		LPRINT USING "|\                  \          SOLICITUD DE ORDEN DE PAGO               |     Fecha    |  Numero   |"; RecordEmp.Nombre

		IF WTipo$ = "1" OR WTipo$ = "5" OR WTipo$ = "6" OR WTipo$ = "9" THEN
			LPRINT USING "|Sres.:\                            \                                   | \\ | \\ | \\ |   ######  |"; RecordPrv.Nombre; LEFT$(Fecha$, 2); MID$(Fecha$, 3, 2); RIGHT$(Fecha$, 2); VAL(WOrden$)
					ELSE
			LPRINT USING "|Sres.:                                                                 | \\ | \\ | \\ |   ######  |"; LEFT$(Fecha$, 2); MID$(Fecha$, 3, 2); RIGHT$(Fecha$, 2); VAL(WOrden$)
		END IF

		LPRINT "|--------------------------------------------------------------------------------------------------|"
		LPRINT "| Comprobante | F.Factura | F.Vencim.  |                 Descripcion             |     Importe     |"
		LPRINT "---------------------------------------------------------------------------------------------------+"

		FOR WCiclo% = 1 TO 8
			IF VAL(Vector$(WCiclo%, 5)) <> 0 THEN
				LPRINT USING "| \  \ \    \ |  \      \ | \      \   |\                                       \|##,###,###,###.##|"; Desc$(VAL(Vector$(WCiclo%, 3))); Vector$(WCiclo%, 2); fnimpredate$(FNREVDATE$(Vector$(WCiclo%, 12)));  _
fnimpredate$(FNREVDATE$(Vector$(WCiclo%, 13))); Vector$(WCiclo%, 4); VAL(Vector$(WCiclo%, 5)) / 100
							  ELSE
				LPRINT "|             |           |            |                                         |                 |"
			END IF
		NEXT WCiclo%

		LPRINT "|--------------------------------------------------------------------------------------------------|"
		LPRINT USING "| RETENCION 2784:###,###,###.##      RETENCION 3125:###,###,###.##      NETO A PAGAR:###,###,###.##|"; Rete2784#; Rete3125#; Total# - Rete2784# - Rete3125#
		LPRINT "|--------------------------------------------------------------------------------------------------|"
		LPRINT USING "| Cond.de Pago  : #### dias                                                                        |"; VAL(RecordPrv.Dias)
		LPRINT "|                                                                                                  |"
		LPRINT USING "| OBSERVACIONES : \                                                                           \    |"; RecordOr4.Observaciones
		'LPRINT "|--------------------------------------------------------------------------------------------------|"
		'LPRINT "|   Obra   |                      | Concepto |                    |  Importe |  Presup. |  Gastado |"
		'LPRINT "|--------------------------------------------------------------------------------------------------|"

		FOR Da% = 1 TO XCo%

			OpCon% = 5
			ClaveCon$ = XObra$(Da%, 2)
			Wst% = FConR

			OpObr% = 5
			ClaveObr$ = XObra$(Da%, 1)
			Wst% = FObrR

			LPRINT USING "|  ####    |\                    \|   ####   |\                  \|###,###.##|###,###.##|###,###.##|"; VAL(XObra$(Da%, 1)); RecordObr.Nombre; VAL(XObra$(Da%, 2)); RecordCon.Nombre; VAL(XObra$(Da%, 3)) / 100; VAL( _
XObra$(Da%, 4)) / 100; VAL(XObra$(Da%, 5)) / 100

		NEXT Da%

		LPRINT "|--------------------------------------------------------------------------------------------------|"
		LPRINT "|                       |                       |                        |                         |"
		LPRINT "|                       |                       |                        |                         |"
		LPRINT "|                       |                       |                        |                         |"
		LPRINT "|                       |                       |                        |                         |"
		LPRINT "|      Confecciono      |         Autorizo      |        1ra Firma       |      2da Firma          |"
		LPRINT "+--------------------------------------------------------------------------------------------------+"

		GOSUB Imprime.CtaCte

		LPRINT CHR$(12)

  NEXT Ciclo%
  RETURN

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

Impre3125:

		IF Rete3125# = 0 THEN RETURN

		EmpNombre$ = RecordEmp.Nombre
		EmpDireccion$ = RecordEmp.Direccion
		EmpLocalidad$ = RecordEmp.Localidad
		EmpCuit$ = RecordEmp.Cuit

		SELECT CASE RecordEmp.Impresora
			CASE "2"
				LPRINT CHR$(27) + CHR$(38) + CHR$(107) + CHR$(50) + CHR$(83);
			CASE ELSE
				LPRINT CHR$(15);
		END SELECT

		Numero# = Numero# + 1
		Nombre.Empre$ = EmpNombre$

		FOR WCiclo% = 1 TO 2

		    IF WCiclo% = 2 THEN
			  Nombre.Empre$ = LEFT$(RecordPrv.Nombre, 25)
		    ELSE
			    IF WCiclo% = 3 THEN
				  Nombre.Empre$ = SPACE$(75)
			    END IF
		    END IF

		    LPRINT
		    LPRINT USING "\                                        \|      RG 3137 - Comprobante de Retencion      "; EmpNombre$
		    LPRINT USING "\                                        \|      Mes de : \                 \            "; EmpDireccion$; WFecha1$
		    LPRINT USING "\                                        \|                                              "; EmpLocalidad$
		    LPRINT USING "C.U.I.T. : \                             \|      Buenos Aires, \                       \ "; EmpCuit$; fnimpredate$(FNREVDATE$(RecordOr4.Fecha))
		    LPRINT "------------------------------------------------------------------------------------------"
		    LPRINT USING "Beneficiario :\                              \ "; RecordPrv.Nombre
		    LPRINT USING "Domicilio    : \                              \ "; RecordPrv.Direccion
		    LPRINT USING "C.U.I.T      : \               \ "; RecordPrv.Cuit
		    LPRINT "------------------------------------------------------------------------------------------"
		    LPRINT USING "Hemos Procedido a retener el importe de $ #,###,###.## por estar encuadrado dentro del "; Rete3125#
		    LPRINT "art. 1ro. de la Resolucion General Nro 3125."
		    LPRINT
		    LPRINT
		    LPRINT TAB(1); "Fecha Fectura"; TAB(20); "Numero Factura"; TAB(40); "Neto Imponible"; TAB(60); "Iva Facturado"; TAB(80); "Iva Retenido"

		    FOR Da% = 1 TO 8

			IF VAL(Vector$(Da%, 9)) <> 0 THEN

				ClaveIvc$ = WProveedor$ + Vector$(Da%, 3) + Vector$(Da%, 2)
				OpIvc% = 5
				Wst% = FIvcR
				IF Wst% = 0 THEN

					LPRINT TAB(1); fnimpredate$(FNREVDATE$(RecordIvc.Emision));
					LPRINT TAB(20); USING "######"; VAL(RecordIvc.Numero);
					LPRINT TAB(40); USING "###,###.##"; RecordIvc.Neto;
					LPRINT TAB(60); USING "###,###.##"; (RecordIvc.iva + RecordIvc.Iva4 + RecordIvc.Iva27);
					LPRINT TAB(80); USING "###,###.##"; VAL(Vector$(Da%, 9))

				END IF

			END IF
		    NEXT Da%

		    LPRINT
		    LPRINT
		    LPRINT

		    LPRINT TAB(70); Nombre.Empre$
		    LPRINT STRING$(90, 45)
		    FOR WBucle% = 1 TO 5
			   LPRINT
		    NEXT WBucle%
		NEXT WCiclo%
		LPRINT CHR$(12)
		RETURN

Selecciona.Datos:

	LOCATE 1, 1
	VidPop "Fill Page 2"
	VidPop "Ayuda"

	LOCATE 9, 28: PRINT SPACE$(40)
	Nombre$ = ""
	CALL Ingreso(Nombre$, "F", "", 40, 0, 9, 28, 13, 0, "", A$)

	FOR Ciclo% = 40 TO 1 STEP -1
		IF MID$(Nombre$, Ciclo%, 1) <> SPACE$(1) THEN
			Nombre$ = LEFT$(Nombre$, Ciclo%)
			EXIT FOR
		END IF
	NEXT Ciclo%

	ERASE Elije$
	Cantidad% = 0

	SELECT CASE XDATO%
		CASE 1
			KynPrv = 1
			OpPrv = 9
			ClavePrv$ = Nombre$ + SPACE$(40)
			Wst% = FPrvR%

			WHILE Wst% = 0 AND (Nombre$ = LEFT$(RecordPrv.Nombre, Ciclo%) OR Ciclo% = 0)

				Cantidad% = Cantidad% + 1
				Elije$(Cantidad%, 1) = RecordPrv.Codigo
				Elije$(Cantidad%, 2) = LEFT$(RecordPrv.Nombre, 25)

				OpPrv = 6
				Wst% = FPrvR%

			WEND

		CASE ELSE
	END SELECT

	GOSUB Elije.Datos1

	RETURN

Elije.Datos1:

	WFila% = 1
	WColumna% = 1
	WDesde% = 1

	GOSUB Impre.Pantalla.Seleccion1

	DO
		WColor% = 13
		GOSUB IMPRE.LINEA.Seleccion1
		DO
			WIngreso$ = ""
			lugar% = (((WFila% - 1) - INT((WFila% - 1) / 11) * 11)) + 13
			CALL Ingreso(WIngreso$, "G", "", 1, 0, lugar%, 13, 13, 0, "", A$)
		LOOP UNTIL VAL(A$) = 0 OR VAL(A$) = 10 OR VAL(A$) = 12 OR VAL(A$) = 11 OR VAL(A$) = 2 OR VAL(A$) = 3
		SELECT CASE VAL(A$)
			CASE 0
				Codigo$ = Elije$(WFila%, 1)
			CASE 11
				IF WFila% < 200 AND Cantidad% > WFila% THEN
					WColor% = 7
					GOSUB IMPRE.LINEA.Seleccion1
					WFila% = WFila% + 1
					Compara% = (INT((WFila% - 1) / 11) * 11) + 1
					IF Compara% = WFila% THEN
						WDesde% = WFila%
						GOSUB Impre.Pantalla.Seleccion1
					END IF
				END IF

			CASE 12
				IF WFila% > 1 THEN
						WColor% = 7
						GOSUB IMPRE.LINEA.Seleccion1
						WFila% = WFila% - 1
						Compara% = (INT((WFila%) / 11) * 11)
						IF Compara% = WFila% THEN
							WDesde% = WFila% - 10
							GOSUB Impre.Pantalla.Seleccion1
						END IF
				END IF

			CASE ELSE

		END SELECT

	LOOP UNTIL VAL(A$) = 0 OR VAL(A$) = 10

	LOCATE 1, 1
	VidPop "Display Page 2"

	A$ = ""

	RETURN

IMPRE.LINEA.Seleccion1:

	lugar% = (((WFila% - 1) - INT((WFila% - 1) / 11) * 11)) + 13
		
	LOCATE lugar%, 16: PRINT Elije$(WFila%, 1)
	LOCATE lugar%, 30: PRINT Elije$(WFila%, 2)
	RETURN

Impre.Pantalla.Seleccion1:

	WHasta% = WDesde% + 10
	IF WHasta% > 200 THEN WHasta% = 200
	FOR WCiclo% = WDesde% TO WHasta%
		lugar% = (((WCiclo% - 1) - INT((WCiclo% - 1) / 11) * 11)) + 13
		IF Elije$(WCiclo%, 1) <> "" THEN
			LOCATE lugar%, 16: PRINT Elije$(WCiclo%, 1)
			LOCATE lugar%, 30: PRINT Elije$(WCiclo%, 2)
						ELSE
			LOCATE lugar%, 16: PRINT SPACE$(11)
			LOCATE lugar%, 30: PRINT SPACE$(35)
		END IF
	NEXT WCiclo%
	RETURN

44000 '***************************************
44010 '* OBTIENE FECHA DE VENCIMIENTO        *
44020 '***************************************
44030 '
44040 'SAVE "FECHA004.BAS",A
44050 '
44060 IF SENAL = 1 THEN 44170
44070 SENAL = 1
44080 REM  ----    DIM DS(12),FR$(15)
44090 DATA 31,28,31,30,31,30,31,31,30,31,30,31
44100 DATA "0101","0105","2505", , ,"0907", ,"1210", ,"2512", , , , , ,
44110    FOR AA = 1 TO 12
44120        READ DS(AA)
44130    NEXT
44140    FOR AA = 1 TO 15
44150       READ FR$(AA)
44160    NEXT
44170 DG% = 0
44180    AN$ = MID$(XFECHA$, 5, 2): AN# = VAL(AN$)
44190    ME$ = MID$(XFECHA$, 3, 2): ME% = VAL(ME$)
44200    DI$ = MID$(XFECHA$, 1, 2): DI% = VAL(DI$)
44210 'CANTIDAD DE DIAS HASTA LA FECHA
44220    DG% = DI% + PLAZO% - 1
44230    FOR AA = ME% TO 12
44240        IF (AN# MOD 4 = 0) AND ME% = 2 THEN DS(2) = 29 ELSE DS(2) = 28
44250        IF DG% <= DS(AA) THEN 44300
44260        DG% = DG% - DS(AA)
44270    NEXT AA
44280    AN# = AN# + 1
44290    ME% = 1: GOTO 44230
44300 DI% = DG%: DI$ = RIGHT$("0" + MID$(STR$(DI%), 2, LEN(STR$(DI%)) - 1), 2)
44310 ME% = AA: ME$ = RIGHT$("0" + MID$(STR$(ME%), 2, LEN(STR$(ME%)) - 1), 2)
44320 AN$ = RIGHT$("0" + MID$(STR$(AN#), 2, LEN(STR$(AN#)) - 1), 2)
44330 FVENC$ = DI$ + ME$ + AN$
44340 RETURN
44350 '

Ingresa.Obra:

	Entra$ = "S"

	FOR XDa% = 1 TO XCo%

		IF XObra$(XDa%, 1) = Obra$ AND XObra$(XDa%, 2) = Concepto$ THEN

			XObra$(XDa%, 3) = STR$(VAL(XObra$(XDa%, 3)) + Importe#)
			Entra$ = "N"

		END IF

	NEXT XDa%

	IF Entra$ = "S" THEN

		XCo% = XCo% + 1

		XObra$(XCo%, 1) = Obra$
		XObra$(XCo%, 2) = Concepto$
		XObra$(XCo%, 3) = STR$(Importe#)

	END IF

	RETURN

Calcula.Obra:

	FOR Da% = 1 TO XCo%

		OpPto% = 5
		ClavePto$ = XObra$(Da%, 1) + XObra$(Da%, 2)
		KynPto% = 1
		Wst% = FPtoR

		IF Wst% = 0 THEN

			XObra$(Da%, 4) = STR$(RecordPto.Importe)

		END IF

		OpIvc% = 9
		ClaveIvc$ = XObra$(Da%, 2)
		KynIvc% = 3
		Wst% = FIvcR

		WHILE Wst% = 0 AND XObra$(Da%, 2) = RecordIvc.Concepto

			Dato$(1, 1) = RecordIvc.Obra1
			Dato$(1, 2) = STR$(RecordIvc.Importe1)

			Dato$(2, 1) = RecordIvc.Obra2
			Dato$(2, 2) = STR$(RecordIvc.Importe2)

			Dato$(3, 1) = RecordIvc.Obra3
			Dato$(3, 2) = STR$(RecordIvc.Importe3)

			Dato$(4, 1) = RecordIvc.Obra4
			Dato$(4, 2) = STR$(RecordIvc.Importe4)

			Dato$(5, 1) = RecordIvc.Obra5
			Dato$(5, 2) = STR$(RecordIvc.Importe5)

			Dato$(6, 1) = RecordIvc.Obra6
			Dato$(6, 2) = STR$(RecordIvc.Importe6)

			Dato$(7, 1) = RecordIvc.Obra7
			Dato$(7, 2) = STR$(RecordIvc.Importe7)

			Dato$(8, 1) = RecordIvc.Obra8
			Dato$(8, 2) = STR$(RecordIvc.Importe8)

			Dato$(9, 1) = RecordIvc.Obra9
			Dato$(9, 2) = STR$(RecordIvc.Importe9)

			Dato$(10, 1) = RecordIvc.Obra10
			Dato$(10, 2) = STR$(RecordIvc.Importe10)

			FOR Xx% = 1 TO 10

				IF XObra$(Da%, 1) = Dato$(Xx%, 1) THEN

					XObra$(Da%, 5) = STR$(VAL(XObra$(Da%, 5)) + VAL(Dato$(Xx%, 2)))

				END IF

			NEXT Xx%

			OpIvc% = 6
			Wst% = FIvcR

		WEND

	NEXT Da%
	RETURN

Imprime.CtaCte:

	Saldo# = 0

	LPRINT "| Comprobante | F.Factura | F.Vencim.  |   Importe    |  Pendiente   |     Saldo     |"
	LPRINT "-------------------------------------------------------------------------------------+"

	OpCtc% = 9
	ClaveCtc$ = WProveedor$ + SPACE$(40)
	KynCtc% = 1
	Wst% = FCtcR

	WHILE Wst% = 0 AND WProveedor$ = RecordCtc.Proveedor

		IF FnRedondeo#(RecordCtc.Saldo) <> 0 THEN

			GET #99, VAL(RecordCtc.Tipo)

			IF Tipotabla$ = "1" THEN
				Saldo# = Saldo# + RecordCtc.Saldo
                        LPRINT USING "| \\ - \    \ |  \      \ |  \      \  |#,###,###.##  | #,###,###.## | #,###,###.##  |"; Impretabla2$; RecordCtc.Numero; fnimpredate$(FNREVDATE$(RecordCtc.Emision)); fnimpredate$(FNREVDATE$(RecordCtc. Vencimiento)); RecordCtc.Importe; RecordCtc.Saldo; Saldo#
					ELSE
				Saldo# = Saldo# - RecordCtc.Saldo
                        LPRINT USING "| \\ - \    \ |  \      \ |  \      \  |#,###,###.##  | #,###,###.## | #,###,###.##  |"; Impretabla2$; RecordCtc.Numero; fnimpredate$(FNREVDATE$(RecordCtc.Emision)); fnimpredate$(FNREVDATE$(RecordCtc. Vencimiento)); RecordCtc.Importe * -1; RecordCtc.Saldo * -1; Saldo#
			END IF

		END IF

		OpCtc% = 6
		Wst% = FCtcR

	WEND

	LPRINT "-------------------------------------------------------------------------------------+"

	RETURN

