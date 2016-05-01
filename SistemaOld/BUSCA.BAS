			ERASE Prove$
			Cantidad% = 0
			Kyn.009% = 1
			Op.009% = 12
			Clave.009$ = SPACE$(100)
			Gosub FPrv009R

			WHILE St.009% = 0

				FOR XCiclo% = 1 TO 30
					If Nombre$ = Mid$(WNombre.009$, XCiclo%,Ciclo%) Then
						Cantidad% = Cantidad% + 1
						Prove$(Cantidad%, 1) = WCodigo.009$
						Prove$(Cantidad%, 2) = left$(WNombre.009$, 25)
					End If
				Next XCiclo%

				Op.009% = 6
				Gosub FPrv009R

			WEND

