impreorden:

	If Val(WEmpresa$) = 1 Then
		ImpreTit$ = "Surfactan S.A."
			Else
		ImpreTit$ = "Pellital S.A."
	End If

	Width "Lpt1:",140

	Retenido# = FNRedondeo#(Val(WRetencion.010$)/100)

	Pagado#   = Val(Wpago.010$)/100
	Fecha$    = WFecha.010$

	Cantidad% = 0
	Total# = 0
	SubTotal# = 0
	Erase Impresion$,Debito$,Credito$
	Clave.009$ = Worden$+Space$(2)
	Op.009% = 9
	Gosub FOrd009r

	While St.009% = 0 And Worden$ = Left$(Clave.009$,6)
		Cantidad% = Cantidad% + 1
		Impresion$(Cantidad%,1) = WFechaCopbte.009$
		Impresion$(Cantidad%,2) = WTipoCopbte.009$
		Impresion$(Cantidad%,3) = WNumeroCopbte.009$
		Impresion$(Cantidad%,4) = WDescripcion.009$ + WDescri.009$
		Impresion$(Cantidad%,5) = WImporteCopbte.009$
		IF Val(Impresion$(Cantidad%,2))= 3 or Val(Impresion$(Cantidad%,2))= 4 Then
			Total# = Total# - Val(WImporteCopbte.009$)/100
											ELSE
			Total# = Total# + Val(WImporteCopbte.009$)/100
		END IF
		Op.009% = 6
		Gosub FOrd009r
	Wend

	Debito$(1) = FnPusing$("##,###,###,###.##",Total#)

	If Retenido# <> 0 Then
		Credito$(1) = FnPusing$("##,###,###,###.##",Retenido#)
	End If

	Credito$(2) = FnPusing$("##,###,###,###.##",Pagado#-Retenido#)
	SubTotal# = Total# - Retenido#

	TotalDebito# = 0
	TotalCredito# = 0

	For Ciclo% = 1 to 8
		IF Val(Impresion$(Ciclo%,2))= 3 or Val(Impresion$(Ciclo%,2))= 4 Then
			TotalDebito#  = TotalDebito#  - Abs(FnEval#(Debito$(Ciclo%)))
			TotalCredito# = TotalCredito# - Abs(FnEval#(Credito$(Ciclo%)))
										  Else
			TotalDebito#  = TotalDebito#  + Abs(FnEval#(Debito$(Ciclo%)))
			TotalCredito# = TotalCredito# + Abs(FnEval#(Credito$(Ciclo%)))
		END IF
	Next Ciclo%

	Erase Campo1%
	XCanti% = 0

	For XCiclo% = 1 to 800

		If Val(Valores$(XCiclo%,1)) <> 0 Then

		       XCanti% = XCanti% + 1

			Campo1%(XCanti%) = XCiclo%

		End If

	Next XCiclo%

	Tipo.Pago$ = "5"
	LPRINT CHR$(15);

	For Ciclo% = 1 to 2

	Impre$ = "N"
	Cierra$ = "S"

	If Tipo.Pago$ =  "5" Then

		Pasa$ = "N"

		For Dada% = 1 to 100
			If Val(Campo$(Dada%,3)) <> 0 Then
				Pasa$ = "S"
			End If
		Next Dada%
		For Dada% = 1 to 100
			If Campo1%(Dada%) <> 0 Then
				Pasa$ = "S"
			End If
		Next Dada%

		If Pasa$ = "S" Then
			lprint tab(2);Chr$(18);Chr$(14);Impretit$;Chr$(15)
			Lprint  "+------------------------------------------------------------------------------------------++-------------------------------------+"
			Lprint "|                                ORDEN DE PAGO                  |     Fecha    |  Numero   ||       Valores o Docucmentos         |"
			Lprint Using "|Sres.:\                            \                           | \\ | \\ | \\ |   ######  ||             Entregados              |";WNombre.007$;Right$(Fecha$,2);Mid$(Fecha$,3,2);Left$(Fecha$,2);Val(WOrden$);
			Lprint  "|------------------------------------------------------------------------------------------||-------------------------------------+"
			Lprint  "| Fecha  | Cte.Nro. | Comprobante  |              Descripcion                  | Importe   ||Nro.  |Banco/Cliente       | Importe |"
			Lprint  "|------------------------------------------------------------------------------------------+|-------------------------------------+"
			Impre$ = "3"
		End If

	End If

	'If Campo1%(1) <> 0 And Tipo.Pago$ <> "5" Then
	'
	'        lprint tab(2);Chr$(18);Chr$(14);Impretit$;Chr$(15)
	'        Lprint  "+------------------------------------------------------------------------------------------++-------------------------------------+"
	'   Lprint "|                                ORDEN DE PAGO                  |     Fecha    |  Numero   ||       Valores o Docucmentos         |"
	'   Lprint  Using"|Sres.:\                            \                           | \\ | \\ | \\ |   ######  ||          Entregados                 |";WNombre.007$;Right$(Fecha$,2);Mid$(Fecha$,3,2);Left$(Fecha$,2);Val(WOrden$);
	'        Lprint  "|------------------------------------------------------------------------------------------||-------------------------------------+"
	'        Lprint  "| Fecha  | Cte.Nro. | Comprobante  |              Descripcion                  | Importe   ||Cheque|Banco               | Importe |"
	'        Lprint  "|------------------------------------------------------------------------------------------+|-------------------------------------+"
	'        Impre$ = "1"
	'
	'End If

	'If Val(Campo$(1,1)) <> 0 And Tipo.Pago$ <> "5" Then
	'
	'        lprint tab(2);Chr$(18);Chr$(14);Impretit$;Chr$(15)
	'        Lprint  "+------------------------------------------------------------------------------------------++-------------------------------------+"
	'   Lprint "|                                ORDEN DE PAGO                  |     Fecha    |  Numero   ||       Valores o Docucmentos         |"
	'   Lprint  Using"|Sres.:\                            \                           | \\ | \\ | \\ |   ######  ||           Entregados                |";WNombre.007$;Right$(Fecha$,2);Mid$(Fecha$,3,2);Left$(Fecha$,2);Val(WOrden$);
	'        Lprint  "|------------------------------------------------------------------------------------------||-------------------------------------+"
	'        Lprint  "| Fecha  | Cte.Nro. | Comprobante  |              Descripcion                  | Importe   ||Docum.|Cliente             | Importe |"
	'        Lprint  "|------------------------------------------------------------------------------------------+|-------------------------------------+"
	'        Impre$ = "2"
	'
	'
	'End If

	If Impre$ = "N" Then

		lprint tab(2);Chr$(18);Chr$(14);Impretit$;Chr$(15)
		Lprint  "+------------------------------------------------------------------------------------------+"
	   Lprint "|                                ORDEN DE PAGO                  |     Fecha    |  Numero   |"
	   Lprint  Using"|Sres.:\                            \                           | \\ | \\ | \\ |   ######  |";WNombre.007$;Right$(Fecha$,2);Mid$(Fecha$,3,2);Left$(Fecha$,2);Val(WOrden$);
		Lprint  "|------------------------------------------------------------------------------------------|"
		Lprint  "| Fecha  | Cte.Nro. | Comprobante  |              Descripcion                  | Importe   |"
		Lprint  "|------------------------------------------------------------------------------------------+"

	End If

    For WCiclo% = 1 to 8

	If Val(Impresion$(WCiclo%,1)) <> 0 Then
		Lprint Using "|\\|\\|\\| \    \   |\            \|\                            \             | ###,###.##|";_
			Right$(Impresion$(WCiclo%,1),2);Mid$(Impresion$(WCiclo%,1),3,2);Left$(Impresion$(WCiclo%,1),2);_
			Impresion$(WCiclo%,3);DESC$(Val(Impresion$(WCiclo%,2)));Impresion$(WCiclo%,4);Val(Impresion$(WCiclo%,5))/100;
							   Else
		  Lprint"|  |  |  |          |              |                                           |           |";
	End if

	XDesde% = WCiclo%
	XHasta% = WCiclo%
	Gosub Anexo

    Next WCiclo%

		Lprint"|------------------------------------------------------------------------------------------|";
		XDesde% = 9
		XHasta% = 9
		Gosub Anexo

	  Lprint Using"|IMPORTE NETO A PAGAR :  ##,###,###,###.##                                     | ###,###.##|";Subtotal#;Total#;
		XDesde% = 10
		XHasta% = 10
		Gosub Anexo

		LPrint"|------------------------------------------------------------------------------------------|";
		XDesde% = 11
		XHasta% = 11
		Gosub Anexo

		Lprint"|     Codigo    |     Importe     |  Codigo       |      Banco     |  Cheque  |  Importe   |";
		XDesde% = 12
		XHasta% = 12
		Gosub Anexo

		Lprint"|------------------------------------------------------------------------------------------|";
		XDesde% = 13
		XHasta% = 13
		Gosub Anexo

		Lprint Using"|\            \ |##,###,###,###.##|\            \ |                |          | ###,###.## |";XCuenta$;TotalDebito#;"2101    ";FNRedondeo#(Val(WRetencion.010$)/100);

		XDesde% = 14
		XHasta% = 14
		Gosub Anexo

		'If Tipo.pago$ <> "5" Then
		'        Lprint Using"|               |                 |\            \ | \           \  |##########| ###,###.## |"; Cuenta$(1);XDescri$;VaL(WNroCheque.010$);(Val(WPago.010$)-Val(WRetencion.010$))/100;
		'        XDesde% = XDesde% + 1
		'        XHasta% = XHasta% + 1
		'        Gosub Anexo
		'                        Else
		'        If Efectivo# <> 0 Then
		'                Lprint Using"|               |                 |\            \ | \           \  |##########| ###,###.## |"; Cuenta1$;XDescri$;VaL(WNroCheque.010$);Efectivo#-(Val(WRetencion.010$)/100);
		'                XDesde% = XDesde% + 1
		'                XHasta% = XHasta% + 1
		'                Gosub Anexo
		'        End If
		'        If Valores# <> 0 Then
		'                Lprint Using"|               |                 |\            \ |\              \|          | ###,###.## |"; Cuenta2$;"VAL. DE TCEROS";Valores#;
		'                XDesde% = XDesde% + 1
		'                XHasta% = XHasta% + 1
		'                Gosub Anexo
		'        End If
		'        If Documentos# <> 0 Then
		'                Lprint Using"|               |                 |\            \ |\              \|          | ###,###.## |"; Cuenta3$;"DOC. DE TCEROS";Documentos#;
		'                XDesde% = XDesde% + 1
		'                XHasta% = XHasta% + 1
		'                Gosub Anexo
		'        End If
		'End If

		For Da% = 1 to 12
			If Val(Campo1$(Da%,6)) <> 0 Then
				Lprint Using"|               |                 |\      \       | \           \  |##########| ###,###.## |"; Campo1$(Da%,3);Campo1$(Da%,4);Val(Campo1$(Da%,5));Val(Campo1$(Da%,6))/100;
				XDesde% = XDesde% + 1
				XHasta% = XHasta% + 1
				Gosub Anexo
			ENd If
		Next da%



		Lprint"|------------------------------------------------------------------------------------------|";
		XDesde% = XDesde% + 1
		XHasta% = XHasta% + 1
		Gosub Anexo

	  Lprint Using"|       TOTAL   |##,###,###,###.##|                                     TOTAL | ###,###.## |";TotalDebito#;TotalDebito#;
		XDesde% = XDesde% + 1
		XHasta% = XHasta% + 1
		Gosub Anexo

		Lprint"|------------------------------------------------------------------------------------------|";
		XDesde% = XDesde% + 1
		XHasta% = XHasta% + 1
		Gosub Anexo

	 Lprint Using "|OBSERVACIONES :\                                                                         \|";WObservaciones.010$;
		XDesde% = XDesde% + 1
		XHasta% = XHasta% + 1
		Gosub Anexo

		Lprint"|------------------------------------------------------------------------------------------+";
		XDesde% = XDesde% + 1
		XHasta% = XHasta% + 1
		Gosub Anexo

		Lprint"|   Confecciono    |    Autorizo    |   1ra Firma     |   2da Firma     |  Recibi Conforme |";
		XDesde% = XDesde% + 1
		XHasta% = XHasta% + 1
		Gosub Anexo

		Lprint"|                  |                |                 |                 |                  |";
		XDesde% = XDesde% + 1
		XHasta% = XHasta% + 1
		Gosub Anexo

		Lprint"|                  |                |                 |                 |                  |";
		XDesde% = XDesde% + 1
		XHasta% = XHasta% + 1
		Gosub Anexo

		Lprint"|------------------------------------------------------------------------------------------+";
		XDesde% = XDesde% + 1
		XHasta% = 99
		WLug% = XDesde%
		Gosub Anexo

		If WLug% > XLug% Then
			XLug% = WLug%
		End If

		For XX% = XLug% to 28
			lprint
		Next XX%

  Next Ciclo%
  Return

Anexo:

	Select  Case Impre$

		Case "1"
			For XCiclo% = XDesde% To XHasta%
				If Campo1%(XCiclo%) <> 0 Then
					XLug% = XCiclo%
					XLugar% = Campo1%(XCiclo%)
					Lprint Tab(93);"|";Left$(Valores$(XLugar%,1),6);
					Lprint Tab(100);"|";Left$(Valores$(XLugar%,2),20);
					Lprint Tab(121);"|";Using"######.##";Val(Valores$(XLugar%,3))/100;
					LPrint Tab(131);"|";
							Else
					If Cierra$ = "S" Then
						LPrint Tab(93);"|-------------------------------------+";
						Cierra$ = "N"
					End If
				End If
			Next XCiclo%

		Case "2"
			For XCiclo% = XDesde% To XHasta%
				If Val(Campo$(XCiclo%,1)) <> 0 Then
					XLug% = XCiclo%
					Lprint Tab(93);"|";Campo$(XCiclo%,1);
					Lprint Tab(100);"|";Left$(Campo$(XCiclo%,2),20);
					Lprint Tab(121);"|";Using"######.##";Val(Campo$(XCiclo%,3))/100;
					LPrint Tab(131);"|";
							Else
					If Cierra$ = "S" Then
						LPrint Tab(93);"|-------------------------------------+";
						Cierra$ = "N"
					End If
				End If
			Next XCiclo%

		Case "3"

			Erase Impre$
			Lugar% = 0
			For Dada% = 1 to 100
				If Val(Campo$(Dada%,3)) <> 0 Then
					XLug% = Dada%
					Lugar% = Lugar%  + 1
					Impre$(Lugar%,1) = Campo$(Dada%,1)
					Impre$(Lugar%,2) = Campo$(Dada%,2)
					Impre$(Lugar%,3) = Campo$(Dada%,3)
				End If
			Next Dada%
			For Dada% = 1 to 100
				If Campo1%(Dada%) <> 0 Then
					XLug% = Dada%
					XLugar% = Campo1%(Dada%)
					Lugar% = Lugar%  + 1
					Impre$(Lugar%,1) = Valores$(XLugar%,1)
					Impre$(Lugar%,2) = Valores$(XLugar%,2)
					Impre$(Lugar%,3) = Valores$(XLugar%,3)
				End If
			Next Dada%

			For XCiclo% = XDesde% To XHasta%
				If Val(Impre$(XCiclo%,3)) <> 0 Then
					Lprint Tab(93);"|";Impre$(XCiclo%,1);
					Lprint Tab(100);"|";Left$(Impre$(XCiclo%,2),20);
					Lprint Tab(121);"|";Using"######.##";Val(Impre$(XCiclo%,3))/100;
					LPrint Tab(131);"|";
							Else
					If Cierra$ = "S" Then
						LPrint Tab(93);"|-------------------------------------+";
						Cierra$ = "N"
					End If
				End If
			Next XCiclo%

		Case Else
	End  Select
	lprint
	Return

