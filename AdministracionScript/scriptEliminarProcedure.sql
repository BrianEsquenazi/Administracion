USE [surfactanSA]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaAuxiliar]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaAuxiliar]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaAuxiliar]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaCtaCte]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaCtaCte]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaCtaCte]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaCtaCtePrv]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaCtaCtePrv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaCtaCtePrv]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaCtaCteUs]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaCtaCteUs]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaCtaCteUs]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaCtaCteUs1]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaCtaCteUs1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaCtaCteUs1]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaCtaCteUs2]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaCtaCteUs2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaCtaCteUs2]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaCtaPrvSaldo]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaCtaPrvSaldo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaCtaPrvSaldo]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaEstadistica]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaEstadistica]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaEstadistica]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaIvaCompras]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaIvaCompras]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaIvaCompras]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaIvaComprasCai]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaIvaComprasCai]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaIvaComprasCai]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaMovGasMarca]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaMovGasMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaMovGasMarca]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRecibos]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRecibos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaRecibos]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRecibosMarca]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRecibosMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaRecibosMarca]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRecibosOtro]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRecibosOtro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaRecibosOtro]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRecibosOtroVI]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRecibosOtroVI]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaRecibosOtroVI]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRecibosSalvaMarca]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRecibosSalvaMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaRecibosSalvaMarca]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRecibosSalvaMarcaII]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRecibosSalvaMarcaII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaRecibosSalvaMarcaII]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRetencion]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRetencion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaRetencion]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRetencionPagos]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRetencionPagos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaRetencionPagos]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaSaldoCtaCteCli]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaSaldoCtaCteCli]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaSaldoCtaCteCli]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaSaldoCtaCtePrv]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaSaldoCtaCtePrv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ActualizaSaldoCtaCtePrv]
GO
/****** Object:  StoredProcedure [dbo].[AltaArticulo]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaArticulo]
GO
/****** Object:  StoredProcedure [dbo].[AltaArticuloII]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaArticuloII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaArticuloII]
GO
/****** Object:  StoredProcedure [dbo].[AltaAtributos]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaAtributos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaAtributos]
GO
/****** Object:  StoredProcedure [dbo].[Altabanco]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Altabanco]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Altabanco]
GO
/****** Object:  StoredProcedure [dbo].[AltaCambio]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCambio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaCambio]
GO
/****** Object:  StoredProcedure [dbo].[AltaCambioAdm]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCambioAdm]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaCambioAdm]
GO
/****** Object:  StoredProcedure [dbo].[AltaCarpeta]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCarpeta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaCarpeta]
GO
/****** Object:  StoredProcedure [dbo].[AltaCliente]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaCliente]
GO
/****** Object:  StoredProcedure [dbo].[AltaCliente1]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCliente1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaCliente1]
GO
/****** Object:  StoredProcedure [dbo].[AltaComposicion]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaComposicion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaComposicion]
GO
/****** Object:  StoredProcedure [dbo].[AltaConsig]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaConsig]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaConsig]
GO
/****** Object:  StoredProcedure [dbo].[AltaCotiza]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCotiza]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaCotiza]
GO
/****** Object:  StoredProcedure [dbo].[AltaCotizaII]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCotizaII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaCotizaII]
GO
/****** Object:  StoredProcedure [dbo].[AltaCtacte]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCtacte]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaCtacte]
GO
/****** Object:  StoredProcedure [dbo].[AltaCtacte1]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCtacte1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaCtacte1]
GO
/****** Object:  StoredProcedure [dbo].[AltaCtaCteCli]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCtaCteCli]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaCtaCteCli]
GO
/****** Object:  StoredProcedure [dbo].[AltaCtaCtePrv]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCtaCtePrv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaCtaCtePrv]
GO
/****** Object:  StoredProcedure [dbo].[AltaCtacteVarios]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCtacteVarios]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaCtacteVarios]
GO
/****** Object:  StoredProcedure [dbo].[AltaCtaPrv]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCtaPrv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaCtaPrv]
GO
/****** Object:  StoredProcedure [dbo].[AltaCuenta]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCuenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaCuenta]
GO
/****** Object:  StoredProcedure [dbo].[AltaDepositos]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaDepositos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaDepositos]
GO
/****** Object:  StoredProcedure [dbo].[AltaDesccomp]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaDesccomp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaDesccomp]
GO
/****** Object:  StoredProcedure [dbo].[AltaDevcon]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaDevcon]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaDevcon]
GO
/****** Object:  StoredProcedure [dbo].[AltaEnsayos]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEnsayos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaEnsayos]
GO
/****** Object:  StoredProcedure [dbo].[AltaEntdev]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEntdev]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaEntdev]
GO
/****** Object:  StoredProcedure [dbo].[AltaEnvase]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEnvase]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaEnvase]
GO
/****** Object:  StoredProcedure [dbo].[AltaEspecif]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEspecif]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaEspecif]
GO
/****** Object:  StoredProcedure [dbo].[AltaEspecificaciones]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEspecificaciones]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaEspecificaciones]
GO
/****** Object:  StoredProcedure [dbo].[AltaEspeCli]    Script Date: 05/01/2016 17:47:02 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEspeCli]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaEspeCli]
GO
/****** Object:  StoredProcedure [dbo].[AltaEstadistica]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEstadistica]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaEstadistica]
GO
/****** Object:  StoredProcedure [dbo].[AltaEstadisticaDev]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEstadisticaDev]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaEstadisticaDev]
GO
/****** Object:  StoredProcedure [dbo].[AltaGasimpo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaGasimpo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaGasimpo]
GO
/****** Object:  StoredProcedure [dbo].[AltaHoja]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaHoja]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaHoja]
GO
/****** Object:  StoredProcedure [dbo].[AltaImputacion]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaImputacion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaImputacion]
GO
/****** Object:  StoredProcedure [dbo].[AltaInforme]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaInforme]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaInforme]
GO
/****** Object:  StoredProcedure [dbo].[AltaInventario]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaInventario]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaInventario]
GO
/****** Object:  StoredProcedure [dbo].[AltaIvaCompras]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaIvaCompras]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaIvaCompras]
GO
/****** Object:  StoredProcedure [dbo].[AltaLaudo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaLaudo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaLaudo]
GO
/****** Object:  StoredProcedure [dbo].[AltaLinea]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaLinea]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaLinea]
GO
/****** Object:  StoredProcedure [dbo].[AltaLineaMp]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaLineaMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaLineaMp]
GO
/****** Object:  StoredProcedure [dbo].[AltaMarcas]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMarcas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaMarcas]
GO
/****** Object:  StoredProcedure [dbo].[AltaMinimo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMinimo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaMinimo]
GO
/****** Object:  StoredProcedure [dbo].[AltaMinimoPlanta]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMinimoPlanta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaMinimoPlanta]
GO
/****** Object:  StoredProcedure [dbo].[AltaMovenv]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMovenv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaMovenv]
GO
/****** Object:  StoredProcedure [dbo].[AltaMovgas]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMovgas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaMovgas]
GO
/****** Object:  StoredProcedure [dbo].[AltaMovgasCon]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMovgasCon]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaMovgasCon]
GO
/****** Object:  StoredProcedure [dbo].[AltaMovguia]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMovguia]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaMovguia]
GO
/****** Object:  StoredProcedure [dbo].[AltaMovlab]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMovlab]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaMovlab]
GO
/****** Object:  StoredProcedure [dbo].[AltaMovvar]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMovvar]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaMovvar]
GO
/****** Object:  StoredProcedure [dbo].[AltaMuestra]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMuestra]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaMuestra]
GO
/****** Object:  StoredProcedure [dbo].[AltaMuestraImpre]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMuestraImpre]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaMuestraImpre]
GO
/****** Object:  StoredProcedure [dbo].[AltaOrden]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaOrden]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaOrden]
GO
/****** Object:  StoredProcedure [dbo].[AltaOrdenII]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaOrdenII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaOrdenII]
GO
/****** Object:  StoredProcedure [dbo].[AltaOrdenIII]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaOrdenIII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaOrdenIII]
GO
/****** Object:  StoredProcedure [dbo].[AltaOt]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaOt]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaOt]
GO
/****** Object:  StoredProcedure [dbo].[AltaPago]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPago]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaPago]
GO
/****** Object:  StoredProcedure [dbo].[AltaPagos]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPagos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaPagos]
GO
/****** Object:  StoredProcedure [dbo].[AltaPedido]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPedido]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaPedido]
GO
/****** Object:  StoredProcedure [dbo].[AltaPedidoDevol]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPedidoDevol]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaPedidoDevol]
GO
/****** Object:  StoredProcedure [dbo].[AltaPedidoDevolII]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPedidoDevolII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaPedidoDevolII]
GO
/****** Object:  StoredProcedure [dbo].[AltaPrecios]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPrecios]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaPrecios]
GO
/****** Object:  StoredProcedure [dbo].[AltaPrecios1]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPrecios1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaPrecios1]
GO
/****** Object:  StoredProcedure [dbo].[AltaPreciosMp]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPreciosMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaPreciosMp]
GO
/****** Object:  StoredProcedure [dbo].[AltaPresCon]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPresCon]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaPresCon]
GO
/****** Object:  StoredProcedure [dbo].[AltaPrestamo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPrestamo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaPrestamo]
GO
/****** Object:  StoredProcedure [dbo].[AltaProveedor]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaProveedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaProveedor]
GO
/****** Object:  StoredProcedure [dbo].[AltaProveedor1]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaProveedor1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaProveedor1]
GO
/****** Object:  StoredProcedure [dbo].[AltaPrueart]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPrueart]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaPrueart]
GO
/****** Object:  StoredProcedure [dbo].[AltaPrueter]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPrueter]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaPrueter]
GO
/****** Object:  StoredProcedure [dbo].[AltaRecibos]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaRecibos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaRecibos]
GO
/****** Object:  StoredProcedure [dbo].[AltaRetencion]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaRetencion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaRetencion]
GO
/****** Object:  StoredProcedure [dbo].[AltaRubro]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaRubro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaRubro]
GO
/****** Object:  StoredProcedure [dbo].[AltaSedronar]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaSedronar]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaSedronar]
GO
/****** Object:  StoredProcedure [dbo].[AltaSolguia]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaSolguia]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaSolguia]
GO
/****** Object:  StoredProcedure [dbo].[AltaSolguiaTotal]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaSolguiaTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaSolguiaTotal]
GO
/****** Object:  StoredProcedure [dbo].[AltaSolHoja]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaSolHoja]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaSolHoja]
GO
/****** Object:  StoredProcedure [dbo].[AltaSolicitud]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaSolicitud]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaSolicitud]
GO
/****** Object:  StoredProcedure [dbo].[AltaSoltot]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaSoltot]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaSoltot]
GO
/****** Object:  StoredProcedure [dbo].[AltaTerminado]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaTerminado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaTerminado]
GO
/****** Object:  StoredProcedure [dbo].[AltaVendedor]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaVendedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AltaVendedor]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorArticulo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorArticulo]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorBanco]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorBanco]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorBanco]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorCambio]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorCambio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorCambio]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorCambioAdm]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorCambioAdm]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorCambioAdm]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorCliente]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorCliente]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorComposicion]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorComposicion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorComposicion]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorCuenta]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorCuenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorCuenta]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorEnsayos]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorEnsayos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorEnsayos]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorEnvases]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorEnvases]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorEnvases]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorEspecif]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorEspecif]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorEspecif]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorEspecificaciones]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorEspecificaciones]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorEspecificaciones]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorGasimpo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorGasimpo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorGasimpo]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorLinea]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorLinea]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorLinea]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorLineaMp]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorLineaMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorLineaMp]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorMuestra]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorMuestra]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorMuestra]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorPago]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorPago]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorPago]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorPrecios]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorPrecios]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorPrecios]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorPreciosMp]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorPreciosMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorPreciosMp]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorProveedor]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorProveedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorProveedor]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorRubro]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorRubro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorRubro]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorTerminado]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorTerminado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorTerminado]
GO
/****** Object:  StoredProcedure [dbo].[AnteriorVendedor]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorVendedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[AnteriorVendedor]
GO
/****** Object:  StoredProcedure [dbo].[BorrarArticulo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarArticulo]
GO
/****** Object:  StoredProcedure [dbo].[BorrarAtributos]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarAtributos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarAtributos]
GO
/****** Object:  StoredProcedure [dbo].[BorrarBanco]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarBanco]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarBanco]
GO
/****** Object:  StoredProcedure [dbo].[BorrarCambio]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCambio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarCambio]
GO
/****** Object:  StoredProcedure [dbo].[BorrarCambioAdm]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCambioAdm]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarCambioAdm]
GO
/****** Object:  StoredProcedure [dbo].[BorrarCarpeta]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCarpeta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarCarpeta]
GO
/****** Object:  StoredProcedure [dbo].[BorrarCliente]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarCliente]
GO
/****** Object:  StoredProcedure [dbo].[BorrarComposicion]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarComposicion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarComposicion]
GO
/****** Object:  StoredProcedure [dbo].[BorrarConsig]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarConsig]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarConsig]
GO
/****** Object:  StoredProcedure [dbo].[BorrarCotiza]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCotiza]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarCotiza]
GO
/****** Object:  StoredProcedure [dbo].[BorrarCotizaTotal]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCotizaTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarCotizaTotal]
GO
/****** Object:  StoredProcedure [dbo].[BorrarCtacte]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCtacte]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarCtacte]
GO
/****** Object:  StoredProcedure [dbo].[BorrarCtacteNumero]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCtacteNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarCtacteNumero]
GO
/****** Object:  StoredProcedure [dbo].[BorrarCtaprv]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCtaprv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarCtaprv]
GO
/****** Object:  StoredProcedure [dbo].[BorrarCuenta]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCuenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarCuenta]
GO
/****** Object:  StoredProcedure [dbo].[BorrarDesccomp]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarDesccomp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarDesccomp]
GO
/****** Object:  StoredProcedure [dbo].[BorrarDevcon]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarDevcon]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarDevcon]
GO
/****** Object:  StoredProcedure [dbo].[BorrarEnsayos]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarEnsayos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarEnsayos]
GO
/****** Object:  StoredProcedure [dbo].[BorrarEntdev]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarEntdev]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarEntdev]
GO
/****** Object:  StoredProcedure [dbo].[BorrarEnvases]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarEnvases]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarEnvases]
GO
/****** Object:  StoredProcedure [dbo].[BorrarEspecif]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarEspecif]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarEspecif]
GO
/****** Object:  StoredProcedure [dbo].[BorrarEspecificaciones]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarEspecificaciones]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarEspecificaciones]
GO
/****** Object:  StoredProcedure [dbo].[BorrarEstadistica]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarEstadistica]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarEstadistica]
GO
/****** Object:  StoredProcedure [dbo].[BorrarGasimpo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarGasimpo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarGasimpo]
GO
/****** Object:  StoredProcedure [dbo].[BorrarHoja]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarHoja]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarHoja]
GO
/****** Object:  StoredProcedure [dbo].[BorrarHojaFecha]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarHojaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarHojaFecha]
GO
/****** Object:  StoredProcedure [dbo].[BorrarImputac]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarImputac]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarImputac]
GO
/****** Object:  StoredProcedure [dbo].[BorrarImputacion]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarImputacion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarImputacion]
GO
/****** Object:  StoredProcedure [dbo].[BorrarInforme]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarInforme]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarInforme]
GO
/****** Object:  StoredProcedure [dbo].[BorrarInventario]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarInventario]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarInventario]
GO
/****** Object:  StoredProcedure [dbo].[BorrarInventarioTotal]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarInventarioTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarInventarioTotal]
GO
/****** Object:  StoredProcedure [dbo].[BorrarIvacomp]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarIvacomp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarIvacomp]
GO
/****** Object:  StoredProcedure [dbo].[BorrarLaudo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarLaudo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarLaudo]
GO
/****** Object:  StoredProcedure [dbo].[BorrarLaudoFecha]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarLaudoFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarLaudoFecha]
GO
/****** Object:  StoredProcedure [dbo].[BorrarLinea]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarLinea]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarLinea]
GO
/****** Object:  StoredProcedure [dbo].[BorrarLineaMp]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarLineaMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarLineaMp]
GO
/****** Object:  StoredProcedure [dbo].[BorrarMarcas]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMarcas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarMarcas]
GO
/****** Object:  StoredProcedure [dbo].[BorrarMinimo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMinimo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarMinimo]
GO
/****** Object:  StoredProcedure [dbo].[BorrarMovenv]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMovenv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarMovenv]
GO
/****** Object:  StoredProcedure [dbo].[BorrarMovgas]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMovgas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarMovgas]
GO
/****** Object:  StoredProcedure [dbo].[BorrarMovGasCon]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMovGasCon]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarMovGasCon]
GO
/****** Object:  StoredProcedure [dbo].[BorrarMovguia]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMovguia]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarMovguia]
GO
/****** Object:  StoredProcedure [dbo].[BorrarMovlab]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMovlab]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarMovlab]
GO
/****** Object:  StoredProcedure [dbo].[BorrarMovvar]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMovvar]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarMovvar]
GO
/****** Object:  StoredProcedure [dbo].[BorrarMuestra]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMuestra]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarMuestra]
GO
/****** Object:  StoredProcedure [dbo].[BorrarMuestraImpre]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMuestraImpre]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarMuestraImpre]
GO
/****** Object:  StoredProcedure [dbo].[BorrarMuestraTotal]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMuestraTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarMuestraTotal]
GO
/****** Object:  StoredProcedure [dbo].[BorrarOrden]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarOrden]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarOrden]
GO
/****** Object:  StoredProcedure [dbo].[BorrarOrdenTotal]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarOrdenTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarOrdenTotal]
GO
/****** Object:  StoredProcedure [dbo].[BorrarOt]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarOt]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarOt]
GO
/****** Object:  StoredProcedure [dbo].[BorrarPago]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPago]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarPago]
GO
/****** Object:  StoredProcedure [dbo].[BorrarPedido]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPedido]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarPedido]
GO
/****** Object:  StoredProcedure [dbo].[BorrarPedidoDevol]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPedidoDevol]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarPedidoDevol]
GO
/****** Object:  StoredProcedure [dbo].[BorrarPrecios]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPrecios]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarPrecios]
GO
/****** Object:  StoredProcedure [dbo].[BorrarPreciosMp]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPreciosMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarPreciosMp]
GO
/****** Object:  StoredProcedure [dbo].[BorrarPreciosMpTotal]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPreciosMpTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarPreciosMpTotal]
GO
/****** Object:  StoredProcedure [dbo].[BorrarPreciosTotal]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPreciosTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarPreciosTotal]
GO
/****** Object:  StoredProcedure [dbo].[BorrarPresCon]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPresCon]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarPresCon]
GO
/****** Object:  StoredProcedure [dbo].[BorrarPrestamo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPrestamo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarPrestamo]
GO
/****** Object:  StoredProcedure [dbo].[BorrarProveedor]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarProveedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarProveedor]
GO
/****** Object:  StoredProcedure [dbo].[BorrarRubro]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarRubro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarRubro]
GO
/****** Object:  StoredProcedure [dbo].[BorrarSedronar]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarSedronar]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarSedronar]
GO
/****** Object:  StoredProcedure [dbo].[BorrarSolGuia]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarSolGuia]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarSolGuia]
GO
/****** Object:  StoredProcedure [dbo].[BorrarSolGuiaTotal]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarSolGuiaTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarSolGuiaTotal]
GO
/****** Object:  StoredProcedure [dbo].[BorrarSolHoja]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarSolHoja]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarSolHoja]
GO
/****** Object:  StoredProcedure [dbo].[BorrarSOlicitud]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarSOlicitud]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarSOlicitud]
GO
/****** Object:  StoredProcedure [dbo].[BorrarSolicitudTotal]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarSolicitudTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarSolicitudTotal]
GO
/****** Object:  StoredProcedure [dbo].[BorrarSoltot]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarSoltot]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarSoltot]
GO
/****** Object:  StoredProcedure [dbo].[BorrarTerminado]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarTerminado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarTerminado]
GO
/****** Object:  StoredProcedure [dbo].[BorrarVendedor]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarVendedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BorrarVendedor]
GO
/****** Object:  StoredProcedure [dbo].[BuscarImputacion]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BuscarImputacion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[BuscarImputacion]
GO
/****** Object:  StoredProcedure [dbo].[CalculaDiferenciaArticulo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CalculaDiferenciaArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[CalculaDiferenciaArticulo]
GO
/****** Object:  StoredProcedure [dbo].[CalculaDiferenciaTerminado]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CalculaDiferenciaTerminado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[CalculaDiferenciaTerminado]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaArticulo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaArticulo]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaAtributo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaAtributo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaAtributo]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaBanco]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaBanco]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaBanco]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaBancos]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaBancos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaBancos]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCambio]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCambio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaCambio]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCambioAdm]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCambioAdm]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaCambioAdm]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCambioAdmOrdFecha]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCambioAdmOrdFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaCambioAdmOrdFecha]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCambioOrdFecha]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCambioOrdFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaCambioOrdFecha]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCliente]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaCliente]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaClienteRazon]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaClienteRazon]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaClienteRazon]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaClientes]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaClientes]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaClientes]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaComposicion]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaComposicion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaComposicion]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaComposicionProducto]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaComposicionProducto]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaComposicionProducto]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaConsig]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaConsig]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaConsig]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaConsig1]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaConsig1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaConsig1]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaConsigArti]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaConsigArti]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaConsigArti]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCotiza]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCotiza]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaCotiza]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCtacte]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCtacte]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaCtacte]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCtaCteCli]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCtaCteCli]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaCtaCteCli]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCtaCteComp]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCtaCteComp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaCtaCteComp]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCtaCtePrv]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCtaCtePrv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaCtaCtePrv]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCtaCtePrv_x_Proveedor]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCtaCtePrv_x_Proveedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaCtaCtePrv_x_Proveedor]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCtaPrv]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCtaPrv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaCtaPrv]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCuentas]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCuentas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaCuentas]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaDepositos]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaDepositos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaDepositos]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaDepositosClave]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaDepositosClave]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaDepositosClave]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaDesccomp]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaDesccomp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaDesccomp]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaDesccomp1]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaDesccomp1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaDesccomp1]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaDesccompTotal]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaDesccompTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaDesccompTotal]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEnsayos]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEnsayos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaEnsayos]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEntdev]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEntdev]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaEntdev]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEntdev1]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEntdev1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaEntdev1]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEntdev2]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEntdev2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaEntdev2]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEnvases]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEnvases]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaEnvases]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEspecif]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEspecif]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaEspecif]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEspecificaciones]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEspecificaciones]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaEspecificaciones]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEspeCli]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEspeCli]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaEspeCli]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEstadistica]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEstadistica]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaEstadistica]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEstadistica1]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEstadistica1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaEstadistica1]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaGasimpo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaGasimpo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaGasimpo]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaHoja]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaHoja]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaHoja]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaHojaEspecial]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaHojaEspecial]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaHojaEspecial]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaImputac]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaImputac]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaImputac]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaInforme]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaInforme]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaInforme]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaInformeOrden]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaInformeOrden]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaInformeOrden]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaInventario]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaInventario]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaInventario]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaIvaComp]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaIvaComp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaIvaComp]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaIvaCompCompro]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaIvaCompCompro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaIvaCompCompro]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaIvaCompras]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaIvaCompras]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaIvaCompras]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaLaudo]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaLaudo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaLaudo]
GO
/****** Object:  StoredProcedure [dbo].[Consultalehman]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Consultalehman]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Consultalehman]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaLinea]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaLinea]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaLinea]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaLineaMp]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaLineaMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaLineaMp]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMarcas]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMarcas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaMarcas]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMinimo]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMinimo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaMinimo]
GO
/****** Object:  StoredProcedure [dbo].[Consultamono]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Consultamono]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Consultamono]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMovenv]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMovenv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaMovenv]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMovgas]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMovgas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaMovgas]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMovguia]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMovguia]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaMovguia]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMovlab]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMovlab]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaMovlab]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMovvar]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMovvar]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaMovvar]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMuestra]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMuestra]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaMuestra]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaNumero]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaNumero]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaOperador]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaOperador]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaOperador]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaOperadorClave]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaOperadorClave]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaOperadorClave]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaOrden]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaOrden]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaOrden]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaOrdenCarpeta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaOrdenCarpeta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaOrdenCarpeta]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaOt]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaOt]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaOt]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPago]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPago]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPago]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPagos]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPagos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPagos]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedido]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedido]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPedido]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedido1]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedido1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPedido1]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedido2]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedido2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPedido2]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedidoDevol]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedidoDevol]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPedidoDevol]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedidoDevol1]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedidoDevol1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPedidoDevol1]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedidoDevol2]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedidoDevol2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPedidoDevol2]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedidoDevolFactura]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedidoDevolFactura]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPedidoDevolFactura]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedidoFactura]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedidoFactura]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPedidoFactura]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPrecios]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPrecios]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPrecios]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPreciosMp]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPreciosMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPreciosMp]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPresCon]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPresCon]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPresCon]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPrestamo]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPrestamo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPrestamo]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaProveedores]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaProveedores]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaProveedores]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaProveedoresSiguiente]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaProveedoresSiguiente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaProveedoresSiguiente]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaProvincia]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaProvincia]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaProvincia]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPrueart]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPrueart]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPrueart]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPrueter]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPrueter]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPrueter]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPrueterMenor]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPrueterMenor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaPrueterMenor]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaRecibos]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaRecibos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaRecibos]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaRecibos_x_Recibo]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaRecibos_x_Recibo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaRecibos_x_Recibo]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaRecibosClave]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaRecibosClave]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaRecibosClave]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaRetencion]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaRetencion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaRetencion]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaRubro]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaRubro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaRubro]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaSolicitud]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaSolicitud]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaSolicitud]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaSolicitud1]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaSolicitud1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaSolicitud1]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaTerminado]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaTerminado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaTerminado]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaUltimoDeposito]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaUltimoDeposito]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaUltimoDeposito]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaUltimoNroInterno]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaUltimoNroInterno]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaUltimoNroInterno]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaUltimoNroRecibo]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaUltimoNroRecibo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaUltimoNroRecibo]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaUltimoPago]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaUltimoPago]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaUltimoPago]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaVendedor]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaVendedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ConsultaVendedor]
GO
/****** Object:  StoredProcedure [dbo].[DepuraMovEnv]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DepuraMovEnv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[DepuraMovEnv]
GO
/****** Object:  StoredProcedure [dbo].[ListaArticulo]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaArticulo]
GO
/****** Object:  StoredProcedure [dbo].[ListaArticuloConsulta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaArticuloConsulta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaArticuloConsulta]
GO
/****** Object:  StoredProcedure [dbo].[ListaArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaArticuloDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaArticuloDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaArticuloDesdeHastaMinimo]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaArticuloDesdeHastaMinimo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaArticuloDesdeHastaMinimo]
GO
/****** Object:  StoredProcedure [dbo].[ListaArticuloStock]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaArticuloStock]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaArticuloStock]
GO
/****** Object:  StoredProcedure [dbo].[ListaBancos]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaBancos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaBancos]
GO
/****** Object:  StoredProcedure [dbo].[ListaCambio]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCambio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCambio]
GO
/****** Object:  StoredProcedure [dbo].[ListaCambioAdm]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCambioAdm]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCambioAdm]
GO
/****** Object:  StoredProcedure [dbo].[ListaCliente]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCliente]
GO
/****** Object:  StoredProcedure [dbo].[ListaClienteConsulta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaClienteConsulta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaClienteConsulta]
GO
/****** Object:  StoredProcedure [dbo].[ListaClienteConsulta1]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaClienteConsulta1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaClienteConsulta1]
GO
/****** Object:  StoredProcedure [dbo].[ListaClientes]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaClientes]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaClientes]
GO
/****** Object:  StoredProcedure [dbo].[ListaComposicion]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaComposicion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaComposicion]
GO
/****** Object:  StoredProcedure [dbo].[ListaComposicionDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaComposicionDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaComposicionDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaConsig]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaConsig]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaConsig]
GO
/****** Object:  StoredProcedure [dbo].[ListaConsigCliente]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaConsigCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaConsigCliente]
GO
/****** Object:  StoredProcedure [dbo].[ListaConsigFactura]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaConsigFactura]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaConsigFactura]
GO
/****** Object:  StoredProcedure [dbo].[ListaConsigNumero]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaConsigNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaConsigNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaConsigRepro]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaConsigRepro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaConsigRepro]
GO
/****** Object:  StoredProcedure [dbo].[ListaConsigTerminado]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaConsigTerminado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaConsigTerminado]
GO
/****** Object:  StoredProcedure [dbo].[ListaConsigTotal]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaConsigTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaConsigTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaCotiza]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCotiza]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCotiza]
GO
/****** Object:  StoredProcedure [dbo].[ListaCotizaArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCotizaArticuloDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCotizaArticuloDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaCotizaNumero]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCotizaNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCotizaNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaCotizaProveedor]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCotizaProveedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCotizaProveedor]
GO
/****** Object:  StoredProcedure [dbo].[ListaCotizaProveedorDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCotizaProveedorDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCotizaProveedorDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaCotizaTotal]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCotizaTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCotizaTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaCtaCte]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtaCte]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCtaCte]
GO
/****** Object:  StoredProcedure [dbo].[ListaCtacteCliente]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtacteCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCtacteCliente]
GO
/****** Object:  StoredProcedure [dbo].[ListaCtacteDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtacteDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCtacteDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaCtacteDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtacteDesdeHastaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCtacteDesdeHastaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaCtacteFecha]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtacteFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCtacteFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaCtaCtePrv]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtaCtePrv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCtaCtePrv]
GO
/****** Object:  StoredProcedure [dbo].[ListaCtaprvDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtaprvDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCtaprvDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaCtaprvDesdeHastaSaldo]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtaprvDesdeHastaSaldo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCtaprvDesdeHastaSaldo]
GO
/****** Object:  StoredProcedure [dbo].[ListaCtaprvDesdeHastaSaldoTotal]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtaprvDesdeHastaSaldoTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCtaprvDesdeHastaSaldoTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaCuentas]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCuentas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaCuentas]
GO
/****** Object:  StoredProcedure [dbo].[ListaDepositos]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaDepositos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaDepositos]
GO
/****** Object:  StoredProcedure [dbo].[ListaDepositosConsulta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaDepositosConsulta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaDepositosConsulta]
GO
/****** Object:  StoredProcedure [dbo].[ListaDepositosFecha]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaDepositosFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaDepositosFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaDepositosMovban]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaDepositosMovban]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaDepositosMovban]
GO
/****** Object:  StoredProcedure [dbo].[ListaDevcon]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaDevcon]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaDevcon]
GO
/****** Object:  StoredProcedure [dbo].[ListaDevconCliente]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaDevconCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaDevconCliente]
GO
/****** Object:  StoredProcedure [dbo].[ListaDevconNumero]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaDevconNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaDevconNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaEnsayos]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEnsayos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEnsayos]
GO
/****** Object:  StoredProcedure [dbo].[ListaEntdev]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEntdev]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEntdev]
GO
/****** Object:  StoredProcedure [dbo].[ListaEntdevNumero]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEntdevNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEntdevNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaEntdevRepro]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEntdevRepro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEntdevRepro]
GO
/****** Object:  StoredProcedure [dbo].[ListaEntdevTerminadoDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEntdevTerminadoDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEntdevTerminadoDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaEntdevTotal]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEntdevTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEntdevTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaEnvases]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEnvases]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEnvases]
GO
/****** Object:  StoredProcedure [dbo].[ListaEspecif]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEspecif]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEspecif]
GO
/****** Object:  StoredProcedure [dbo].[ListaEspecificaciones]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEspecificaciones]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEspecificaciones]
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaArticuloDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEstadisticaArticuloDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaArticuloDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaArticuloDesdeHastaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEstadisticaArticuloDesdeHastaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaCliente]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEstadisticaCliente]
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEstadisticaDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaDesdeHastaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEstadisticaDesdeHastaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaFecha]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEstadisticaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaRepro]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaRepro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEstadisticaRepro]
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaReproDy]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaReproDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaEstadisticaReproDy]
GO
/****** Object:  StoredProcedure [dbo].[ListaFeriado]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaFeriado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaFeriado]
GO
/****** Object:  StoredProcedure [dbo].[ListaGasimpo]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaGasimpo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaGasimpo]
GO
/****** Object:  StoredProcedure [dbo].[ListaHoja]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHoja]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHoja]
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaArticuloDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHojaArticuloDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaArticuloDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaArticuloDesdeHastaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHojaArticuloDesdeHastaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaDesdeHastaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHojaDesdeHastaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaDesdeHastaFefcha]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaDesdeHastaFefcha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHojaDesdeHastaFefcha]
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaNumero]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHojaNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaProducto]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaProducto]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHojaProducto]
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaProductoDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaProductoDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHojaProductoDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaProductoDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaProductoDesdeHastaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHojaProductoDesdeHastaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaRepro]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaRepro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHojaRepro]
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaRepro1]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaRepro1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHojaRepro1]
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaRepro2]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaRepro2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHojaRepro2]
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaReProceso]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaReProceso]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHojaReProceso]
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaTerminadoDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaTerminadoDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHojaTerminadoDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaTotal]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaHojaTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaImputac]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaImputac]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaImputac]
GO
/****** Object:  StoredProcedure [dbo].[ListaImputacDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaImputacDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaImputacDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaInforme]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInforme]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaInforme]
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeArticulo]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaInformeArticulo]
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeDesdeHastaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaInformeDesdeHastaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeListado]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeListado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaInformeListado]
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeNumero]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaInformeNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeOrden]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeOrden]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaInformeOrden]
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeOrdenArticulo]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeOrdenArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaInformeOrdenArticulo]
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeTotal]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaInformeTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeTotalDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeTotalDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaInformeTotalDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaInsumos]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInsumos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaInsumos]
GO
/****** Object:  StoredProcedure [dbo].[ListaInventario]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInventario]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaInventario]
GO
/****** Object:  StoredProcedure [dbo].[ListaInventarioNumero]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInventarioNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaInventarioNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaInventarioTotal]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInventarioTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaInventarioTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaIvaComp]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaIvaComp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaIvaComp]
GO
/****** Object:  StoredProcedure [dbo].[ListaIvaCompDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaIvaCompDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaIvaCompDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaIvaCompMenor]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaIvaCompMenor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaIvaCompMenor]
GO
/****** Object:  StoredProcedure [dbo].[ListaIvaCompNumero]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaIvaCompNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaIvaCompNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaIvacompTotal]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaIvacompTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaIvacompTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudo]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLaudo]
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoArticulo]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLaudoArticulo]
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoArticuloDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLaudoArticuloDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoArticuloDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoArticuloDesdeHastaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLaudoArticuloDesdeHastaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoArticuloPartiOri]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoArticuloPartiOri]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLaudoArticuloPartiOri]
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoDevol]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoDevol]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLaudoDevol]
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoDy]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLaudoDy]
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoFecha]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLaudoFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoInforme]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoInforme]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLaudoInforme]
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoNumero]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLaudoNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoOrden]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoOrden]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLaudoOrden]
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoPartiOri]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoPartiOri]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLaudoPartiOri]
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoRepro]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoRepro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLaudoRepro]
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoTotal]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLaudoTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaLinea]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLinea]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLinea]
GO
/****** Object:  StoredProcedure [dbo].[ListaLineaMp]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLineaMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaLineaMp]
GO
/****** Object:  StoredProcedure [dbo].[ListaMarcas]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMarcas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMarcas]
GO
/****** Object:  StoredProcedure [dbo].[ListaMarcasArticulo]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMarcasArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMarcasArticulo]
GO
/****** Object:  StoredProcedure [dbo].[ListaMarcasConsulta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMarcasConsulta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMarcasConsulta]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovenv]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovenv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovenv]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovenvDesdeHastaCliente]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovenvDesdeHastaCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovenvDesdeHastaCliente]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovenvDesdeHastaEnvases]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovenvDesdeHastaEnvases]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovenvDesdeHastaEnvases]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovenvTotal]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovenvTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovenvTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovgas]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovgas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovgas]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovgasTotal]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovgasTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovgasTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguia]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguia]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovguia]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaArticuloDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovguiaArticuloDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaArticuloDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaArticuloDesdeHastaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovguiaArticuloDesdeHastaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaLote]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaLote]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovguiaLote]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaLote1]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaLote1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovguiaLote1]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaLoteSolo]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaLoteSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovguiaLoteSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaNumero]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovguiaNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaRepro]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaRepro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovguiaRepro]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaRepro1]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaRepro1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovguiaRepro1]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaTerminadoDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaTerminadoDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovguiaTerminadoDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaTerminadoDesdeHastaFecha]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaTerminadoDesdeHastaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovguiaTerminadoDesdeHastaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovguiaTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlab]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlab]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovlab]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabArticuloDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovlabArticuloDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabArticuloDesdeHastaFecha]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabArticuloDesdeHastaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovlabArticuloDesdeHastaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabNumero]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovlabNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabRepro]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabRepro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovlabRepro]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabRepro1]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabRepro1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovlabRepro1]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabTerminadoDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabTerminadoDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovlabTerminadoDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabTerminadoDesdeHastaFecha]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabTerminadoDesdeHastaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovlabTerminadoDesdeHastaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovlabTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvar]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvar]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovvar]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarArticuloDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovvarArticuloDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarArticuloDesdeHastaFecha]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarArticuloDesdeHastaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovvarArticuloDesdeHastaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarNumero]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovvarNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarRepro]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarRepro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovvarRepro]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarRepro1]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarRepro1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovvarRepro1]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarTerminadoDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarTerminadoDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovvarTerminadoDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarTerminadoDesdeHastaFecha]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarTerminadoDesdeHastaFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovvarTerminadoDesdeHastaFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMovvarTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraArticulo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraArticulo]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraArticuloSolo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraArticuloSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraArticuloSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraCantidad]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraCantidad]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraCantidad]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraCantidadSolo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraCantidadSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraCantidadSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraCliente]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraCliente]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraClienteSolo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraClienteSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraClienteSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraDescriCliente]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraDescriCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraDescriCliente]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraDescriClienteSolo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraDescriClienteSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraDescriClienteSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraEnsayoSolo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraEnsayoSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraEnsayoSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraFecha]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraFechaSolo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraFechaSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraFechaSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraNombre]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraNombre]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraNombre]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraNombreSolo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraNombreSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraNombreSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraNumero]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraObservaciones]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraObservaciones]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraObservaciones]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraObservacionesSolo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraObservacionesSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraObservacionesSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraProductoSolo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraProductoSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraProductoSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaMuestraTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaOrden]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrden]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOrden]
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenArticulo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOrdenArticulo]
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenCarpeta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenCarpeta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOrdenCarpeta]
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOrdenDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenImpresion]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenImpresion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOrdenImpresion]
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenNumero]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOrdenNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenProveedor]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenProveedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOrdenProveedor]
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOrdenTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenTotalDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenTotalDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOrdenTotalDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaOtCliente]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOtCliente]
GO
/****** Object:  StoredProcedure [dbo].[ListaOtClienteSolo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtClienteSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOtClienteSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaOtFCompro]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtFCompro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOtFCompro]
GO
/****** Object:  StoredProcedure [dbo].[ListaOtFComproSolo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtFComproSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOtFComproSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaOtFecha]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOtFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaOtFechaSalida]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtFechaSalida]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOtFechaSalida]
GO
/****** Object:  StoredProcedure [dbo].[ListaOtFechaSolo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtFechaSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOtFechaSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaOtFSalidaSolo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtFSalidaSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOtFSalidaSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaOtNumero]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOtNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaOtObservaciones1]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtObservaciones1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOtObservaciones1]
GO
/****** Object:  StoredProcedure [dbo].[ListaOtObservaciones1Solo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtObservaciones1Solo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOtObservaciones1Solo]
GO
/****** Object:  StoredProcedure [dbo].[ListaOtSolicitante]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtSolicitante]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOtSolicitante]
GO
/****** Object:  StoredProcedure [dbo].[ListaOtSolicitanteSolo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtSolicitanteSolo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOtSolicitanteSolo]
GO
/****** Object:  StoredProcedure [dbo].[ListaOtTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaOtTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaPago]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPago]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPago]
GO
/****** Object:  StoredProcedure [dbo].[ListaPagos]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPagos]
GO
/****** Object:  StoredProcedure [dbo].[ListaPagosCarpeta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagosCarpeta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPagosCarpeta]
GO
/****** Object:  StoredProcedure [dbo].[ListaPagosCarpetaTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagosCarpetaTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPagosCarpetaTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaPagosConsulta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagosConsulta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPagosConsulta]
GO
/****** Object:  StoredProcedure [dbo].[ListaPagosConsultaII]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagosConsultaII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPagosConsultaII]
GO
/****** Object:  StoredProcedure [dbo].[ListaPagosFecha]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagosFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPagosFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaPagosMovban]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagosMovban]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPagosMovban]
GO
/****** Object:  StoredProcedure [dbo].[ListaPagosNumero]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagosNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPagosNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedido]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedido]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedido]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoCentro]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoCentro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoCentro]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoDevol]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoDevol]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoDevol]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoDevolFecha]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoDevolFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoDevolFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoDevolFechaMarca]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoDevolFechaMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoDevolFechaMarca]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoDevolNumero]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoDevolNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoDevolNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoDevolTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoDevolTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoDevolTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoDevolTotalListado]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoDevolTotalListado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoDevolTotalListado]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoFecha]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoFechaMarca]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoFechaMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoFechaMarca]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoFechaMarcaColor]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoFechaMarcaColor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoFechaMarcaColor]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoII]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoII]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoNumero]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoPend]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoPend]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoPend]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoPendDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoPendDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoPendDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoPigmentos]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoPigmentos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoPigmentos]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTerminado]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTerminado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoTerminado]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTerminadoPendiente]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTerminadoPendiente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoTerminadoPendiente]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTotalListado]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTotalListado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoTotalListado]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTotalListado1]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTotalListado1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoTotalListado1]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTotalListado2]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTotalListado2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoTotalListado2]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTotalListado3]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTotalListado3]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoTotalListado3]
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTotalListado4]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTotalListado4]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPedidoTotalListado4]
GO
/****** Object:  StoredProcedure [dbo].[ListaPrecios]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPrecios]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPrecios]
GO
/****** Object:  StoredProcedure [dbo].[ListaPreciosCliente]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPreciosCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPreciosCliente]
GO
/****** Object:  StoredProcedure [dbo].[ListaPreciosClienteMp]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPreciosClienteMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPreciosClienteMp]
GO
/****** Object:  StoredProcedure [dbo].[ListaPreciosMp]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPreciosMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPreciosMp]
GO
/****** Object:  StoredProcedure [dbo].[ListaPrestamoTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPrestamoTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPrestamoTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaProveedores]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaProveedores]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaProveedores]
GO
/****** Object:  StoredProcedure [dbo].[ListaProveedoresOrd]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaProveedoresOrd]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaProveedoresOrd]
GO
/****** Object:  StoredProcedure [dbo].[ListaProveedoresOrdConsulta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaProveedoresOrdConsulta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaProveedoresOrdConsulta]
GO
/****** Object:  StoredProcedure [dbo].[ListaProveedoresOrdConsultaII]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaProveedoresOrdConsultaII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaProveedoresOrdConsultaII]
GO
/****** Object:  StoredProcedure [dbo].[ListaPrueba]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPrueba]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPrueba]
GO
/****** Object:  StoredProcedure [dbo].[ListaPruebaConsulta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPruebaConsulta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPruebaConsulta]
GO
/****** Object:  StoredProcedure [dbo].[ListaPrueter]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPrueter]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPrueter]
GO
/****** Object:  StoredProcedure [dbo].[ListaPrueterConsulta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPrueterConsulta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaPrueterConsulta]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibos]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibos]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosBusqueda]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosBusqueda]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosBusqueda]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosCartera]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosCartera]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosCartera]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosCliente]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosCliente]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDeposito]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDeposito]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosDeposito]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDifeI]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDifeI]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosDifeI]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDifeOtroI]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDifeOtroI]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosDifeOtroI]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDifeOtroII]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDifeOtroII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosDifeOtroII]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDifeOtroIII]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDifeOtroIII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosDifeOtroIII]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDifeOtroIV]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDifeOtroIV]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosDifeOtroIV]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDifeOtroV]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDifeOtroV]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosDifeOtroV]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDifeOtroVI]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDifeOtroVI]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosDifeOtroVI]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosFactura]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosFactura]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosFactura]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosFecha]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosFecha]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosMovban]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosMovban]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosMovban]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosNroCheque]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosNroCheque]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosNroCheque]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosNumero]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRecibosTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaRetencion]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRetencion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRetencion]
GO
/****** Object:  StoredProcedure [dbo].[ListaRubro]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRubro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaRubro]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolguia]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolguia]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolguia]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolGuiaNumero]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolGuiaNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolGuiaNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolGuiaPendiente]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolGuiaPendiente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolGuiaPendiente]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolGuiaTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolGuiaTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolGuiaTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolGuiaTotalTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolGuiaTotalTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolGuiaTotalTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolHoja]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolHoja]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolHoja]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolHojaNumero]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolHojaNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolHojaNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolHojaTotalListado]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolHojaTotalListado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolHojaTotalListado]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolicitud]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolicitud]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolicitud]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolicitudArticulo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolicitudArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolicitudArticulo]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolicitudBajaArticulo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolicitudBajaArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolicitudBajaArticulo]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolicitudNumero]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolicitudNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolicitudNumero]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolicitudPendiente]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolicitudPendiente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolicitudPendiente]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolicitudPendienteArticulo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolicitudPendienteArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolicitudPendienteArticulo]
GO
/****** Object:  StoredProcedure [dbo].[ListaSolicitudTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolicitudTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSolicitudTotal]
GO
/****** Object:  StoredProcedure [dbo].[ListaSoltot]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSoltot]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSoltot]
GO
/****** Object:  StoredProcedure [dbo].[ListaSoltotnan]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSoltotnan]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSoltotnan]
GO
/****** Object:  StoredProcedure [dbo].[ListaSoltotnan1]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSoltotnan1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSoltotnan1]
GO
/****** Object:  StoredProcedure [dbo].[ListaSoltotnan2]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSoltotnan2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaSoltotnan2]
GO
/****** Object:  StoredProcedure [dbo].[ListaTerminado]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaTerminado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaTerminado]
GO
/****** Object:  StoredProcedure [dbo].[ListaTerminadoArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaTerminadoArticuloDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaTerminadoArticuloDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaTerminadoConsulta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaTerminadoConsulta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaTerminadoConsulta]
GO
/****** Object:  StoredProcedure [dbo].[ListaTerminadoDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaTerminadoDesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaTerminadoDesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ListaTerminadoDesdeHastaMinimo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaTerminadoDesdeHastaMinimo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaTerminadoDesdeHastaMinimo]
GO
/****** Object:  StoredProcedure [dbo].[ListaVendedor]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaVendedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ListaVendedor]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticulo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticulo]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticulo1]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticulo1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticulo1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloCosto]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloCosto]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloCosto]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloCosto1]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloCosto1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloCosto1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloCostoDolares]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloCostoDolares]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloCostoDolares]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloCostoImportacion]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloCostoImportacion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloCostoImportacion]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloCostoImpre]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloCostoImpre]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloCostoImpre]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloCostoPesos]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloCostoPesos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloCostoPesos]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloCostoTotal]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloCostoTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloCostoTotal]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloDescriComercial]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloDescriComercial]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloDescriComercial]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloDy]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloDy]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloFacturas]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloFacturas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloFacturas]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloFecha]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloFecha]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloII]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloII]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloInforme]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloInforme]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloInforme]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloInformeLaboratorio]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloInformeLaboratorio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloInformeLaboratorio]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloInicial0]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloInicial0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloInicial0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloLaboratorio]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloLaboratorio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloLaboratorio]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloLaboratorio0]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloLaboratorio0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloLaboratorio0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloLaboratorio0DesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloLaboratorio0DesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloLaboratorio0DesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloLaudo]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloLaudo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloLaudo]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloLaudoDolares]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloLaudoDolares]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloLaudoDolares]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloLaudoPesos]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloLaudoPesos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloLaudoPesos]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloLeyenda]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloLeyenda]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloLeyenda]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloMinimo1]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloMinimo1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloMinimo1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloMovi]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloMovi]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloMovi]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloMovimientos]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloMovimientos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloMovimientos]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloMovimientosSuma]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloMovimientosSuma]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloMovimientosSuma]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloOrden]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloOrden]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloOrden]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloOrdenLaboratorio]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloOrdenLaboratorio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloOrdenLaboratorio]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloPedido]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloPedido]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloPedido]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloPedido0]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloPedido0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloPedido0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloPedido0DesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloPedido0DesdeHasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloPedido0DesdeHasta]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloSalidas]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloSalidas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloSalidas]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloStock]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloStock]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloStock]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloStock0]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloStock0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloStock0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloVarios]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloVarios]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloVarios]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloVarios1]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloVarios1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloVarios1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloVariosII]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloVariosII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloVariosII]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloVenta]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloVenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloVenta]
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloVenta0]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloVenta0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaArticuloVenta0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaBanco]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaBanco]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaBanco]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCambio]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCambio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCambio]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCambioAdm]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCambioAdm]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCambioAdm]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCliente]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCliente]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCliente1]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCliente1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCliente1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaClienteIb]    Script Date: 05/01/2016 17:47:05 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaClienteIb]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaClienteIb]
GO
/****** Object:  StoredProcedure [dbo].[ModificaClienteImporte0]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaClienteImporte0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaClienteImporte0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaComposicion]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaComposicion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaComposicion]
GO
/****** Object:  StoredProcedure [dbo].[ModificaComposicionCosto]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaComposicionCosto]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaComposicionCosto]
GO
/****** Object:  StoredProcedure [dbo].[ModificaComposicionDy]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaComposicionDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaComposicionDy]
GO
/****** Object:  StoredProcedure [dbo].[ModificaConsig]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaConsig]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaConsig]
GO
/****** Object:  StoredProcedure [dbo].[ModificaConsigFacturado]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaConsigFacturado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaConsigFacturado]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCotiza]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCotiza]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCotiza]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCotizaDy]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCotizaDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCotizaDy]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCotizaMoneda]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCotizaMoneda]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCotizaMoneda]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacte]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte1]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacte1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte10]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte10]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacte10]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte11]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte11]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacte11]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte12]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte12]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacte12]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte2]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacte2]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte3]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte3]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacte3]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte4]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte4]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacte4]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte5]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte5]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacte5]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte6]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte6]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacte6]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte7]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte7]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacte7]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte8]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte8]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacte8]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte9]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte9]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacte9]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtaCteCli]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtaCteCli]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtaCteCli]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteIb]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteIb]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacteIb]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteIbCiudad]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteIbCiudad]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacteIbCiudad]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteImporte]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteImporte]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacteImporte]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteImporte0]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteImporte0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacteImporte0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteImporteIb]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteImporteIb]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacteImporteIb]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteImporteIva0]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteImporteIva0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacteImporteIva0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteIva1]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteIva1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacteIva1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteIva2]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteIva2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacteIva2]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteIva3]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteIva3]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacteIva3]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteIva4]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteIva4]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacteIva4]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteSalva]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteSalva]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacteSalva]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteTipo1]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteTipo1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacteTipo1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteTipo2]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteTipo2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtacteTipo2]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtaprv]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtaprv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCtaprv]
GO
/****** Object:  StoredProcedure [dbo].[ModificaCuenta]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCuenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaCuenta]
GO
/****** Object:  StoredProcedure [dbo].[ModificaDepositoImpolista]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaDepositoImpolista]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaDepositoImpolista]
GO
/****** Object:  StoredProcedure [dbo].[ModificaDepositoImpolista0]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaDepositoImpolista0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaDepositoImpolista0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaDesccomp]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaDesccomp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaDesccomp]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEnsayos]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEnsayos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEnsayos]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEntdev]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEntdev]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEntdev]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEntdev2]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEntdev2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEntdev2]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEntdevMarca]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEntdevMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEntdevMarca]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEntdevMarcaant]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEntdevMarcaant]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEntdevMarcaant]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEnvases]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEnvases]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEnvases]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEspecif]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEspecif]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEspecif]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEspecificaciones]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEspecificaciones]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEspecificaciones]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEspecificacionesDy]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEspecificacionesDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEspecificacionesDy]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEspeCli]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEspeCli]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEspeCli]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEstadistica]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEstadistica]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEstadistica]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEstadisticaLinea]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEstadisticaLinea]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEstadisticaLinea]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEstadisticaMarca]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEstadisticaMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEstadisticaMarca]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEstadisticaMarcaant]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEstadisticaMarcaant]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEstadisticaMarcaant]
GO
/****** Object:  StoredProcedure [dbo].[ModificaEstadisticaSalva]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEstadisticaSalva]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaEstadisticaSalva]
GO
/****** Object:  StoredProcedure [dbo].[ModificaGasimpo]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaGasimpo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaGasimpo]
GO
/****** Object:  StoredProcedure [dbo].[ModificaGuiaDy]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaGuiaDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaGuiaDy]
GO
/****** Object:  StoredProcedure [dbo].[ModificaHoja]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHoja]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaHoja]
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaDy]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaHojaDy]
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaFechaOrd]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaFechaOrd]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaHojaFechaOrd]
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaMarca]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaHojaMarca]
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaMarcaant]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaMarcaant]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaHojaMarcaant]
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaSaldo]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaSaldo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaHojaSaldo]
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaSaldo1]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaSaldo1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaHojaSaldo1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaSaldo2]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaSaldo2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaHojaSaldo2]
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaSaldo3]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaSaldo3]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaHojaSaldo3]
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaSaldoCierre]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaSaldoCierre]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaHojaSaldoCierre]
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaWImporte]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaWImporte]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaHojaWImporte]
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaWImporte0]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaWImporte0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaHojaWImporte0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaInforme]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInforme]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaInforme]
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeDatos]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeDatos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaInformeDatos]
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeDy]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaInformeDy]
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeLaboratorio]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeLaboratorio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaInformeLaboratorio]
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeListado]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeListado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaInformeListado]
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeListadoII]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeListadoII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaInformeListadoII]
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformePartida]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformePartida]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaInformePartida]
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeProceso]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeProceso]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaInformeProceso]
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeProceso0]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeProceso0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaInformeProceso0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeProcesoDife]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeProcesoDife]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaInformeProcesoDife]
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeProcesoSaldo]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeProcesoSaldo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaInformeProcesoSaldo]
GO
/****** Object:  StoredProcedure [dbo].[ModificaInventario]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInventario]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaInventario]
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudo]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaLaudo]
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoDy]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaLaudoDy]
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoFechaOrd]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoFechaOrd]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaLaudoFechaOrd]
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoMarca]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaLaudoMarca]
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoMarcaant]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoMarcaant]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaLaudoMarcaant]
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoProceso1]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoProceso1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaLaudoProceso1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoProceso2]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoProceso2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaLaudoProceso2]
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoSaldo]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoSaldo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaLaudoSaldo]
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoSaldo1]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoSaldo1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaLaudoSaldo1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoSaldoCierre]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoSaldoCierre]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaLaudoSaldoCierre]
GO
/****** Object:  StoredProcedure [dbo].[ModificaLinea]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLinea]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaLinea]
GO
/****** Object:  StoredProcedure [dbo].[ModificaLineaMp]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLineaMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaLineaMp]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMarcas]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMarcas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMarcas]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMinimo]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMinimo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMinimo]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMinimoDife]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMinimoDife]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMinimoDife]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMinimoDifePlanta]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMinimoDifePlanta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMinimoDifePlanta]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMinimoDisponible]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMinimoDisponible]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMinimoDisponible]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMinimoPlanta]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMinimoPlanta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMinimoPlanta]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovenv]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovenv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovenv]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovenvMovi]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovenvMovi]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovenvMovi]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovgas]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovgas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovgas]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovgasImpoDerechos]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovgasImpoDerechos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovgasImpoDerechos]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovgasProceso]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovgasProceso]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovgasProceso]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovgasProceso0]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovgasProceso0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovgasProceso0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovguiaMarca]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovguiaMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovguiaMarca]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovguiaMarcaant]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovguiaMarcaant]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovguiaMarcaant]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovguiaSaldo]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovguiaSaldo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovguiaSaldo]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovguiaSaldoCierre]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovguiaSaldoCierre]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovguiaSaldoCierre]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovlab]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovlab]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovlab]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovlabDy]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovlabDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovlabDy]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovlabMarca]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovlabMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovlabMarca]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovlabMarcaant]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovlabMarcaant]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovlabMarcaant]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovvar]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovvar]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovvar]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovvarDy]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovvarDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovvarDy]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovvarMarca]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovvarMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovvarMarca]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovvarMarcaant]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovvarMarcaant]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMovvarMarcaant]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMuestraI]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMuestraI]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMuestraI]
GO
/****** Object:  StoredProcedure [dbo].[ModificaMuestraII]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMuestraII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaMuestraII]
GO
/****** Object:  StoredProcedure [dbo].[ModificaNumero]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaNumero]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaNumero]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrden]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrden]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOrden]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenDerechos]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenDerechos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOrdenDerechos]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenDy]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOrdenDy]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenFecha2]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenFecha2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOrdenFecha2]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenFechaLLegada]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenFechaLLegada]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOrdenFechaLLegada]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenImpresion]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenImpresion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOrdenImpresion]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenInforme]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenInforme]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOrdenInforme]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenLaboratorio]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenLaboratorio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOrdenLaboratorio]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenLaudo]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenLaudo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOrdenLaudo]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenLeyenda]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenLeyenda]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOrdenLeyenda]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenPrueba]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenPrueba]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOrdenPrueba]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenReproceso]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenReproceso]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOrdenReproceso]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenReproceso0]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenReproceso0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOrdenReproceso0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenSaldo]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenSaldo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOrdenSaldo]
GO
/****** Object:  StoredProcedure [dbo].[ModificaOt]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOt]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaOt]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPago]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPago]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPago]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPagos]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPagos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPagos]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPagosNroCertificado]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPagosNroCertificado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPagosNroCertificado]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPagosParidad]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPagosParidad]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPagosParidad]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedido]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedido]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedido]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoActualiza]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoActualiza]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoActualiza]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoActualizaExpo]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoActualizaExpo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoActualizaExpo]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoActualizaII]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoActualizaII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoActualizaII]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoAnulacion]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoAnulacion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoAnulacion]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoArticulo]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoArticulo]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoAutoriza]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoAutoriza]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoAutoriza]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoCantidad]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoCantidad]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoCantidad]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoCantidadSaldo]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoCantidadSaldo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoCantidadSaldo]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoDevolAnulacion]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoDevolAnulacion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoDevolAnulacion]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoDevolAutoriza]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoDevolAutoriza]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoDevolAutoriza]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoDevolImpresion]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoDevolImpresion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoDevolImpresion]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoEspecificaciones]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoEspecificaciones]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoEspecificaciones]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoFacturas]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoFacturas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoFacturas]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoFechaEntrega]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoFechaEntrega]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoFechaEntrega]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoImpresion]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoImpresion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoImpresion]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoImpresion1]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoImpresion1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoImpresion1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoImpresion3]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoImpresion3]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoImpresion3]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoMarca]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoMarca]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoNousar]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoNousar]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoNousar]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoParcial]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoParcial]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoParcial]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoPigmentos]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoPigmentos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoPigmentos]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoProceso1]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoProceso1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoProceso1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoProceso2]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoProceso2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoProceso2]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoTerminado]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoTerminado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoTerminado]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoTipoPedido]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoTipoPedido]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoTipoPedido]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoVersion]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoVersion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedidoVersion]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedpen]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedpen]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedpen]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedpen0]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedpen0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedpen0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedpenDy]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedpenDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPedpenDy]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrecios]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrecios]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPrecios]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrecios1]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrecios1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPrecios1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrecios2]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrecios2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPrecios2]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrecios3]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrecios3]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPrecios3]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPreciosFactura]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPreciosFactura]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPreciosFactura]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPreciosFacturaMp]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPreciosFacturaMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPreciosFacturaMp]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPreciosMp]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPreciosMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPreciosMp]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPreciosMp1]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPreciosMp1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPreciosMp1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPreciosMp3]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPreciosMp3]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPreciosMp3]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrestamoDy]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrestamoDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPrestamoDy]
GO
/****** Object:  StoredProcedure [dbo].[ModificaProveedor]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaProveedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaProveedor]
GO
/****** Object:  StoredProcedure [dbo].[ModificaProveedor1]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaProveedor1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaProveedor1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaProveedorIb]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaProveedorIb]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaProveedorIb]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrueart]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrueart]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPrueart]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrueartDy]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrueartDy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPrueartDy]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrueartObservaciones]    Script Date: 05/01/2016 17:47:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrueartObservaciones]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPrueartObservaciones]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrueter]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrueter]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPrueter]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrueterObservaciones]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrueterObservaciones]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPrueterObservaciones]
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrueterValores]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrueterValores]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaPrueterValores]
GO
/****** Object:  StoredProcedure [dbo].[ModificaReciboDifeI]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaReciboDifeI]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaReciboDifeI]
GO
/****** Object:  StoredProcedure [dbo].[ModificaReciboDifeOtroI]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaReciboDifeOtroI]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaReciboDifeOtroI]
GO
/****** Object:  StoredProcedure [dbo].[ModificaReciboDifeOtroII]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaReciboDifeOtroII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaReciboDifeOtroII]
GO
/****** Object:  StoredProcedure [dbo].[ModificaReciboDifeOtroV]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaReciboDifeOtroV]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaReciboDifeOtroV]
GO
/****** Object:  StoredProcedure [dbo].[ModificaReciboImpolista]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaReciboImpolista]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaReciboImpolista]
GO
/****** Object:  StoredProcedure [dbo].[ModificaReciboImpolista0]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaReciboImpolista0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaReciboImpolista0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaReciboParidad]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaReciboParidad]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaReciboParidad]
GO
/****** Object:  StoredProcedure [dbo].[ModificaRecibos]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaRecibos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaRecibos]
GO
/****** Object:  StoredProcedure [dbo].[ModificaRubro]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaRubro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaRubro]
GO
/****** Object:  StoredProcedure [dbo].[ModificaSolGuiaMarca]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaSolGuiaMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaSolGuiaMarca]
GO
/****** Object:  StoredProcedure [dbo].[ModificaSolHojaImpresion]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaSolHojaImpresion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaSolHojaImpresion]
GO
/****** Object:  StoredProcedure [dbo].[ModificaSolicitud]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaSolicitud]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaSolicitud]
GO
/****** Object:  StoredProcedure [dbo].[ModificaSolicitudEntregado]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaSolicitudEntregado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaSolicitudEntregado]
GO
/****** Object:  StoredProcedure [dbo].[ModificaSolicitudMarca]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaSolicitudMarca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaSolicitudMarca]
GO
/****** Object:  StoredProcedure [dbo].[ModificaSolicitudMarcaTotal]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaSolicitudMarcaTotal]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaSolicitudMarcaTotal]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminado]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminado]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoCosto]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoCosto]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoCosto]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoEntradas]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoEntradas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoEntradas]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoFacturas]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoFacturas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoFacturas]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoFecha]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoFecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoFecha]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoHoja]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoHoja]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoHoja]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoInicial0]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoInicial0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoInicial0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoMinimo1]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoMinimo1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoMinimo1]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoMovimientos]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoMovimientos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoMovimientos]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoPedido]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoPedido]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoPedido]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoPedido0]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoPedido0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoPedido0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoProceso]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoProceso]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoProceso]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoProceso0]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoProceso0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoProceso0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoSalidas]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoSalidas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoSalidas]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoStock]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoStock]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoStock]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoStock0]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoStock0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoStock0]
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoVersion]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoVersion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaTerminadoVersion]
GO
/****** Object:  StoredProcedure [dbo].[ModificaVendedor]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaVendedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ModificaVendedor]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorArticulo]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorArticulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorArticulo]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorBanco]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorBanco]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorBanco]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorCambio]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorCambio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorCambio]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorCambioAdm]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorCambioAdm]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorCambioAdm]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorCliente]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorCliente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorCliente]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorComposicion]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorComposicion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorComposicion]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorCuenta]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorCuenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorCuenta]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorEnsayos]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorEnsayos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorEnsayos]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorEnvases]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorEnvases]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorEnvases]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorEspecif]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorEspecif]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorEspecif]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorEspecificaciones]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorEspecificaciones]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorEspecificaciones]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorGasimpo]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorGasimpo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorGasimpo]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorLinea]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorLinea]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorLinea]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorLineaMp]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorLineaMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorLineaMp]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorPago]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorPago]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorPago]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorPrecios]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorPrecios]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorPrecios]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorPreciosMp]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorPreciosMp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorPreciosMp]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorProveedor]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorProveedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorProveedor]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorRubro]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorRubro]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorRubro]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorTerminado]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorTerminado]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorTerminado]
GO
/****** Object:  StoredProcedure [dbo].[PosteriorVendedor]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorVendedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PosteriorVendedor]
GO
/****** Object:  StoredProcedure [dbo].[Limpia0]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Limpia0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Limpia0]
GO
/****** Object:  StoredProcedure [dbo].[Limpia1]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Limpia1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Limpia1]
GO
/****** Object:  StoredProcedure [dbo].[Limpia2]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Limpia2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Limpia2]
GO
/****** Object:  StoredProcedure [dbo].[Limpia3]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Limpia3]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Limpia3]
GO
/****** Object:  StoredProcedure [dbo].[Limpia4]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Limpia4]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Limpia4]
GO
/****** Object:  StoredProcedure [dbo].[ArticuloMinimo0]    Script Date: 05/01/2016 17:47:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ArticuloMinimo0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ArticuloMinimo0]
GO
/****** Object:  StoredProcedure [dbo].[TerminadoMinimo0]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TerminadoMinimo0]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[TerminadoMinimo0]
GO
/****** Object:  StoredProcedure [dbo].[SP_ConsultaBancos]    Script Date: 05/01/2016 17:47:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SP_ConsultaBancos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[SP_ConsultaBancos]
GO
/****** Object:  StoredProcedure [dbo].[Empresas]    Script Date: 05/01/2016 17:47:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Empresas]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Empresas]
GO
