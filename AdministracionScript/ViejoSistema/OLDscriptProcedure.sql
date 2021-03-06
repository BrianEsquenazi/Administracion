USE [sufractanSA]
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
/****** Object:  StoredProcedure [dbo].[Empresas]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Empresas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'Create procedure [dbo].[Empresas]
as
select * from empresa' 
END
GO
/****** Object:  StoredProcedure [dbo].[SP_ConsultaBancos]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SP_ConsultaBancos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[SP_ConsultaBancos]
AS

SELECT * FROM Bancos' 
END
GO
/****** Object:  StoredProcedure [dbo].[TerminadoMinimo0]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TerminadoMinimo0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[TerminadoMinimo0] AS

UPDATE  Terminado
	SET
		Minimo  	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ArticuloMinimo0]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ArticuloMinimo0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ArticuloMinimo0] AS

UPDATE  Articulo
	SET
		Minimo  	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[Limpia4]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Limpia4]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[Limpia4]
as

UPDATE  Laudo
	SET
		Origen	 	= ""  ,
		PartiOri 		= ""  ,
		Envase = 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[Limpia3]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Limpia3]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[Limpia3]
as

UPDATE  Orden
	SET
		Origen	 	= ""' 
END
GO
/****** Object:  StoredProcedure [dbo].[Limpia2]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Limpia2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[Limpia2]
as

UPDATE  Articulo
	SET
		Venta	 	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[Limpia1]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Limpia1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[Limpia1]
as

UPDATE  Estadistica
	SET
		TipoProDy 	=  "T"  ,
		ArticuloDy	= ""' 
END
GO
/****** Object:  StoredProcedure [dbo].[Limpia0]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Limpia0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[Limpia0]
as

UPDATE  Pedido
	SET
		TipoPro 	=  "T"  ,
		Articulo	 	= ""  ,
		Descripcion = ""' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorVendedor]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorVendedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorVendedor]
	@Vendedor  int
 AS

SELECT * FROM Vendedor
WHERE
	Vendedor > @Vendedor

Order by Vendedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorTerminado]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorTerminado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorTerminado]
	@Codigo	char(12)
 AS

SELECT * FROM Terminado
WHERE
	Codigo > @Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorRubro]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorRubro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorRubro]
	@Rubro  int
 AS

SELECT * FROM Rubros
WHERE
	Rubro > @Rubro

Order by Rubro' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorProveedor]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorProveedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorProveedor]
	@Proveedor  Char(11)
 AS

SELECT * FROM Proveedor
WHERE
	Proveedor > @Proveedor

Order by Proveedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorPreciosMp]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorPreciosMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorPreciosMp]
	@Clave  Char(16)
 AS

SELECT clave, Cliente, ArticuloFROM PreciosMp
WHERE
	Clave > @Clave

Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorPrecios]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorPrecios]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorPrecios]
	@Clave  Char(18)
 AS

SELECT clave, Cliente, terminadoFROM Precios
WHERE
	Clave > @Clave

Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorPago]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorPago]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorPago]
	@Pago  int
 AS

SELECT * FROM Pago
WHERE
	Pago > @Pago

Order by Pago' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorLineaMp]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorLineaMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorLineaMp]
	@Linea  int
 AS

SELECT * FROM LineasMp
WHERE
	Linea > @Linea

Order by Linea' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorLinea]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorLinea]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorLinea]
	@Linea  int
 AS

SELECT * FROM Lineas
WHERE
	Linea > @Linea

Order by Linea' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorGasimpo]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorGasimpo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorGasimpo]
	@Codigo  int
 AS

SELECT * FROM Gasimpo
WHERE
	Codigo > @Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorEspecificaciones]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorEspecificaciones]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorEspecificaciones]
	@Producto  char(10)
 AS

SELECT * FROM Especificaciones
WHERE
	Producto > @Producto

Order by Producto' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorEspecif]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorEspecif]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorEspecif]
	@Producto  char(12)
 AS

SELECT * FROM Especif
WHERE
	Producto > @Producto

Order by Producto' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorEnvases]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorEnvases]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorEnvases]
	@Envases  int
 AS

SELECT * FROM Envases
WHERE
	Envases > @Envases

Order by Envases' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorEnsayos]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorEnsayos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorEnsayos]
	@Codigo  int
 AS

SELECT * FROM Ensayos
WHERE
	Codigo > @Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorCuenta]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorCuenta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorCuenta]
	@Cuenta  int
 AS

SELECT * FROM Cuenta
WHERE
	Cuenta > @Cuenta

Order by Cuenta' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorComposicion]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorComposicion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorComposicion]
	@Clave char(14)
 AS

SELECT * FROM Composicion
WHERE
	Clave > @Clave

Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorCliente]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorCliente]
	@Cliente char(6)
 AS

SELECT * FROM Cliente
WHERE
	Cliente > @Cliente

Order by Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorCambioAdm]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorCambioAdm]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorCambioAdm]
	@Fecha char(10)
 AS

SELECT * FROM CambioAdm
WHERE
	Fecha > @Fecha

Order by Fecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorCambio]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorCambio]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorCambio]
	@Fecha char(10)
 AS

SELECT * FROM Cambios
WHERE
	Fecha > @Fecha

Order by Fecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorBanco]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorBanco]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorBanco]
	@Banco smallint
 AS

SELECT * FROM Banco
WHERE
	Banco > @Banco

Order by Banco' 
END
GO
/****** Object:  StoredProcedure [dbo].[PosteriorArticulo]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PosteriorArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PosteriorArticulo]
	@Codigo	char(10)
 AS

SELECT * FROM Articulo
WHERE
	Codigo > @Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaVendedor]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaVendedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaVendedor]
             @Vendedor int,
	@Nombre char(50)


 AS

UPDATE  Vendedor
	SET
		Vendedor    	= @Vendedor,
		Nombre       	= @Nombre

WHERE
	Vendedor = @Vendedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoVersion]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoVersion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoVersion]
             @Codigo char(12),
	@Version int,
	@FechaVersion char(10)
 AS

UPDATE  Terminado
	SET
		Codigo   	= @Codigo,
		Version      	= @Version,                                    
		FechaVersion    	= @FechaVersion                                   
	
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoStock0]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoStock0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoStock0]

 AS

UPDATE  Terminado
	SET
		Stock	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoStock]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoStock]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoStock]
             @Codigo char(12),
	@Stock float,
	@Costo float
 AS

UPDATE  Terminado
	SET
		Stock      	= @Stock,                                    
		Costo      	= @Costo
	
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoSalidas]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoSalidas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoSalidas]
             @Codigo char(12),
	@Salidas float,
	@WDate char(10)
 AS

UPDATE  Terminado
	SET
		Codigo   	= @Codigo,
		Salidas      	= @Salidas,                                    
		WDate      	= @WDate
	
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoProceso0]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoProceso0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoProceso0]  AS

UPDATE  Terminado
	SET
		Proceso		= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoProceso]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoProceso]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoProceso]
             @Codigo char(12),
	@Proceso float
 AS

UPDATE  Terminado
	SET
		Proceso      	= @Proceso + Proceso                                    
	
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoPedido0]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoPedido0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoPedido0]  AS

UPDATE  Terminado
	SET
		Pedido		= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoPedido]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoPedido]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoPedido]
             @Codigo char(12),
	@Pedido float,
	@WDate char(10)
 AS

UPDATE  Terminado
	SET
		Codigo   	= @Codigo,
		Pedido      	= @Pedido,                                    
		WDate      	= @WDate
	
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoMovimientos]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoMovimientos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoMovimientos]
             @Codigo char(12),
	@Entradas float,
	@Salidas float,
	@WDate char(10)
 AS

UPDATE  Terminado
	SET
		Codigo   	= @Codigo,
		Entradas      	= @Entradas,                                    
		Salidas      	= @Salidas,                                    
		WDate      	= @WDate
	
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoMinimo1]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoMinimo1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoMinimo1]
             @Codigo char(12),
	@Minimo1 Float
 AS

UPDATE  Terminado
	SET
		Codigo   	= @Codigo,
		Minimo1      	= @Minimo1
	
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoInicial0]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoInicial0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoInicial0]  AS

UPDATE  Terminado
	SET
		Inicial	  	= 0  ,
		Proceso		= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoHoja]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoHoja]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoHoja]
             @Codigo char(12),
	@Entradas float,
	@Proceso float,
	@WDate char(10)
 AS

UPDATE  Terminado
	SET
		Codigo   	= @Codigo,
		Entradas      	= @Entradas,                                    
		Proceso      	= @Proceso,                                    
		WDate      	= @WDate
	
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoFecha]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoFecha]

	@FechaCierre  char(10),
	@OrdFechaCierre char(8)
 AS

UPDATE  Terminado
	SET
		FechaCierre     	 	= @FechaCierre,                                    
		OrdFechaCierre      	= @OrdFechaCierre' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoFacturas]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoFacturas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoFacturas]
             @Codigo char(12),
	@Pedido float,
	@Salidas float,
	@WDate char(10)
 AS

UPDATE  Terminado
	SET
		Codigo   	= @Codigo,
		Pedido      	= @Pedido,                                    
		Salidas      	= @Salidas,                                    
		WDate      	= @WDate
	
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoEntradas]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoEntradas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoEntradas]
             @Codigo char(12),
	@Entradas float,
	@WDate char(10)
 AS

UPDATE  Terminado
	SET
		Codigo   	= @Codigo,
		Entradas      	= @Entradas,                                    
		WDate      	= @WDate
	
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminadoCosto]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminadoCosto]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminadoCosto]
             @Codigo char(12),
	@Costo Float
 AS

UPDATE  Terminado
	SET
		Codigo   	= @Codigo,
		Costo      	= @Costo
	
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaTerminado]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaTerminado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaTerminado]
             @Codigo char(12),
	@Descripcion char(30),
	@Linea int,
	@Unidad char(10),
	@Inicial float,
	@Entradas float,
	@Salidas float,
	@Minimo float,
	@Deposito char(10),
	@Pedido float,
	@Envase1 int,
	@Envase2 int,
	@Envase3 int,
	@Envase4 int,
	@Envase5 int,
	@Envase6 int,
	@Proceso float,
	@Costo Float,
	@Factor float,
	@WDate char(10),
	@Impreadi char(1),
	@Clase char(30),
	@Intervencion char(10),
	@Naciones char(10),
	@Embalaje char(10),
	@Version int,
	@FechaVersion char(10),
	@Controla int,
	@Observaciones char(50)  ,
	@Tipoeti char(50)   ,
	@Escrito  int

 AS

UPDATE  Terminado
	SET
		Codigo   	= @Codigo,
		Descripcion       	= @Descripcion,                                     
		Linea      	= @Linea,                                    
		Unidad      	= @Unidad,                                    
		Inicial      	= @Inicial,                                    
		Entradas      	= @Entradas,                                    
		Salidas      	= @Salidas,                                    
		Minimo      	= @Minimo,                                    
		Deposito      	= @Deposito,                                    
		Pedido      	= @Pedido,                                    
		Envase1      	= @Envase1,                                    
		Envase2      	= @Envase2,                                    
		Envase3      	= @Envase3,                                    
		Envase4      	= @Envase4,                                    
		Envase5      	= @Envase5,                                    
		Envase6      	= @Envase6,                                    
		Proceso      	= @Proceso,                                    
		Costo      	= @Costo,                                    
		Factor      	= @Factor,                                    
		WDate      	= @WDate,
		Impreadi      	= @Impreadi,                                    
		Clase      	= @Clase,                                    
		Intervencion      	= @Intervencion,                                    
		Naciones      	= @Naciones,                                    
		Embalaje      	= @Embalaje,                                    
		Version      	= @Version,                                    
		FechaVersion    	= @FechaVersion      ,
		Controla	             = @Controla  ,
		Observaciones 	= @Observaciones  ,
		Tipoeti		= @Tipoeti   ,
		Escrito		= @Escrito
	
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaSolicitudMarcaTotal]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaSolicitudMarcaTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaSolicitudMarcaTotal]
             @Solicitud int,
	@Marca char(1)

 AS

UPDATE  Solic
	SET
		Marca 	= @Marca  ,
		Entregado = Cantidad

Where
	Solicitud = @Solicitud' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaSolicitudMarca]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaSolicitudMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaSolicitudMarca]
             @Solicitud int,
	@Articulo char(10) ,
	@Marca char(1)

 AS

UPDATE  Solic
	SET
		Marca 	= @Marca  ,
		Entregado = Cantidad

Where
	Solicitud = @Solicitud And Articulo = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaSolicitudEntregado]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaSolicitudEntregado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaSolicitudEntregado]
             @Clave Char(8),
	@Entregado float ,
	@Marca char(1)

 AS

UPDATE  Solic
	SET
		Entregado = @Entregado  ,
		Marca = @Marca

Where
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaSolicitud]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaSolicitud]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaSolicitud]
             @Clave char(8),
	@Solicitud int,
	@Renglon int,
	@Fecha char(10),
	@Fechaord char(8), 
	@Observaciones char(100), 
	@Articulo char(10),
	@Cantidad float,
	@Entrega char(10),
	@OrdEntrega char(8),
	@Planta char(30),
	@Solicitante char(30),
	@WDate Char(10),
	@Marca  char(1) ,
	@Obser char(50),	
	@Entregado float


 AS

UPDATE  Solic
	SET
		Clave   		= @Clave,
		Solicitud	= @Solicitud,
		Renglon	= @Renglon,   
		Fecha 		= @Fecha,	
		Fechaord	= @Fechaord,	
		Observaciones 	= @Observaciones,	
		Articulo  	= @Articulo,	
		Cantidad	= @Cantidad,	
		Entrega 	= @Entrega,	
		OrdEntrega 	= @OrdEntrega,	
		Planta          	= @Planta,	
		Solicitante	= @Solicitante  ,
		WDate		= @WDate  ,
		Marca		= @Marca   ,
		Obser		= @Obser  ,
		Entregado          = @Entregado

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaSolHojaImpresion]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaSolHojaImpresion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaSolHojaImpresion]
             @Hoja int,
	@Marca char(1)

 AS

UPDATE  SolHoja
	SET
		Marca 	= @Marca


WHERE
	Hoja = @Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaSolGuiaMarca]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaSolGuiaMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaSolGuiaMarca]
             @Clave char(8) ,
	@Marca char(1)

 AS

UPDATE  SolGuia
	SET
		Marca 	= @Marca 

Where
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaRubro]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaRubro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaRubro]
             @Rubro int,
	@Nombre char(50)


 AS

UPDATE  Rubros
	SET
		Rubro    	= @Rubro,
		Nombre       	= @Nombre

WHERE
	Rubro = @Rubro' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaRecibos]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaRecibos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'Create procedure [dbo].[ModificaRecibos]
	@Clave   	varchar(8), 
	@Recibo  	varchar(6),
	@Renglon 	varchar(2),
	@Cliente 	varchar(6),
	@Fecha      	varchar(10),
	@Fechaord 	varchar(8),
	@TipoRec 	varchar(1),
	@RetGanancias   real,          
	@RetIva         real,          
	@RetOtra        real,          
	@Retencion      real,          
	@TipoReg        varchar(1),
	@Tipo1 		varchar(2),
	@Letra1 	varchar(1),
	@Punto1 	varchar(4),
	@Numero1  	varchar(8),
	@Importe1       real,          
	@Tipo2 		varchar(2),
	@Numero2  	varchar(8),
	@Fecha2     	varchar(10),
	@banco2         varchar(20),
	@Importe2       real,          
	@Estado2 	varchar(1),
	@Empresa 	integer,
	@FechaOrd2 	varchar(8),
	@Importe        float,                            
	@Observaciones  varchar(50),                                    
	@Impolist       float,                           
	@Impo1list      float,                                       
	@Destino        varchar(50),                                    
	@Cuenta     	varchar(10)
AS	
UPDATE 	Recibos
	SET
--		Recibo ,
		Renglon = @renglon,
		Cliente = @cliente,
		Fecha   = @fecha ,  
		Fechaord = @fechaord ,
		TipoRec = @tiporec,
		RetGanancias = @retganancias,
		RetIva   = @retiva    ,            
		RetOtra  = @retotra    ,            
		Retencion = @retencion   ,            
		TipoReg = @tiporeg, 
		Tipo1 = @tipo1,
		Letra1 = @letra1,
		Punto1 = @punto1,
		Numero1 = @numero1, 
		Importe1 = @importe1,                
		Tipo2 = @tipo2,
		Numero2 = @numero2,  
		Fecha2  = @fecha2,   
		banco2  = @banco2 ,            
		Importe2  = @importe2,               
		Estado2 = @estado2 ,
		Empresa = @empresa,
		FechaOrd2 = @fechaord2, 
		Importe  = @importe ,                                            
		Observaciones  = @observaciones,                                     
		Impolist = @impolist      ,                                       
		Impo1list = @impo1list     ,                                       
		Destino = @destino       ,                                    
		Cuenta  = @cuenta  
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaReciboParidad]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaReciboParidad]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaReciboParidad]
	@Paridad float
	
 AS

UPDATE  Recibos
	SET
		Paridad    	= @Paridad' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaReciboImpolista0]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaReciboImpolista0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaReciboImpolista0]  AS

UPDATE  Recibos
	SET
		Impolist  	= 0,
		Impo1list 	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaReciboImpolista]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaReciboImpolista]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaReciboImpolista]  
	@Desde char(8),
	@Hasta Char(8)

AS

UPDATE  Recibos
	SET
		Impolist  	= Importe1,
		Impo1list 	= Importe2

WHERE

	Fechaord  >= @Desde and FechaOrd <=  @Hasta' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaReciboDifeOtroV]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaReciboDifeOtroV]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaReciboDifeOtroV]
	@Clave char(8) ,  
	@Importe2 float ,
	@Paridad float 
	
 AS

UPDATE  Recibos
	SET
		Importe2	= @Importe2 ,
		Impo1list    	= @Paridad  

WHERE
	Clave= @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaReciboDifeOtroII]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaReciboDifeOtroII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaReciboDifeOtroII]
             @Clave char(8),
	@Paridad float 
	
 AS

UPDATE  Recibos
	SET
		Impo1list    	= @Paridad  

WHERE
	Clave= @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaReciboDifeOtroI]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaReciboDifeOtroI]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaReciboDifeOtroI]
             @Clave char(8),
	@Paridad float  ,
	@FechaFactura char(20)
	
 AS

UPDATE  Recibos
	SET
		Impolist    	= @Paridad  ,
		Banco2		= @Fechafactura

WHERE
	Clave= @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaReciboDifeI]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaReciboDifeI]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaReciboDifeI]
             @Clave char(8),
	@Paridad1 float  ,
	@Paridad2 float,
	@FechaFactura char(20)
	
 AS

UPDATE  Recibos
	SET
		Impolist    	= @Paridad1,
		Impo1list       	= @Paridad2  ,
		Banco2		= @Fechafactura

WHERE
	Clave= @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrueterValores]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrueterValores]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPrueterValores]
             @Prueba char(7),
	@Valor1 char(20),
	@Valor2 char(20),
	@Valor3 char(20),
	@Valor4 char(20),
	@Valor5 char(20),
	@Valor6 char(20),
	@Valor7 char(20),
	@Valor8 char(20),
	@Valor9 char(20),
	@Valor10 char(20),
	@Ensayo char(50),
	@Aspecto char(50),
	@Observaciones char(50),
	@Confecciono char(50),
	@WDate char(10)

 AS

UPDATE  Prueter
	SET
		Prueba   	= @Prueba,
		Valor1      	= @Valor1,                                    
		Valor2      	= @Valor2,                                    
		Valor3      	= @Valor3,                                    
		Valor4      	= @Valor4,                                    
		Valor5      	= @Valor5,                                    
		Valor6      	= @Valor6,                                    
		Valor7      	= @Valor7,                                    
		Valor8      	= @Valor8,                                    
		Valor9      	= @Valor9,                                    
		Valor10      	= @Valor10,                                    
		Ensayo      	= @Ensayo,                                    
		Aspecto 	= @Aspecto,
		Observaciones	= @Observaciones,
		Confecciono       = @Confecciono,       
		Wdate  		= @Wdate
WHERE
	Prueba = @Prueba' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrueterObservaciones]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrueterObservaciones]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPrueterObservaciones]
             @Prueba char(7),
	@Ensayo char(50),
	@Aspecto char(50),
	@Observaciones char(50),
	@Confecciono char(50)

 AS

UPDATE  Prueter
	SET
		Prueba   	= @Prueba,
		Ensayo      	= @Ensayo,                                    
		Aspecto 	= @Aspecto,
		Observaciones	= @Observaciones,
		Confecciono       = @Confecciono
WHERE
	Prueba = @Prueba' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrueter]    Script Date: 05/01/2016 17:47:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrueter]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPrueter]
             @Prueba char(7),
	@Producto char(12),
	@Fecha char(10),
	@Valor1 char(50),
	@Valor2 char(50),
	@Valor3 char(50),
	@Valor4 char(50),
	@Valor5 char(50),
	@Valor6 char(50),
	@Valor7 char(50),
	@Valor8 char(50),
	@Valor9 char(50),
	@Valor10 char(50),
	@Ensayo char(50),
	@Aspecto char(50),
	@Observaciones char(50),
	@Confecciono char(50),
	@LIberada float,
	@Lote int,
	@Rechazo int,
	@Fechaord char(8),
	@WDate char(10)

 AS

UPDATE  Prueter
	SET
		Prueba   	= @Prueba,
		Producto       	= @Producto,                                     
		Fecha        	= @Fecha,                                     
		Valor1      	= @Valor1,                                    
		Valor2      	= @Valor2,                                    
		Valor3      	= @Valor3,                                    
		Valor4      	= @Valor4,                                    
		Valor5      	= @Valor5,                                    
		Valor6      	= @Valor6,                                    
		Valor7      	= @Valor7,                                    
		Valor8      	= @Valor8,                                    
		Valor9      	= @Valor9,                                    
		Valor10      	= @Valor10,                                    
		Ensayo      	= @Ensayo,                                    
		Aspecto 	= @Aspecto,
		Observaciones	= @Observaciones,
		Confecciono       = @Confecciono,       
		Liberada	= @Liberada                      ,
		Lote         	= @Lote              ,
		Rechazo	= @Rechazo            ,                          
		Fechaord	= @Fechaord,
		Wdate  		= @Wdate
WHERE
	Prueba = @Prueba' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrueartObservaciones]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrueartObservaciones]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPrueartObservaciones]
             @Prueba char(7),
	@Ensayo char(50),
	@Aspecto char(50),
	@Observaciones char(50),
	@Confecciono char(50)

 AS

UPDATE  Prueart
	SET
		Prueba   	= @Prueba,
		Ensayo      	= @Ensayo,                                    
		Aspecto 	= @Aspecto,
		Observaciones	= @Observaciones,
		Confecciono       = @Confecciono
WHERE
	Prueba = @Prueba' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrueartDy]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrueartDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPrueartDy]
             @Articulo char(10),
	@ArticuloDy Char(10)

 AS

uPDATE  Prueart
	SET
		Producto   		= @ArticuloDy

		

WHERE
	Producto = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrueart]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrueart]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPrueart]
             @Prueba char(7),
	@Producto char(10),
	@Fecha char(10),
	@Orden char(6),
	@Valor1 char(50),
	@Valor2 char(50),
	@Valor3 char(50),
	@Valor4 char(50),
	@Valor5 char(50),
	@Valor6 char(50),
	@Valor7 char(50),
	@Valor8 char(50),
	@Valor9 char(50),
	@Valor10 char(50),
	@Ensayo char(50),
	@Aspecto char(50),
	@Observaciones char(50),
	@Observa2 char(50),
	@Confecciono char(50),
	@LIberada float,
	@Devuelta float,
	@Lote int,
	@Rechazo int,
	@Nueva char(1),
	@Fechaord char(8),
	@WDate char(10)

 AS

UPDATE  Prueart
	SET
		Prueba   	= @Prueba,
		Producto       	= @Producto,                                     
		Fecha        	= @Fecha,                                     
		Orden        	= @Orden,                                     
		Valor1      	= @Valor1,                                    
		Valor2      	= @Valor2,                                    
		Valor3      	= @Valor3,                                    
		Valor4      	= @Valor4,                                    
		Valor5      	= @Valor5,                                    
		Valor6      	= @Valor6,                                    
		Valor7      	= @Valor7,                                    
		Valor8      	= @Valor8,                                    
		Valor9      	= @Valor9,                                    
		Valor10      	= @Valor10,                                    
		Ensayo      	= @Ensayo,                                    
		Aspecto      	= @Aspecto,                                    
		Observaciones	= @Observaciones,
		Observa2 	= @Observa2,
		Confecciono       = @Confecciono,       
		Liberada	= @Liberada                      ,
		Devuelta	= @Devuelta                      ,
		Lote         	= @Lote              ,
		Rechazo	= @Rechazo            ,                          
		Nueva		= @Nueva                      ,
		Fechaord	= @Fechaord,
		Wdate  		= @Wdate

WHERE
	Prueba = @Prueba' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaProveedorIb]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaProveedorIb]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaProveedorIb]  AS

UPDATE  Proveedor
	SET
		CodIb  	= 0   ,
		NroIb 	= " "' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaProveedor1]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaProveedor1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaProveedor1]
            @Proveedor char(11),
	@Nombre char(50),
	@Direccion char(50),
	@Localidad char(50),
	@Provincia char(2),
	@Postal char(4),
	@Cuit char(15),
	@Telefono    char(30),
	@EMail char(200),
	@Observaciones char(50),
	@Tipo char(1),
	@Iva char(1),
	@Dias char(20),
	@Empresa smallint,
	@Cuenta char(10),
	@Importe1 float,
	@Importe2 float,
	@Importe3 float,
	@Importe4 float,
	@Importe5 float,
	@Importe6 float,
	@NombreCheque char(50)  ,
	@WDate char(10)    ,
	@CodIb    int    ,
	@NroIb  char(20)  ,
	@NroInsc  char(15)
 AS

UPDATE  Proveedor
	SET
		Proveedor   	= @Proveedor,
		Nombre       	= @Nombre,                                     
		Direccion      	= @Direccion,                                    
		Localidad      	= @Localidad,                                    
		Provincia 	= @Provincia,
		Postal 		= @Postal,
		Cuit    		= @Cuit,       
		Telefono	= @Telefono                      ,
		Email         	= @Email               ,
		Observaciones	= @Observaciones            ,                          
		Tipo 		= @Tipo,
		Iva  		= @Iva,
		Dias		= @Dias,                
		Empresa	= @Empresa,                
		Cuenta 		= @Cuenta   ,
		Importe1	= @Importe1   ,
		Importe2	= @Importe2   ,
		Importe3	= @Importe3   ,
		Importe4	= @Importe4   ,
		Importe5	= @Importe5   ,
		Importe6	= @Importe6   ,
		NombreCheque 	= @NombreCheque  ,
		WDate		 = @WDate   ,
		CodIb		= @CodIb  ,
		NroIb		= @NroIb ,
		NroInsc		= @NroInsc

WHERE
	Proveedor = @Proveedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaProveedor]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaProveedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaProveedor]
            @Proveedor char(11),
	@Nombre char(50),
	@Direccion char(50),
	@Localidad char(50),
	@Provincia char(2),
	@Postal char(4),
	@Cuit char(15),
	@Telefono    char(30),
	@EMail char(200),
	@Observaciones char(50),
	@Tipo char(1),
	@Iva char(1),
	@Dias char(20),
	@Empresa smallint,
	@Cuenta char(10),
	@Importe1 float,
	@Importe2 float,
	@Importe3 float,
	@Importe4 float,
	@Importe5 float,
	@Importe6 float,
	@NombreCheque char(50)   ,
	@WDate char(10)    ,
	@CodIb   int   ,
	@NroIb  char(20)  ,
	@NroInsc  char(15)
 AS

UPDATE  Proveedor
	SET
		Proveedor   	= @Proveedor,
		Nombre       	= @Nombre,                                     
		Direccion      	= @Direccion,                                    
		Localidad      	= @Localidad,                                    
		Provincia 	= @Provincia,
		Postal 		= @Postal,
		Cuit    		= @Cuit,       
		Telefono	= @Telefono                      ,
		Email         	= @Email               ,
		Observaciones	= @Observaciones            ,                          
		Tipo 		= @Tipo,
		Iva  		= @Iva,
		Dias		= @Dias,                
		Empresa	= @Empresa,                
		Cuenta 		= @Cuenta   ,
		Importe1	= @Importe1   ,
		Importe2	= @Importe2   ,
		Importe3	= @Importe3   ,
		Importe4	= @Importe4   ,
		Importe5	= @Importe5   ,
		Importe6	= @Importe6   ,
		NombreCheque 	= @NombreCheque                                         ,
		WDate		= @WDate    ,
		CodIb		= @CodIb  ,
		NroIb		= @NroIb  ,
		NroInsc		= @NroInsc
WHERE
	Proveedor = @Proveedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrestamoDy]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrestamoDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPrestamoDy]
             @Articulo char(10),
	@ArticuloDy Char(10)

 AS

uPDATE  Prestamo
	SET
		Articulo   		= @ArticuloDy

		

WHERE
	Articulo = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPreciosMp3]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPreciosMp3]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPreciosMp3]
             @Clave char(16),
             @Precio  float,
	@Fecha char(10)
 AS

UPDATE  PreciosMp
	SET
		Clave   		= @Clave,
		Precio 		= @Precio   ,
		Fecha		= @Fecha


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPreciosMp1]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPreciosMp1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPreciosMp1]
             @Clave char(16),
             @Precio  float
 AS

UPDATE  PreciosMp
	SET
		Clave   		= @Clave,
		Precio 		= @Precio 


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPreciosMp]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPreciosMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPreciosMp]
             @Clave char(16),
             @Cliente char(6),
             @Articulo char(10),   
             @Precio  float,	
             @Fecha1 char(10),
             @Factura1 char(6),
             @Precio1 float,
             @Cantidad1  float,
             @Fecha2 char(10),
             @Factura2 char(6),
             @Precio2 float,
             @Cantidad2  float,
             @Fecha3 char(10),
             @Factura3 char(6),
             @Precio3 float,
             @Cantidad3  float,
             @Fecha4 char(10),
             @Factura4 char(6),
             @Precio4 float,
             @Cantidad4  float,
             @Fecha5 char(10),
             @Factura5 char(6),
             @Precio5 float,
             @Cantidad5  float,
             @WDate Char(10)  ,
	@Fecha char(10) ,
	@Pago int
	


 AS

UPDATE  PreciosMp
	SET
		Clave   		= @Clave,
		Cliente		= @Cliente,
		Articulo		= @Articulo,   
		Precio 		= @Precio,	
		Fecha1 		= @Fecha1,
		Factura1	= @Factura1,
		Precio1		= @Precio1,
		Cantidad1	= @Cantidad1,
		Fecha2 		= @Fecha2,
		Factura2	= @Factura2,
		Precio2		= @Precio2,
		Cantidad2	= @Cantidad2,
		Fecha3 		= @Fecha3,
		Factura3	= @Factura3,
		Precio3		= @Precio3,
		Cantidad3	= @Cantidad3,		
		Fecha4 		= @Fecha4,
		Factura4	= @Factura4,
		Precio4		= @Precio4,
		Cantidad4	= @Cantidad4,
		Fecha5 		= @Fecha5,
		Factura5	= @Factura5,
		Precio5		= @Precio5,
		Cantidad5	= @Cantidad5,
		WDate		=  @WDate  ,
		Fecha		=  @Fecha  ,
		Pago		= @Pago
		

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPreciosFacturaMp]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPreciosFacturaMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPreciosFacturaMp]
             @Clave char(16),
             @Fecha1 char(10),
             @Factura1 char(6),
             @Precio1 float,
             @Cantidad1  float,
             @Fecha2 char(10),
             @Factura2 char(6),
             @Precio2 float,
             @Cantidad2  float,
             @Fecha3 char(10),
             @Factura3 char(6),
             @Precio3 float,
             @Cantidad3  float,
             @Fecha4 char(10),
             @Factura4 char(6),
             @Precio4 float,
             @Cantidad4  float,
             @Fecha5 char(10),
             @Factura5 char(6),
             @Precio5 float,
             @Cantidad5  float,
             @WDate Char(10)


 AS

UPDATE  PreciosMp
	SET
		Clave   		= @Clave,
		Fecha1 		= @Fecha1,
		Factura1	= @Factura1,
		Precio1		= @Precio1,
		Cantidad1	= @Cantidad1,
		Fecha2 		= @Fecha2,
		Factura2	= @Factura2,
		Precio2		= @Precio2,
		Cantidad2	= @Cantidad2,
		Fecha3 		= @Fecha3,
		Factura3	= @Factura3,
		Precio3		= @Precio3,
		Cantidad3	= @Cantidad3,		
		Fecha4 		= @Fecha4,
		Factura4	= @Factura4,
		Precio4		= @Precio4,
		Cantidad4	= @Cantidad4,
		Fecha5 		= @Fecha5,
		Factura5	= @Factura5,
		Precio5		= @Precio5,
		Cantidad5	= @Cantidad5,
		WDate		=  @WDate

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPreciosFactura]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPreciosFactura]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPreciosFactura]
             @Clave char(18),
             @Fecha1 char(10),
             @Factura1 char(6),
             @Precio1 float,
             @Cantidad1  float,
             @Fecha2 char(10),
             @Factura2 char(6),
             @Precio2 float,
             @Cantidad2  float,
             @Fecha3 char(10),
             @Factura3 char(6),
             @Precio3 float,
             @Cantidad3  float,
             @Fecha4 char(10),
             @Factura4 char(6),
             @Precio4 float,
             @Cantidad4  float,
             @Fecha5 char(10),
             @Factura5 char(6),
             @Precio5 float,
             @Cantidad5  float,
             @WDate Char(10)


 AS

UPDATE  Precios
	SET
		Clave   		= @Clave,
		Fecha1 		= @Fecha1,
		Factura1	= @Factura1,
		Precio1		= @Precio1,
		Cantidad1	= @Cantidad1,
		Fecha2 		= @Fecha2,
		Factura2	= @Factura2,
		Precio2		= @Precio2,
		Cantidad2	= @Cantidad2,
		Fecha3 		= @Fecha3,
		Factura3	= @Factura3,
		Precio3		= @Precio3,
		Cantidad3	= @Cantidad3,		
		Fecha4 		= @Fecha4,
		Factura4	= @Factura4,
		Precio4		= @Precio4,
		Cantidad4	= @Cantidad4,
		Fecha5 		= @Fecha5,
		Factura5	= @Factura5,
		Precio5		= @Precio5,
		Cantidad5	= @Cantidad5,
		WDate		=  @WDate

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrecios3]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrecios3]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPrecios3]
             @Clave char(18),
             @Precio  float,
	@Fecha char(10)
 AS

UPDATE  Precios
	SET
		Clave   		= @Clave,
		Precio 		= @Precio   ,
		Fecha		= @Fecha


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrecios2]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrecios2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPrecios2]
             @Clave char(18),
             @Cliente char(6),
             @Terminado char(12),   
             @Precio  float,	
             @Descripcion char(50),
             @Fecha1 char(10),
             @Factura1 char(6),
             @Precio1 float,
             @Cantidad1  float,
             @Fecha2 char(10),
             @Factura2 char(6),
             @Precio2 float,
             @Cantidad2  float,
             @Fecha3 char(10),
             @Factura3 char(6),
             @Precio3 float,
             @Cantidad3  float,
             @Fecha4 char(10),
             @Factura4 char(6),
             @Precio4 float,
             @Cantidad4  float,
             @Fecha5 char(10),
             @Factura5 char(6),
             @Precio5 float,
             @Cantidad5  float,
             @WDate Char(10)  ,
	@Fecha char(10) ,
	@Pago int
	


 AS

UPDATE  Precios
	SET
		Clave   		= @Clave,
		Cliente		= @Cliente,
		Terminado	= @Terminado,   
		Precio 		= @Precio,	
		Descripcion	= @Descripcion,
		Fecha1 		= @Fecha1,
		Factura1	= @Factura1,
		Precio1		= @Precio1,
		Cantidad1	= @Cantidad1,
		Fecha2 		= @Fecha2,
		Factura2	= @Factura2,
		Precio2		= @Precio2,
		Cantidad2	= @Cantidad2,
		Fecha3 		= @Fecha3,
		Factura3	= @Factura3,
		Precio3		= @Precio3,
		Cantidad3	= @Cantidad3,		
		Fecha4 		= @Fecha4,
		Factura4	= @Factura4,
		Precio4		= @Precio4,
		Cantidad4	= @Cantidad4,
		Fecha5 		= @Fecha5,
		Factura5	= @Factura5,
		Precio5		= @Precio5,
		Cantidad5	= @Cantidad5,
		WDate		=  @WDate  ,
		Fecha		=  @Fecha  ,
		Pago		= @Pago
		

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrecios1]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrecios1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPrecios1]
             @Clave char(18),
             @Precio  float
 AS

UPDATE  Precios
	SET
		Clave   		= @Clave,
		Precio 		= @Precio 


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPrecios]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPrecios]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPrecios]
             @Clave char(18),
             @Cliente char(6),
             @Terminado char(12),   
             @Precio  float,	
             @Descripcion char(50),
             @Fecha1 char(10),
             @Factura1 char(6),
             @Precio1 float,
             @Cantidad1  float,
             @Fecha2 char(10),
             @Factura2 char(6),
             @Precio2 float,
             @Cantidad2  float,
             @Fecha3 char(10),
             @Factura3 char(6),
             @Precio3 float,
             @Cantidad3  float,
             @Fecha4 char(10),
             @Factura4 char(6),
             @Precio4 float,
             @Cantidad4  float,
             @Fecha5 char(10),
             @Factura5 char(6),
             @Precio5 float,
             @Cantidad5  float,
             @WDate Char(10) 
	


 AS

UPDATE  Precios
	SET
		Clave   		= @Clave,
		Cliente		= @Cliente,
		Terminado	= @Terminado,   
		Precio 		= @Precio,	
		Descripcion	= @Descripcion,
		Fecha1 		= @Fecha1,
		Factura1	= @Factura1,
		Precio1		= @Precio1,
		Cantidad1	= @Cantidad1,
		Fecha2 		= @Fecha2,
		Factura2	= @Factura2,
		Precio2		= @Precio2,
		Cantidad2	= @Cantidad2,
		Fecha3 		= @Fecha3,
		Factura3	= @Factura3,
		Precio3		= @Precio3,
		Cantidad3	= @Cantidad3,		
		Fecha4 		= @Fecha4,
		Factura4	= @Factura4,
		Precio4		= @Precio4,
		Cantidad4	= @Cantidad4,
		Fecha5 		= @Fecha5,
		Factura5	= @Factura5,
		Precio5		= @Precio5,
		Cantidad5	= @Cantidad5,
		WDate		=  @WDate 

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedpenDy]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedpenDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedpenDy]
             @Desde char(10),
             @Hasta char(10)

 AS

UPDATE  Pedido
	SET
		Importe		=Cantidad-Facturado


WHERE
	Articulo >= @Desde AND Articulo <= @Hasta' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedpen0]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedpen0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedpen0]

 AS

UPDATE  Pedido
	SET
		Importe		= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedpen]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedpen]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedpen]
             @Desdefec char(8),
             @Hastafec char(8)

 AS

UPDATE  Pedido
	SET
		Importe		=Cantidad-Facturado


WHERE
	OrdFecEntrega >= @DesdeFec AND OrdFecentrega <= @HastaFec' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoVersion]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoVersion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoVersion]
             @Pedido int

 AS

UPDATE  Pedido
	SET
		Autorizo		= "N",
		Impresion 	= "N"  ,
		Impresion1 	= "X"  ,
		Version		= Version + 1,
		Cantidad1	= 0  ,
		Cantidad2	= 0  ,
		Lote1		= 0  ,
		Cantilote1	= 0  ,
		Lote2		= 0  ,
		Cantilote2	= 0  ,
		Lote3		= 0  ,
		Cantilote3	= 0  ,
		Lote4		= 0  ,
		Cantilote4	= 0  ,
		Lote5		= 0  ,
		Cantilote5	= 0  ,
		Env1		= 0  ,
		CantiEnv1	= 0  ,
		Env2		= 0  ,
		CantiEnv2	= 0  ,
		Env3		= 0  ,
		CantiEnv3	= 0  ,
		Env4		= 0  ,
		CantiEnv4	= 0  ,
		Env5		= 0  ,
		CantiEnv5	= 0  

WHERE
	Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoTipoPedido]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoTipoPedido]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoTipoPedido]
             @Pedido int,
	@TipoPedido int

 AS

UPDATE  Pedido
	SET
		TipoPedido 	= @TipoPedido


WHERE
	Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoTerminado]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoTerminado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoTerminado]
             @Terminado char(12),
	@Descripcion char(50)

 AS

UPDATE  Pedido
	SET
		Descripcion 	= @Descripcion


WHERE
	Terminado = @Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoProceso2]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoProceso2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoProceso2]
             @Clave char(8),
	@Proceso1 int,
	@Proceso2 int,
	@ProcesoCanti float,
	@DesdeProceso int,
	@HastaProceso int ,
	@UsuarioProceso int  ,
	@Observa char(50)

 AS

UPDATE  Pedido
	SET
		Proceso1	= @Proceso1,
		Proceso2	= @Proceso2,   
		CantiProceso	= @ProcesoCanti  ,
		DesdeProceso  	= @DesdeProceso  ,
		HastaProceso	= @HastaProceso  ,
		UsuarioProceso	= @UsuarioProceso  ,
		Observa	= @Observa

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoProceso1]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoProceso1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoProceso1]
             @Pedido int ,
	@Proceso1 int

 AS

UPDATE  Pedido
	SET
		Proceso1	= @Proceso1

WHERE
	Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoPigmentos]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoPigmentos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoPigmentos]
             @Pedido int,
	@Impresion2 char(1)

 AS

UPDATE  Pedido
	SET
		Impresion2 	= @Impresion2


WHERE
	Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoParcial]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoParcial]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoParcial]
             @Clave char(8),
	@Cantidad1 float  ,
	@Cantidad2   float  ,
	@Lote1 int  ,
	@CantiLote1  float  ,
	@Lote2 int ,
	@CantiLote2  float  ,
	@Lote3 int  ,
	@CantiLote3  float  ,
	@Lote4 int  ,
	@CantiLote4  float  ,
	@Lote5 int  ,
	@CantiLote5  float  ,
	@Env1 int  ,
	@CantiEnv1  float  ,
	@Env2 int  ,
	@CantiEnv2  float  ,
	@Env3 int  ,
	@CantiEnv3  float  ,
	@Env4 int  ,
	@CantiEnv4  float  ,
	@Env5 int  ,
	@CantiEnv5  float  

 AS

UPDATE  Pedido
	SET
		Clave   		= @Clave,
		Cantidad1	= @Cantidad1,
		Cantidad2        	= @Cantidad2  ,
		Lote1        	= @Lote1  ,
		CantiLote1	= @CantiLote1  ,
		Lote2       	= @Lote2  ,
		CantiLote2	= @CantiLote2  ,
		Lote3        	= @Lote3  ,
		CantiLote3	= @CantiLote3  ,
		Lote4        	= @Lote4  ,
		CantiLote4	= @CantiLote4  ,
		Lote5        	= @Lote5  ,
		CantiLote5	= @CantiLote5  ,
		Env1        	= @Env1  ,
		CantiEnv1	= @CantiEnv1  ,
		Env2        	= @Env2  ,
		CantiEnv2	= @CantiEnv2  ,
		Env3        	= @Env3  ,
		CantiEnv3	= @CantiEnv3  ,
		Env4        	= @Env4  ,
		CantiEnv4	= @CantiEnv4  ,
		Env5        	= @Env5  ,
		CantiEnv5	= @CantiEnv5  


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoNousar]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoNousar]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoNousar]
             @Desdefec char(8),
             @Hastafec char(8)

 AS

UPDATE  Pedido
	SET
		Facturado =Cantidad  ,
		Autorizo = "S"  ,
		Impresion = "S"


WHERE
	Fechaord >= @DesdeFec AND Fechaord <= @HastaFec' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoMarca]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoMarca]
             @Pedido int

 AS

UPDATE  Pedido
	SET
		Autorizo		= "X",
		Impresion 	= "X" 

WHERE
	Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoImpresion3]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoImpresion3]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoImpresion3]
             @Pedido int,
	@Impresion3 char(1)

 AS

UPDATE  Pedido
	SET
		Impresion3 	= @Impresion3


WHERE
	Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoImpresion1]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoImpresion1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoImpresion1]
             @Pedido int,
	@Impresion1 char(1)

 AS

UPDATE  Pedido
	SET
		Impresion1 	= @Impresion1


WHERE
	Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoImpresion]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoImpresion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoImpresion]
             @Pedido int,
	@Impresion char(1)

 AS

UPDATE  Pedido
	SET
		Impresion 	= @Impresion


WHERE
	Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoFechaEntrega]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoFechaEntrega]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoFechaEntrega]
             @Clave char(8),
	@OrdFecEntrega char(8)

 AS

UPDATE  Pedido
	SET
		Clave   		= @Clave,
		OrdFecEntrega	= @OrdFecEntrega

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoFacturas]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoFacturas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoFacturas]
             @Clave char(8),
	@Facturado float

 AS

UPDATE  Pedido
	SET
		Facturado 	= @Facturado,
		Cantidad1	= 0  ,
		Cantidad2	= 0  ,
		Lote1		= 0  ,
		Cantilote1	= 0  ,
		Lote2		= 0  ,
		Cantilote2	= 0  ,
		Lote3		= 0  ,
		Cantilote3	= 0  ,
		Lote4		= 0  ,
		Cantilote4	= 0  ,
		Lote5		= 0  ,
		Cantilote5	= 0  ,
		Env1		= 0  ,
		CantiEnv1	= 0  ,
		Env2		= 0  ,
		CantiEnv2	= 0  ,
		Env3		= 0  ,
		CantiEnv3	= 0  ,
		Env4		= 0  ,
		CantiEnv4	= 0  ,
		Env5		= 0  ,
		CantiEnv5	= 0  



WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoEspecificaciones]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoEspecificaciones]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoEspecificaciones]
             @Clave char(8),
	@Especificaciones char(30)

 AS

UPDATE  Pedido
	SET
		Especificaciones = @Especificaciones

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoDevolImpresion]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoDevolImpresion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoDevolImpresion]
             @Pedido int,
	@Impresion char(1)

 AS

UPDATE  PedidoDevol
	SET
		Impresion 	= @Impresion


WHERE
	Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoDevolAutoriza]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoDevolAutoriza]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoDevolAutoriza]
             @Pedido int,
	@Autorizo char(1),
	@Fecha char(10)  ,
	@Fechaord char(8)

 AS

UPDATE  PedidoDevol
	SET
		Autorizo 	= @Autorizo  ,
		Fecha		= @Fecha  ,
		Fechaord	= @Fechaord
		

WHERE
	Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoDevolAnulacion]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoDevolAnulacion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoDevolAnulacion]
             @Pedido int

 AS

UPDATE  PedidoDevol
	SET
		Autorizo 	= "X"  ,
		Impresion 	= "X"  ,
		Facturado	= Cantidad


WHERE
	Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoCantidadSaldo]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoCantidadSaldo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoCantidadSaldo]
             @Clave char(8),
	@Cantidad float

 AS

UPDATE  Pedido
	SET
		Cantidad = @Cantidad,
		Facturado = 0

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoCantidad]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoCantidad]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoCantidad]
             @Clave char(8),
	@Cantidad float

 AS

UPDATE  Pedido
	SET
		Cantidad = @Cantidad

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoAutoriza]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoAutoriza]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoAutoriza]
             @Pedido int,
	@Autorizo char(1),
	@Fecha char(10)  ,
	@Fechaord char(8)

 AS

UPDATE  Pedido
	SET
		Autorizo 	= @Autorizo  ,
		Fecha		= @Fecha  ,
		Fechaord	= @Fechaord
		

WHERE
	Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoArticulo]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoArticulo]
             @Articulo char(10),
	@Descripcion char(50)

 AS

UPDATE  Pedido
	SET
		Descripcion 	= @Descripcion


WHERE
	Articulo = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoAnulacion]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoAnulacion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoAnulacion]
             @Pedido int

 AS

UPDATE  Pedido
	SET
		Autorizo 	= "X"  ,
		Impresion 	= "X"  ,
		Facturado	= Cantidad


WHERE
	Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoActualizaII]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoActualizaII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoActualizaII]
             @Clave char(8),
	@Cantidad1 float  ,
	@Cantidad2 float  ,
	@Lote1 int  ,
	@CantiLote1 float  ,
	@Lote2 int  ,
	@CantiLote2 float  ,
	@Lote3 int  ,
	@CantiLote3 float  ,
	@Lote4 int  ,
	@CantiLote4 float  ,
	@Lote5 int  ,
	@CantiLote5 float  ,
	@Env1 int  ,
	@CantiEnv1  float ,
	@Env2 int  ,
	@CantiEnv2  float ,
	@Env3 int  ,
	@CantiEnv3  float ,
	@Env4 int  ,
	@CantiEnv4  float ,
	@Env5 int  ,
	@CantiEnv5  float 


 AS

UPDATE  Pedido
	SET
		Clave   		= @Clave,
		Cantidad1	=@Cantidad1  ,
		Cantidad2	=@Cantidad2  ,
		Lote1		=@Lote1  ,
		CantiLote1	=@CantiLote1  ,
		Lote2		=@Lote2  ,
		CantiLote2	=@CantiLote2  ,
		Lote3		=@Lote3  ,
		CantiLote3	=@CantiLote3  ,
		Lote4		=@Lote4  ,
		CantiLote4	=@CantiLote4  ,
		Lote5		=@Lote5  ,
		CantiLote5	=@CantiLote5  ,
		Env1		=@Env1  ,
		CantiEnv1	=@CantiEnv1  ,
		Env2		=@Env2  ,
		CantiEnv2	=@CantiEnv2  ,
		Env3		=@Env3  ,
		CantiEnv3	=@CantiEnv3  ,
		Env4		=@Env4  ,
		CantiEnv4	=@CantiEnv4  ,
		Env5		=@Env5  ,
		CantiEnv5	=@CantiEnv5 

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoActualizaExpo]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoActualizaExpo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoActualizaExpo]
             @Clave char(8),
	@Cantidad1 float  ,
	@Cantidad2 float  ,
	@Lote1 int  ,
	@CantiLote1 float  ,
	@Lote2 int  ,
	@CantiLote2 float  ,
	@Lote3 int  ,
	@CantiLote3 float  ,
	@Lote4 int  ,
	@CantiLote4 float  ,
	@Lote5 int  ,
	@CantiLote5 float 

 AS

UPDATE  Pedido
	SET
		Clave   		= @Clave,
		Cantidad1	=@Cantidad1  ,
		Cantidad2	=@Cantidad2  ,
		Lote1		=@Lote1  ,
		CantiLote1	=@CantiLote1  ,
		Lote2		=@Lote2  ,
		CantiLote2	=@CantiLote2  ,
		Lote3		=@Lote3  ,
		CantiLote3	=@CantiLote3  ,
		Lote4		=@Lote4  ,
		CantiLote4	=@CantiLote4  ,
		Lote5		=@Lote5  ,
		CantiLote5	=@CantiLote5 


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedidoActualiza]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedidoActualiza]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedidoActualiza]
             @Clave char(8),
	@Cantidad1 float  ,
	@Cantidad2 float  ,
	@Lote1 int  ,
	@CantiLote1 float  ,
	@Lote2 int  ,
	@CantiLote2 float  ,
	@Lote3 int  ,
	@CantiLote3 float  ,
	@Lote4 int  ,
	@CantiLote4 float  ,
	@Lote5 int  ,
	@CantiLote5 float  ,
	@Env1 int  ,
	@CantiEnv1  float ,
	@Env2 int  ,
	@CantiEnv2  float ,
	@Env3 int  ,
	@CantiEnv3  float ,
	@Env4 int  ,
	@CantiEnv4  float ,
	@Env5 int  ,
	@CantiEnv5  float ,
	@Eti1 int  ,
	@Tipo1 char(1) ,
	@Eti2 int  ,
	@Tipo2 char(1)  ,
	@Eti3 int  ,
	@Tipo3 char(1)  ,
	@Eti4 int  ,
	@Tipo4 char(1)  ,
	@Eti5 int  ,
	@Tipo5 char(1)  



 AS

UPDATE  Pedido
	SET
		Clave   		= @Clave,
		Cantidad1	=@Cantidad1  ,
		Cantidad2	=@Cantidad2  ,
		Lote1		=@Lote1  ,
		CantiLote1	=@CantiLote1  ,
		Lote2		=@Lote2  ,
		CantiLote2	=@CantiLote2  ,
		Lote3		=@Lote3  ,
		CantiLote3	=@CantiLote3  ,
		Lote4		=@Lote4  ,
		CantiLote4	=@CantiLote4  ,
		Lote5		=@Lote5  ,
		CantiLote5	=@CantiLote5  ,
		Env1		=@Env1  ,
		CantiEnv1	=@CantiEnv1  ,
		Env2		=@Env2  ,
		CantiEnv2	=@CantiEnv2  ,
		Env3		=@Env3  ,
		CantiEnv3	=@CantiEnv3  ,
		Env4		=@Env4  ,
		CantiEnv4	=@CantiEnv4  ,
		Env5		=@Env5  ,
		CantiEnv5	=@CantiEnv5  ,
		Eti1		=@Eti1  ,
		Tipo1		=@Tipo1  ,
		Eti2		=@Eti2  ,
		Tipo2		=@Tipo2  ,
		Eti3		=@Eti3  ,
		Tipo3		=@Tipo3  ,
		Eti4		=@Eti4  ,
		Tipo4		=@Tipo4  ,
		Eti5		=@Eti5  ,
		Tipo5		=@Tipo5  


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPedido]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPedido]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPedido]
             @Clave char(8),
	@Pedido int,
	@Renglon int,
	@Cliente char(6),
	@Fecha char(10),
	@Fecentrega char(10),
	@Hora char(5),
	@Observaciones char(100),
	@Terminado char(12),
	@Cantidad float,
	@Envase1 int,
	@Canti1 int,
	@Envase2 int,
	@Canti2 int,
	@Envase3 int,
	@Canti3 int,
	@Envase4 int,
	@Canti4 int,
	@Fechaord char(8),
	@Precio float,
	@Linea int,
	@Facturado float,
	@Importe int

 AS

UPDATE  Pedido
	SET
		Clave   		= @Clave,
		Pedido		= @Pedido,
		Renglon	= @Renglon,   
		Cliente 		= @Cliente,	
		Fecha 		= @Fecha,	
		Fecentrega 	= @Fecentrega,	
		Hora		= @Hora,	
		Observaciones 	= @Observaciones,	
		Terminado 	= @Terminado,	
		Cantidad	= @Cantidad,	
		Envase1 	= @Envase1,	
		Canti1 		= @Canti1,	
		Envase2 	= @Envase2,	
		Canti2 		= @Canti2,	
		Envase3 	= @Envase3,	
		Canti3 		= @Canti3,	
		Envase4 	= @Envase4,	
		Canti4		= @Canti4,	
		Fechaord 	= @Fechaord,	
		Precio 		= @Precio,	
		Linea 		= @Linea,	
		Facturado 	= @Facturado,	
		Importe 		= @Importe


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPagosParidad]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPagosParidad]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPagosParidad]
             @Orden char(6),
	@Paridad Float

 AS

UPDATE  Pagos
	SET
		Orden     	= @Orden,
		Paridad		= @Paridad

WHERE
	Orden = @Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPagosNroCertificado]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPagosNroCertificado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPagosNroCertificado]
             @Orden char(6),
	@CertificadoGan  int  ,
	@CertificadoIb  int

 AS

UPDATE  Pagos
	SET
		Orden     	= @Orden,
		CertificadoGan   	= @CertificadoGan  ,
		CertificadoIb	= @CertificadoIb

WHERE
	Orden = @Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPagos]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPagos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPagos]
             @Orden char(6),
	@Carpeta int  ,
	@ImpoCarpeta  float  ,
	@Carpeta1 int  ,
	@ImpoCarpeta1  float  ,
	@Carpeta2 int  ,
	@ImpoCarpeta2 float  ,
	@Carpeta3 int  ,
	@ImpoCarpeta3  float  ,
	@Carpeta4 int  ,
	@ImpoCarpeta4  float


 AS

UPDATE  Pagos
	SET
		Orden     	= @Orden,
		Carpeta       	= @Carpeta  ,
		ImpoCarpeta	= @ImpoCarpeta  ,
		Carpeta1       	= @Carpeta1  ,
		ImpoCarpeta1	= @ImpoCarpeta1  ,
		Carpeta2      	= @Carpeta2  ,
		ImpoCarpeta2	= @ImpoCarpeta2  ,
		Carpeta3     	= @Carpeta3  ,
		ImpoCarpeta3	= @ImpoCarpeta3  ,
		Carpeta4       	= @Carpeta4  ,
		ImpoCarpeta4	= @ImpoCarpeta4  

WHERE
	Orden = @Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaPago]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaPago]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaPago]
             @Pago int,
	@Nombre char(50),
	@Dias real,	
	@Plazo real,	
	@Tasa real,
	@Descuento real


 AS

UPDATE  Pago
	SET
		Pago     	= @Pago,
		Nombre       	= @Nombre,
		Dias       	= @Dias,
		Plazo       	= @Plazo,
		Tasa       	= @Tasa,
	 	Descuento       	= @Descuento


WHERE
	Pago = @Pago' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOt]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOt]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOt]
             @Codigo int,
	@Fecha char(10),
	@Cliente char(6),
	@Razon char(50),
	@Preparacion char(50) ,
	@Solidez char(50) ,
	@Observaciones1 char(50)   ,
	@Observaciones2 char(50)   ,
	@Observaciones3  char(50)   ,
	@Solicitante  char(50)   ,
	@Compo char(50),
	@Compo1 int,
	@Compo2 int,
	@Compo3 int,
	@Compo4 int,
	@Compo5 int,
	@Compo6 int,
	@Compo7 int,
	@Compo8 int,
	@Compo9 int,
	@Compo10 int,
	@Compo11 int,
	@Compo12 int,
	@Compo13 int,
	@Compo14 int,
	@Traba char(50),
	@Trabajo1 int,
	@Trabajo2 int,
	@Trabajo3 int,
	@Trabajo4 int,
	@Trabajo5 int,
	@Trabajo6 int,
	@Trabajo7 int,
	@Trabajo8 int,
	@Trabajo9 int,
	@Trabajo10 int,
	@Trabajo11 int,
	@Trabajo12 int,
	@Trabajo13 int,
	@Trabajo14 int,
	@Color char(50),
	@Color1 int,
	@Color2 int,
	@Color3 int,
	@Color4 int,
	@Color5 int,
	@Color6 int,
	@Color7 int,
	@Color8 int,
	@Color9 int,
	@Color10 int,
	@Color11 int,
	@Color12 int,
	@Color13 int,
	@Color14 int,
	@Color15 int,
	@Color16 int,
	@Color17 int,
	@Color18 int,
	@Color19 int,
	@Color20 int,
	@Color21 int,
	@Maqui char(50),
	@Maquina1 int,
	@Maquina2 int,
	@Maquina3 int,
	@Maquina4 int,
	@Maquina5 int,
	@Maquina6 int,
	@Maquina7 int,
	@Maquina8 int,
	@Maquina9 int,
	@Maquina10 int,
	@Maquina11 int,
	@Maquina12 int,
	@Maquina13 int,
	@Maquina14 int,
	@FechaCompro char(10),
	@FechaSalida char(10),
	@OrdFecha char(8),
	@OrdFechaCompro char(8),
	@OrdFechaSalida char(8)  ,
	@Clave int

 AS

UPDATE  Ot
	SET
		Codigo   		= @Codigo,
		Fecha	  		= @Fecha,                                     
		Cliente 			= @Cliente,                                     
		Razon       		= @Razon,                                     
		Preparacion       		= @Preparacion,                                     
		Solidez       		= @Solidez,                                     
		Observaciones1		= @Observaciones1,                                     
		Observaciones2 	= @Observaciones2,                                     
		Observaciones3 	= @Observaciones3,                                     
		Solicitante  		= @Solicitante,                                     
		Compo       		= @Compo,                                     
		Compo1       		= @Compo1,                                     
		Compo2       		= @Compo2,                                     
		Compo3       		= @Compo3,                                     
		Compo4       		= @Compo4,                                     
		Compo5       		= @Compo5,                                     
		Compo6       		= @Compo6,                                     
		Compo7       		= @Compo7,                                     
		Compo8       		= @Compo8,                                     
		Compo9       		= @Compo9,                                     
		Compo10       		= @Compo10,                                     
		Compo11       		= @Compo11,                                     
		Compo12      		= @Compo12,                                     
		Compo13       		= @Compo13,                                     
		Compo14       		= @Compo14,                                     
		Traba       		= @Traba,                                     
		Trabajo1       		= @Trabajo1,                                     
		Trabajo2       		= @Trabajo2,                                     
		Trabajo3       		= @Trabajo3,                                     
		Trabajo4       		= @Trabajo4,                                     
		Trabajo5       		= @Trabajo5,                                     
		Trabajo6       		= @Trabajo6,                                     
		Trabajo7       		= @Trabajo7,                                     
		Trabajo8       		= @Trabajo8,                                     
		Trabajo9       		= @Trabajo9,                                     
		Trabajo10       		= @Trabajo10,                                     
		Trabajo11       		= @Trabajo11,                                     
		Trabajo12       		= @Trabajo12,                                     
		Trabajo13       		= @Trabajo13,                                     
		Trabajo14       		= @Trabajo14,                                     
		Color       		= @Color,                                     
		Color1       		= @Color1,                                     
		Color2       		= @Color2,                                     
		Color3       		= @Color3,                                     
		Color4       		= @Color4,                                     
		Color5       		= @Color5,                                     
		Color6       		= @Color6,                                     
		Color7       		= @Color7,                                     
		Color8       		= @Color8,                                     
		Color9       		= @Color9,                                     
		Color10       		= @Color10,                                     
		Color11       		= @Color11,                                     
		Color12       		= @Color12,                                     
		Color13       		= @Color13,                                     
		Color14       		= @Color14,                                     
		Color15       		= @Color15,                                     
		Color16       		= @Color16,                                     
		Color17       		= @Color17,                                     
		Color18       		= @Color18,                                     
		Color19       		= @Color19,                                     
		Color20       		= @Color20,                                     
		Color21       		= @Color21,                                     
		Maqui       		= @Maqui,                                     
		Maquina1       		= @Maquina1,                                     
		Maquina2       		= @Maquina2,                                     
		Maquina3       		= @Maquina3,                                     
		Maquina4       		= @Maquina4,                                     
		Maquina5       		= @Maquina5,                                     
		Maquina6       		= @Maquina6,                                     
		Maquina7       		= @Maquina7,                                     
		Maquina8       		= @Maquina8,                                     
		Maquina9       		= @Maquina9,                                     
		Maquina10       		= @Maquina10,                                     
		Maquina11       		= @Maquina11,                                     
		Maquina12      		= @Maquina12,                                     
		Maquina13       		= @Maquina13,                                     
		Maquina14       		= @Maquina14,                                     
		FechaCompro    		= @FechaCompro,                                     
		FechaSalida     		= @FechaSalida,                                    
		OrdFecha     		= @OrdFecha,
		OrdFechaCompro  	= @OrdFechaCompro,
		OrdFechaSalida 		= @FechaSalida  ,
		Clave 			= @Clave

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenSaldo]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenSaldo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOrdenSaldo]
             @Clave char(8)

 AS

UPDATE  Orden
	SET
		Saldo 	   	= Cantidad-Recibida' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenReproceso0]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenReproceso0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOrdenReproceso0]
 AS

UPDATE  Orden
	SET
		Recibida 	= Cantidad' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenReproceso]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenReproceso]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOrdenReproceso]
             @Clave char(8) ,
	@Recibida float

 AS

UPDATE  Orden
	SET
		Recibida	= @Recibida

where
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenPrueba]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenPrueba]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOrdenPrueba]
             @Clave char(8),
	@LIberada float,
	@Devuelta float,
	@Fechaentrega char(10),
	@WDate char(10)  


 AS

UPDATE  Orden
	SET
		Clave   		= @Clave,
		Liberada	= @Liberada,	
		Devuelta	= @Devuelta,	
	 	Fechaentrega	= @Fechaentrega,	
		WDate 	   	= @WDate


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenLeyenda]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenLeyenda]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOrdenLeyenda]
             @Orden int,
	@Leyenda int


 AS

UPDATE  Orden
	SET
		Leyenda	= @Leyenda


WHERE
	Orden = @Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenLaudo]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenLaudo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOrdenLaudo]
             @Clave char(8),
	@Liberada float,
	@Devuelta float,
	@Fechaentrega char(10),
	@WDate char(10)  


 AS

UPDATE  Orden
	SET
		Clave   		= @Clave,
		Liberada	= @Liberada,	
		Devuelta	= @Devuelta,	
	 	Fechaentrega	= @Fechaentrega,	
		WDate 	   	= @WDate


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenLaboratorio]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenLaboratorio]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOrdenLaboratorio]
             @Clave char(8),
	@Articulo char(10),
	@WDate char(10)  


 AS

UPDATE  Orden
	SET
		Clave   		= @Clave,
		Articulo  	= @Articulo,	
		WDate 	   	= @WDate

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenInforme]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenInforme]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOrdenInforme]
             @Clave char(8),
	@Recibida float,
	@WDate char(10)  


 AS

UPDATE  Orden
	SET
		Clave   		= @Clave,
		Recibida	= @Recibida,	
		WDate 	   	= @WDate


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenImpresion]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenImpresion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOrdenImpresion]
             @Orden int,
	@Impresion char(1)


 AS

UPDATE  Orden
	SET
		Impresion   	= @Impresion

WHERE
	Orden = @Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenFechaLLegada]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenFechaLLegada]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOrdenFechaLLegada]
             @Orden int,
	@Fecha1  char(10)  ,
	@Fecha2  char(10)  


 AS

UPDATE  Orden
	SET
		Orden   		= @Orden,
		Fecha1		= @Fecha1  ,
		Fecha2		= @Fecha2


WHERE
	Orden = @Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenFecha2]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenFecha2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOrdenFecha2]
             @Orden int,
	@OrdFecha2  char(8)


 AS

UPDATE  Orden
	SET
		Orden   		= @Orden,
		OrdFecha2	= @OrdFecha2


WHERE
	Orden = @Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenDy]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOrdenDy]
             @Articulo char(10),
	@ArticuloDy Char(10)

 AS

uPDATE  Orden
	SET
		Articulo   		= @ArticuloDy

		

WHERE
	Articulo = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrdenDerechos]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrdenDerechos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOrdenDerechos]
             @Clave char(8),
	@Derechos float


 AS

UPDATE  Orden
	SET
		Derechos	= @Derechos


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaOrden]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaOrden]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaOrden]
             @Clave char(8),
	@Orden int,
	@Renglon int,
	@Fecha char(10),
	@Proveedor char(11),
	@Articulo char(10),
	@Cantidad float,
	@Precio float,
	@Fecha1 char(10),
	@Fecha2 char(10),
	@Condicion char(40),
	@Recibida float,
	@Saldo float,
	@Fechaord char(8),
	@Liberada float,
	@Devuelta float,
	@Fechaentrega char(10),
	@WDate char(10)  


 AS

UPDATE  Orden
	SET
		Clave   		= @Clave,
		Orden		= @Orden,
		Renglon	= @Renglon,   
		Fecha 		= @Fecha,	
		Proveedor	= @Proveedor,	
		Articulo  	= @Articulo,	
		Cantidad  	= @Cantidad,	
		Precio  		= @Precio,	
		Fecha1	 	= @Fecha1,	
		Fecha2	 	= @Fecha2,	
		Condicion	= @Condicion,	
		Recibida	= @Recibida,	
		Saldo		= @Saldo,	
		Fechaord 	= @Fechaord,	
		Liberada	= @Liberada,	
		Devuelta	= @Devuelta,	
		Fechaentrega 	= @Fechaentrega,	
		WDate 	   	= @WDate


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaNumero]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaNumero]
             @Codigo char(2),
	@Numero int

 AS

UPDATE  Numero
	SET
		Numero      	= @Numero
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMuestraII]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMuestraII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMuestraII]
             @Codigo int,
	@Producto2 char(12) ,
	@Articulo2 char(10) ,
	@Ensayo2 char(15) ,
	@Nombre2 char(50),
	@Fecha2 char(10) ,
	@OrdFecha2 char(8) ,
	@Cantidad2 char(15),
	@Lote2 char(10),
	@Observaciones2 char(50) ,
	@Stock2 int
 AS

UPDATE  Muestra
	SET
		Producto2	= @Producto2,
		Articulo2	= @Articulo2,
		Ensayo2	= @Ensayo2,   
		Nombre2 	= @Nombre2 ,
		Fecha2 		= @Fecha2,	
		OrdFecha2	= @OrdFecha2,	
		Cantidad2  	= @Cantidad2,	
		Lote2		= @Lote2,	
		Observaciones2	= @Observaciones2  ,
		Stock2  	= @Stock2

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMuestraI]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMuestraI]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMuestraI]
             @Codigo int,
	@Producto char(12) ,
	@Articulo char(10) ,
	@Ensayo char(15) ,  
	@Nombre char(50),
	@Fecha char(10) ,
	@OrdFecha char(8) ,
	@Cantidad char(15) ,
	@Cliente char(6) ,
	@Razon char(50),
	@DescriCliente char(50) ,
	@Vendedor int ,
	@DesVendedor char(50) ,
	@Observaciones char(50)  

 AS

UPDATE  Muestra
	SET
		Producto	= @Producto,
		Articulo		= @Articulo,
		Ensayo		= @Ensayo,   
		Nombre		= @Nombre ,
		Fecha 		= @Fecha,	
		OrdFecha	= @OrdFecha,	
		Cantidad  	= @Cantidad,	
		Cliente		= @Cliente,	
		Razon 		= @Razon  ,
		DescriCliente	= @DescriCliente,	
		Vendedor  	= @Vendedor,	
		DesVendedor 	= @DesVendedor ,
		Observaciones  	= @Observaciones


WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovvarMarcaant]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovvarMarcaant]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovvarMarcaant]  AS

UPDATE  Movvar
	SET
		Marcaant  	= Marca' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovvarMarca]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovvarMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovvarMarca]  AS

UPDATE  Movvar
	SET
		Marca	  	= "X"' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovvarDy]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovvarDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovvarDy]
             @Articulo char(10),
	@ArticuloDy Char(10)

 AS

uPDATE  Movvar
	SET
		Articulo   		= @ArticuloDy

		

WHERE
	Articulo = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovvar]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovvar]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovvar]
             @Clave char(8),
	@Codigo int,
	@Renglon int,
	@Fecha char(10),
	@Tipo char(1),
	@Articulo char(10),
	@Terminado char(12),
	@Cantidad float,
	@Fechaord char(8),
	@Movi char(1),
	@Tipomov char(1),
	@Observaciones char(50),
	@WDate char(10),
	@Marca char(1)  ,
	@Lote int

 AS

UPDATE  Movvar
	SET
		Clave   		= @Clave,
		Codigo		= @Codigo,
		Renglon	= @Renglon,   
		Fecha 		= @Fecha,	
		Tipo		= @Tipo,	
		Articulo  	= @Articulo,	
		Terminado  	= @Terminado,	
		Cantidad  	= @Cantidad,	
		Fechaord  	= @Fechaord,	
		Movi		= @Movi,	
		Tipomov	= @Tipomov,	
		Observaciones	= @Observaciones,	
		WDate 	   	= @WDate,
		Marca		= @Marca  ,
		Lote		= @Lote




WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovlabMarcaant]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovlabMarcaant]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovlabMarcaant]  AS

UPDATE  Movlab
	SET
		Marcaant  	= Marca' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovlabMarca]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovlabMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovlabMarca]  AS

UPDATE  Movlab
	SET
		Marca	  	= "X"' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovlabDy]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovlabDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovlabDy]
             @Articulo char(10),
	@ArticuloDy Char(10)

 AS

uPDATE  Movlab
	SET
		Articulo   		= @ArticuloDy

		

WHERE
	Articulo = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovlab]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovlab]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovlab]
             @Clave char(8),
	@Codigo int,
	@Renglon int,
	@Fecha char(10),
	@Tipo char(1),
	@Articulo char(10),
	@Terminado char(12),
	@Cantidad float,
	@Fechaord char(8),
	@Movi char(1),
	@Tipomov char(1),
	@Observaciones char(50),
	@WDate char(10),
	@Marca char(1) ,
	@Lote int

 AS

UPDATE  Movlab
	SET
		Clave   		= @Clave,
		Codigo		= @Codigo,
		Renglon	= @Renglon,   
		Fecha 		= @Fecha,	
		Tipo		= @Tipo,	
		Articulo  	= @Articulo,	
		Terminado  	= @Terminado,	
		Cantidad  	= @Cantidad,	
		Fechaord  	= @Fechaord,	
		Movi		= @Movi,	
		Tipomov	= @Tipomov,	
		Observaciones	= @Observaciones,	
		WDate 	   	= @WDate,
		Marca		= @Marca  ,
		Lote		= @Lote




WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovguiaSaldoCierre]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovguiaSaldoCierre]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovguiaSaldoCierre]
             @Clave char(10),
	@Saldo Float, 
	@Cantidad Float  ,
	@Marca  char(1)


 AS

UPDATE  Guia
	SET
		Clave   		= @Clave,
		Saldo 		= @Saldo   ,
		Cantidad	= @Cantidad  ,
		Marca		= @Marca


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovguiaSaldo]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovguiaSaldo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovguiaSaldo]
             @Clave char(10),
	@WDate char(10) , 
	@Saldo Float


 AS

UPDATE  Guia
	SET
		Clave   		= @Clave,
		WDate 	   	= @WDate  ,
		Saldo 		= @Saldo


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovguiaMarcaant]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovguiaMarcaant]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovguiaMarcaant]  
          @Clave char(10)

 AS
UPDATE  Guia
	SET
		Marcaant  	= Marca,
		Saldoant	= Saldo  ,
		Cantidadant	= Cantidad


WHERE 
		Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovguiaMarca]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovguiaMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovguiaMarca]  AS

UPDATE  Guia
	SET
		Marca	  	= "X",
		Saldo 		= 0 ,
		Cantidad	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovgasProceso0]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovgasProceso0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovgasProceso0]  AS

UPDATE  Movgas
	SET
		FechaLlegada = ""  ,
		OrdFechaLlegada = ""  ,
		CostoFlete = 0  ,
		Gastos = 0  ,
		Pagado = 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovgasProceso]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovgasProceso]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovgasProceso]
             @Carpeta int,
             @FechaLLegada char(10)  ,
	@OrdFechaLlegada  char(8)  ,
	@CostoFlete float  ,
	@Gastos float  ,
	@Pagado float


 AS

UPDATE  Movgas
	SET
		FechaLLegada		= @FechaLlegada  ,
		OrdFechaLlegada	= @OrdFechaLlegada  ,
		CostoFlete		= @CostoFlete  ,
		Gastos			= @Gastos  ,
		Pagado			= @Pagado


WHERE
	Carpeta = @Carpeta' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovgasImpoDerechos]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovgasImpoDerechos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovgasImpoDerechos]
             @Carpeta int,
             @ImpoDerechos float


 AS

UPDATE  Movgas
	SET
		ImpoDerechos	= @ImpoDerechos


WHERE
	Carpeta = @Carpeta' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovgas]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovgas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovgas]
             @Clave char(8),
             @Carpeta int,
             @Renglon int,
             @Fecha char(10),
             @Derechos float ,
             @Orden int  ,
	@Concepto int,
             @Importe float ,
             @Auxiliar float  ,
             @OrdFecha char(8)  ,
	@Proveedor char(11)  ,
	@Origen char(30)  ,
	@Moneda  int


 AS

UPDATE  Movgas
	SET
		Clave   		= @Clave,
		Carpeta		= @Carpeta,
		Renglon	= @Renglon,   
		Fecha 		= @Fecha,	
		Derechos  	= @Derechos,	
		Orden		= @Orden,	
		Concepto  	= @Concepto,	
		Importe  	= @Importe,	
		Auxiliar   	= @Auxiliar,
		OrdFecha	= @OrdFecha  ,
		Proveedor	= @Proveedor  ,
		Origen		= @Origen   ,
		Moneda		= @Moneda


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovenvMovi]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovenvMovi]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovenvMovi]
             @Fechaord char(8),
	@Movimiento char(1),
	@Graba char(1)

 AS

UPDATE  Movenv
	SET
		Movimiento   	= @Graba
WHERE
	Fechaord < @Fechaord and Movimiento = @Movimiento' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMovenv]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMovenv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMovenv]
             @Clave char(9),
             @Tipo char(1),
             @Codigo int,
             @Renglon int,
             @Fecha char(10),
             @Fechaord char(8),
             @Cliente char(6),
             @Envase int,
             @Movimiento char(1),
             @Cantidad float


 AS

UPDATE  Movenv
	SET
		Clave   		= @Clave,
		Tipo 		= @Tipo,
		Codigo		= @Codigo,
		Renglon	= @Renglon,   
		Fecha 		= @Fecha,	
		Fechaord  	= @Fechaord,	
		Cliente		= @Cliente,	
		Envase  	= @Envase,	
		Movimiento  	= @Movimiento,	
		Cantidad  	= @Cantidad


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMinimoPlanta]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMinimoPlanta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMinimoPlanta]
             @Codigo char(12),
	@Stock1 float,
	@Stock2 float,
	@Stock3 float,
	@Stock4 float,
	@Stock5 float,
	@MInimo1 float,
	@MInimo2 float,
	@MInimo3 float,
	@MInimo4 float,
	@MInimo5 float


 AS

UPDATE  Minimo
	SET
		Codigo    	= @Codigo,
		Stock1    	= @Stock1,
		Stock2    	= @Stock2,
		Stock3    	= @Stock3,
		Stock4    	= @Stock4,
		Stock5    	= @Stock5  ,
		Minimo1    	= @Minimo1  ,
		Minimo2    	= @Minimo2,
		Minimo3    	= @Minimo3,
		Minimo4    	= @Minimo4,
		Minimo5    	= @Minimo5

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMinimoDisponible]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMinimoDisponible]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMinimoDisponible]
             @Codigo char(12),
	@Stock1 float,
	@Stock2 float,
	@Stock3 float,
	@Stock4 float,
	@Stock5 float ,
	@MInimo float


 AS

UPDATE  Minimo
	SET
		Codigo    	= @Codigo,
		Stock1    	= @Stock1,
		Stock2    	= @Stock2,
		Stock3    	= @Stock3,
		Stock4    	= @Stock4,
		Stock5    	= @Stock5 ,
		Minimo    	= @Minimo

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMinimoDifePlanta]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMinimoDifePlanta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMinimoDifePlanta]
 AS

UPDATE  Minimo
	SET
		Dife1	    	= Stock1 - Minimo1  ,
		Dife2	    	= Stock2 - Minimo2,
		Dife3	    	= Stock3 - Minimo3,
		Dife4	    	= Stock4 - Minimo4,
		Dife5	    	= Stock5 - Minimo5' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMinimoDife]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMinimoDife]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMinimoDife]
 AS

UPDATE  Minimo
	SET
		Dife	    	= Stock1 + Stock2 + Stock3 + Stock4 + Stock5 - Minimo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMinimo]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMinimo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMinimo]
             @Codigo char(12),
	@Stock1 float,
	@Stock2 float,
	@Stock3 float,
	@Stock4 float,
	@Stock5 float


 AS

UPDATE  Minimo
	SET
		Codigo    	= @Codigo,
		Stock1    	= @Stock1,
		Stock2    	= @Stock2,
		Stock3    	= @Stock3,
		Stock4    	= @Stock4,
		Stock5    	= @Stock5

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaMarcas]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaMarcas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaMarcas]
             @Clave char(21),
	@Articulo char(10),
	@Proveedor char(11),
	@Descripcion char(50)

 AS

UPDATE  Marcas
	SET
		Clave	   	= @Clave,
		Articulo       	= @Articulo,                                     
		Proveedor       	= @Proveedor,                                     
		Descripcion      	= @Descripcion
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaLineaMp]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLineaMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaLineaMp]
             @Linea int,
	@Nombre char(50)


 AS

UPDATE  LineasMp
	SET
		Linea     	= @Linea,
		Nombre       	= @Nombre

WHERE
	Linea = @Linea' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaLinea]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLinea]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaLinea]
             @Linea int,
	@Nombre char(50)


 AS

UPDATE  Lineas
	SET
		Linea     	= @Linea,
		Nombre       	= @Nombre

WHERE
	Linea = @Linea' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoSaldoCierre]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoSaldoCierre]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaLaudoSaldoCierre]
             @Clave char(8),
	@WDate char(10) , 
	@Saldo Float, 
	@Liberada Float  ,
	@Marca  char(1)


 AS

UPDATE  Laudo
	SET
		Clave   		= @Clave,
		WDate 	   	= @WDate  ,
		Saldo 		= @Saldo   ,
		Liberada	= @Liberada  ,
		Marca		= @Marca


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoSaldo1]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoSaldo1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaLaudoSaldo1]
             @Clave char(8),
	@WDate char(10) , 
	@Saldo Float, 
	@Liberada Float


 AS

UPDATE  Laudo
	SET
		Clave   		= @Clave,
		WDate 	   	= @WDate  ,
		Saldo 		= @Saldo   ,
		Liberada	= @Liberada


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoSaldo]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoSaldo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaLaudoSaldo]
             @Clave char(8),
	@WDate char(10) , 
	@Saldo Float


 AS

UPDATE  Laudo
	SET
		Clave   		= @Clave,
		WDate 	   	= @WDate  ,
		Saldo 		= @Saldo


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoProceso2]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoProceso2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaLaudoProceso2]
             @Clave char(8),
	@PartiOri char(20)

 AS

UPDATE  Laudo
	SET
		PartiOri 		= @PartiOri


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoProceso1]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoProceso1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaLaudoProceso1]  AS

UPDATE  Laudo
	SET
		PartiOriAnterior  	= PartiOri' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoMarcaant]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoMarcaant]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaLaudoMarcaant]  
            @Clave char(8)

 AS

UPDATE  Laudo
	SET
		Marcaant	= Marca,
		Saldoant 	= Saldo  ,
		Liberadaant	= Liberada  ,
		Devueltaant	= Devuelta
WHERE 
		Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoMarca]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaLaudoMarca]  AS

UPDATE  Laudo
	SET
		Marca	  	= "X",
		Saldo 		= 0,
		Liberada	= 0 ,
		Devuelta 	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoFechaOrd]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoFechaOrd]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaLaudoFechaOrd]
             @Laudo int,
	@Fechaord char(8)

 AS

uPDATE  Laudo
	SET
		FechaOrd    		= @FechaOrd

		

WHERE
	Laudo = @Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudoDy]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudoDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaLaudoDy]
             @Articulo char(10),
	@ArticuloDy Char(10)

 AS

uPDATE  Laudo
	SET
		Articulo    		= @ArticuloDy

		

WHERE
	Articulo = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaLaudo]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaLaudo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaLaudo]
             @Clave char(8),
	@Laudo int,
	@Renglon int,
	@Fecha char(10),
	@Articulo char(10),
	@Liberada float,
	@Devuelta float,
	@Orden int,
	@Marca char(1),
	@Lote int,
	@Rechazo int,
	@Informe int,
	@Actualiza char(1),
	@WDate char(10)  ,
	@Saldo Float    ,
	@Origen  char(50)   ,
	@PartiOri  char(20)    ,
	@Envase  int



 AS

UPDATE  Laudo
	SET
		Clave   		= @Clave,
		Laudo		= @Laudo,
		Renglon	= @Renglon,   
		Fecha 		= @Fecha,	
		Articulo  	= @Articulo,	
		Liberada	= @Liberada,	
		Devuelta	= @Devuelta,	
		Orden  		= @Orden,	
		Marca	 	= @Marca,	
		Rechazo 	= @Rechazo,	
		Informe		= @Informe,	
		Actualiza	= @Actualiza,	
		WDate 	   	= @WDate ,
		Saldo 		= @Saldo    ,
		Origen		= @Origen   ,
		PartiOri		= @PartiOri   ,
		Envase		= @Envase

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaInventario]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInventario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaInventario]
             @Clave char(8),
	@Numero int,
	@Renglon int,
	@Tipo  char(1)  ,
	@Articulo char(10),
	@Terminado char(12),
	@Cantidad float,
	@Lote int,
	@Talon int,
	@Observaciones char(30),
	@Ubicacion char(20)

 AS

UPDATE  Inventario
	SET
		Clave   		= @Clave,
		Numero		= @Numero,
		Renglon	= @Renglon,   
		Tipo		= @Tipo,	
		Articulo  	= @Articulo,	
		Terminado  	= @Terminado,	
		Cantidad  	= @Cantidad,	
		Lote   	  	= @Lote,	
		Talon		= @Talon,	
		Observaciones	= @Observaciones,	
		Ubicacion	= @Ubicacion


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeProcesoSaldo]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeProcesoSaldo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaInformeProcesoSaldo]  AS

UPDATE  Informe
	SEt
		CantidadLaudo 	= Cantidad' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeProcesoDife]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeProcesoDife]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaInformeProcesoDife]  AS

UPDATE  Informe
	SET
		Dife 	= Cantidad - CantidadLaudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeProceso0]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeProceso0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaInformeProceso0]
             @OrdFecha char(8)

 AS

UPDATE  Informe
	SET
		CantidadLaudo 	= 0  ,
		Dife 		= 0


WHERE
	FechaOrd >= @OrdFecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeProceso]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeProceso]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaInformeProceso]
             @Clave char(8),
	@CantidadLaudo  Float

 AS

UPDATE  Informe
	SET
		CantidadLaudo	=  @CantidadLaudo

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformePartida]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformePartida]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaInformePartida]
             @Clave char(8),
	@Lote1 char(10),
	@Canti1 float ,
	@Lote2 char(10),
	@Canti2 float ,
	@Lote3 char(10),
	@Canti3 float ,
	@Lote4 char(10),
	@Canti4 float ,
	@Lote5 char(10),
	@Canti5 float

 AS

UPDATE  Informe
	SET
		Lote1		= @Lote1,
		Canti1 	 	= @Canti1,	
		Lote2		= @Lote2,
		Canti2 	 	= @Canti2,	
		Lote3		= @Lote3,
		Canti3 	 	= @Canti3,	
		Lote4		= @Lote4,
		Canti4 	 	= @Canti4,	
		Lote5		= @Lote5,
		Canti5 	 	= @Canti5

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeListadoII]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeListadoII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaInformeListadoII]
             @Clave char(8),
	@Liberada   float,
	@Partida1 int,
	@Devuelta   float ,
	@Partida2  int

 AS

UPDATE  Informe
	SET
		Clave   		= @Clave,
		Liberada  	= @Liberada,	
		Partida1 	= @Partida1   ,
		Devuelta 	= @Devuelta   ,
		Partida2		= @Partida2



WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeListado]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeListado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaInformeListado]
             @Clave char(8),
	@FechaOrden char(10),
	@OrdFechaOrden char(8) ,
	@DifeFecha  int

 AS

UPDATE  Informe
	SET
		Clave   		= @Clave,
		FechaOrden  	= @FechaOrden,	
		OrdFechaOrden 	= @OrdFechaOrden   ,
		DifeFecha	= @DifeFecha



WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeLaboratorio]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeLaboratorio]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaInformeLaboratorio]
             @Clave char(8),
	@Articulo char(10),
	@WDate char(10)

 AS

UPDATE  Informe
	SET
		Clave   		= @Clave,
		Articulo  	= @Articulo,	
		WDate 	   	= @WDate



WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeDy]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaInformeDy]
             @Articulo char(10),
	@ArticuloDy Char(10)

 AS

uPDATE  Informe
	SET
		Articulo    		= @ArticuloDy

		

WHERE
	Articulo = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaInformeDatos]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInformeDatos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaInformeDatos]
             @Informe int,
	@Certificado1 int,
	@Certificado2 char(50) ,
	@Estado1 int,
	@Estado2 char(50)

 AS

UPDATE  Informe
	SET
		Certificado1 	= @Certificado1,
		Certificado2  	= @Certificado2,	
		Estado1 	= @Estado1   ,
		Estado2		= @Estado2



WHERE
	Informe = @Informe' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaInforme]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaInforme]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaInforme]
             @Clave char(8),
	@Informe int,
	@Renglon int,
	@Fecha char(10),
	@Remito int,
	@Proveedor char(11),
	@Orden int,
	@Articulo char(10),
	@Cantidad int,
	@Resta int,
	@Fechaord char(8),
	@WDate char(10),
	@Envase int

 AS

UPDATE  Informe
	SET
		Clave   		= @Clave,
		Informe		= @Informe,
		Renglon	= @Renglon,   
		Fecha 		= @Fecha,	
		Remito		= @Remito,	
		Proveedor	= @Proveedor,	
		Orden		= @Orden,	
		Articulo  	= @Articulo,	
		Cantidad  	= @Cantidad,	
		Resta  		= @Resta,	
		Fechaord 	= @Fechaord,	
		WDate 	   	= @WDate,
		Envase		= @Envase



WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaWImporte0]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaWImporte0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaHojaWImporte0]  AS

UPDATE  Hoja
	SET
		WImporte  	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaWImporte]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaWImporte]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaHojaWImporte]  
	@Desde char(8),
	@Hasta char(8)
as

UPDATE  Hoja
	SET
		WImporte  	= Real

WHERE
	Fechaingord >= @Desde and Fechaingord <= @Hasta' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaSaldoCierre]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaSaldoCierre]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaHojaSaldoCierre]
             @Clave char(8),
	@Real Float, 
	@Saldo Float  ,
	@Marca  char(1)


 AS

UPDATE  Hoja
	SET
		Clave   		= @Clave,
		Saldo 		= @Saldo   ,
		Real		= @Real  ,
		Marca		= @Marca


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaSaldo3]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaSaldo3]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaHojaSaldo3]
             @Clave char(8),
	@WDate char(10) , 
	@Saldo Float,
	@Teorico Float,
	@Real Float,
	@Cantidad Float,
	@Marca char(1)


 AS

UPDATE  Hoja
	SET
		Clave   		= @Clave,
		WDate 	   	= @WDate  ,
		Saldo 		= @Saldo,
		Teorico 		= @Teorico,
		Real 		= @Real,
		Cantidad	= @Cantidad,
		Marca		= @Marca


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaSaldo2]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaSaldo2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaHojaSaldo2]
             @Clave char(8),
	@WDate char(10) , 
	@Saldo Float,
	@Marca char(1)


 AS

UPDATE  Hoja
	SET
		Clave   		= @Clave,
		WDate 	   	= @WDate  ,
		Saldo 		= @Saldo,
		Marca		= @Marca


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaSaldo1]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaSaldo1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaHojaSaldo1]
             @Clave char(8),
	@WDate char(10) , 
	@Saldo Float ,
	@Real float


 AS

UPDATE  Hoja
	SET
		Clave   		= @Clave,
		WDate 	   	= @WDate  ,
		Saldo 		= @Saldo ,
		Real		= @Real


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaSaldo]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaSaldo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaHojaSaldo]
             @Clave char(8),
	@WDate char(10) , 
	@Saldo Float


 AS

UPDATE  Hoja
	SET
		Clave   		= @Clave,
		WDate 	   	= @WDate  ,
		Saldo 		= @Saldo


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaMarcaant]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaMarcaant]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaHojaMarcaant]  
            @Clave char(8)

 AS

UPDATE  Hoja
	SET
		Marcaant  	= Marca,
		Saldoant	= Saldo  ,
		Realant		= Real


WHERE 
		Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaMarca]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaHojaMarca]  AS

UPDATE  Hoja
	SET
		Marca	  	= "X",
		Saldo		= 0  , 
		Real		= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaFechaOrd]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaFechaOrd]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaHojaFechaOrd]
             @Hoja int,
	@FechaOrd char(8)


 AS

UPDATE  Hoja
	SET
		FechaOrd		= @FechaOrd


WHERE
	Hoja = @Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaHojaDy]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHojaDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaHojaDy]
             @Articulo char(10),
	@ArticuloDy Char(10)

 AS

uPDATE  Hoja
	SET
		Articulo    		= @ArticuloDy

		

WHERE
	Articulo = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaHoja]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaHoja]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaHoja]
             @Clave char(8),
	@Hoja int,
	@Renglon int,
	@Fecha char(10),
	@Producto char(12),
	@Cantidad float,
	@Tipo char(1),
	@Lote int,
	@Articulo char(10),
	@Terminado char(12),
	@Teorico float,
	@Real float,
	@Fechaing char(10),
	@Fechaingord char(8),
	@WDate char(10),
	@WImporte float,
	@Marca char(1) ,
	@Saldo float,
	@Lote1 int,
	@Canti1 float ,
	@Lote2 int,
	@Canti2 float ,
	@Lote3 int ,
	@Canti3 float


 AS

UPDATE  Hoja
	SET
		Clave   		= @Clave,
		Hoja		= @Hoja,
		Renglon	= @Renglon,   
		Fecha 		= @Fecha,	
		Producto	= @Producto,	
		Cantidad  	= @Cantidad,	
		Tipo		= @Tipo,	
		Lote		= @Lote,	
		Articulo  	= @Articulo,	
		Terminado  	= @Terminado,	
		Teorico		= @Teorico,	
		Real		= @Real,	
		Fechaing  	= @Fechaing,	
		Fechaingord 	= @Fechaingord,	
		WDate 	   	= @WDate,
		WImporte	= @WImporte,	
		Marca		= @Marca ,
		Saldo  		= @Saldo,
		Lote1  		= @Lote1,
		Canti1  		= @Canti1,
		Lote2  		= @Lote2,
		Canti2  		= @Canti2,
		Lote3  		= @Lote3,
		Canti3  		= @Canti3

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaGuiaDy]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaGuiaDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaGuiaDy]
             @Articulo char(10),
	@ArticuloDy Char(10)

 AS

uPDATE  Guia
	SET
		Articulo    		= @ArticuloDy

		

WHERE
	Articulo = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaGasimpo]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaGasimpo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaGasimpo]
             @Codigo int,
	@Nombre char(50)


 AS

UPDATE  Gasimpo
	SET
		Codigo     	= @Codigo,
		Nombre       	= @Nombre

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEstadisticaSalva]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEstadisticaSalva]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEstadisticaSalva]
             @Clave char(12),
             @Precio float,
             @Importe float ,
             @Paridad float
 AS

UPDATE  Estadistica
	SET
		Clave	   	= @Clave,
		Precio      	= @Precio,
		Importe      	= @Importe,
		Paridad      	= @Paridad

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEstadisticaMarcaant]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEstadisticaMarcaant]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEstadisticaMarcaant]  AS

UPDATE  Estadistica
	SET
		Marcaant  	= Marca' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEstadisticaMarca]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEstadisticaMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEstadisticaMarca]  AS

UPDATE  Estadistica
	SET
		Marca	  	= "X"' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEstadisticaLinea]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEstadisticaLinea]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEstadisticaLinea]
             @Clave char(12), 
	@Linea int

 AS

UPDATE  Estadistica
	SET
		Clave		= @Clave,
		Linea	 	= @Linea

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEstadistica]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEstadistica]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEstadistica]
             @Clave char(12),
             @Tipo int,
             @Numero int,
             @Renglon int,
             @Articulo char(12),
             @Cantidad float,
             @Precio float,
             @PrecioUS float,
             @Importe float,
             @ImporteUS float,
             @Cliente char(6),
             @Paridad float,
             @Vendedor int,
             @Rubro int,
             @Linea int,
             @Costo1 float,
             @Costo2 float,
             @Coeficiente float,
             @Pedido int,
             @Fecha char(10),
             @Importe1 float,
             @Importe2 float,
             @Importe3 float,
             @Importe4 float,
             @Ordfecha char(8),
             @WArticulo char(8),
             @Remito char(10),
             @WDate char(10),
             @WCantidad float,
             @WImporte float,
             @WImporteUs float,
             @Marca char(1),
	@Lote1 int,
	@Canti1 float,
	@Lote2 int ,
	@Canti2 float ,
	@Lote3 int ,
	@Canti3 float,
	@Lote4 int ,
	@Canti4 float ,
	@Lote5 int ,
	@Canti5 float
 AS

UPDATE  Estadistica
	SET
		Clave	   	= @Clave,
		Tipo       	= @Tipo,                                     
		Numero      	= @Numero,                                    
		Renglon      	= @Renglon,
		Articulo      	= @Articulo,		
		Cantidad      	= @Cantidad,		
		Precio      	= @Precio,
		PrecioUs      	= @PrecioUs,
		Importe      	= @Importe,
		ImporteUs      	= @ImporteUs,
		Cliente      	= @Cliente,
		Paridad      	= @Paridad,
		Vendedor      	= @Vendedor,
		Rubro      	= @Rubro,
		Linea      	= @Linea,
		Costo1      	= @Costo1,
		Costo2      	= @Costo2,
		Coeficiente      	= @Coeficiente,
		Pedido      	= @Pedido,
		Fecha      	= @Fecha,
		Importe1     	= @Importe1,
		Importe2      	= @Importe2,
		Importe3      	= @Importe3,
		Importe4      	= @Importe4,
		OrdFecha      	= @OrdFecha,
		WArticulo      	= @WArticulo,
		Remito      	= @Remito,
		WDate      	= @WDAte,
		WCantidad      	= @WCantidad,
		WImporteUs      	= @WImporteUs,
		Marca      	= @Marca,
		Lote1		 = @Lote1 ,
		Canti1		= @Canti1 ,
		Lote2		 = @Lote2 ,
		Canti2		= @Canti2 ,
		Lote3		 = @Lote3 ,
		Canti3		= @Canti3,
		Lote4		 = @Lote4 ,
		Canti4		= @Canti4 ,
		Lote5		 = @Lote5 ,
		Canti5		= @Canti5

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEspeCli]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEspeCli]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEspeCli]
             @Cliente char(6),
	@Terminado char(12),
	@Especificaciones char(30)

 AS

UPDATE  EspeCli
	SET
		Cliente    		= @Cliente,
	 	Terminado    		= @Terminado,
		Especificaciones       	= @Especificaciones

WHERE
	Cliente = @Cliente and Terminado = @Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEspecificacionesDy]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEspecificacionesDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEspecificacionesDy]
             @Articulo char(10),
	@ArticuloDy Char(10)

 AS

uPDATE  especificaciones
	SET
		Producto    		= @ArticuloDy

		

WHERE
	Producto = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEspecificaciones]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEspecificaciones]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEspecificaciones]
             @Producto char(10),
	@Ensayo1 int,
	@Valor1 char(50),
	@Ensayo2 int,
	@Valor2 char(50),
	@Ensayo3 int,
	@Valor3 char(50),
	@Ensayo4 int,
	@Valor4 char(50),
	@Ensayo5 int,
	@Valor5 char(50),
	@Ensayo6 int,
	@Valor6 char(50),
	@Ensayo7 int,
	@Valor7 char(50),
	@Ensayo8 int,
	@Valor8 char(50),
	@Ensayo9 int,
	@Valor9 char(50),
	@Ensayo10 int,
	@Valor10 char(50),
	@WDate char(10)

 AS

UPDATE  Especificaciones
	SET
		Producto   	= @Producto,
		Ensayo1       	= @Ensayo1,                                     
		Valor1      	= @Valor1,                                    
		Ensayo2       	= @Ensayo2,                                     
		Valor2      	= @Valor2,                                    
		Ensayo3       	= @Ensayo3,                                     
		Valor3      	= @Valor3,                                    
		Ensayo4       	= @Ensayo4,                                     
		Valor4      	= @Valor4,                                    
		Ensayo5       	= @Ensayo5,                                     
		Valor5      	= @Valor5,                                    
		Ensayo6       	= @Ensayo6,                                     
		Valor6      	= @Valor6,                                    
		Ensayo7       	= @Ensayo7,                                     
		Valor7      	= @Valor7,                                    
		Ensayo8       	= @Ensayo8,                                     
		Valor8      	= @Valor8,                                    
		Ensayo9       	= @Ensayo9,                                     
		Valor9      	= @Valor9,                                    
		Ensayo10       	= @Ensayo10,                                     
		Valor10      	= @Valor10,                                    
		WDate      	= @WDate
WHERE
	Producto = @Producto' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEspecif]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEspecif]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEspecif]
             @Producto char(12),
	@Ensayo1 int,
	@Valor1 char(50),
	@Ensayo2 int,
	@Valor2 char(50),
	@Ensayo3 int,
	@Valor3 char(50),
	@Ensayo4 int,
	@Valor4 char(50),
	@Ensayo5 int,
	@Valor5 char(50),
	@Ensayo6 int,
	@Valor6 char(50),
	@Ensayo7 int,
	@Valor7 char(50),
	@Ensayo8 int,
	@Valor8 char(50),
	@Ensayo9 int,
	@Valor9 char(50),
	@Ensayo10 int,
	@Valor10 char(50),
	@WDate char(10)

 AS

UPDATE  Especif
	SET
		Producto   	= @Producto,
		Ensayo1       	= @Ensayo1,                                     
		Valor1      	= @Valor1,                                    
		Ensayo2       	= @Ensayo2,                                     
		Valor2      	= @Valor2,                                    
		Ensayo3       	= @Ensayo3,                                     
		Valor3      	= @Valor3,                                    
		Ensayo4       	= @Ensayo4,                                     
		Valor4      	= @Valor4,                                    
		Ensayo5       	= @Ensayo5,                                     
		Valor5      	= @Valor5,                                    
		Ensayo6       	= @Ensayo6,                                     
		Valor6      	= @Valor6,                                    
		Ensayo7       	= @Ensayo7,                                     
		Valor7      	= @Valor7,                                    
		Ensayo8       	= @Ensayo8,                                     
		Valor8      	= @Valor8,                                    
		Ensayo9       	= @Ensayo9,                                     
		Valor9      	= @Valor9,                                    
		Ensayo10       	= @Ensayo10,                                     
		Valor10      	= @Valor10,                                    
		WDate      	= @WDate
WHERE
	Producto = @Producto' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEnvases]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEnvases]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEnvases]
             @Envases int,
	@Descripcion char(50),
	@Abreviatura char(10),
	@Kilos int

 AS

UPDATE  Envases
	SET
		Envases   	= @Envases,
		Descripcion       	= @Descripcion,                                     
		Abreviatura      	= @Abreviatura,                                    
		Kilos      	= @Kilos
WHERE
	Envases = @Envases' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEntdevMarcaant]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEntdevMarcaant]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEntdevMarcaant]  AS

UPDATE  Entdev
	SET
		Marcaant  	= Marca' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEntdevMarca]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEntdevMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEntdevMarca]  AS

UPDATE  Entdev
	SET
		Marca	  	= "X"' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEntdev2]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEntdev2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEntdev2]
             @Cliente char(6),
	@Lote int,
	@Terminado char(12),
	@Saldo float   ,
	@Laboratorio   float

 AS

UPDATE  Entdev
	SET
		Saldo   		= @Saldo,
		Laboratorio	= @Laboratorio

WHERE
	Cliente = @Cliente and Lote = @Lote and Terminado = @Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEntdev]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEntdev]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEntdev]
             @Clave char(8),
	@Codigo int,
	@Renglon int,
	@Fecha char(10),
	@Terminado char(12),
	@Cantidad float,
	@Fechaord char(8),
	@Observaciones char(50),
	@Marca char(1)  ,
	@Lote int  ,
	@Cliente char(6)   ,
	@Saldo float   ,
	@Laboratorio   float

 AS

UPDATE  Entdev
	SET
		Clave   		= @Clave,
		Codigo		= @Codigo,
		Renglon	= @Renglon,   
		Fecha 		= @Fecha,	
		Terminado  	= @Terminado,	
		Cantidad  	= @Cantidad,	
		Fechaord  	= @Fechaord,	
		Observaciones	= @Observaciones,	
		Marca		= @Marca  ,
		Lote		= @Lote  ,
		Cliente		= @Cliente  ,
		Saldo		= @Saldo  ,
		Laboratorio	= @Laboratorio

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaEnsayos]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaEnsayos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaEnsayos]
             @Codigo int,
	@Descripcion char(50),
	@WDate Char(10)

 AS

UPDATE  Ensayos
	SET
		Codigo   	= @Codigo,
		Descripcion       	= @Descripcion,                                     
		WDate      	= @WDate
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaDesccomp]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaDesccomp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaDesccomp]
             @Clave char(12),
             @Tipo char(2),
             @Numero char(8),
             @Renglon char(2),
             @Descripcion char(50),
             @Importe float,
             @Empresa smallint,
             @WDate char(10)


 AS
	INSERT INTO  Desccomp
			(
		             Clave,
		             Tipo,
		             Numero,
		             Renglon,
		             Descripcion,
		             Importe,
		             Empresa,
		             WDate
			)

VALUES
			(
		             @Clave,
		             @Tipo,
		             @Numero,
		             @Renglon,
		             @Descripcion,
		             @Importe,
		             @Empresa,
		             @WDate
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaDepositoImpolista0]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaDepositoImpolista0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaDepositoImpolista0]
as

UPDATE  Depositos
	SET
		Impolista =  0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaDepositoImpolista]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaDepositoImpolista]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaDepositoImpolista]
	@DesdeBanco int,
	@HastaBanco int,
	@DesdeFecha char(8),
	@HastaFecha char(8)

as

UPDATE  Depositos
	SET
		Impolista  	= Importe2

WHERE
	Fechaord >= @DesdeFecha and FechaOrd <= @HastaFecha and 	Banco >= @DesdeBanco and Banco <= @HastaBanco' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCuenta]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCuenta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCuenta]
             @Cuenta char(20),
	@Descripcion char(100),
	@Nivel int,
	@Empresa int

 AS

UPDATE  Cuenta
	SET
		Cuenta   	= @Cuenta,
		Descripcion       	= @Descripcion,                                     
		Nivel       	= @Nivel,                                     
		Empresa      	= @Empresa
WHERE
	Cuenta = @Cuenta' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtaprv]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtaprv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtaprv]
        	@Clave varchar(26),                     
	@Proveedor varchar(11),   
	@Letra varchar(1), 
	@Tipo varchar(2),
	@Punto varchar(4),
	@Numero varchar(8),  
	@fecha   varchar(10),   
	@Estado varchar(1),
	@Vencimiento varchar(10), 
	@Vencimiento1 varchar(50),                                       
	@Total   float    ,                                          
	@Saldo   float    ,                                          
	@OrdFecha  varchar(8),
	@OrdVencimiento varchar(8), 
	@Impre varchar(2),
	@Empresa smallint, 
	@SaldoList float,                                             
	@NroInterno int,  
	@Lista  varchar(1),
	@Acumulado float  ,
	@Paridad  float   ,
	@Pago  int
 AS

UPDATE   CtaCteprv
	SET
        	Clave  =   @Clave ,
	Proveedor = @Proveedor,   
	Letra  = 	@Letra,
	Tipo = @Tipo ,
	Punto = @Punto ,
	Numero = @Numero ,
	fecha   = @fecha  ,
	Estado = @Estado ,
	Vencimiento = @Vencimiento ,
	Vencimiento1 =   @Vencimiento1 ,
	Total   = @Total  ,
	Saldo  = @Saldo ,
	OrdFecha  = @OrdFecha  ,
	OrdVencimiento = @OrdVencimiento ,
	Impre = 	@Impre ,
	Empresa = @Empresa , 
	SaldoList = @SaldoList ,
	NroInterno = @NroInterno ,
	Lista  = 	@Lista ,
	Acumulado = @Acumulado   ,
	Paridad = @Paridad  ,
	Pago = @Pago
		
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteTipo2]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteTipo2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacteTipo2]  AS

UPDATE  Ctacte
	SET
		Tipo 		="07"

WHERE 
	Tipo = 7' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteTipo1]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteTipo1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacteTipo1]  AS

UPDATE  Ctacte
	SET
		Tipo 		="06"

WHERE 
	Tipo = 6' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteSalva]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteSalva]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacteSalva]
             @Clave char(12),
             @Neto float,
             @Iva1 float,
             @Iva2 float,
             @Paridad Float

 AS

UPDATE   Ctacte
	SET
	             Neto  		= @Neto,
	             Iva1  		= @Iva1,
	             Iva2  		= @Iva2,
	             Paridad  	= @Paridad
		
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteIva4]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteIva4]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacteIva4]
 AS

UPDATE  Ctacte
	SET
		Importe4	= Saldo

WHERE
	Impre = "DC"' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteIva3]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteIva3]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacteIva3]
             @Desde char(8),
             @Hasta char(8)

 AS

UPDATE  Ctacte
	SET
		Importe4	= Saldo

WHERE
	OrdFecha >= @Desde AND OrdFecha <= @Hasta' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteIva2]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteIva2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacteIva2]
             @Desde char(8),
             @Hasta char(8)

 AS

UPDATE  Ctacte
	SET
		Importe4	=0 ,
		Importe5	=0 ,
		Importe6	=0 ,
		Importe7	=Neto ,
		Importe8	=ImpoIb


WHERE
	OrdFecha >= @Desde AND OrdFecha <= @Hasta and Iva1 = 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteIva1]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteIva1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacteIva1]
             @Desde char(8),
             @Hasta char(8)

 AS

UPDATE  Ctacte
	SET
		Importe4	=Neto ,
		Importe5	=Iva1 ,
		Importe6	=Iva2 ,
		Importe7	=0  ,
		Importe8	=ImpoIb


WHERE
	OrdFecha >= @Desde AND OrdFecha <= @Hasta and Iva1 <> 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteImporteIva0]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteImporteIva0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacteImporteIva0]  AS

UPDATE  Ctacte
	SET
		Importe4  	= 0,
		Importe5 	= 0,
		Importe6 	= 0,
		Importe7 	= 0  ,
		Importe8	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteImporteIb]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteImporteIb]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacteImporteIb]  AS

UPDATE  Ctacte
	SET
		ImpoIb  	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteImporte0]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteImporte0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacteImporte0]  AS

UPDATE  Ctacte
	SET
		Importe1  	= 0,
		Importe2 	= 0,
		Importe3 	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteImporte]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteImporte]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacteImporte]
             @Clave char(12),
	@Importe1 float,
	@Importe2 float,
	@Importe3 float

 AS

UPDATE  Ctacte
	SET
		Clave     	= @Clave,
		Importe1	=@Importe1 ,
		Importe2	=@Importe2 ,
		Importe3	=@Importe3 


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteIbCiudad]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteIbCiudad]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacteIbCiudad]
             @Clave char(12)  ,
	@ClaveRecibo char(8)


 AS

UPDATE  Ctacte
	SET
		ClaveRecibo	= @ClaveRecibo  ,
		Importe4	= Neto ,
		Importe5	= Iva1 ,
		Importe6	= Iva2 ,
		Importe7	= 0  ,
		Importe8	= ImpoIbCiudad


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacteIb]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacteIb]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacteIb]
             @Clave char(12)  ,
	@ClaveRecibo char(8)


 AS

UPDATE  Ctacte
	SET
		ClaveRecibo	= @ClaveRecibo  ,
		Importe4	= Neto ,
		Importe5	= Iva1 ,
		Importe6	= Iva2 ,
		Importe7	= 0  ,
		Importe8	= ImpoIb


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtaCteCli]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtaCteCli]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ModificaCtaCteCli]
	@Clave 		varchar(12),       
	@Tipo 		varchar(2),
	@Numero      	varchar(4),
	@Renglon 	varchar(2),
	@Cliente 	varchar(6),
	@fecha      	varchar(10),
	@Estado 	varchar(1),
	@Vencimiento 	varchar(10),
	@Vencimiento1 	varchar(10),
	@Total          float,                           
	@TotalUs        float,                                       
	@Saldo          float,                                       
	@SaldoUs        float,                                       
	@OrdFecha 	varchar(8),
	@OrdVencimiento varchar(8),
	@OrdVencimiento1 varchar(8),
	@Impre 		varchar(2),
	@Empresa 	integer,
	@Neto           float,                                       
	@Iva1           float,                                       
	@Iva2           float,                                       
	@Pedido 	varchar(6),
	@Remito     	varchar(10),
	@Orden      	varchar(10),
	@Paridad        float,                           
	@Provincia 	varchar(2),
	@Vendedor    	integer,
	@Rubro       	integer,
	@Comprobante 	varchar(8),
	@Aceptada 	varchar(1),
	@Costo          float,                            
	@Importe1       float,                                       
	@Importe2       float,                                       
	@Importe3       float,                                       
	@Importe4       float,                                       
	@Importe5       float,                                       
	@Importe6       float,                                       
	@Importe7       float,                                       
	@WDate     	varchar(10)   ,
	@Seguro float  ,
	@Flete float  ,
	@ImpoIb  float



AS
UPDATE	CtaCte
	SET
		Tipo = @Tipo,
		Numero = @Numero,      
		Renglon = @Renglon, 
		Cliente = @Cliente ,
		fecha   = @fecha ,  
		Estado  = @estado,
		Vencimiento = @vencimiento,
		Vencimiento1 = @vencimiento1,
		Total = @total       ,                                        
		TotalUs = @totalus      ,                                       
		Saldo = @saldo         ,                                      
		SaldoUs = @saldous        ,                                     
		OrdFecha = @ordfecha ,
		OrdVencimiento = @ordvencimiento,
		OrdVencimiento1 = @ordvencimiento1,
		Impre = @impre,
		Empresa = @empresa,
		Neto = @neto   ,                                             
		Iva1 = @iva1    ,                                            
		Iva2 = @iva2     ,                                           
		Pedido = @pedido,
		Remito = @remito ,  
		Orden  = @orden  , 
		Paridad  = @paridad ,                                           
		Provincia = @provincia ,
		Vendedor = @vendedor  ,
		Rubro  = @rubro     ,
		Comprobante = @comprobante ,
		Aceptada = @aceptada,
		Costo  = @costo   ,                                           
		Importe1 = @importe1  ,                                          
		Importe2 = @importe2   ,                                         
		Importe3 = @importe3    ,                                        
		Importe4 = @importe4     ,                                       
		Importe5 = @importe5      ,                                      
		Importe6 = @importe6       ,                                     
		Importe7 = @importe7        ,                                    
		WDate = @wdate      	 ,
		Seguro = @Seguro  ,
		Flete = @Flete  ,
		ImpoIb = @ImpoIb
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte9]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte9]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacte9]
             @Desde char(6),
             @Hasta char(6)

 AS

UPDATE  Ctacte
	SET
		Importe1	=Total ,
		Importe2	=0 ,
		Importe3	=Saldo 


WHERE
	Cliente >= @Desde AND Cliente <= @Hasta and Total > 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte8]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte8]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacte8]
             @Desde char(6),
             @Hasta char(6)

 AS

UPDATE  Ctacte
	SET
		Importe1	=0,
		Importe2	=TotalUs ,
		Importe3	=SaldoUs 


WHERE
	Cliente >= @Desde AND Cliente <= @Hasta and Tipo >= 50 and TotalUs < 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte7]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte7]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacte7]
             @Desde char(6),
             @Hasta char(6)

 AS

UPDATE  Ctacte
	SET
		Importe1	=TotalUs ,
		Importe2	=0 ,
		Importe3	=SaldoUs


WHERE
	Cliente >= @Desde AND Cliente <= @Hasta and Tipo >= 50 and TotalUs > 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte6]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte6]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacte6]
             @Desde char(6),
             @Hasta char(6)

 AS

UPDATE  Ctacte
	SET
		Importe1	=0,
		Importe2	=Total ,
		Importe3	=Saldo 


WHERE
	Cliente >= @Desde AND Cliente <= @Hasta and Tipo >= 50 and Total < 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte5]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte5]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacte5]
             @Desde char(6),
             @Hasta char(6)

 AS

UPDATE  Ctacte
	SET
		Importe1	=Total ,
		Importe2	=0 ,
		Importe3	=Saldo 


WHERE
	Cliente >= @Desde AND Cliente <= @Hasta and Tipo >= 50 and Total > 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte4]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte4]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacte4]
             @Desde char(6),
             @Hasta char(6)

 AS

UPDATE  Ctacte
	SET
		Importe1	=0,
		Importe2	=TotalUs ,
		Importe3	=SaldoUs 


WHERE
	Cliente >= @Desde AND Cliente <= @Hasta and Tipo < 50 and TotalUs < 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte3]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte3]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacte3]
             @Desde char(6),
             @Hasta char(6)

 AS

UPDATE  Ctacte
	SET
		Importe1	=TotalUs ,
		Importe2	=0 ,
		Importe3	=SaldoUs


WHERE
	Cliente >= @Desde AND Cliente <= @Hasta and Tipo < 50 and TotalUs > 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte2]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacte2]
             @Desde char(6),
             @Hasta char(6)

 AS

UPDATE  Ctacte
	SET
		Importe1	=0,
		Importe2	=Total ,
		Importe3	=Saldo 


WHERE
	Cliente >= @Desde AND Cliente <= @Hasta and Tipo < 50 and Total < 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte12]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte12]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacte12]
             @Desde char(6),
             @Hasta char(6)

 AS

UPDATE  Ctacte
	SET
		Importe1	=0,
		Importe2	=TotalUs ,
		Importe3	=SaldoUs 


WHERE
	Cliente >= @Desde AND Cliente <= @Hasta and TotalUs < 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte11]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte11]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacte11]
             @Desde char(6),
             @Hasta char(6)

 AS

UPDATE  Ctacte
	SET
		Importe1	=TotalUs ,
		Importe2	=0 ,
		Importe3	=SaldoUs


WHERE
	Cliente >= @Desde AND Cliente <= @Hasta and TotalUs > 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte10]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte10]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacte10]
             @Desde char(6),
             @Hasta char(6)

 AS

UPDATE  Ctacte
	SET
		Importe1	=0,
		Importe2	=Total ,
		Importe3	=Saldo 


WHERE
	Cliente >= @Desde AND Cliente <= @Hasta and Total < 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte1]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacte1]
             @Desde char(6),
             @Hasta char(6)

 AS

UPDATE  Ctacte
	SET
		Importe1	=Total ,
		Importe2	=0 ,
		Importe3	=Saldo 


WHERE
	Cliente >= @Desde AND Cliente <= @Hasta and Tipo < 50 and Total > 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCtacte]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCtacte]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCtacte]
             @Clave char(12),
             @Tipo char(2),
             @Numero int,
             @Renglon char(2),
             @Cliente char(6),
             @Fecha char(10),
             @Estado char(1),
             @Vencimiento char(10),
             @Vencimiento1 char(10),
             @Total float,
             @TotalUs float,
             @Saldo float,
             @SaldoUs float,
             @Ordfecha char(8),
             @Ordvencimiento char(8),
             @Ordvencimiento1 char(8),
             @Impre char(2),
             @Empresa smallint,
             @Neto float,
             @Iva1 float,
             @Iva2 float,
             @Pedido char(6),
             @Remito char(10),
             @Orden char(10),
             @Paridad Float,
             @Provincia char(2),
             @Vendedor int,
             @Rubro int,
             @Comprobante char(8),
             @Aceptada char(1),
             @Costo float,
             @Importe1 float,
             @Importe2 float,
             @Importe3 float,
             @Importe4 float,
             @Importe5 float,
             @Importe6 float,
             @Importe7 float,
             @WDate char(10)  ,
	@Seguro float  ,
	@Flete float  ,
	@ImpoIb  float



 AS

UPDATE   Ctacte
	SET
	             Clave 		= @Clave, 		
	             Tipo  		= @Tipo,
	             Numero 		= @Numero,
	             Renglon  	= @Renglon,
	             Cliente  		= @Cliente,
	             Fecha  		= @Fecha,
	             Estado  		= @Estado,
	             Vencimiento  	= @Vencimiento,
	             Vencimiento1  	= @Vencimiento1,
	             Total  		= @Total,
	             TotalUs  	= @TotalUs,
	             Saldo  		= @Saldo,
	             SaldoUs  	= @SaldoUs,
	             Ordfecha  	= @Ordfecha,
	             Ordvencimiento  	= @Ordvencimiento,
	             Ordvencimiento1	= @Ordvencimiento1,
	             Impre  		= @Impre,
	             Empresa  	= @Empresa,
	             Neto  		= @Neto,
	             Iva1  		= @Iva1,
	             Iva2  		= @Iva2,
	             Pedido  		= @Pedido,
	             Remito  		= @Remito,
	             Orden  		= @Orden,
	             Paridad  	= @Paridad,
	             Provincia  	= @Provincia,
	             Vendedor  	= @Vendedor,
	             Rubro  		= @Rubro,
	             Comprobante  	= @Comprobante,
	             Aceptada  	= @Aceptada,
	             Costo  		= @Costo,
	             Importe1  	= @Importe1,
	             Importe2  	= @Importe2,
	             Importe3  	= @Importe3,
	             Importe4  	= @Importe4,
	             Importe5  	= @Importe5,
	             Importe6  	= @Importe6,
	             Importe7  	= @Importe7,
	             WDate  		= @WDate  ,
		Seguro		= @Seguro  ,
		Flete		= @Flete   ,
		ImpoIb		= @ImpoIb
		
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCotizaMoneda]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCotizaMoneda]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCotizaMoneda]  AS

UPDATE  Cotiza
	SET
		Moneda 		=0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCotizaDy]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCotizaDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCotizaDy]
             @Articulo char(10),
	@ArticuloDy Char(10)

 AS

uPDATE  Cotiza
	SET
		Articulo    		= @ArticuloDy

		

WHERE
	Articulo = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCotiza]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCotiza]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCotiza]
             @Clave char(8),
	@Cotiza int,
	@Renglon int,
	@Fecha char(10),
	@Proveedor char(11),
	@Articulo char(10),
	@Precio float,
	@Condicion char(40),
	@Observaciones char(40),
	@Fechaord char(8),
	@WDate char(10)


 AS

UPDATE  Cotiza
	SET
		Clave   		= @Clave,
		Cotiza		= @Cotiza,
		Renglon	= @Renglon,   
		Fecha 		= @Fecha,	
		Proveedor	= @Proveedor,	
		Articulo  	= @Articulo,	
		Precio  		= @Precio,	 
		Condicion	= @Condicion,	
		Observaciones	= @Observaciones,	
		Fechaord 	= @Fechaord,	
		WDate 	   	= @WDate



WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaConsigFacturado]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaConsigFacturado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaConsigFacturado]
             @Clave char(8),
	@Facturado int

 AS

UPDATE  Consig
	SET
		Facturado 	= @Facturado


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaConsig]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaConsig]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaConsig]
             @Clave char(8),
	@Numero int,
	@Renglon int,
	@Cliente char(6),
	@Fecha char(10),
	@Observaciones char(50),
	@Terminado char(12),
	@Cantidad int,
	@Envase1 int,
	@Canti1 int,
	@Envase2 int,
	@Canti2 int,
	@Envase3 int,
	@Canti3 int,
	@Envase4 int,
	@Canti4 int,
	@Fechaord char(8),
	@Precio float,
	@Linea int,
	@Facturado int,
	@Importe int

 AS

UPDATE  Consig
	SET
		Clave   		= @Clave,
		Numero		= @Numero,
		Renglon	= @Renglon,   
		Cliente 		= @Cliente,	
		Fecha 		= @Fecha,	
		Observaciones 	= @Observaciones,	
		Terminado 	= @Terminado,	
		Cantidad	= @Cantidad,	
		Envase1 	= @Envase1,	
		Canti1 		= @Canti1,	
		Envase2 	= @Envase2,	
		Canti2 		= @Canti2,	
		Envase3 	= @Envase3,	
		Canti3 		= @Canti3,	
		Envase4 	= @Envase4,	
		Canti4		= @Canti4,	
		Fechaord 	= @Fechaord,	
		Precio 		= @Precio,	
		Linea 		= @Linea,	
		Facturado 	= @Facturado,	
		Importe 		= @Importe


WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaComposicionDy]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaComposicionDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaComposicionDy]
             @Articulo char(10),
	@ArticuloDy Char(10)

 AS

uPDATE  composicion
	SET
		Articulo1    		= @ArticuloDy

		

WHERE
	Articulo1 = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaComposicionCosto]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaComposicionCosto]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaComposicionCosto]
             @Clave char(14),
	@Costo1 float,
	@Costo2 float,
	@DescriTerminado Char(30),
	@DescriArticulo1 Char(30),
	@DescriArticulo2 Char(30)

 AS

uPDATE  composicion
	SET
		Clave     		= @Clave,
		Costo1       		= @Costo1,
		Costo2      	   	= @Costo2, 
		DescriTerminado      	= @DescriTerminado,
		DescriArticulo1      	= @DescriArticulo1,
		DescriArticulo2      	= @DescriArticulo2
		

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaComposicion]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaComposicion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaComposicion]
             @Clave char(14),
	@Terminado char(12),
	@Renglon char(2),
	@Tipo char(1),
	@Articulo1 char(10),
	@Articulo2 char(12),
	@Cantidad float,
	@WDate char(10),
	@Costo1 float,
	@Costo2 float  


 AS

UPDATE  Composicion
	SET
		Clave     	= @Clave,
		Terminado       	= @Terminado,
		Renglon       	= @Renglon,
		Tipo       	= @Tipo,
		Articulo1       	= @Articulo1,
		Articulo2       	= @Articulo2,
		Cantidad      	= @Cantidad,
		Wdate       	= @WDate,
		Costo1       	= @Costo1,
		Costo2      	= @Costo2

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaClienteImporte0]    Script Date: 05/01/2016 17:47:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaClienteImporte0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaClienteImporte0]  AS

UPDATE  Cliente
	SET
		Importe1  	= 0,
		Importe2 	= 0,
		Importe3	= 0,
		Importe4 	= 0,
		Importe5 	= 0,
		Importe6 	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaClienteIb]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaClienteIb]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaClienteIb]  AS

UPDATE  Cliente
	SET
		Ib  	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCliente1]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCliente1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCliente1]
             @Cliente char(6),
	@Razon char(50),
	@Direccion char(50),
	@Localidad char(50),
	@Provincia char(2),	
	@Postal char(10),	
	@EMail char(40),	
	@Fax char(20),
	@Telefono char(40),
	@Cuit char(15),
	@Contacto char(50),
	@Observaciones char(100),	
	@Vendedor int,
	@Iva char(1),
	@Rubro int,
	@Horario char(20),
	@Pago1 int,
	@Pago2 int,
	@Limite float,
	@Minimo float,
	@Direntrega char(50),
	@Importe1 float,
	@Importe2 float,
	@Importe3 float,
	@Importe4 float,
	@Importe5 float,
	@Importe6 float,
	@WDate char(10)  ,
	@Precio char(2)  ,
	@Ib int


 AS

UPDATE  Cliente
	SET
		Cliente     	= @Cliente,
		Razon		=@Razon,
		Direccion	=@Direccion,
		Localidad	=@Localidad,
		Provincia	=@Provincia ,	
		Postal		=@Postal,	
		EMail		=@EMail,	
		Fax		=@Fax ,
		Telefono	=@Telefono ,
		Cuit		=@Cuit ,
		Contacto	=@Contacto ,
		Observaciones	=@Observaciones ,	
		Vendedor	=@Vendedor ,
		Iva		=@Iva ,
		Rubro		=@Rubro ,
		Horario		=@Horario ,
		Pago1		=@Pago1 ,
		Pago2		=@Pago2 ,
		Limite		=@Limite ,
		Minimo		=@Minimo ,
		Direntrega	=@Direntrega ,
		Importe1	=@Importe1 ,
		Importe2	=@Importe2 ,
		Importe3	=@Importe3 ,
		Importe4	=@Importe4 ,
		Importe5	=@Importe5 ,
		Importe6	=@Importe6 ,
		WDate		=@WDate  ,
		Precio		=@Precio  ,
		Ib		=@Ib


WHERE
	Cliente = @Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCliente]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCliente]
             @Cliente char(6),
	@Razon char(50),
	@Direccion char(50),
	@Localidad char(50),
	@Provincia char(2),	
	@Postal char(10),	
	@EMail char(40),	
	@Fax char(20),
	@Telefono char(40),
	@Cuit char(15),
	@Contacto char(50),
	@Observaciones char(100),	
	@Vendedor int,
	@Iva char(1),
	@Rubro int,
	@Horario char(20),
	@Pago1 int,
	@Pago2 int,
	@Limite float,
	@Minimo float,
	@Direntrega char(50),
	@Importe1 float,
	@Importe2 float,
	@Importe3 float,
	@Importe4 float,
	@Importe5 float,
	@Importe6 float,
	@WDate char(10)  ,
	@Precio char(2)  ,
	@Ib int


 AS

UPDATE  Cliente
	SET
		Cliente     	= @Cliente,
		Razon		=@Razon,
		Direccion	=@Direccion,
		Localidad	=@Localidad,
		Provincia	=@Provincia ,	
		Postal		=@Postal,	
		EMail		=@EMail,	
		Fax		=@Fax ,
		Telefono	=@Telefono ,
		Cuit		=@Cuit ,
		Contacto	=@Contacto ,
		Observaciones	=@Observaciones ,	
		Vendedor	=@Vendedor ,
		Iva		=@Iva ,
		Rubro		=@Rubro ,
		Horario		=@Horario ,
		Pago1		=@Pago1 ,
		Pago2		=@Pago2 ,
		Limite		=@Limite ,
		Minimo		=@Minimo ,
		Direntrega	=@Direntrega ,
		Importe1	=@Importe1 ,
		Importe2	=@Importe2 ,
		Importe3	=@Importe3 ,
		Importe4	=@Importe4 ,
		Importe5	=@Importe5 ,
		Importe6	=@Importe6 ,
		WDate		=@WDate  ,
		Precio		=@Precio  ,
		Ib		=@Ib


WHERE
	Cliente = @Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCambioAdm]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCambioAdm]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCambioAdm]
             @Fecha char(10),
	@Cambio float,
	@Ordfecha char(10)


 AS

UPDATE  CambioAdm
	SET
		Fecha    	= @Fecha,
		Cambio    	= @Cambio,
		OrdFecha       	= @OrdFecha

WHERE
	Fecha = @Fecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaCambio]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaCambio]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaCambio]
             @Fecha char(10),
	@Cambio float,
	@Ordfecha char(10)


 AS

UPDATE  Cambios
	SET
		Fecha    	= @Fecha,
		Cambio    	= @Cambio,
		OrdFecha       	= @OrdFecha

WHERE
	Fecha = @Fecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaBanco]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaBanco]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaBanco]
             @Banco smallint,
	@Nombre char(50),
	@Cuenta char(10),
	@Empresa int

 AS

UPDATE  Banco
	SET
		Banco   	= @Banco,
		Nombre       	= @Nombre,                                     
		Cuenta       	= @Cuenta,                                     
		Empresa      	= @Empresa
WHERE
	Banco = @Banco' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloVenta0]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloVenta0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloVenta0]  AS

UPDATE  Articulo
	SET
		Venta		= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloVenta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloVenta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloVenta]
             @Codigo char(10),
	@Venta float,
	@WDate char(10)
 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Venta      	= @Venta,                                    
		WDate      	= @WDate
	
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloVariosII]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloVariosII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloVariosII]
             @Codigo char(10),
	@Descripcion char(50),
	@Costo1 float,
	@Costo2 float,
	@Unidad char(10),
	@Envase int,
	@Rs char(1),
	@Proveedor char(11),
	@Flete float,
	@Moneda char(3),
	@Controla  int ,
	@Densidad char(20)  ,
	@Costo3 float  ,
	@WCosto1 float  ,
	@WCosto2 float  ,

             @WCosto3 float
 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Descripcion       	= @Descripcion,                                     
		Costo1      	= @Costo1,                                    
		Costo2      	= @Costo2,                                    
		Unidad      	= @Unidad,                                    
		Envase      	= @Envase,                                    
		Rs      	             = @Rs,                                    
		Proveedor      	= @Proveedor,                                    
		Flete      	= @Flete,                                    
		Moneda      	= @Moneda,
		Controla    	= @Controla  ,
		Densidad	= @Densidad  ,
		Costo3		= @Costo3  ,
		WCosto1	= @WCosto1  ,
		WCosto2	= @WCosto2  ,
		WCosto3	= @WCosto3  
                          

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloVarios1]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloVarios1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloVarios1]
             @Codigo char(10),
	@Envase int,
	@Proveedor char(11)
 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Envase      	= @Envase,                                    
		Proveedor      	= @Proveedor
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloVarios]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloVarios]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloVarios]
             @Codigo char(10),
	@Descripcion char(50),
	@Costo1 float,
	@Costo2 float,
	@Unidad char(10),
	@Envase int,
	@Rs char(1),
	@Proveedor char(11),
	@Flete float,
	@Moneda char(3),
	@Controla  int ,
	@Densidad char(20)  ,
	@Costo3 float
 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Descripcion       	= @Descripcion,                                     
		Costo1      	= @Costo1,                                    
		Costo2      	= @Costo2,                                    
		Unidad      	= @Unidad,                                    
		Envase      	= @Envase,                                    
		Rs      	             = @Rs,                                    
		Proveedor      	= @Proveedor,                                    
		Flete      	= @Flete,                                    
		Moneda      	= @Moneda,
		Controla    	= @Controla  ,
		Densidad	= @Densidad  ,
		Costo3		= @Costo3
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloStock0]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloStock0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloStock0]

 AS

UPDATE  Articulo
	SET
		Stock	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloStock]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloStock]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloStock]
             @Codigo char(10),
	@Stock float  ,
	@Costo  Float

 AS

UPDATE  Articulo
	SET
		Stock		= @Stock  ,
		Costo		= @Costo

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloSalidas]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloSalidas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloSalidas]
             @Codigo char(10),
	@Salidas float,
	@WDate char(10)

 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Salidas      	= @Salidas,                                    
		WDate      	= @WDate

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloPedido0DesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloPedido0DesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloPedido0DesdeHasta] 
             @Desde char(10)  ,
	@Hasta char(10)

as

UPDATE  Articulo
	SET
		Pedido		= 0


WHERE
	@Desde <= Codigo and @Hasta >= Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloPedido0]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloPedido0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloPedido0]  AS

UPDATE  Articulo
	SET
		Pedido		= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloPedido]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloPedido]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloPedido]
             @Codigo char(10),
	@Pedido  Float



 AS

UPDATE  Articulo
	SET
		Pedido		= Pedido + @Pedido

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloOrdenLaboratorio]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloOrdenLaboratorio]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloOrdenLaboratorio]
             @Codigo char(10),
	@Pedido float,
	@WDate char(10)

 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Pedido      	= @Pedido,                                    
		WDate      	= @WDate

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloOrden]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloOrden]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloOrden]
             @Codigo char(10),
	@Costo1 float,
	@Pedido float,
	@Fecha char(10),
	@Orden int,
	@Proveedor char(11),
	@WDate char(10)

 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Costo1      	= @Costo1,                                    
		Pedido      	= @Pedido,                                    
		Fecha      	= @Fecha,                                    
		Orden      	= @Orden,                                    
		Proveedor      	= @Proveedor,                                    
		WDate      	= @WDate

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloMovimientosSuma]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloMovimientosSuma]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloMovimientosSuma]
             @Codigo char(10),
	@Entradas float,
	@Salidas float,
	@WDate char(10)

 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Entradas      	= Entradas + @Entradas,                                    
		Salidas      	= Salidas + @Salidas,                                    
		WDate      	= @WDate

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloMovimientos]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloMovimientos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloMovimientos]
             @Codigo char(10),
	@Entradas float,
	@Salidas float,
	@WDate char(10)

 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Entradas      	= @Entradas,                                    
		Salidas      	= @Salidas,                                    
		WDate      	= @WDate

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloMovi]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloMovi]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloMovi]
             @Codigo char(10),
	@Inicial float,
	@Entradas float,
	@Salidas float,
	@Laboratorio float,
	@Pedido float,
	@WDate char(10)

 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Inicial      	= @Inicial,                                    
		Entradas      	= @Entradas,                                    
		Salidas      	= @Salidas,                                    
		Laboratorio      	= @Laboratorio,                                    
		Pedido      	= @Pedido,                                    
		WDate      	= @WDate

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloMinimo1]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloMinimo1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloMinimo1]
             @Codigo char(10),
	@Minimo1 float

 AS

UPDATE  Articulo
	SET
		Minimo1      	= @Minimo1

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloLeyenda]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloLeyenda]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloLeyenda]
             @Codigo char(10),
	@Leyenda int

 AS

UPDATE  Articulo
	SET
		Leyenda      	= @Leyenda

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloLaudoPesos]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloLaudoPesos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloLaudoPesos]
             @Codigo char(10),
	@Laboratorio float,
	@Entradas float,
	@WDate char(10)  ,
	@Costo1 float  ,
	@Costo3  Float



 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Laboratorio      	= @Laboratorio,                                    
		Entradas      	= @Entradas,                                    
		WDate      	= @WDate  ,
		WCosto1	= @Costo1  ,
		WCosto3	= @Costo3

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloLaudoDolares]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloLaudoDolares]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloLaudoDolares]
             @Codigo char(10),
	@Laboratorio float,
	@Entradas float,
	@WDate char(10)  ,
	@Costo1 float  ,
	@Costo3  Float



 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Laboratorio      	= @Laboratorio,                                    
		Entradas      	= @Entradas,                                    
		WDate      	= @WDate  ,
		Costo1		= @Costo1  ,
		Costo3		= @Costo3

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloLaudo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloLaudo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloLaudo]
             @Codigo char(10),
	@Laboratorio float,
	@Entradas float,
	@WDate char(10)  ,
	@Costo1 float  ,
	@Costo3  Float



 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Laboratorio      	= @Laboratorio,                                    
		Entradas      	= @Entradas,                                    
		WDate      	= @WDate  ,
		Costo1		= @Costo1  ,
		Costo3		= @Costo3

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloLaboratorio0DesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloLaboratorio0DesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloLaboratorio0DesdeHasta] 
             @Desde char(10)  ,
	@Hasta char(10)

as

UPDATE  Articulo
	SET
		Laboratorio	= 0

WHERE
	@Desde <= Codigo and @Hasta >= Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloLaboratorio0]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloLaboratorio0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloLaboratorio0]  AS

UPDATE  Articulo
	SET
		Laboratorio	= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloLaboratorio]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloLaboratorio]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloLaboratorio]
             @Codigo char(10),
	@Laboratorio float
 AS

UPDATE  Articulo
	SET
		Laboratorio      	= Laboratorio + @Laboratorio

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloInicial0]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloInicial0]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloInicial0]  AS

UPDATE  Articulo
	SET
		Inicial	  	= 0  ,
		Laboratorio	= 0  ,
		Pedido		= 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloInformeLaboratorio]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloInformeLaboratorio]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloInformeLaboratorio]
             @Codigo char(10),
	@Pedido float,
	@Laboratorio float,
	@WDate char(10)

 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Pedido      	= @Pedido,                                    
		Laboratorio      	= @Laboratorio,                                    
		WDate      	= @WDate

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloInforme]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloInforme]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloInforme]
             @Codigo char(10),
	@Pedido float,
	@Laboratorio float,
	@WDate char(10),
	@Envase int,
	@Proveedor char(11)

 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Pedido      	= @Pedido,                                    
		Laboratorio      	= @Laboratorio,                                    
		WDate      	= @WDate,
		Envase		= @Envase,
		Proveedor	= @Proveedor

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloII]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloII]
             @Codigo char(10),
	@Descripcion char(50),
	@Costo1 float,
	@Costo2 float,
	@Inicial float,
	@Entradas float,
	@Salidas float,
	@Minimo float,
	@Laboratorio float,
	@Unidad char(10),
	@Pedido float,
	@Deposito char(20),
	@Envase int,
	@Rs char(1),
	@Fecha char(10),
	@Orden int,
	@Dife float,
	@Proveedor char(11),
	@WDate char(10),
	@Flete float,
	@Moneda char(3),
	@Controla  int ,
	@Densidad char(20)  ,
	@Costo3 float  ,
	@WCosto1   float   ,
	@WCosto2   float  ,
	@WCosto3   float  ,
             @Venta   float
 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Descripcion       	= @Descripcion,                                     
		Costo1      	= @Costo1,                                    
		Costo2      	= @Costo2,                                    
		Inicial      	= @Inicial,                                    
		Entradas      	= @Entradas,                                    
		Salidas      	= @Salidas,                                    
		Minimo      	= @MInimo,                                    
		Laboratorio      	= @Laboratorio,                                    
		Unidad      	= @Unidad,                                    
		Pedido      	= @Pedido,                                    
		Deposito      	= @Deposito,                                    
		Envase      	= @Envase,                                    
		Rs      	             = @Rs,                                    
		Fecha      	= @Fecha,                                    
		Orden      	= @Orden,                                    
		Dife      		= @Dife,                                    
		Proveedor      	= @Proveedor,                                    
		WDate      	= @WDate,                                    
		Flete      	= @Flete,                                    
		Moneda      	= @Moneda,
		Controla    	= @Controla  ,
		Densidad	= @Densidad  ,
		Costo3		= @Costo3  ,
		WCosto1	= @WCosto1  ,
		WCosto2	= @WCosto2  ,
		WCosto3	= @WCosto3   ,
		Venta		= @Venta 

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloFecha]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloFecha]
	@FechaCierre char(10),
	@OrdFechaCierre char(8)
	

 AS

UPDATE  Articulo
	SET
		FechaCierre      	= @FechaCierre ,                                    
		OrdFechaCierre 	= @OrdFechaCierre' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloFacturas]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloFacturas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloFacturas]
             @Codigo char(10),
	@Venta float,
	@Salidas float,
	@WDate char(10)
 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Venta      	= @Venta,                                    
		Salidas      	= @Salidas,                                    
		WDate      	= @WDate
	
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloDy]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloDy]
             @Articulo char(10),
	@ArticuloDy Char(10)

 AS

uPDATE  Articulo
	SET
		Codigo   		= @ArticuloDy

		

WHERE
	Codigo = @Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloDescriComercial]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloDescriComercial]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloDescriComercial]
             @Codigo char(10),
	@DescriComercial char(50)
 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		DescriComercial 	= @DescriComercial
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloCostoTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloCostoTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloCostoTotal]
 AS

UPDATE  Articulo
	SET
		Costo3 		= Costo1' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloCostoPesos]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloCostoPesos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloCostoPesos]
             @Codigo char(10),
	@Costo1 float
	

 AS

UPDATE  Articulo
	SET
		WCosto1      	= @Costo1

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloCostoImpre]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloCostoImpre]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloCostoImpre]
             @Codigo char(10),
	@Costo float

 AS

UPDATE  Articulo
	SET
		Costo		= @Costo

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloCostoImportacion]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloCostoImportacion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloCostoImportacion]
             @Codigo char(10),
	@Costo1 float  ,
	@Costo2 float  ,
	@Costo3 float  ,
	@Flete float
	

 AS

UPDATE  Articulo
	SET
		Costo1      	= @Costo1  ,
		Costo2     	= @Costo2 ,
		Costo3     	= @Costo3 ,
		Flete 		= @Flete

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloCostoDolares]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloCostoDolares]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloCostoDolares]
             @Codigo char(10),
	@Costo1 float
	

 AS

UPDATE  Articulo
	SET
		Costo1      	= @Costo1

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloCosto1]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloCosto1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloCosto1]
             @Codigo char(10),
	@Costo1 float
	

 AS

UPDATE  Articulo
	SET
		Costo1      	= @Costo1

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticuloCosto]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticuloCosto]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticuloCosto]
             @Codigo char(10),
	@Costo1 float,
	@Costo2 float  ,
	@Costo3 float
	

 AS

UPDATE  Articulo
	SET
		Costo1      	= @Costo1,                                    
		Costo2      	= @Costo2 ,
		Costo3 		= @Costo3

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticulo1]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticulo1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticulo1]

             @Codigo char(10),
	@Descripcion char(50),
	@Costo1 float,
	@Costo2 float,
	@Inicial float,
	@Entradas float,
	@Salidas float,
	@Minimo float,
	@Laboratorio float,
	@Unidad char(10),
	@Pedido float,
	@Deposito char(20),
	@Envase int,
	@Rs char(1)



 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Descripcion       	= @Descripcion,                                     
		Costo1      	= @Costo1,                                    
		Costo2      	= @Costo2,                                    
		Inicial      	= @Inicial,                                    
		Entradas      	= @Entradas,                                    
		Salidas      	= @Salidas,                                    
		Minimo      	= @MInimo,                                    
		Laboratorio      	= @Laboratorio,                                    
		Unidad      	= @Unidad,                                    
		Pedido      	= @Pedido,                                    
		Deposito      	= @Deposito,                                    
		Envase      	= @Envase,                                    
		Rs      	             = @Rs


WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ModificaArticulo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ModificaArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ModificaArticulo]
             @Codigo char(10),
	@Descripcion char(50),
	@Costo1 float,
	@Costo2 float,
	@Inicial float,
	@Entradas float,
	@Salidas float,
	@Minimo float,
	@Laboratorio float,
	@Unidad char(10),
	@Pedido float,
	@Deposito char(20),
	@Envase int,
	@Rs char(1),
	@Fecha char(10),
	@Orden int,
	@Dife float,
	@Proveedor char(11),
	@WDate char(10),
	@Flete float,
	@Moneda char(3),
	@Controla  int ,
	@Densidad char(20)  ,
	@Costo3 float
 AS

UPDATE  Articulo
	SET
		Codigo   	= @Codigo,
		Descripcion       	= @Descripcion,                                     
		Costo1      	= @Costo1,                                    
		Costo2      	= @Costo2,                                    
		Inicial      	= @Inicial,                                    
		Entradas      	= @Entradas,                                    
		Salidas      	= @Salidas,                                    
		Minimo      	= @MInimo,                                    
		Laboratorio      	= @Laboratorio,                                    
		Unidad      	= @Unidad,                                    
		Pedido      	= @Pedido,                                    
		Deposito      	= @Deposito,                                    
		Envase      	= @Envase,                                    
		Rs      	             = @Rs,                                    
		Fecha      	= @Fecha,                                    
		Orden      	= @Orden,                                    
		Dife      		= @Dife,                                    
		Proveedor      	= @Proveedor,                                    
		WDate      	= @WDate,                                    
		Flete      	= @Flete,                                    
		Moneda      	= @Moneda,
		Controla    	= @Controla  ,
		Densidad	= @Densidad  ,
		Costo3		= @Costo3
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaVendedor]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaVendedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaVendedor] AS

Select * From Vendedor
Order by Vendedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaTerminadoDesdeHastaMinimo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaTerminadoDesdeHastaMinimo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaTerminadoDesdeHastaMinimo]
	@Desde char(12),
	@Hasta char(12)
		 AS

Select Codigo,Descripcion,Inicial,Salidas,Entradas,Minimo,Minimo1,Pedido From Terminado
WHERE
	Codigo >= @Desde and Codigo <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaTerminadoDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaTerminadoDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaTerminadoDesdeHasta] 
	@Desde char(12),
	@Hasta char(12)
		 AS

Select * From Terminado
WHERE
	Codigo >= @Desde and Codigo <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaTerminadoConsulta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaTerminadoConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaTerminadoConsulta] AS

Select Codigo, Descripcion

From Terminado

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaTerminadoArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaTerminadoArticuloDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaTerminadoArticuloDesdeHasta]
	@Desde char(10),
	@Hasta char(10)
		 AS

SELECT * FROM Movvar
WHERE
	Terminado >= @Desde and Terminado <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaTerminado]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaTerminado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaTerminado] AS

Select * From Terminado
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSoltotnan2]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSoltotnan2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSoltotnan2] AS

SELECT * FROM Soltot


Order by ordentrega  desc' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSoltotnan1]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSoltotnan1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSoltotnan1] AS

SELECT * FROM Soltot


Order by planta  desc' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSoltotnan]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSoltotnan]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSoltotnan] AS

SELECT * FROM Soltot


Order by fechaord  desc' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSoltot]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSoltot]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSoltot] AS

SELECT * FROM Soltot


Order by Articulo,OrdEntrega' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolicitudTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolicitudTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolicitudTotal] AS

SELECT * FROM Solic
Order by Solicitud' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolicitudPendienteArticulo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolicitudPendienteArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolicitudPendienteArticulo]
	@Articulo char(10) ,
	@Marca char(1)
 AS

SELECT * FROM Solic
WHERE
	Marca <> @Marca and Articulo = @Articulo
Order by Articulo,clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolicitudPendiente]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolicitudPendiente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolicitudPendiente]
	@Marca char(1)
 AS

SELECT * FROM Solic
WHERE
	Marca <> @Marca
Order by Articulo,clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolicitudNumero]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolicitudNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolicitudNumero] AS

SELECT Max(Solicitud), Solicitud

FROM Solic

Group By Solicitud

Order by Solicitud' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolicitudBajaArticulo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolicitudBajaArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolicitudBajaArticulo]
	@Articulo char(10) ,
	@Marca char(1)
 AS

SELECT * FROM Solic
WHERE
	Articulo = @Articulo
Order by Articulo,clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolicitudArticulo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolicitudArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolicitudArticulo]
	@Solicitud   int ,
	@Articulo char(10)
 AS

SELECT * FROM Solic
WHERE
	Solicitud = @Solicitud and Articulo = @Articulo
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolicitud]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolicitud]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolicitud]
	@Solicitud  int
 AS

SELECT * FROM SOlic
WHERE
	Solicitud = @Solicitud
Order by Solicitud' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolHojaTotalListado]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolHojaTotalListado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolHojaTotalListado] AS

SELECT *

FROM SolHoja

where Marca <> "X" 

Order by Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolHojaNumero]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolHojaNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolHojaNumero]
	
 AS

SELECT MAX(Hoja), hoja

FROM SolHoja

group by hoja




order by hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolHoja]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolHoja]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolHoja]
	@Hoja  int
 AS

SELECT * FROM SolHoja
WHERE
	Hoja = @Hoja
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolGuiaTotalTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolGuiaTotalTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolGuiaTotalTotal] AS

SELECT * FROM SolGuiaTotal
Order by Articulo,Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolGuiaTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolGuiaTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolGuiaTotal] AS

SELECT * FROM SolGuiaTotal
Order by Desde,Hasta,Articulo,Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolGuiaPendiente]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolGuiaPendiente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolGuiaPendiente]
	@Marca char(1)
 AS

SELECT * FROM SolGuia
WHERE
	Marca = @Marca
Order by Articulo,Terminado,clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolGuiaNumero]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolGuiaNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolGuiaNumero] AS

SELECT Max(Codigo), Codigo

FROM SolGuia

Group By Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaSolguia]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSolguia]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaSolguia]
	@Codigo int
 AS

SELECT * FROM SolGuia
WHERE
	Codigo = @Codigo 
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRubro]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRubro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaRubro] AS

Select * From Rubros
Order by Rubro' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRetencion]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRetencion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRetencion] AS

SELECT * FROM Retencion' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosTotal] AS

SELECT * FROM Recibos
Order by Recibo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosNumero]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosNumero] AS

SELECT MAX(Recibo), Recibo

FROM Recibos

Group By Recibo

Order by Recibo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosNroCheque]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosNroCheque]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosNroCheque] AS

SELECT * FROM Recibos

WHERE Tiporeg = "2" and Estado2 <> "X"

ORDER BY Fechaord2' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosMovban]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosMovban]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaRecibosMovban]
	@Cuenta1 char(10),
	@Cuenta2 char(10),
	@Cuenta3 char(10)
		 AS

Select * From Recibos

WHERE
	Cuenta = @Cuenta1 or Cuenta = @Cuenta2 or Cuenta = @Cuenta3

Order by Recibo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosFecha]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaRecibosFecha]
	@DesdeFecha char(8),
	@HastaFecha char(8)
		 AS

Select * From Recibos

WHERE
	FechaOrd >= @DesdeFecha and FechaOrd <= @HastaFecha 

Order by Recibo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosFactura]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosFactura]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosFactura]
	@Tipo char(2),
	@Numero char(8)
		 AS

SELECT * FROM Recibos
WHERE
	Tipo1 = @Tipo and Numero1 = @Numero


Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDifeOtroVI]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDifeOtroVI]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosDifeOtroVI]
		 AS

SELECT * FROM Recibos
WHERE
	Tiporeg = "2" and Marca = "X"  and Fechadepoord < FechaOrd2 and Fechaord > "20011231"

Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDifeOtroV]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDifeOtroV]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosDifeOtroV]
	@Desde char(8),
	@Hasta char(8)
		 AS

SELECT * FROM Recibos
WHERE
	fechadepoord >= @Desde and fechadepoord <= @Hasta and Renglon = "01" and Marca ="X"

Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDifeOtroIV]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDifeOtroIV]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosDifeOtroIV]
		 AS

SELECT * FROM Recibos
WHERE
	Tiporeg = "2" and Marca <> "X" and Fechaord > "20011231"

Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDifeOtroIII]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDifeOtroIII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosDifeOtroIII]
		 AS

SELECT * FROM Recibos
WHERE
	Tiporeg = "2" and Estado2 <> "X"

Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDifeOtroII]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDifeOtroII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosDifeOtroII]
	@Desde char(8),
	@Hasta char(8)
		 AS

SELECT * FROM Recibos
WHERE
	fechadepoord >= @Desde and fechadepoord <= @Hasta and Tiporeg = "2" and Marca ="X"

Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDifeOtroI]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDifeOtroI]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosDifeOtroI]
	@Desde char(8),
	@Hasta char(8)
		 AS

SELECT * FROM Recibos
WHERE
	fechadepoord >= @Desde and fechadepoord <= @Hasta and Tiporeg = "1" and Marca ="X"

Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDifeI]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDifeI]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosDifeI]
	@Desde char(8),
	@Hasta char(8)
		 AS

SELECT * FROM Recibos
WHERE
	fechaord >= @Desde and fechaord <= @Hasta and Tiporeg = "1"

Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosDeposito]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosDeposito]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosDeposito] AS

SELECT * FROM Recibos


ORDER BY  Fechaord2,Numero2' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosCliente]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosCliente]
	@Desde char(6),
	@Hasta char(6)
		 AS

SELECT * FROM Recibos
WHERE
	Cliente >= @Desde and Cliente <= @Hasta


Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosCartera]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosCartera]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosCartera] AS

SELECT * FROM Recibos

WHere	Tiporeg= "2" and  Estado2 <> "X" 

ORDER BY  Fechaord2,Numero2' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibosBusqueda]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibosBusqueda]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibosBusqueda] AS

SELECT Numero2, Banco2, Importe2, Fecha, Fecha2, Recibo, Cliente FROM Recibos

WHere	Tiporeg= "2" and  (Tipo2 = "2" or Tipo2 = "02")

ORDER BY  Fechaord2,Numero2' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaRecibos]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaRecibos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaRecibos] AS

SELECT * FROM Recibos


ORDER BY RECIBO' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPrueterConsulta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPrueterConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaPrueterConsulta] AS

Select Prueba, Producto, Fecha, Lote

From Prueter

Order by Prueba' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPrueter]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPrueter]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaPrueter] AS

Select * From Prueter
Order by Prueba' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPruebaConsulta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPruebaConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaPruebaConsulta] AS

Select Prueba, Producto, Fecha

From Prueart

Order by fechaord,Prueba' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPrueba]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPrueba]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaPrueba] AS

Select * From Prueart

Order by Prueba' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaProveedoresOrdConsultaII]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaProveedoresOrdConsultaII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaProveedoresOrdConsultaII]
	@Nombre  char varying(50)
 AS

SELECT * FROM Proveedor
WHERE
	Nombre like("%"+@Nombre+"%")
Order by Nombre' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaProveedoresOrdConsulta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaProveedoresOrdConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaProveedoresOrdConsulta] AS

Select Proveedor,Nombre From Proveedor
Order by Nombre' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaProveedoresOrd]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaProveedoresOrd]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaProveedoresOrd] AS

Select * From Proveedor
Order by Nombre' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaProveedores]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaProveedores]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaProveedores] AS

Select * From Proveedor
Order by Proveedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPrestamoTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPrestamoTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPrestamoTotal] AS

SELECT * FROM Prestamo
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPreciosMp]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPreciosMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaPreciosMp] AS

Select * From PreciosMp
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPreciosClienteMp]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPreciosClienteMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaPreciosClienteMp]
	@Cliente char(6)
AS

Select  Cliente, Articulo, Precio

From PreciosMp

Where
	Cliente =  @Cliente

Order by Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPreciosCliente]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPreciosCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaPreciosCliente] 
	@Cliente char(6)
AS

Select  Cliente, Terminado, Descripcion , Precio

From Precios

Where
	Cliente =  @Cliente

Order by Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPrecios]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPrecios]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaPrecios] AS

Select * From Precios
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTotalListado4]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTotalListado4]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoTotalListado4] AS

SELECT *

FROM Pedido

where Proceso1 = 2

Order by Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTotalListado3]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTotalListado3]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoTotalListado3] AS

SELECT *

FROM Pedido

where Autorizo = "X" and Impresion1 = "N" and TipoPedido = 3

Order by Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTotalListado2]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTotalListado2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoTotalListado2] AS

SELECT *

FROM Pedido

where Autorizo = "X" and Impresion1 = "N" and TipoPedido = 1

Order by Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTotalListado1]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTotalListado1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoTotalListado1] AS

SELECT *

FROM Pedido

where Autorizo = "X" and Impresion1 = "N"

Order by Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTotalListado]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTotalListado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoTotalListado] AS

SELECT *

FROM Pedido

where Autorizo = "X" and Impresion <> "X"

Order by Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoTotal] AS

SELECT * FROM Pedido
Order by Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTerminadoPendiente]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTerminadoPendiente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoTerminadoPendiente]
	@Terminado char(12)
 AS

SELECT * FROM Pedido
WHERE
	Terminado = @Terminado AND Importe > 0 
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoTerminado]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoTerminado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoTerminado]
	@Terminado char(12)
 AS

SELECT * FROM Pedido
WHERE
	Terminado = @Terminado
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoPigmentos]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoPigmentos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoPigmentos] AS

SELECT *

FROM Pedido

where  TipoPedido = 5 and Impresion2 <> "S" 

Order by Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoPendDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoPendDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoPendDesdeHasta]
             @Desde char(10)  ,
	@Hasta char(10)
as

SELECT Terminado,Importe

FROM Pedido

WHERE
	@Desde <= Articulo and @Hasta >= Articulo and Importe > 0

Order by Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoPend]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoPend]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoPend] AS

SELECT *

FROM Pedido

where Importe > 0

Order by Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoNumero]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoNumero] AS

SELECT max(Pedido), Pedido 
FROM Pedido
group by Pedido
Order by Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoII]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoII]
	@Desde char(8),
	@Marca char(1)
 AS

SELECT * FROM Pedido
WHERE
	TipoPedido = 1 
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoFechaMarcaColor]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoFechaMarcaColor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoFechaMarcaColor]
	@Desde char(8)  ,
	@Marca char(1)
 AS

SELECT * FROM Pedido
WHERE
	@Desde > Fechaord and TipoPedido = 1 and Impresion2 <> @Marca
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoFechaMarca]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoFechaMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoFechaMarca]
	@Desde char(8) ,
	@Marca char(1)
 AS

SELECT * FROM Pedido
WHERE
	@Desde > Fechaord and Autorizo = @Marca
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoFecha]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoFecha]
	@Desde char(8) ,
	@Hasta char(8)
 AS

SELECT * FROM Pedido
WHERE
	@Desde <= Fechaord and @Hasta >= Fechaord
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoDevolTotalListado]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoDevolTotalListado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoDevolTotalListado] AS

SELECT *

FROM PedidoDevol

where Autorizo = "X" and Impresion <> "X"

Order by Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoDevolTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoDevolTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoDevolTotal] AS

SELECT * FROM PedidoDevol
Order by Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoDevolNumero]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoDevolNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoDevolNumero] AS

SELECT max(Pedido), Pedido 
FROM PedidoDevol
group by Pedido
Order by Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoDevolFechaMarca]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoDevolFechaMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoDevolFechaMarca]
	@Desde char(8) ,
	@Marca char(1)
 AS

SELECT * FROM PedidoDevol
WHERE
	@Desde > Fechaord and Autorizo = @Marca
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoDevolFecha]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoDevolFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoDevolFecha]
	@Desde char(8) ,
	@Hasta char(8)
 AS

SELECT * FROM PedidoDevol
WHERE
	@Desde <= Fechaord and @Hasta >= Fechaord
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoDevol]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoDevol]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoDevol]
	@Pedido  int
 AS

SELECT * FROM PedidoDevol
WHERE
	Pedido = @Pedido
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedidoCentro]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedidoCentro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedidoCentro] AS

SELECT Pedido, Fecha, Cliente, FecEntrega, TipoPed, Autorizo, Impresion, Cantidad, Facturado, Precio, Impresion3, Terminado

FROM Pedido

Where Proceso1 = 1 and Impresion = "N" 

Order by Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPedido]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPedido]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPedido]
	@Pedido  int
 AS

SELECT * FROM Pedido
WHERE
	Pedido = @Pedido
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPagosNumero]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagosNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPagosNumero] AS

SELECT  MAX(Orden), Orden

FROM Pagos

GROUP BY Orden

ORDER BY Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPagosMovban]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagosMovban]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaPagosMovban]
	@Fecha char(8),
	@DesdeBanco int,
	@HastaBanco int
		 AS

Select Fechaord,Tipo2,Banco2,Orden,Fecha,Fecha2,Fechaord2,Observaciones,Numero2,Importe2,Proveedor,TipoReg,Importe1 From Pagos

WHERE
	FechaOrd > @Fecha and Banco2 >= @DesdeBanco and Banco2 <= @HastaBanco

Order by Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPagosFecha]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagosFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaPagosFecha]
	@DesdeFecha char(8),
	@HastaFecha char(8)
		 AS

Select * From Pagos

WHERE
	FechaOrd >= @DesdeFecha and FechaOrd <= @HastaFecha 

Order by Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPagosConsultaII]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagosConsultaII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPagosConsultaII] AS

SELECT Numero2, Observaciones2, Importe2, Fecha, Fecha2, Orden, Proveedor, Observaciones FROM Pagos

Where  Tiporeg = "2" and Tipo2 = "02"

ORDER BY Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPagosConsulta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagosConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPagosConsulta] AS

SELECT Numero2, Observaciones2, Importe2, Fecha, Fecha2, Orden, Proveedor, Observaciones FROM Pagos

Where  Tiporeg = "2" and (Tipo2 = "3" Or Tipo2= "03")

ORDER BY Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPagosCarpetaTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagosCarpetaTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaPagosCarpetaTotal]
	
AS

Select Orden,Renglon,Carpeta,Carpeta1,Carpeta2,Carpeta3,Carpeta4,ImpoCarpeta,ImpoCarpeta1,ImpoCarpeta2,ImpoCarpeta3,ImpoCarpeta4, Importe, Fecha

From Pagos

Where
	Renglon = 1 or Renglon = "01" and (Carpeta <> 0 or Carpeta1 <> 0  or Carpeta2 <> 0 or Carpeta3 <> 0 Or Carpeta4 <> 0)' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPagosCarpeta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagosCarpeta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaPagosCarpeta]
	@Carpeta int
AS

Select *

From Pagos

Where
	Carpeta =  @Carpeta 

Order by clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPagos]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPagos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaPagos] AS

SELECT * FROM Pagos
ORDER BY Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaPago]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaPago]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaPago] AS

Select * From Pago
Order by Pago' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOtTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOtTotal] AS

SELECT * FROM Ot
Order by ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOtSolicitanteSolo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtSolicitanteSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOtSolicitanteSolo] 
	@Solicitante  char(50)
 AS

SELECT * FROM Ot

WHERE
	Solicitante = @Solicitante

Order by Solicitante,ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOtSolicitante]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtSolicitante]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOtSolicitante] AS

SELECT Solicitante FROM Ot
Order by Solicitante' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOtObservaciones1Solo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtObservaciones1Solo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOtObservaciones1Solo] 
	@Observaciones1 char(50)
 AS

SELECT * FROM Ot

WHERE
	Observaciones1 = @Observaciones1

Order by Observaciones1,ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOtObservaciones1]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtObservaciones1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOtObservaciones1] AS

SELECT Observaciones1 FROM Ot
Order by Observaciones1' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOtNumero]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOtNumero] AS

SELECT max(Codigo), Codigo
FROM Ot
group by Codigo
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOtFSalidaSolo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtFSalidaSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOtFSalidaSolo] 
	@FechaSalida  char(10)
 AS

SELECT * FROM Ot

WHERE
	FechaSalida = @FechaSalida

Order by ordfechaSalida,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOtFechaSolo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtFechaSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOtFechaSolo] 
	@Fecha  char(10)
 AS

SELECT * FROM Ot

WHERE
	Fecha = @Fecha

Order by ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOtFechaSalida]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtFechaSalida]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOtFechaSalida] AS

SELECT OrdFechaSalida,FechaSalida FROM Ot
Order by ordfechaSalida' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOtFecha]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOtFecha] AS

SELECT OrdFecha,Fecha FROM Ot
Order by ordfecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOtFComproSolo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtFComproSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOtFComproSolo] 
	@FechaCompro  char(10)
 AS

SELECT * FROM Ot

WHERE
	FechaCompro = @FechaCompro

Order by ordfechaCompro,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOtFCompro]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtFCompro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOtFCompro] AS

SELECT OrdFechaCompro,FechaCompro FROM Ot
Order by ordfechaCompro' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOtClienteSolo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtClienteSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOtClienteSolo] 
	@Razon  char(50)
 AS

SELECT * FROM Ot

WHERE
Razon = @Razon

Order by Cliente,ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOtCliente]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOtCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOtCliente] AS

SELECT Razon FROM Ot
Order by Razon' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenTotalDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenTotalDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOrdenTotalDesdeHasta]
             @Desde char(10)  ,
	@Hasta char(10)

 AS

SELECT Clave,Orden,Articulo,Cantidad,FechaOrd FROM Orden

WHERE
	@Desde <= Articulo and @Hasta >= Articulo

Order by Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOrdenTotal] AS

SELECT * FROM Orden
Order by Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenProveedor]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenProveedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOrdenProveedor]
	@Proveedor char(11)
 AS

SELECT * FROM Orden
WHERE
	Proveedor = @Proveedor
Order by Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenNumero]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOrdenNumero] AS

SELECT Max(Orden), Orden

FROM Orden

Group By Orden

Order by Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenImpresion]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenImpresion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOrdenImpresion]
	@Impresion char(1)
 AS

SELECT * FROM Orden
WHERE
	Impresion = @Impresion
Order by Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOrdenDesdeHasta]
	@Desde  int,
	@Hasta int
 AS

SELECT * FROM Orden
WHERE
	Orden >= @Desde and Orden <= @Hasta
	
Order by Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenCarpeta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenCarpeta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOrdenCarpeta]
	@Carpeta  int
 AS

SELECT * FROM Orden
WHERE
	Carpeta = @Carpeta
Order by Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOrdenArticulo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrdenArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOrdenArticulo]
	@Orden  int,
	@Articulo char(10)
 AS

SELECT * FROM Orden
WHERE
	Orden = @Orden and Articulo = @Articulo
	
Order by Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaOrden]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaOrden]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaOrden]
	@Orden  int
 AS

SELECT * FROM Orden
WHERE
	Orden = @Orden
Order by Orden,Renglon' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraTotal] AS

SELECT * FROM Muestra
Order by ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraProductoSolo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraProductoSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraProductoSolo] 
	@Producto  char(12)
 AS

SELECT * FROM Muestra

WHERE
	Producto = @Producto

Order by ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraObservacionesSolo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraObservacionesSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraObservacionesSolo] 
	@Observaciones char(50)
 AS

SELECT * FROM Muestra

WHERE
	Observaciones = @Observaciones

Order by ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraObservaciones]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraObservaciones]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraObservaciones] AS

SELECT Observaciones FROM Muestra
Order by Observaciones' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraNumero]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraNumero] AS

SELECT max(Codigo), Codigo
FROM Muestra
group by Codigo
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraNombreSolo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraNombreSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraNombreSolo] 
	@Nombre  char(50)
 AS

SELECT * FROM Muestra

WHERE
	Nombre = @Nombre

Order by ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraNombre]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraNombre]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraNombre] AS

SELECT Nombre FROM Muestra
Order by Nombre' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraFechaSolo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraFechaSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraFechaSolo] 
	@Fecha  char(10)
 AS

SELECT * FROM Muestra

WHERE
	Fecha = @Fecha

Order by ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraFecha]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraFecha] AS

SELECT OrdFecha,Fecha FROM Muestra
Order by ordfecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraEnsayoSolo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraEnsayoSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraEnsayoSolo] 
	@Ensayo  char(15)
 AS

SELECT * FROM Muestra

WHERE
	Ensayo = @Ensayo

Order by ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraDescriClienteSolo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraDescriClienteSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraDescriClienteSolo] 
	@DescriCliente char(50)
 AS

SELECT * FROM Muestra

WHERE
	DescriCliente = @DescriCliente

Order by ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraDescriCliente]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraDescriCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraDescriCliente] AS

SELECT DescriCliente FROM Muestra
Order by DescriCliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraClienteSolo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraClienteSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraClienteSolo] 
	@Razon  char(50)
 AS

SELECT * FROM Muestra

WHERE
	Razon = @Razon

Order by ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraCliente]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraCliente] AS

SELECT Razon FROM Muestra
Order by Razon' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraCantidadSolo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraCantidadSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraCantidadSolo] 
	@Cantidad  char(15)
 AS

SELECT * FROM Muestra

WHERE
	Cantidad = @Cantidad

Order by ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraCantidad]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraCantidad]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraCantidad] AS

SELECT Cantidad FROM Muestra
Order by Cantidad' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraArticuloSolo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraArticuloSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraArticuloSolo] 
	@Articulo  char(10)
 AS

SELECT * FROM Muestra

WHERE
	Articulo = @Articulo

Order by ordfecha,Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMuestraArticulo]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMuestraArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMuestraArticulo] AS

SELECT Articulo,Producto,Ensayo FROM Muestra
Order by Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovvarTotal] AS

SELECT * FROM Movvar
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarTerminadoDesdeHastaFecha]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarTerminadoDesdeHastaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovvarTerminadoDesdeHastaFecha]
	@Desde char(12),
	@Hasta char(12),
	@FechaOrd char(8)
		 AS

SELECT * FROM Movvar
WHERE
	Terminado >= @Desde and Terminado <= @Hasta and FechaOrd > @FechaOrd
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarTerminadoDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarTerminadoDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovvarTerminadoDesdeHasta]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT * FROM Movvar
WHERE
	Terminado >= @Desde and Terminado <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarRepro1]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarRepro1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovvarRepro1]
	@Desde char(10),
	@Hasta char(10)
		 AS

SELECT Articulo,Marca,Tipo,Movi,Cantidad,Fecha FROM Movvar
WHERE
	Articulo >= @Desde and Articulo <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarRepro]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarRepro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovvarRepro]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT Terminado,Marca,Tipo,Movi,Cantidad,Fecha FROM Movvar
WHERE
	Terminado >= @Desde and Terminado <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarNumero]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovvarNumero] AS

SELECT Max(Codigo), Codigo

FROM Movvar

Group By Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarArticuloDesdeHastaFecha]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarArticuloDesdeHastaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovvarArticuloDesdeHastaFecha]
	@Desde char(10),
	@Hasta char(10)  ,
	@FechaOrd char(8)
	
		 AS

SELECT Movi,Cantidad FROM Movvar
WHERE
	Articulo >= @Desde and Articulo <= @Hasta and FechaOrd > @FechaOrd
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvarArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvarArticuloDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovvarArticuloDesdeHasta]
	@Desde char(10),
	@Hasta char(10)
		 AS

SELECT * FROM Movvar
WHERE
	Articulo >= @Desde and Articulo <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovvar]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovvar]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovvar]
	@Codigo  int
 AS

SELECT * FROM Movvar
WHERE
	Codigo = @Codigo
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovlabTotal] AS

SELECT * FROM Movlab
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabTerminadoDesdeHastaFecha]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabTerminadoDesdeHastaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovlabTerminadoDesdeHastaFecha]
	@Desde char(12),
	@Hasta char(12)  ,
	@FechaOrd char(8)
		 AS

SELECT * FROM Movlab
WHERE
	Terminado >= @Desde and Terminado <= @Hasta and FechaOrd > @FechaOrd
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabTerminadoDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabTerminadoDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovlabTerminadoDesdeHasta]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT * FROM Movlab
WHERE
	Terminado >= @Desde and Terminado <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabRepro1]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabRepro1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovlabRepro1]
	@Desde char(10),
	@Hasta char(10)
		 AS

SELECT Articulo,Marca,Tipo,Movi,Cantidad,Fecha FROM Movlab
WHERE
	Articulo >= @Desde and Articulo <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabRepro]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabRepro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovlabRepro]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT Terminado,Marca,Tipo,Movi,Cantidad,Fecha FROM Movlab
WHERE
	Terminado >= @Desde and Terminado <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabNumero]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovlabNumero] AS

SELECT Max(Codigo), Codigo

FROM Movlab

Group By Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabArticuloDesdeHastaFecha]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabArticuloDesdeHastaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovlabArticuloDesdeHastaFecha]
	@Desde char(10),
	@Hasta char(10)  ,
	@FechaOrd char(8)

		 AS

SELECT Movi,Cantidad FROM Movlab
WHERE
	Articulo >= @Desde and Articulo <= @Hasta and FechaOrd > @FechaOrd
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlabArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlabArticuloDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovlabArticuloDesdeHasta]
	@Desde char(10),
	@Hasta char(10)
		 AS

SELECT * FROM Movlab
WHERE
	Articulo >= @Desde and Articulo <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovlab]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovlab]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovlab]
	@Codigo  int
 AS

SELECT * FROM Movlab
WHERE
	Codigo = @Codigo
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaTotal]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovguiaTotal] AS

SELECT * FROM Guia
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaTerminadoDesdeHastaFecha]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaTerminadoDesdeHastaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovguiaTerminadoDesdeHastaFecha]
	@Desde char(12),
	@Hasta char(12)  ,
	@FechaOrd char(8)
		 AS

SELECT * FROM Guia
WHERE
	Terminado >= @Desde and Terminado <= @Hasta and FechaOrd > @FechaOrd
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaTerminadoDesdeHasta]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaTerminadoDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovguiaTerminadoDesdeHasta]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT * FROM Guia
WHERE
	Terminado >= @Desde and Terminado <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaRepro1]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaRepro1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovguiaRepro1]
	@Desde char(10),
	@Hasta char(10)
		 AS

SELECT Articulo,Marca,Tipo,Movi,Cantidad,Fecha,Saldo FROM Guia
WHERE
	Articulo >= @Desde and Articulo <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaRepro]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaRepro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovguiaRepro]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT Terminado,Marca,Tipo,Movi,Cantidad,Fecha,Saldo FROM Guia
WHERE
	Terminado >= @Desde and Terminado <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaNumero]    Script Date: 05/01/2016 17:47:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovguiaNumero]
	@Tipomov int
 AS

SELECT Max(Codigo), Codigo

FROM Guia

Where
	Tipomov = @Tipomov

Group By Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaLoteSolo]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaLoteSolo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovguiaLoteSolo]
	@Lote int
		 AS

SELECT * FROM Guia
WHERE
	Lote = @Lote
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaLote1]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaLote1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovguiaLote1]
	@Terminado char(12),
	@Lote int
		 AS

SELECT * FROM Guia
WHERE
	Terminado = @Terminado and Lote = @Lote
Order by saldo desc, FechaOrd' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaLote]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaLote]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovguiaLote]
	@Articulo char(10),
	@Lote int
		 AS

SELECT * FROM Guia
WHERE
	Articulo = @Articulo and Lote = @Lote
Order by saldo desc, FechaOrd' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaArticuloDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaArticuloDesdeHastaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovguiaArticuloDesdeHastaFecha]
	@Desde char(10),
	@Hasta char(10)  ,
	@FechaOrd char(8)
		 AS

SELECT Cantidad,CantidadAnt,Movi FROM Guia
WHERE
	Articulo >= @Desde and Articulo <= @Hasta and FechaOrd > @FechaOrd
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguiaArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguiaArticuloDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovguiaArticuloDesdeHasta]
	@Desde char(10),
	@Hasta char(10)
		 AS

SELECT * FROM Guia
WHERE
	Articulo >= @Desde and Articulo <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovguia]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovguia]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovguia]
	@Tipomov  int,
	@Codigo int
 AS

SELECT * FROM Guia
WHERE
	Codigo = @Codigo and Tipomov  = @Tipomov
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovgasTotal]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovgasTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovgasTotal] AS

SELECT * FROM Movgas
Order by Carpeta' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovgas]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovgas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovgas]
	@Carpeta  int
 AS

SELECT * FROM Movgas
WHERE
	Carpeta = @Carpeta
Order by Carpeta' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovenvTotal]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovenvTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovenvTotal] AS

SELECT * FROM Movenv
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovenvDesdeHastaEnvases]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovenvDesdeHastaEnvases]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovenvDesdeHastaEnvases]
	@Desde int,
	@Hasta int
		 AS

SELECT * FROM Movenv
WHERE
	Envase >= @Desde and Envase <= @Hasta
Order by Envase, Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovenvDesdeHastaCliente]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovenvDesdeHastaCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovenvDesdeHastaCliente]
	@Desde char(6),
	@Hasta char(6)
		 AS

SELECT * FROM Movenv
WHERE
	Cliente >= @Desde and Cliente <= @Hasta
Order by Cliente,Envase' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMovenv]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMovenv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMovenv]
	@Codigo  int
 AS

SELECT * FROM Movenv
WHERE
	Codigo = @Codigo
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMarcasConsulta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMarcasConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMarcasConsulta]
	@Descripcion  char varying(50)
 AS

SELECT * FROM Marcas
WHERE
	Descripcion like("%"+@Descripcion+"%")
Order by Descripcion' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMarcasArticulo]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMarcasArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaMarcasArticulo]
	@Articulo char(10)
		 AS

Select * From Marcas

WHERE
	Articulo = @Articulo
Order by Proveedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaMarcas]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaMarcas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaMarcas]
AS

SELECT * FROM Marcas' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLineaMp]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLineaMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaLineaMp] AS

Select * From LineasMp
Order by Linea' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLinea]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLinea]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaLinea] AS

Select * From Lineas
Order by Linea' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoTotal]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaLaudoTotal] AS

SELECT * FROM Laudo
Order by Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoRepro]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoRepro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaLaudoRepro]
	@Desde char(10),
	@Hasta char(10)
		 AS

SELECT Articulo,Marca,Liberada,Saldo,Fecha FROM Laudo
WHERE
	Articulo >= @Desde and Articulo <= @Hasta

Order by Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoPartiOri]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoPartiOri]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaLaudoPartiOri]
	@PartiOri char(20)
 AS

SELECT * FROM Laudo
WHERE
	PartiOri = @PartiOri


Order by Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoOrden]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoOrden]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaLaudoOrden]
	@Orden  int
 AS

SELECT * FROM Laudo
WHERE
	Orden = @Orden
Order by Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoNumero]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaLaudoNumero] AS

SELECT Max(Laudo), Laudo

FROM Laudo

Group By Laudo

Order by Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoInforme]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoInforme]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaLaudoInforme]
	@Informe  int  ,
	@Articulo  char(10)
 AS

SELECT * FROM Laudo
WHERE
	Informe = @Informe and Articulo = @Articulo


Order by Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoFecha]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaLaudoFecha]
	@Fecha char(8)
		 AS

SELECT * FROM Laudo
WHERE
	Fechaord = null

Order by Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoDy]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaLaudoDy] AS

SELECT * FROM Laudo

Where
	Laudo > 949999 and Laudo < 989999

Order by Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoDevol]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoDevol]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaLaudoDevol] AS

SELECT * FROM Laudo

Where
	Laudo > 995000 and Laudo < 999999

Order by Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoArticuloPartiOri]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoArticuloPartiOri]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaLaudoArticuloPartiOri]
	@PartiOri char(20),
	@Articulo Char(10)
		 AS

SELECT * FROM Laudo
WHERE
	Articulo = @Articulo and PArtiOri = @PartiOri

Order by Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoArticuloDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoArticuloDesdeHastaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaLaudoArticuloDesdeHastaFecha]
	@Desde char(10),
	@Hasta char(10) ,
	@Fecha char(8)
		 AS

SELECT Liberada,LiberadaAnt FROM Laudo
WHERE
	Articulo >= @Desde and Articulo <= @Hasta and Fechaord > @Fecha

Order by Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoArticuloDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaLaudoArticuloDesdeHasta]
	@Desde char(10),
	@Hasta char(10)
		 AS

SELECT * FROM Laudo
WHERE
	Articulo >= @Desde and Articulo <= @Hasta

Order by Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudoArticulo]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudoArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaLaudoArticulo]
	@Laudo int,
	@Articulo Char(10)
		 AS

SELECT * FROM Laudo
WHERE
	Articulo = @Articulo and Laudo = @Laudo

Order by Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaLaudo]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaLaudo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaLaudo]
	@Laudo  int
 AS

SELECT * FROM Laudo
WHERE
	Laudo = @Laudo
Order by Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaIvacompTotal]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaIvacompTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaIvacompTotal] AS

SELECT * FROM Ivacomp
Order by NroInterno' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaIvaCompNumero]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaIvaCompNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaIvaCompNumero]
AS

SELECT Max(NroInterno), NroInterno

FROM IvaComp

Group By NroInterno

order by NroInterno' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaIvaCompMenor]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaIvaCompMenor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaIvaCompMenor]
	@NroInterno int
AS
SELECT * FROM IvaComp

where 
	Nrointerno < @NroInterno

order by NroInterno' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaIvaCompDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaIvaCompDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaIvaCompDesdeHasta]  
	@DesdeFecha char(10),
	@HastaFecha char(10)
		 AS

Select * From Ivacomp

WHERE
	OrdFecha >= @DesdeFecha and OrdFecha <= @HastaFecha 

Order by OrdFecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaIvaComp]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaIvaComp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaIvaComp]
AS
SELECT * FROM IvaComp' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaInventarioTotal]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInventarioTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaInventarioTotal] AS

SELECT * FROM Inventario
Order by Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaInventarioNumero]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInventarioNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaInventarioNumero] AS

SELECT Max(Numero), Numero

FROM Inventario

Group By Numero

Order by Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaInventario]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInventario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaInventario]
	@Numero  int
 AS

SELECT * FROM Inventario
WHERE
	Numero = @Numero
Order by Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaInsumos]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInsumos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaInsumos]
	@Solicitante  char varying(50)
 AS

SELECT * FROM Insumos
WHERE
	Renglon = 1 and Estado <> 2 And Estado <> 6 and  Solicitante like("%"+@Solicitante+"%")
Order by Solicitante' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeTotalDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeTotalDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaInformeTotalDesdeHasta]
             @Desde char(10)  ,
	@Hasta char(10)

 AS

SELECT FechaOrd,Clave,Informe,Articulo,Cantidad FROM Informe

WHERE
	@Desde <= Articulo and @Hasta >= Articulo

Order by Informe' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeTotal]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaInformeTotal] AS

SELECT * FROM Informe
Order by Informe' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeOrdenArticulo]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeOrdenArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaInformeOrdenArticulo]
	@Orden  int,
	@Articulo char(10)
 AS

SELECT * FROM Informe
WHERE
	Orden = @Orden and Articulo = @Articulo
Order by Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeOrden]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeOrden]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaInformeOrden]
	@Orden  int
 AS

SELECT * FROM Informe
WHERE
	Orden = @Orden
Order by Informe' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeNumero]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaInformeNumero] AS

SELECT Max(Informe), Informe

FROM Informe

Group By Informe

Order by Informe' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeListado]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeListado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaInformeListado]
	@DesdeFecha char(12),
	@HastaFecha char(12)  ,
	@DesdeProv char(11)  ,
	@HastaProv char(11)

		 AS

SELECT * FROM Informe
WHERE
	FechaOrd >= @DesdeFecha and FechaOrd <= @HastaFecha and 	Proveedor >= @DesdeProv and Proveedor <= @HastaProv
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeDesdeHastaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaInformeDesdeHastaFecha]
	@Desde char(8)  ,
	@Hasta char(8)
 AS

SELECT * FROM Informe
WHERE
	@Desde <= FechaOrd and @Hasta >= FechaOrd
Order by Informe' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaInformeArticulo]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInformeArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaInformeArticulo]
	@Articulo char(10)
 AS

SELECT * FROM Informe
WHERE
	Articulo = @Articulo
Order by Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaInforme]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaInforme]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaInforme]
	@Informe  int
 AS

SELECT * FROM Informe
WHERE
	Informe = @Informe
Order by Informe' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaImputacDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaImputacDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaImputacDesdeHasta]  
	@DesdeFecha char(10),
	@HastaFecha char(10),
	@DesdeCuenta  char(10),
	@HastaCuenta char(10)
		 AS

Select * From Imputac

WHERE
	Fechaord >= @DesdeFecha and Fechaord <= @HastaFecha  and Cuenta >= @DesdeCuenta and Cuenta <= @HastaCuenta
Order by Cuenta' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaImputac]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaImputac]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaImputac]
AS
SELECT * FROM Imputac' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaTotal]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHojaTotal] AS

SELECT * FROM Hoja
Order by Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaTerminadoDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaTerminadoDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHojaTerminadoDesdeHasta]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT * FROM Hoja
WHERE
	Terminado >= @Desde and Terminado <= @Hasta
Order by Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaReProceso]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaReProceso]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHojaReProceso]
		 AS

SELECT * FROM Hoja
WHERE
	Real = 0 and  Renglon = 1 and Marca <> "X" and Teorico <> 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaRepro2]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaRepro2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHojaRepro2]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT Producto,Marca,Renglon,Real,Saldo,Fecha FROM Hoja
WHERE
	Producto >= @Desde and Producto <= @Hasta
Order by Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaRepro1]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaRepro1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHojaRepro1]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT hoja,Terminado,Marca,Tipo,Cantidad,Saldo,Fecha FROM Hoja
WHERE
	Terminado >= @Desde and Terminado <= @Hasta
Order by Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaRepro]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaRepro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHojaRepro]
	@Desde char(10),
	@Hasta char(10)
		 AS

SELECT Articulo,Marca,Tipo,Cantidad,Saldo,Fecha,Clave FROM Hoja
WHERE
	Articulo >= @Desde and Articulo <= @Hasta
Order by Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaProductoDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaProductoDesdeHastaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHojaProductoDesdeHastaFecha]
	@Desde char(12),
	@Hasta char(12)  ,
	@FechaOrd char(8)
		 AS

SELECT * FROM Hoja
WHERE
	Producto >= @Desde and Producto <= @Hasta and FechaOrd > @FechaOrd
Order by Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaProductoDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaProductoDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHojaProductoDesdeHasta]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT * FROM Hoja
WHERE
	Producto >= @Desde and Producto <= @Hasta
Order by Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaProducto]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaProducto]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHojaProducto]
	@Hoja int,
	@Producto char(12)
		 AS

SELECT * FROM Hoja
WHERE
	Producto = @Producto and Hoja = @Hoja and Renglon = 1
Order by Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaNumero]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHojaNumero]
	
 AS

SELECT MAX(Hoja), hoja

FROM Hoja

group by hoja




order by hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaDesdeHastaFefcha]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaDesdeHastaFefcha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHojaDesdeHastaFefcha]
	@Desde char(12),
	@Hasta char(12),
	@FechaOrd char(8)
		 AS

SELECT Terminado,Marca,Tipo,Cantidad,Saldo,Fecha FROM Hoja
WHERE
	Terminado >= @Desde and Terminado <= @Hasta and FechaOrd > @FechaOrd
Order by Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaDesdeHastaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHojaDesdeHastaFecha]
	@Desde char(12),
	@Hasta char(12),
	@FechaOrd char(8)
		 AS

SELECT Terminado,Marca,Tipo,Cantidad,Saldo,Fecha FROM Hoja
WHERE
	Terminado >= @Desde and Terminado <= @Hasta and FechaOrd > @FechaOrd
Order by Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaArticuloDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaArticuloDesdeHastaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHojaArticuloDesdeHastaFecha]
	@Desde char(10),
	@Hasta char(10),
	@FechaOrd char(8)
		 AS

SELECT Cantidad,Canti1,Canti2,Canti3 FROM Hoja
WHERE
	Articulo >= @Desde and Articulo <= @Hasta and FechaOrd > @FechaOrd
Order by Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHojaArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHojaArticuloDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHojaArticuloDesdeHasta]
	@Desde char(10),
	@Hasta char(10)
		 AS

SELECT * FROM Hoja
WHERE
	Articulo >= @Desde and Articulo <= @Hasta
Order by Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaHoja]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaHoja]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaHoja]
	@Hoja  int
 AS

SELECT * FROM Hoja
WHERE
	Hoja = @Hoja
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaGasimpo]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaGasimpo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaGasimpo] AS

Select * From Gasimpo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaFeriado]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaFeriado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaFeriado] AS

SELECT * FROM Feriado
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaReproDy]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaReproDy]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaEstadisticaReproDy]
	@Desde char(10),
	@Hasta char(10)
		 AS

SELECT Marca,Tipo,Lote1,Canti1,Lote2,Canti2,Lote3,Canti3,Lote4,Canti4,Lote5,Canti5,LoteAdicional FROM Estadistica
WHERE
	ArticuloDy >= @Desde and ArticuloDy <= @Hasta and TipoProDy = "M"


Order by ArticuloDy' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaRepro]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaRepro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaEstadisticaRepro]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT Articulo,Clave,Marca,Tipo,Cantidad,Fecha FROM Estadistica
WHERE
	Articulo >= @Desde and Articulo <= @Hasta


Order by Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaFecha]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaEstadisticaFecha]
	@DesdeFec char(8),
	@HastaFec char(8)
		 AS

SELECT * FROM Estadistica
WHERE
	OrdFecha >= @DesdeFec and OrdFecha <= @HastaFec 

Order by OrdFecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaDesdeHastaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaEstadisticaDesdeHastaFecha]
	@Desde char(12),
	@Hasta char(12) ,
	@OrdFecha char(8)

		 AS

SELECT * FROM Estadistica
WHERE
	Articulo >= @Desde and Articulo <= @Hasta and OrdFecha > @OrdFecha


Order by Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaEstadisticaDesdeHasta]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT * FROM Estadistica
WHERE
	Articulo >= @Desde and Articulo <= @Hasta


Order by Articulo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaCliente]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaEstadisticaCliente]
	@Desde char(6),
	@Hasta char(6)
		 AS

SELECT * FROM Estadistica
WHERE
	Cliente >= @Desde and Cliente <= @Hasta

Order by Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaArticuloDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaArticuloDesdeHastaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaEstadisticaArticuloDesdeHastaFecha]
	@Desde char(10),
	@Hasta char(10)  ,
	@FechaOrd char(8)
		 AS

SELECT Tipo,Cantidad FROM Estadistica
WHERE
	ArticuloDy >= @Desde and ArticuloDy <= @Hasta and OrdFecha > @FechaOrd


Order by ArticuloDy' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEstadisticaArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEstadisticaArticuloDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaEstadisticaArticuloDesdeHasta]
	@Desde char(10),
	@Hasta char(10)
		 AS

SELECT * FROM Estadistica
WHERE
	ArticuloDy >= @Desde and ArticuloDy <= @Hasta


Order by ArticuloDy' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEspecificaciones]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEspecificaciones]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaEspecificaciones] AS

Select * From Especificaciones
Order by Producto' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEspecif]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEspecif]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaEspecif] AS

Select * From Especif
Order by Producto' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEnvases]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEnvases]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaEnvases] AS

Select * From Envases
Order by Envases' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEntdevTotal]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEntdevTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaEntdevTotal] AS

SELECT * FROM Entdev
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEntdevTerminadoDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEntdevTerminadoDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaEntdevTerminadoDesdeHasta]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT * FROM Entdev
WHERE
	Terminado >= @Desde and Terminado <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEntdevRepro]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEntdevRepro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaEntdevRepro]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT Terminado,Marca,Cantidad FROM Entdev
WHERE
	Terminado >= @Desde and Terminado <= @Hasta
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEntdevNumero]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEntdevNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaEntdevNumero] AS

SELECT Max(Codigo), Codigo

FROM Entdev

Group By Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEntdev]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEntdev]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaEntdev]
	@Codigo  int
 AS

SELECT * FROM Entdev

WHERE
	Codigo = @Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaEnsayos]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaEnsayos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaEnsayos] AS

Select * From Ensayos
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaDevconNumero]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaDevconNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaDevconNumero] AS

SELECT max(Numero), Numero
FROM Devcon
group by Numero
Order by Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaDevconCliente]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaDevconCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaDevconCliente]
	@Desde char(6),
	@Hasta char(6)
		 AS

SELECT * FROM Devcon
WHERE
	Cliente >= @Desde and Cliente <= @Hasta


Order by Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaDevcon]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaDevcon]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaDevcon]
	@Numero  int
 AS

SELECT * FROM Devcon
WHERE
	Numero = @Numero
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaDepositosMovban]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaDepositosMovban]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaDepositosMovban]
	@Fecha char(8),
	@DesdeBanco int,
	@HastaBanco int ,
	@Renglon char(2)
		 AS

Select Fechaord,Banco,Renglon,Fecha,Acredita,AcreditaOrd,Deposito,Importe From Depositos

WHERE
	FechaOrd > @Fecha  and Banco >= @DesdeBanco and Banco <= @HastaBanco and Renglon = @Renglon

Order by Deposito' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaDepositosFecha]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaDepositosFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaDepositosFecha]
	@DesdeFecha char(8),
	@HastaFecha char(8)
		 AS

Select * From Depositos

WHERE
	FechaOrd >= @DesdeFecha and FechaOrd <= @HastaFecha 

Order by Deposito' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaDepositosConsulta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaDepositosConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaDepositosConsulta] AS

SELECT Numero2, Observaciones2, Importe2, Fecha, Fecha2, Deposito, Banco FROM Depositos

Where Tipo2 = "3" Or Tipo2= "03"

Order By Deposito' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaDepositos]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaDepositos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaDepositos] AS

SELECT * FROM Depositos

Order By Deposito' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCuentas]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCuentas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaCuentas]  AS

Select * From Cuenta
Order by Cuenta' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCtaprvDesdeHastaSaldoTotal]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtaprvDesdeHastaSaldoTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaCtaprvDesdeHastaSaldoTotal]
	@Desde char(11),
	@Hasta char(11)
		 AS

Select * From CtaCteprv

WHERE
	Proveedor >= @Desde and Proveedor <= @Hasta and Saldo <> 0 


ORDER BY PROVEEDOR' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCtaprvDesdeHastaSaldo]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtaprvDesdeHastaSaldo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaCtaprvDesdeHastaSaldo]
	@Desde char(11),
	@Hasta char(11)
		 AS

Select Proveedor,Saldo,Clave From CtaCteprv

WHERE
	Proveedor >= @Desde and Proveedor <= @Hasta and Saldo <> 0 


ORDER BY PROVEEDOR' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCtaprvDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtaprvDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaCtaprvDesdeHasta]  
	@Desde char(11),
	@Hasta char(11)
		 AS

Select * From CtaCteprv

WHERE
	Proveedor >= @Desde and Proveedor <= @Hasta


ORDER BY PROVEEDOR, ORDFECHA' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCtaCtePrv]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtaCtePrv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaCtaCtePrv]
AS
SELECT * FROM CtaCtePrv

ORDER BY PROVEEDOR, ORDFECHA' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCtacteFecha]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtacteFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaCtacteFecha]
	@Desde char(8),
	@Hasta char(8)
		 AS

Select * From Ctacte

WHERE
	OrdFecha >= @Desde and OrdFecha <= @Hasta

Order by Ordfecha, Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCtacteDesdeHastaFecha]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtacteDesdeHastaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaCtacteDesdeHastaFecha]
	@Desde char(8),
	@Hasta char(8)
		 AS

Select * From Ctacte

WHERE
	OrdFecha >= @Desde and OrdFecha <= @Hasta

Order by Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCtacteDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtacteDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaCtacteDesdeHasta]  
	@Desde char(10),
	@Hasta char(10)
		 AS

Select * From Ctacte

WHERE
	Cliente >= @Desde and Cliente <= @Hasta

Order by Cliente,  Ordfecha, Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCtacteCliente]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtacteCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaCtacteCliente]
	@Cliente char(10)
		 AS

Select * From Ctacte

WHERE
	Cliente = @Cliente

Order by Ordfecha,  Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCtaCte]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCtaCte]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE procedure [dbo].[ListaCtaCte]
AS
SELECT * FROM CtaCte


Order by Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCotizaTotal]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCotizaTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaCotizaTotal]
	@Cotiza  int
 AS

SELECT * FROM Cotiza
WHERE
	Cotiza = @Cotiza

Order by Cotiza' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCotizaProveedorDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCotizaProveedorDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaCotizaProveedorDesdeHasta]
	@Desde char(11),
	@Hasta char(11)
		 AS

SELECT * FROM Cotiza
WHERE
	Proveedor >= @Desde and Proveedor <= @Hasta


Order by Proveedor,Articulo,Fechaord' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCotizaProveedor]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCotizaProveedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaCotizaProveedor]
	@Proveedor char(11)
		 AS

SELECT * FROM Cotiza
WHERE
	Proveedor = @Proveedor

Order by Proveedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCotizaNumero]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCotizaNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaCotizaNumero]

 AS

SELECT MAX(Cotiza), Cotiza

FROM Cotiza

Group By Cotiza

Order by Cotiza' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCotizaArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCotizaArticuloDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaCotizaArticuloDesdeHasta]
	@Desde char(10),
	@Hasta char(10)
		 AS

SELECT * FROM Cotiza
WHERE
	Articulo >= @Desde and Articulo <= @Hasta


Order by Articulo,Proveedor,Fechaord' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCotiza]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCotiza]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaCotiza] AS

SELECT * FROM Cotiza
Order by Cotiza' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaConsigTotal]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaConsigTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaConsigTotal] AS

SELECT * FROM Consig
Order by Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaConsigTerminado]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaConsigTerminado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaConsigTerminado]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT * FROM Consig
WHERE
	Terminado >= @Desde and Terminado <= @Hasta


Order by Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaConsigRepro]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaConsigRepro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaConsigRepro]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT Terminado,Marca,Facturado,Cantidad,Fecha FROM Consig
WHERE
	Terminado >= @Desde and Terminado <= @Hasta


Order by Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaConsigNumero]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaConsigNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaConsigNumero] AS

SELECT max(Numero), Numero
FROM Consig
group by Numero
Order by Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaConsigFactura]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaConsigFactura]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaConsigFactura]
	@Numero Int,
	@Terminado char(12)
		 AS

SELECT * FROM Consig
WHERE
	Numero = @Numero and Terminado = @Terminado

Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaConsigCliente]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaConsigCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaConsigCliente]
	@Desde char(6),
	@Hasta char(6)
		 AS

SELECT * FROM Consig
WHERE
	Cliente >= @Desde and Cliente <= @Hasta


Order by Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaConsig]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaConsig]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaConsig]
	@Numero  int
 AS

SELECT * FROM Consig
WHERE
	Numero = @Numero
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaComposicionDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaComposicionDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaComposicionDesdeHasta]
	@Desde char(12),
	@Hasta char(12)
		 AS

SELECT * FROM Composicion
WHERE
	Terminado >= @Desde and Terminado <= @Hasta


Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaComposicion]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaComposicion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaComposicion] AS

Select * From Composicion
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaClientes]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaClientes]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE procedure [dbo].[ListaClientes]
AS
SELECT * FROM Cliente
order by Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaClienteConsulta1]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaClienteConsulta1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaClienteConsulta1] AS

Select Cliente,Razon From Cliente

Order by Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaClienteConsulta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaClienteConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaClienteConsulta] AS

Select Cliente,Razon From Cliente

Order by Razon' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCliente]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaCliente] AS

Select * From Cliente

Order by Razon' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCambioAdm]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCambioAdm]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaCambioAdm] AS

Select * From CambioAdm
Order by ORDFecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaCambio]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaCambio]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaCambio] AS

Select * From Cambios
Order by Fecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaBancos]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaBancos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ListaBancos]
AS

SELECT * FROM Banco' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaArticuloStock]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaArticuloStock]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaArticuloStock] AS

Select Codigo, Descripcion, Entradas, Salidas, Costo2, Costo3

From Articulo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaArticuloDesdeHastaMinimo]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaArticuloDesdeHastaMinimo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaArticuloDesdeHastaMinimo]
	@Desde char(10),
	@Hasta char(10)
		 AS

Select Codigo,Descripcion,Inicial,Salidas,Entradas,Minimo,Minimo1 From Articulo

WHERE
	Codigo >= @Desde and Codigo <= @Hasta

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaArticuloDesdeHasta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaArticuloDesdeHasta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaArticuloDesdeHasta]  
	@Desde char(10),
	@Hasta char(10)
		 AS

Select * From Articulo

WHERE
	Codigo >= @Desde and Codigo <= @Hasta

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaArticuloConsulta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaArticuloConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaArticuloConsulta] AS

Select Codigo, Descripcion

From Articulo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ListaArticulo]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ListaArticulo] AS

Select * From Articulo
Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[DepuraMovEnv]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DepuraMovEnv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[DepuraMovEnv]
	@FechaOrd char(8)
 AS

DELETE	MovEnv
WHERE
		FechaOrd < @FechaOrd' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaVendedor]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaVendedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaVendedor]
	@Vendedor  int
 AS

SELECT * FROM Vendedor
WHERE
	Vendedor = @Vendedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaUltimoPago]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaUltimoPago]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaUltimoPago] AS

SELECT isnull(MAX(Orden),0) Orden FROM Pagos' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaUltimoNroRecibo]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaUltimoNroRecibo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaUltimoNroRecibo]
AS
SELECT	IsNull(Max(Recibo),0) Ultimo
FROM	Recibos' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaUltimoNroInterno]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaUltimoNroInterno]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaUltimoNroInterno] 
AS
SELECT	IsNull(Max(NroInterno),0) Ultimo
FROM	IvaComp' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaUltimoDeposito]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaUltimoDeposito]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaUltimoDeposito] AS

SELECT isnull(MAX(Deposito),0) Deposito FROM Depositos' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaTerminado]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaTerminado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaTerminado]
	@Codigo  char(12)
 AS

SELECT * FROM Terminado
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaSolicitud1]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaSolicitud1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaSolicitud1]
	@Solicitud int
 AS

SELECT * FROM Solic
WHERE
	Solicitud = @Solicitud' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaSolicitud]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaSolicitud]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaSolicitud]
	@Clave Char(8)
 AS

SELECT * FROM Solic
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaRubro]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaRubro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaRubro]
	@Rubro  int
 AS

SELECT * FROM Rubros
WHERE
	Rubro = @Rubro' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaRetencion]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaRetencion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaRetencion]
@Clave Varchar(20)
 AS
SELECT * FROM Retencion
Where
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaRecibosClave]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaRecibosClave]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'Create Procedure [dbo].[ConsultaRecibosClave]
	@Clave varchar(8)
AS
SELECT * FROM Recibos
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaRecibos_x_Recibo]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaRecibos_x_Recibo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'create procedure [dbo].[ConsultaRecibos_x_Recibo]
	@Recibo int
AS
SELECT 	*
FROM
	Recibos
WHERE
	Recibo = @Recibo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaRecibos]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaRecibos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'Create Procedure [dbo].[ConsultaRecibos]
	@Recibo varchar(6)
AS
SELECT * FROM Recibos
WHERE
	Recibo = @Recibo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPrueterMenor]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPrueterMenor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPrueterMenor]
	@Prueba char(7)
 AS

SELECT * FROM Prueter
WHERE
	Prueba < @Prueba

Order by Prueba' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPrueter]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPrueter]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPrueter]
	@Prueba Char(7)
 AS

SELECT * FROM Prueter
WHERE
	Prueba = @Prueba' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPrueart]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPrueart]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPrueart]
	@Prueba Char(7)
 AS

SELECT * FROM Prueart
WHERE
	Prueba = @Prueba' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaProvincia]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaProvincia]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaProvincia]
	@Provincia int
 AS

SELECT * FROM Provincia
WHERE
	Provincia = @Provincia' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaProveedoresSiguiente]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaProveedoresSiguiente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaProveedoresSiguiente]
	@Proveedor  char(11)
 AS

SELECT * FROM Proveedor
WHERE
	Proveedor > @Proveedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaProveedores]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaProveedores]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaProveedores]
	@Proveedor  char(11)
 AS

SELECT * FROM Proveedor
WHERE
	Proveedor = @Proveedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPrestamo]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPrestamo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPrestamo]
	@Codigo  int
 AS

SELECT * FROM Prestamo
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPresCon]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPresCon]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPresCon]
	@Codigo  int
 AS

SELECT * FROM Prescon
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPreciosMp]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPreciosMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPreciosMp]
	@Clave char(16)
 AS

SELECT * FROM PreciosMp
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPrecios]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPrecios]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPrecios]
	@Clave char(18)
 AS

SELECT * FROM Precios
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedidoFactura]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedidoFactura]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPedidoFactura]
	@Pedido int,
	@Terminado char(12)
 AS

SELECT * FROM Pedido
WHERE
	Pedido = @Pedido and Terminado = @Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedidoDevolFactura]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedidoDevolFactura]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPedidoDevolFactura]
	@Pedido int,
	@Terminado char(12)
 AS

SELECT * FROM PedidoDevol
WHERE
	Pedido = @Pedido and Terminado = @Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedidoDevol2]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedidoDevol2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPedidoDevol2]
	@Pedido Int  ,
	@Renglon int
 AS

SELECT * FROM PedidoDevol
WHERE
	Pedido = @Pedido and Renglon = @Renglon
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedidoDevol1]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedidoDevol1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPedidoDevol1]
	@Pedido Int
 AS

SELECT * FROM PedidoDevol
WHERE
	Pedido = @Pedido
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedidoDevol]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedidoDevol]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPedidoDevol]
	@Clave Char(8)
 AS

SELECT * FROM PedidoDevol
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedido2]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedido2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPedido2]
	@Pedido Int  ,
	@Renglon int
 AS

SELECT * FROM Pedido
WHERE
	Pedido = @Pedido and Renglon = @Renglon
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedido1]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedido1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPedido1]
	@Pedido Int
 AS

SELECT * FROM Pedido
WHERE
	Pedido = @Pedido
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPedido]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPedido]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPedido]
	@Clave Char(8)
 AS

SELECT * FROM Pedido
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPagos]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPagos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPagos]
@Clave char(8)
 AS
SELECT * FROM Pagos
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaPago]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaPago]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaPago]
	@Pago  int
 AS

SELECT * FROM Pago
WHERE
	Pago = @Pago' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaOt]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaOt]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaOt]
	@Codigo int
AS

SELECT 	* From Ot

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaOrdenCarpeta]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaOrdenCarpeta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaOrdenCarpeta]
	@Carpeta  int
 AS

SELECT * FROM Orden
WHERE
	Carpeta = @Carpeta' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaOrden]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaOrden]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaOrden]
	@Clave Char(8)
 AS

SELECT * FROM Orden
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaOperadorClave]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaOperadorClave]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaOperadorClave]
	@Clave char(10)
 AS

SELECT * FROM Operador
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaOperador]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaOperador]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaOperador]
	@Operador  int
 AS

SELECT * FROM Operador
WHERE
	Operador = @Operador' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaNumero]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaNumero]
	@Numero Char(2)
 AS

SELECT * FROM Numero
WHERE
	Codigo = @Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMuestra]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMuestra]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaMuestra]
	@Codigo int
AS

SELECT 	* From Muestra

WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMovvar]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMovvar]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaMovvar]
	@Clave Char(8)
 AS

SELECT * FROM Movvar
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMovlab]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMovlab]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaMovlab]
	@Clave Char(8)
 AS

SELECT * FROM Movlab
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMovguia]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMovguia]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaMovguia]
	@Clave Char(9)
 AS

SELECT * FROM Guia
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMovgas]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMovgas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaMovgas]
	@Clave Char(8)
 AS

SELECT * FROM Movgas
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMovenv]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMovenv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaMovenv]
	@Clave Char(9)
 AS

SELECT * FROM Movenv
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[Consultamono]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Consultamono]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'

CREATE PROCEDURE [dbo].[Consultamono]
	@Codigo  char(12)
 AS

SELECT * FROM codigomono
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMinimo]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMinimo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaMinimo]
	@Codigo  char(12)
 AS

SELECT * FROM Minimo
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaMarcas]    Script Date: 05/01/2016 17:47:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaMarcas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaMarcas]
	@Clave char(21)
AS

SELECT 	* From Marcas

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaLineaMp]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaLineaMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaLineaMp]
	@Linea  int
 AS

SELECT * FROM LineasMp
WHERE
	Linea = @Linea' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaLinea]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaLinea]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaLinea]
	@Linea  int
 AS

SELECT * FROM Lineas
WHERE
	Linea = @Linea' 
END
GO
/****** Object:  StoredProcedure [dbo].[Consultalehman]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Consultalehman]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[Consultalehman]
	@Codigo  char(12)
 AS

SELECT * FROM etilehman
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaLaudo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaLaudo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaLaudo]
	@Clave Char(8)
 AS

SELECT * FROM Laudo
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaIvaCompras]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaIvaCompras]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaIvaCompras]
	@NroInterno  int
AS
SELECT * FROM IvaComp
WHERE
	NroInterno = @NroInterno' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaIvaCompCompro]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaIvaCompCompro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaIvaCompCompro]
	@Proveedor char(11),
	@Tipo char(2)  ,
	@Punto char(4) ,
	@Numero char(8)
AS
SELECT * FROM IvaComp
WHERE
	Proveedor = @Proveedor and Tipo = @Tipo  and 	Punto = @Punto  and Numero = @Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaIvaComp]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaIvaComp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaIvaComp]
@NroInterno int
AS
SELECT * FROM IvaComp
WHERE
	NroInterno = @NroInterno' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaInventario]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaInventario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaInventario]
	@Clave Char(8)
 AS

SELECT * FROM Inventario
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaInformeOrden]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaInformeOrden]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaInformeOrden]
	@Orden Int
 AS

SELECT * FROM Informe
WHERE
	Orden = @Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaInforme]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaInforme]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaInforme]
	@Clave Char(8)
 AS

SELECT * FROM Informe
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaImputac]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaImputac]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaImputac]
@Clave char(24)
AS
SELECT * FROM Imputac
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaHojaEspecial]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaHojaEspecial]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaHojaEspecial]
	@Producto char(12)  
 AS

SELECT * FROM Hoja
WHERE
	Producto = @Producto

order by producto, fechaingord' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaHoja]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaHoja]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaHoja]
	@Clave Char(8)
 AS

SELECT * FROM Hoja
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaGasimpo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaGasimpo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaGasimpo]
	@Codigo  int
 AS

SELECT * FROM Gasimpo
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEstadistica1]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEstadistica1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaEstadistica1]
	@Tipo int,
	@Numero int
 AS
 
SELECT * FROM Estadistica
WHERE
	Tipo = @Tipo and Numero = @Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEstadistica]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEstadistica]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaEstadistica]
	@Clave Char(12)
 AS

SELECT * FROM Estadistica
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEspeCli]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEspeCli]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaEspeCli]
	@Cliente char(6) ,
	@Terminado char(12)

 AS

SELECT * FROM EspeCli
WHERE
	Cliente = @Cliente and Terminado = @Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEspecificaciones]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEspecificaciones]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaEspecificaciones]
	@Producto char(10)
 AS

SELECT * FROM Especificaciones
WHERE
	Producto = @Producto' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEspecif]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEspecif]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaEspecif]
	@Producto char(12)
 AS

SELECT * FROM Especif
WHERE
	Producto = @Producto' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEnvases]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEnvases]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaEnvases]
	@Envases  int
 AS

SELECT * FROM Envases
WHERE
	Envases = @Envases' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEntdev2]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEntdev2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaEntdev2]
	@Cliente char(6) ,
	@Lote int ,
	@Terminado char(12) 
	AS

SELECT * FROM Entdev
WHERE
	Cliente = @Cliente  and Lote = @Lote and Terminado = @Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEntdev1]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEntdev1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaEntdev1]
	@Codigo int ,
	@Terminado char(12) 
	AS

SELECT * FROM Entdev
WHERE
	Codigo = @Codigo and Terminado = @Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEntdev]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEntdev]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaEntdev]
	@Clave Char(8)
 AS

SELECT * FROM Entdev
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaEnsayos]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaEnsayos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaEnsayos]
	@Codigo  int
 AS

SELECT * FROM Ensayos
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaDesccompTotal]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaDesccompTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaDesccompTotal]
	
 AS

SELECT * FROM Desccomp' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaDesccomp1]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaDesccomp1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaDesccomp1]
	@Tipo char(2),
	@Numero char(8)
 AS

SELECT * FROM Desccomp
WHERE
	Tipo = @Tipo and Numero = @Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaDesccomp]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaDesccomp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaDesccomp]
	@Clave Char(12)
 AS

SELECT * FROM Desccomp
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaDepositosClave]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaDepositosClave]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaDepositosClave]
	@Clave Char(8)
 AS

SELECT * FROM Depositos
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaDepositos]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaDepositos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaDepositos]
	@Deposito Char(6)
 AS

SELECT * FROM Depositos
WHERE
	Deposito = @Deposito' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCuentas]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCuentas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaCuentas]
	@Cuenta char(10)	
 AS

SELECT * FROM Cuenta
WHERE
	Cuenta = @Cuenta' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCtaPrv]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCtaPrv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ConsultaCtaPrv]
	@Clave varchar(26)
AS
SELECT * FROM CtaCtePrv
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCtaCtePrv_x_Proveedor]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCtaCtePrv_x_Proveedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'Create Procedure [dbo].[ConsultaCtaCtePrv_x_Proveedor]
	@Proveedor varchar(11)
AS
SELECT	* 
FROM
	CtaCtePrv
WHERE
	Proveedor = @Proveedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCtaCtePrv]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCtaCtePrv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ConsultaCtaCtePrv]
	@Clave varchar(26)
AS
SELECT * FROM CtaCtePrv
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCtaCteComp]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCtaCteComp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'Create Procedure [dbo].[ConsultaCtaCteComp]
	@Clave varchar(12)
AS
SELECT * FROM CtaCte
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCtaCteCli]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCtaCteCli]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'Create Procedure [dbo].[ConsultaCtaCteCli]
	@Cliente varchar(12)
AS
SELECT * FROM CtaCte
WHERE
	Cliente = @Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCtacte]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCtacte]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaCtacte]
	@Clave Char(12)
 AS

SELECT * FROM Ctacte
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCotiza]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCotiza]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaCotiza]
	@Clave Char(8)
 AS

SELECT * FROM Cotiza
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaConsigArti]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaConsigArti]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaConsigArti]
	@Numero int,
	@Terminado char(12)
 AS

SELECT * FROM Consig
WHERE
	Numero = @Numero and Terminado = @Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaConsig1]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaConsig1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaConsig1]
	@Numero Int
 AS

SELECT * FROM Consig
WHERE
	Numero = @Numero
Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaConsig]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaConsig]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaConsig]
	@Clave Char(8)
 AS

SELECT * FROM Consig
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaComposicionProducto]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaComposicionProducto]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaComposicionProducto]
	@Terminado char(12)

 AS

SELECT * FROM Composicion
WHERE
	Terminado = @Terminado 






ORDER BY CLAVE' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaComposicion]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaComposicion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaComposicion]
	@Clave  char(14)
 AS

SELECT * FROM Composicion
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaClientes]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaClientes]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaClientes]
	@Cliente  char(12)
 AS

SELECT * FROM Cliente
WHERE
	Cliente = @Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaClienteRazon]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaClienteRazon]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaClienteRazon]
	@Cliente Char(6)
 AS

SELECT Cliente,Razon FROM Cliente
WHERE
	Cliente = @Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCliente]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaCliente]
	@Cliente Char(6)
 AS

SELECT * FROM Cliente
WHERE
	Cliente = @Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCambioOrdFecha]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCambioOrdFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaCambioOrdFecha]
	@OrdFecha char(8)
 AS

SELECT * FROM Cambios
WHERE
	OrdFecha <= @OrdFecha
ORDER BY 	
	OrdFecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCambioAdmOrdFecha]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCambioAdmOrdFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaCambioAdmOrdFecha]
	@OrdFecha char(8)
 AS

SELECT * FROM CambioAdm
WHERE
	OrdFecha <= @OrdFecha
ORDER BY 	
	OrdFecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCambioAdm]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCambioAdm]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaCambioAdm]
	@Fecha char(10)
 AS

SELECT * FROM CambioAdm
WHERE
	Fecha = @Fecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaCambio]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaCambio]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaCambio]
	@Fecha char(10)
 AS

SELECT * FROM Cambios
WHERE
	Fecha = @Fecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaBancos]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaBancos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaBancos]
	@Banco Int
AS

SELECT * FROM Banco Where Banco = @Banco' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaBanco]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaBanco]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaBanco]
@Banco int
AS
SELECT 	* From Banco
WHERE
	Banco = @Banco' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaAtributo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaAtributo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaAtributo]
	@Operador  int ,
	@Proceso int
 AS

SELECT * FROM Atributos
WHERE
	Operador = @Operador and Proceso = @Proceso' 
END
GO
/****** Object:  StoredProcedure [dbo].[ConsultaArticulo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ConsultaArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ConsultaArticulo]
	@Codigo  char(10)
 AS

SELECT * FROM Articulo
WHERE
	Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[CalculaDiferenciaTerminado]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CalculaDiferenciaTerminado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[CalculaDiferenciaTerminado]

 AS

UPDATE  Terminado
	SET
		Dife 	   	= Minimo-Inicial-Entradas+Salidas' 
END
GO
/****** Object:  StoredProcedure [dbo].[CalculaDiferenciaArticulo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CalculaDiferenciaArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[CalculaDiferenciaArticulo]

 AS

UPDATE  Articulo
	SET
		Dife 	   	= Minimo-Inicial-Entradas+Salidas' 
END
GO
/****** Object:  StoredProcedure [dbo].[BuscarImputacion]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BuscarImputacion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BuscarImputacion]
@Clave Varchar(10)
AS

SELECT	* 
FROM 
	Imputac 
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarVendedor]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarVendedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarVendedor]
	@Vendedor  int
 AS

DELETE	Vendedor
WHERE
		Vendedor = @Vendedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarTerminado]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarTerminado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarTerminado]
	@Codigo	char(12)
 AS

DELETE	Terminado
WHERE
		Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarSoltot]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarSoltot]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarSoltot]

 AS

DELETE	Soltot
WHERE
		Baja = 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarSolicitudTotal]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarSolicitudTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarSolicitudTotal]
	@Solicitud int
 AS

DELETE	Solic
WHERE
		Solicitud = @Solicitud' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarSOlicitud]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarSOlicitud]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarSOlicitud]
	@Clave	char(8)
 AS

DELETE	Solic
WHERE
		Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarSolHoja]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarSolHoja]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarSolHoja]
	@Hoja int
 AS

DELETE	SolHoja
WHERE
		Hoja = @Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarSolGuiaTotal]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarSolGuiaTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarSolGuiaTotal]

 AS

DELETE	SolGuiaTotal' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarSolGuia]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarSolGuia]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarSolGuia]
	@Codigo int
 AS

DELETE	SolGuia

WHERE
		Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarSedronar]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarSedronar]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarSedronar]

 AS

DELETE	Sedronar' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarRubro]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarRubro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarRubro]
	@Rubro   int
 AS

DELETE	Rubros
WHERE
		Rubro = @Rubro' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarProveedor]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarProveedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarProveedor]
	@Proveedor   char(11)
 AS

DELETE	Proveedor
WHERE
		Proveedor = @Proveedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarPrestamo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPrestamo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarPrestamo]
	@Codigo int
 AS

DELETE	Prestamo

WHERE
		Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarPresCon]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPresCon]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarPresCon]
 AS

DELETE	PresCon' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarPreciosTotal]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPreciosTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarPreciosTotal]
 AS

DELETE	Precios' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarPreciosMpTotal]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPreciosMpTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarPreciosMpTotal]
 AS

DELETE	PreciosMp' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarPreciosMp]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPreciosMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarPreciosMp]
	@Clave  char(16)
 AS

DELETE	PreciosMp
WHERE
		Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarPrecios]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPrecios]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarPrecios]
	@Clave  char(18)
 AS

DELETE	Precios
WHERE
		Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarPedidoDevol]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPedidoDevol]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarPedidoDevol]
	@Pedido int
 AS

DELETE	PedidoDevol
WHERE
		Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarPedido]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPedido]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarPedido]
	@Pedido int
 AS

DELETE	Pedido
WHERE
		Pedido = @Pedido' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarPago]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarPago]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarPago]
	@Pago   int
 AS

DELETE	Pago
WHERE
		Pago = @Pago' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarOt]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarOt]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarOt]
	@Codigo int
 AS

DELETE	Ot
WHERE
		Codigo  = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarOrdenTotal]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarOrdenTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarOrdenTotal]
	@Orden int
 AS

DELETE	Orden
WHERE
		Orden = @Orden' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarOrden]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarOrden]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarOrden]
	@Clave	char(8)
 AS

DELETE	Orden
WHERE
		Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarMuestraTotal]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMuestraTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarMuestraTotal]
 AS

DELETE	Muestra' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarMuestraImpre]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMuestraImpre]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarMuestraImpre]
 AS

DELETE	MuestraImpre' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarMuestra]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMuestra]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarMuestra]
	@Codigo int
 AS

DELETE	Muestra
WHERE
		Codigo  = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarMovvar]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMovvar]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarMovvar]
	@Codigo int
 AS

DELETE	Movvar

WHERE
		Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarMovlab]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMovlab]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarMovlab]
	@Codigo int
 AS

DELETE	Movlab

WHERE
		Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarMovguia]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMovguia]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarMovguia]
	@Tipomov int,
	@Codigo int
 AS

DELETE	Guia

WHERE
		Codigo = @Codigo and Tipomov = @Tipomov' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarMovGasCon]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMovGasCon]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarMovGasCon]
 AS

DELETE	MovGasCon' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarMovgas]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMovgas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarMovgas]
	@Carpeta int
 AS

DELETE	Movgas
WHERE
		Carpeta = @Carpeta' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarMovenv]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMovenv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarMovenv]
	@Codigo int
 AS

DELETE	Movenv
WHERE
		Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarMinimo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMinimo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarMinimo]
 AS

DELETE	Minimo' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarMarcas]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarMarcas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarMarcas]
	@Clave char(21)
 AS

DELETE	Marcas
WHERE
		Clave  = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarLineaMp]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarLineaMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarLineaMp]
	@Linea   int
 AS

DELETE	LineasMp
WHERE
		Linea = @Linea' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarLinea]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarLinea]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarLinea]
	@Linea   int
 AS

DELETE	Lineas
WHERE
		Linea = @Linea' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarLaudoFecha]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarLaudoFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarLaudoFecha]
	@Fecha char(10)

 AS

DELETE	Laudo
WHERE
		Fecha= @Fecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarLaudo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarLaudo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarLaudo]
	@Laudo int

 AS

DELETE	Laudo
WHERE
		Laudo= @Laudo' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarIvacomp]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarIvacomp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarIvacomp]
	@NroInterno int
 AS

DELETE	Ivacomp
WHERE
		NroInterno = @NroInterno' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarInventarioTotal]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarInventarioTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarInventarioTotal]

 AS

DELETE	Inventario' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarInventario]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarInventario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarInventario]
	@Numero int
 AS

DELETE	Inventario

WHERE
		Numero = @Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarInforme]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarInforme]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarInforme]
	@Informe int
 AS

DELETE	Informe
WHERE
		Informe = @Informe' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarImputacion]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarImputacion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarImputacion]
@Clave Varchar(24)
AS

DELETE	Imputac 
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarImputac]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarImputac]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarImputac]
	@NroInterno int
 AS

DELETE	Imputac
WHERE
		NroInterno = @NroInterno' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarHojaFecha]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarHojaFecha]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarHojaFecha]
	@Fecha char(10)
 AS

DELETE	Hoja
WHERE
		Fecha = @Fecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarHoja]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarHoja]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarHoja]
	@Hoja int
 AS

DELETE	Hoja
WHERE
		Hoja = @Hoja' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarGasimpo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarGasimpo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarGasimpo]
	@Codigo  int
 AS

DELETE	Gasimpo
WHERE
		Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarEstadistica]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarEstadistica]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarEstadistica]
	@Clave	char(12)
 AS

DELETE	Estadistica
WHERE
		Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarEspecificaciones]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarEspecificaciones]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarEspecificaciones]
	@Producto   char(10)
 AS

DELETE	Especificaciones
WHERE
		Producto = @Producto' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarEspecif]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarEspecif]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarEspecif]
	@Producto   char(12)
 AS

DELETE	Especif
WHERE
		Producto = @Producto' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarEnvases]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarEnvases]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarEnvases]
	@Envases   int
 AS

DELETE	Envases
WHERE
		Envases = @Envases' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarEntdev]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarEntdev]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarEntdev]
	@Codigo int
 AS

DELETE	Entdev

WHERE
		Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarEnsayos]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarEnsayos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarEnsayos]
	@Codigo   int
 AS

DELETE	Ensayos
WHERE
		Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarDevcon]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarDevcon]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarDevcon]
	@Numero int
 AS

DELETE	Devcon
WHERE
		Numero = @Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarDesccomp]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarDesccomp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarDesccomp]
	@Clave	char(12)
 AS

DELETE	Desccomp
WHERE
		Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarCuenta]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCuenta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarCuenta]
	@Cuenta   char(11)
 AS

DELETE	Cuenta
WHERE
		Cuenta = @Cuenta' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarCtaprv]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCtaprv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarCtaprv]
	@Clave char(26)
 AS

DELETE	Ctacteprv 
WHERE
		Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarCtacteNumero]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCtacteNumero]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarCtacteNumero]
 AS

DELETE	Ctacte
WHERE
		Numero = 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarCtacte]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCtacte]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarCtacte]
	@Clave	char(12)
 AS

DELETE	Ctacte
WHERE
		Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarCotizaTotal]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCotizaTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarCotizaTotal]
	@Cotiza int

 AS

DELETE	Cotiza
WHERE
		Cotiza = @Cotiza' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarCotiza]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCotiza]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarCotiza]
	@Clave	char(8)
 AS

DELETE	Cotiza
WHERE
		Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarConsig]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarConsig]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarConsig]
	@Numero int
 AS

DELETE	Consig
WHERE
		Numero = @Numero' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarComposicion]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarComposicion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarComposicion]
	@Terminado char(14)
 AS

DELETE	Composicion
WHERE
		Terminado = @Terminado' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarCliente]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarCliente]
	@Cliente   char(6)
 AS

DELETE	CLiente
WHERE
		Cliente = @Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarCarpeta]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCarpeta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarCarpeta]
as
DELETE	Carpeta' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarCambioAdm]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCambioAdm]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarCambioAdm]
	@Fecha  char(10)
 AS

DELETE	CambioAdm
WHERE
		Fecha = @Fecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarCambio]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarCambio]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarCambio]
	@Fecha  char(10)
 AS

DELETE	Cambios
WHERE
		Fecha = @Fecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarBanco]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarBanco]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarBanco]
	@Banco   smallint
 AS

DELETE	Banco
WHERE
		Banco = @Banco' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarAtributos]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarAtributos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarAtributos]
	@Operador  int,
	@Proceso int
 AS

DELETE	Atributos
WHERE
		Operador = @Operador and Proceso = @Proceso' 
END
GO
/****** Object:  StoredProcedure [dbo].[BorrarArticulo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BorrarArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[BorrarArticulo]
	@Codigo	char(10)
 AS

DELETE	Articulo
WHERE
		Codigo = @Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorVendedor]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorVendedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorVendedor]
	@Vendedor  int
 AS

SELECT * FROM Vendedor
WHERE
	Vendedor < @Vendedor

Order by Vendedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorTerminado]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorTerminado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorTerminado]
	@Codigo	char(12)
 AS

SELECT * FROM Terminado
WHERE
	Codigo < @Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorRubro]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorRubro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorRubro]
	@Rubro  int
 AS

SELECT * FROM Rubros
WHERE
	Rubro < @Rubro

Order by Rubro' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorProveedor]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorProveedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorProveedor]
	@Proveedor char(11)
 AS

SELECT * FROM Proveedor
WHERE
	Proveedor < @Proveedor

Order by Proveedor' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorPreciosMp]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorPreciosMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorPreciosMp]
	@Clave char(16)
 AS

SELECT clave, Cliente, Articulo
FROM PreciosMp
WHERE
	Clave < @Clave

Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorPrecios]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorPrecios]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorPrecios]
	@Clave char(18)
 AS

SELECT clave, Cliente, terminado
FROM Precios
WHERE
	Clave < @Clave

Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorPago]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorPago]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorPago]
	@Pago  int
 AS

SELECT * FROM Pago
WHERE
	Pago < @Pago

Order by Pago' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorMuestra]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorMuestra]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorMuestra]
	@Codigo int
 AS

SELECT * FROM Muestra
WHERE
	Codigo < @Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorLineaMp]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorLineaMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorLineaMp]
	@Linea  int
 AS

SELECT * FROM LineasMp
WHERE
	Linea < @Linea

Order by Linea' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorLinea]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorLinea]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorLinea]
	@Linea  int
 AS

SELECT * FROM Lineas
WHERE
	Linea < @Linea

Order by Linea' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorGasimpo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorGasimpo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorGasimpo]
	@Codigo  int
 AS

SELECT * FROM Gasimpo
WHERE
	Codigo < @Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorEspecificaciones]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorEspecificaciones]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorEspecificaciones]
	@Producto char(10)
 AS

SELECT * FROM Especificaciones
WHERE
	Producto < @Producto

Order by Producto' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorEspecif]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorEspecif]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorEspecif]
	@Producto char(12)
 AS

SELECT * FROM Especif
WHERE
	Producto < @Producto

Order by Producto' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorEnvases]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorEnvases]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorEnvases]
	@Envases  int
 AS

SELECT * FROM Envases
WHERE
	Envases < @Envases

Order by Envases' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorEnsayos]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorEnsayos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorEnsayos]
	@Codigo  int
 AS

SELECT * FROM Ensayos
WHERE
	Codigo < @Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorCuenta]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorCuenta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorCuenta]
	@Cuenta char(20)
 AS

SELECT * FROM Cuenta
WHERE
	Cuenta < @Cuenta

Order by Cuenta' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorComposicion]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorComposicion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorComposicion]
	@Clave char(14)
 AS

SELECT * FROM Composicion
WHERE
	Clave < @Clave

Order by Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorCliente]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorCliente]
	@Cliente char(6)
 AS

SELECT * FROM Cliente
WHERE
	Cliente < @Cliente

Order by Cliente' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorCambioAdm]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorCambioAdm]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorCambioAdm]
	@Fecha  char(10)

 AS

SELECT * FROM CambioAdm
WHERE
	Fecha < @Fecha

Order by Fecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorCambio]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorCambio]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorCambio]
	@Fecha  char(10)

 AS

SELECT * FROM Cambios
WHERE
	Fecha < @Fecha

Order by Fecha' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorBanco]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorBanco]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorBanco]
	@Banco smallint
 AS

SELECT * FROM Banco
WHERE
	Banco < @Banco

Order by Banco' 
END
GO
/****** Object:  StoredProcedure [dbo].[AnteriorArticulo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AnteriorArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AnteriorArticulo]
	@Codigo	char(10)
 AS

SELECT * FROM Articulo
WHERE
	Codigo < @Codigo

Order by Codigo' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaVendedor]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaVendedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaVendedor]
             @Vendedor int,
	@Nombre char(50)


 AS
	INSERT INTO  Vendedor
			(
			Vendedor   ,
		             Nombre
			)

VALUES
			(
		             @Vendedor,
			@Nombre
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaTerminado]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaTerminado]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaTerminado]
             @Codigo char(12),
	@Descripcion char(30),
	@Linea int,
	@Unidad char(10),
	@Inicial float,
	@Entradas float,
	@Salidas float,
	@MInimo float,
	@Deposito char(10),
	@Pedido float,
	@Envase1 int,
	@Envase2 int,
	@Envase3 int,
	@Envase4 int,
	@Envase5 int,
	@Envase6 int,
	@Proceso float,
	@Costo Float,
	@Factor float,
	@WDate char(10),
	@Impreadi char(1),
	@Clase char(30),
	@Intervencion char(10),
	@Naciones char(10),
	@Embalaje char(10),
	@Version int,
	@FechaVersion char(10),
	@Controla int,
	@Observaciones char(50)  ,
	@Tipoeti char(20)  ,
	@Escrito   int



 AS
	INSERT INTO  Terminado
			(
			Codigo   ,
		             Descripcion        ,                                     
			Linea      ,                                    
			Unidad      ,                                    
			Inicial      ,                                    
			Entradas      ,                                    
			Salidas      ,                                    
			Minimo      ,                                    
			Deposito      ,                                    
			Pedido      ,                                    
			Envase1      ,                                    
			Envase2      ,                                    
			Envase3      ,                                    
			Envase4      ,                                    
			Envase5      ,                                    
			Envase6      ,                                    
			Proceso      ,                                    
			Costo      ,                                    
			Factor      ,                                    
			WDate    ,
			Impreadi      ,                                    
			Clase      ,                                    
			Intervencion      ,                                    
			Naciones      ,                                    
			Embalaje      ,                                    
			Version      ,                                    
			FechaVersion      ,
			Controla,
			Observaciones  ,
			Tipoeti   ,
			Escrito
			)

VALUES
			(
			@Codigo   ,
		             @Descripcion        ,                                     
			@Linea      ,                                    
			@Unidad      ,                                    
			@Inicial      ,                                    
			@Entradas      ,                                    
			@Salidas      ,                                    
			@Minimo      ,                                    
			@Deposito      ,                                    
			@Pedido      ,                                    
			@Envase1      ,                                    
			@Envase2      ,                                    
			@Envase3      ,                                    
			@Envase4      ,                                    
			@Envase5      ,                                    
			@Envase6      ,                                    
			@Proceso      ,                                    
			@Costo      ,                                    
			@Factor      ,                                    
			@WDate   , 
			@Impreadi      ,                                    
			@Clase      ,                                    
			@Intervencion      ,                                    
			@Naciones      ,                                    
			@Embalaje      ,                                    
			@Version      ,                                    
			@FechaVersion     ,
			@Controla    ,
			@Observaciones   ,
			@Tipoeti   ,
			@Escrito
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaSoltot]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaSoltot]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaSoltot]
             @Clave char(8),
	@Solicitud int,
	@Renglon int,
	@Fecha char(10),
	@Fechaord char(8), 
	@Observaciones char(100), 
	@Articulo char(10),
	@Cantidad float,
	@Entrega char(10),
	@OrdEntrega char(8),
	@Planta char(30),
	@Solicitante char(30),
	@WDate char(10)  ,
	@Marca  char(1),
	@Obser char(50)  ,
	@Entregado float  ,
	@Baja float
	

 AS
	INSERT INTO  Soltot
			(
		             Clave ,
			Solicitud,
			Renglon ,
			Fecha ,
			Fechaord ,
			Observaciones ,
			Articulo ,
			Cantidad ,
			Entrega ,
			OrdEntrega ,
			Planta ,
			Solicitante  ,
			WDate  ,
			Marca  ,
			Obser  ,
			Entregado  ,
			Baja
			)

VALUES
			(
		             @Clave ,
			@Solicitud,
			@Renglon ,
			@Fecha ,
			@Fechaord ,
			@Observaciones ,
			@Articulo ,
			@Cantidad ,
			@Entrega ,
			@OrdEntrega ,
			@Planta ,
			@Solicitante  ,
			@WDate  ,
			@Marca  ,
			@Obser   ,
			@Entregado  ,
			@Baja
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaSolicitud]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaSolicitud]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaSolicitud]
             @Clave char(8),
	@Solicitud int,
	@Renglon int,
	@Fecha char(10),
	@Fechaord char(8), 
	@Observaciones char(100), 
	@Articulo char(10),
	@Cantidad float,
	@Entrega char(10),
	@OrdEntrega char(8),
	@Planta char(30),
	@Solicitante char(30),
	@WDate char(10)  ,
	@Marca  char(1),
	@Obser char(50)  ,
	@Entregado float
	

 AS
	INSERT INTO  Solic
			(
		             Clave ,
			Solicitud,
			Renglon ,
			Fecha ,
			Fechaord ,
			Observaciones ,
			Articulo ,
			Cantidad ,
			Entrega ,
			OrdEntrega ,
			Planta ,
			Solicitante  ,
			WDate  ,
			Marca  ,
			Obser  ,
			Entregado
			)

VALUES
			(
		             @Clave ,
			@Solicitud,
			@Renglon ,
			@Fecha ,
			@Fechaord ,
			@Observaciones ,
			@Articulo ,
			@Cantidad ,
			@Entrega ,
			@OrdEntrega ,
			@Planta ,
			@Solicitante  ,
			@WDate  ,
			@Marca  ,
			@Obser   ,
			@Entregado
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaSolHoja]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaSolHoja]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaSolHoja]
             @Clave char(8),
	@Hoja int,
	@Renglon int,
	@Fecha char(10),
	@Producto char(12),
	@Cantidad float,
	@Tipo char(1),
	@Articulo char(10),
	@Terminado char(12),
	@Teorico float,
	@Lote1 int  ,
	@Canti1 float  ,
	@Lote2 int   ,
	@Canti2 float  ,
	@Lote3 int  ,
	@Canti3 float  ,
	@Marca char(1)

 AS
	INSERT INTO  SolHoja
			(
		             Clave ,
			Hoja ,
			Renglon ,
			Fecha ,
			Producto ,
			Cantidad ,
			Tipo,
			Articulo,
			Terminado ,
			Teorico ,
			Lote1  ,
			Canti1  ,
			Lote2  ,
			Canti2  ,
			Lote3  ,
			Canti3  ,
			Marca
			)

VALUES
			(
		             @Clave ,
			@Hoja ,
			@Renglon ,
			@Fecha ,
			@Producto ,
			@Cantidad ,
			@Tipo,
			@Articulo,
			@Terminado ,
			@Teorico ,
			@Lote1  ,
			@Canti1  ,
			@Lote2  ,
			@Canti2  ,
			@Lote3  ,
			@Canti3 ,
			@Marca
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaSolguiaTotal]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaSolguiaTotal]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaSolguiaTotal]
             @Empresa int,
             @Clave Char(8),
	@Codigo int,
	@Fecha char(10),
	@OrdFecha char(8),
	@Tipo char(1),
	@Articulo char(10),
	@Terminado char(12),
	@Cantidad float,
	@Desde int  ,
	@Hasta int,
	@Observaciones char(50),
	@Usuario int  ,
	@DescriArticulo char(50)  ,
	@DescriTerminado char(50)

 AS
	INSERT INTO  SolGuiaTotal
			(
		             Empresa ,
		             Clave ,
			Codigo ,
			Fecha ,
			OrdFecha ,
			Tipo ,
			Articulo ,
			Terminado ,
			Cantidad ,  
			Desde   ,
			Hasta ,
			Observaciones  ,
			Usuario  ,
			DescriArticulo  ,
			DescriTerminado
			)

VALUES
			(
			@Empresa ,
			@Clave ,
			@Codigo ,
			@Fecha ,
			@OrdFecha ,
			@Tipo ,
			@Articulo ,
			@Terminado ,
			@Cantidad ,  
			@Desde ,
			@Hasta ,
			@Observaciones  ,
			@Usuario  ,
			@DescriArticulo  ,
			@DescriTerminado
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaSolguia]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaSolguia]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaSolguia]
             @Clave char(8),
	@Codigo int,
	@Renglon int,
	@Fecha char(10),
	@Tipo char(1),
	@Articulo char(10),
	@Terminado char(12),
	@Cantidad float,
	@Fechaord char(8),
	@Observaciones char(50),
	@Desde int ,
	@Hasta int ,
	@Marca char(1) ,
	@Usuario int  ,
	@Aviso int


 AS
	INSERT INTO  SolGuia
			(
		             Clave ,
			Codigo ,
			Renglon ,
			Fecha ,
			Tipo ,
			Articulo ,
			Terminado ,
			Cantidad ,
			Fechaord ,
			Observaciones ,
			Desde ,
			Hasta ,
			Marca  ,
			Usuario  ,
			Aviso
			)

VALUES
			(
		             @Clave ,
			@Codigo ,
			@Renglon ,
			@Fecha ,
			@Tipo ,
			@Articulo ,
			@Terminado ,
			@Cantidad ,
			@Fechaord ,
			@Observaciones ,
			@Desde  ,
			@Hasta ,
			@Marca   ,
			@Usuario   ,
			@Aviso
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaSedronar]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaSedronar]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaSedronar]
             @Articulo char(10)  ,
	@Renglon int,
	@Inicial Float,
	@Comprada Float  ,
	@Final Float ,
	@Periodo  char(15) ,
	@Ano char(4)  ,
	@NroInsc   char(15)
	
 AS
	INSERT INTO  Sedronar
			(
		             Articulo ,
			Renglon,
			Inicial ,
			Comprada  , 
			Final ,
			Periodo ,
			Ano  ,
			NroInsc  
			)

VALUES
			(
		             @Articulo ,
			@Renglon,
			@Inicial ,
			@Comprada ,
			@Final ,
			@Periodo ,
			@Ano  ,
			@NroInsc
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaRubro]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaRubro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaRubro]
             @Rubro int,
	@Nombre char(50)


 AS
	INSERT INTO  Rubros
			(
			Rubro   ,
		             Nombre
			)

VALUES
			(
		             @Rubro,
			@Nombre
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaRetencion]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaRetencion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaRetencion]
	@Clave     varchar(15),
	@Fecha 	   varchar(8),
	@Proveedor varchar(11),  
	@Neto      float,                                            
	@Retenido  float,                                            
	@Anticipo  float,                                            
	@Bruto     float,                                            
	@Iva       float
AS      
INSERT INTO
		Retencion
		(
		Clave , 
		Fecha ,
		Proveedor ,
		Neto      ,                                         
		Retenido  ,                                        
		Anticipo  ,                                            
		Bruto     ,                                            
		Iva        
		)
VALUES
		(
		@Clave	   ,
		@Fecha 	   ,
		@Proveedor ,  
		@Neto      ,                                            
		@Retenido  ,                                            
		@Anticipo  ,                                            
		@Bruto     ,                                            
		@Iva       
		)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaRecibos]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaRecibos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE procedure [dbo].[AltaRecibos]
	@Clave   	varchar(8), 
	@Recibo  	varchar(6),
	@Renglon 	varchar(2),
	@Cliente 	varchar(6),
	@Fecha      	varchar(10),
	@Fechaord 	varchar(8),
	@TipoRec 	varchar(1),
	@RetGanancias   FLOAT,          
	@RetIva         FLOAT,          
	@RetOtra        FLOAT,          
	@Retencion      FLOAT,          
	@TipoReg        varchar(1),
	@Tipo1 	varchar(2),
	@Letra1 	varchar(1),
	@Punto1 	varchar(4),
	@Numero1  	varchar(8),
	@Importe1       FLOAT,          
	@Tipo2 	varchar(2),
	@Numero2  	varchar(8),
	@Fecha2     	varchar(10),
	@banco2         varchar(20),
	@Importe2       FLOAT,          
	@Estado2 	varchar(1),
	@Empresa 	smallint,
	@FechaOrd2 	varchar(8),
	@Importe        float,                            
	@Observaciones  varchar(50),                                    
	@Impolist       float,                           
	@Impo1list      float,                                       
	@Destino        varchar(50),                                    
	@Cuenta     	varchar(10),
	@Marca  char(1),
	@FechaDepo Char(10) ,
	@FechaDepoOrd char(8)



AS	
INSERT INTO
		Recibos
		(
		Clave    ,
		Recibo ,
		Renglon ,
		Cliente ,
		Fecha    ,  
		Fechaord ,
		TipoRec ,
		RetGanancias ,
		RetIva       ,            
		RetOtra      ,            
		Retencion    ,            
		TipoReg, 
		Tipo1 ,
		Letra1 ,
		Punto1 ,
		Numero1 , 
		Importe1 ,                
		Tipo2 ,
		Numero2,  
		Fecha2  ,   
		banco2   ,            
		Importe2  ,               
		Estado2 ,
		Empresa ,
		FechaOrd2, 
		Importe   ,                                            
		Observaciones ,                                     
		Impolist       ,                                       
		Impo1list      ,                                       
		Destino        ,                                    
		Cuenta    ,
		Marca  ,
		FechaDepo  ,
		FechaDepoOrd
		)
VALUES
		(
		@Clave,    
		@Recibo, 
		@Renglon, 
		@Cliente ,
		@Fecha   ,   
		@Fechaord, 
		@TipoRec ,
		@RetGanancias,             
		@RetIva      ,             
		@RetOtra     ,            
		@Retencion   ,             
		@TipoReg ,
		@Tipo1 ,
		@Letra1 ,
		@Punto1 ,
		@Numero1 , 
		@Importe1 ,                
		@Tipo2 ,
		@Numero2,  
		@Fecha2  ,   
		@banco2   ,            
		@Importe2  ,               
		@Estado2 ,
		@Empresa ,
		@FechaOrd2, 
		@Importe   ,                                            
		@Observaciones,                                      
		@Impolist      ,                                        
		@Impo1list     ,                                        
		@Destino       ,                                     
		@Cuenta    ,
		@Marca  ,
		@FechaDepo  ,
		@FechaDepoOrd 	
		)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaPrueter]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPrueter]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaPrueter]
             @Prueba char(7),
	@Producto char(12),
	@Fecha char(10),
	@Valor1 char(50),
	@Valor2 char(50),
	@Valor3 char(50),
	@Valor4 char(50),
	@Valor5 char(50),
	@Valor6 char(50),
	@Valor7 char(50),
	@Valor8 char(50),
	@Valor9 char(50),
	@Valor10 char(50),
	@Ensayo char(50),
	@Aspecto char(50),
	@Observaciones char(50),
	@Confecciono char(50),
	@LIberada float,
	@Lote int,
	@Rechazo int,
	@Fechaord char(8),
	@WDate char(10)

 AS
	INSERT INTO  Prueter
			(
		             Prueba ,
			Producto ,
			Fecha ,
			Valor1,
			Valor2 ,
			Valor3 ,
			Valor4 ,
			Valor5 ,
			Valor6 ,
			Valor7 ,
			Valor8 ,
			Valor9 ,
			Valor10 ,
			Ensayo,
			Aspecto,
			Observaciones,
			Confecciono ,
			LIberada ,
			Lote ,
			Rechazo ,
			Fechaord ,
			WDate
			)

VALUES
			(
		             @Prueba ,
			@Producto ,
			@Fecha ,
			@Valor1,
			@Valor2 ,
			@Valor3 ,
			@Valor4 ,
			@Valor5 ,
			@Valor6 ,
			@Valor7 ,
			@Valor8 ,
			@Valor9 ,
			@Valor10 ,
			@Ensayo,
			@Aspecto,
			@Observaciones,
			@Confecciono ,
			@LIberada ,
			@Lote ,
			@Rechazo ,
			@Fechaord ,
			@WDate
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaPrueart]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPrueart]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaPrueart]
             @Prueba char(7),
	@Producto char(10),
	@Fecha char(10),
	@Orden char(6),
	@Valor1 char(50),
	@Valor2 char(50),
	@Valor3 char(50),
	@Valor4 char(50),
	@Valor5 char(50),
	@Valor6 char(50),
	@Valor7 char(50),
	@Valor8 char(50),
	@Valor9 char(50),
	@Valor10 char(50),
	@Ensayo char(50),
	@Aspecto char(50),
	@Observaciones char(50),
	@Observa2 char(50),
	@Confecciono char(50),
	@LIberada float,
	@Devuelta float,
	@Lote int,
	@Rechazo int,
	@Nueva char(1),
	@Fechaord char(8),
	@WDate char(10)

 AS
	INSERT INTO  Prueart
			(
		             Prueba ,
			Producto ,
			Fecha ,
			Orden ,
			Valor1,
			Valor2 ,
			Valor3 ,
			Valor4 ,
			Valor5 ,
			Valor6 ,
			Valor7 ,
			Valor8 ,
			Valor9 ,
			Valor10 ,
			Ensayo,
			Aspecto,
			Observaciones,
			Observa2 ,
			Confecciono ,
			LIberada ,
			Devuelta ,
			Lote ,
			Rechazo ,
			Nueva ,
			Fechaord ,
			WDate
			)

VALUES
			(
		             @Prueba ,
			@Producto ,
			@Fecha ,
			@Orden ,
			@Valor1,
			@Valor2 ,
			@Valor3 ,
			@Valor4 ,
			@Valor5 ,
			@Valor6 ,
			@Valor7 ,
			@Valor8 ,
			@Valor9 ,
			@Valor10 ,
			@Ensayo,
			@Aspecto,
			@Observaciones,
			@Observa2 ,
			@Confecciono ,
			@LIberada ,
			@Devuelta ,
			@Lote ,
			@Rechazo ,
			@Nueva ,
			@Fechaord ,
			@WDate
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaProveedor1]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaProveedor1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaProveedor1]
             @Proveedor char(11),
	@Nombre char(50),
	@Direccion char(50),
	@Localidad char(50),
	@Provincia char(2),
	@Postal char(4),
	@Cuit char(15),
	@Telefono    char(30),
	@EMail char(200),
	@Observaciones char(50),
	@Tipo char(1),
	@Iva char(1),
	@Dias char(20),
	@Empresa smallint,
	@Cuenta char(10),
	@Importe1 float,
	@Importe2 float,
	@Importe3 float,
	@Importe4 float,
	@Importe5 float,
	@Importe6 float,
	@NombreCheque char(50)  ,
	@WDate   char(10)    ,
 	@CodIb   int     ,
	@NroIb    char(20)  ,
	@NroInsc    char(15)


 AS
	INSERT INTO  Proveedor
			(
			Proveedor   ,
			Nombre        ,                                     
			Direccion      ,                                    
			Localidad      ,                                    
			Provincia ,
			Postal ,
			Cuit     ,       
			Telefono                       ,
			Email                          ,
			Observaciones            ,                          
			Tipo ,
			Iva  ,
			Dias ,                
			Empresa     ,
			Cuenta     ,
			Importe1    ,
			Importe2    ,
			Importe3    ,
			Importe4    ,
			Importe5    ,
			Importe6    ,
			NombreCheque ,
			WDate    ,
			CodIb   ,
			NroIb  ,
			NroInsc
			)
VALUES
			(
		             @Proveedor,
			@Nombre,
			@Direccion,
			@Localidad,
			@Provincia,
			@Postal,
			@Cuit,
			@Telefono,
			@EMail,
			@Observaciones,
			@Tipo,
			@Iva,
			@Dias,
			@Empresa,
			@Cuenta,
			@Importe1,
			@Importe2,
			@Importe3,
			@Importe4,
			@Importe5,
			@Importe6,
			@NombreCheque  ,
			@WDate   ,
			@CodIb   ,
			@NroIb  ,
			@NroInsc
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaProveedor]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaProveedor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaProveedor]
             @Proveedor char(11),
	@Nombre char(50),
	@Direccion char(50),
	@Localidad char(50),
	@Provincia char(2),
	@Postal char(4),
	@Cuit char(15),
	@Telefono    char(30),
	@EMail char(200),
	@Observaciones char(50),
	@Tipo char(1),
	@Iva char(1),
	@Dias char(20),
	@Empresa smallint,
	@Cuenta char(10),
	@Importe1 float,
	@Importe2 float,
	@Importe3 float,
	@Importe4 float,
	@Importe5 float,
	@Importe6 float,
	@NombreCheque char(50)  ,
	@WDate   char(10)  ,
	@CodIb Int  ,
	@NroIb char(20),
	@NroInsc char(15)


 AS
	INSERT INTO  Proveedor
			(
			Proveedor   ,
			Nombre        ,                                     
			Direccion      ,                                    
			Localidad      ,                                    
			Provincia ,
			Postal ,
			Cuit     ,       
			Telefono                       ,
			Email                          ,
			Observaciones            ,                          
			Tipo ,
			Iva  ,
			Dias ,                
			Empresa     ,
			Cuenta     ,
			Importe1    ,
			Importe2    ,
			Importe3    ,
			Importe4    ,
			Importe5    ,
			Importe6    ,
			NombreCheque                                       ,
			WDate   ,
			CodIb  ,
			NroIb  ,
			NroInsc
			)
VALUES
			(
		             @Proveedor,
			@Nombre,
			@Direccion,
			@Localidad,
			@Provincia,
			@Postal,
			@Cuit,
			@Telefono,
			@EMail,
			@Observaciones,
			@Tipo,
			@Iva,
			@Dias,
			@Empresa,
			@Cuenta,
			@Importe1,
			@Importe2,
			@Importe3,
			@Importe4,
			@Importe5,
			@Importe6,
			@NombreCheque     ,
			@WDate   ,
			@CodIb  ,
			@NroIb  ,
			@NroInsc
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaPrestamo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPrestamo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaPrestamo]
             @Clave char(8),
	@Codigo int,
	@Renglon int,
	@Fecha char(10),
	@OrdFecha char(8)  ,
	@Observaciones char(50),
	@Tipo char(1),
	@Articulo char(10),
	@Terminado char(12),
	@Cantidad float,
	@Costo float ,
	@Destino Int


 AS
	INSERT INTO Prestamo
			(
		             Clave ,
			Codigo ,
			Renglon ,
			Fecha ,
			OrdFecha ,
			Observaciones  ,
			Tipo ,
			Articulo ,
			Terminado ,
			Cantidad ,
			Costo , 
			Destino
			)

VALUES
			(
		             @Clave ,
			@Codigo ,
			@Renglon ,
			@Fecha ,
			@OrdFecha ,
			@Observaciones  ,
			@Tipo ,
			@Articulo ,
			@Terminado ,
			@Cantidad ,
			@Costo  ,
			@Destino
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaPresCon]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPresCon]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaPresCon]
	@Codigo int,
	@Fecha char(10),
	@OrdFecha char(8)  ,
	@Tipo char(1),
	@Articulo char(10),
	@Terminado char(12),
	@Cantidad1 float,
	@Cantidad2 float,
	@Costo float ,
	@Destino Int  ,
	@Observaciones char(50)  , 
	@Descripcion char(50)


 AS
	INSERT INTO PresCon
			(
			Codigo ,
			Fecha ,
			OrdFecha ,
			Tipo ,
			Articulo ,
			Terminado ,
			Cantidad1 ,
			Cantidad2,
			Costo , 
			Destino  ,
			Observaciones   ,
			Descripcion
			)

VALUES
			(
			@Codigo ,
			@Fecha ,
			@OrdFecha ,
			@Tipo ,
			@Articulo ,
			@Terminado ,
			@Cantidad1 ,
			@Cantidad2,
			@Costo  ,
			@Destino  ,
			@Observaciones   ,
			@Descripcion
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaPreciosMp]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPreciosMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaPreciosMp]
             @Clave char(16),
             @Cliente char(6),
             @Articulo char(10),   
             @Precio  float,	
             @Fecha1 char(10),
             @Factura1 char(6),
             @Precio1 float,
             @Cantidad1  float,
             @Fecha2 char(10),
             @Factura2 char(6),
             @Precio2 float,
             @Cantidad2  float,
             @Fecha3 char(10),
             @Factura3 char(6),
             @Precio3 float,
             @Cantidad3  float,
             @Fecha4 char(10),
             @Factura4 char(6),  
             @Precio4 float,
             @Cantidad4  float,
             @Fecha5 char(10),
             @Factura5 char(6),
             @Precio5 float,
             @Cantidad5  float,
             @WDate Char(10),
	@Fecha char(10),
	@Pago int


 AS
	INSERT INTO  PreciosMp
			(
		             Clave,
		             Cliente,
		             Articulo,   
		             Precio,	
		             Fecha1,
		             Factura1,
		             Precio1,
		             Cantidad1,
		             Fecha2,
		             Factura2,
		             Precio2,
		             Cantidad2,
		             Fecha3,
		             Factura3,
		             Precio3,
		             Cantidad3,
		             Fecha4,
		             Factura4,
		             Precio4,
		             Cantidad4,
		             Fecha5,
		             Factura5,
		             Precio5,
		             Cantidad5,
		             WDate ,
			Fecha   ,
			Pago
			)

VALUES
			(
		             @Clave,
		             @Cliente,
		             @Articulo,   
		             @Precio,	
		             @Fecha1,
		             @Factura1,
		             @Precio1,
		             @Cantidad1,
		             @Fecha2,
		             @Factura2,
		             @Precio2,
		             @Cantidad2,
		             @Fecha3,
		             @Factura3,
		             @Precio3,
		             @Cantidad3,
		             @Fecha4,
		             @Factura4,
		             @Precio4,
		             @Cantidad4,
		             @Fecha5,
		             @Factura5,
		             @Precio5,
		             @Cantidad5,
		             @WDate  ,
			@Fecha   ,
			@Pago	
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaPrecios1]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPrecios1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaPrecios1]
             @Clave char(18),
             @Cliente char(6),
             @Terminado char(12),   
             @Precio  float,	
             @Descripcion char(50),
             @Fecha1 char(10),
             @Factura1 char(6),
             @Precio1 float,
             @Cantidad1  float,
             @Fecha2 char(10),
             @Factura2 char(6),
             @Precio2 float,
             @Cantidad2  float,
             @Fecha3 char(10),
             @Factura3 char(6),
             @Precio3 float,
             @Cantidad3  float,
             @Fecha4 char(10),
             @Factura4 char(6),  
             @Precio4 float,
             @Cantidad4  float,
             @Fecha5 char(10),
             @Factura5 char(6),
             @Precio5 float,
             @Cantidad5  float,
             @WDate Char(10),
	@Fecha char(10),
	@Pago int


 AS
	INSERT INTO  Precios
			(
		             Clave,
		             Cliente,
		             Terminado,   
		             Precio,	
		             Descripcion,
		             Fecha1,
		             Factura1,
		             Precio1,
		             Cantidad1,
		             Fecha2,
		             Factura2,
		             Precio2,
		             Cantidad2,
		             Fecha3,
		             Factura3,
		             Precio3,
		             Cantidad3,
		             Fecha4,
		             Factura4,
		             Precio4,
		             Cantidad4,
		             Fecha5,
		             Factura5,
		             Precio5,
		             Cantidad5,
		             WDate ,
			Fecha   ,
			Pago
			)

VALUES
			(
		             @Clave,
		             @Cliente,
		             @Terminado,   
		             @Precio,	
		             @Descripcion,
		             @Fecha1,
		             @Factura1,
		             @Precio1,
		             @Cantidad1,
		             @Fecha2,
		             @Factura2,
		             @Precio2,
		             @Cantidad2,
		             @Fecha3,
		             @Factura3,
		             @Precio3,
		             @Cantidad3,
		             @Fecha4,
		             @Factura4,
		             @Precio4,
		             @Cantidad4,
		             @Fecha5,
		             @Factura5,
		             @Precio5,
		             @Cantidad5,
		             @WDate  ,
			@Fecha   ,
			@Pago	
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaPrecios]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPrecios]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaPrecios]
             @Clave char(18),
             @Cliente char(6),
             @Terminado char(12),   
             @Precio  float,	
             @Descripcion char(50),
             @Fecha1 char(10),
             @Factura1 char(6),
             @Precio1 float,
             @Cantidad1  float,
             @Fecha2 char(10),
             @Factura2 char(6),
             @Precio2 float,
             @Cantidad2  float,
             @Fecha3 char(10),
             @Factura3 char(6),
             @Precio3 float,
             @Cantidad3  float,
             @Fecha4 char(10),
             @Factura4 char(6),  
             @Precio4 float,
             @Cantidad4  float,
             @Fecha5 char(10),
             @Factura5 char(6),
             @Precio5 float,
             @Cantidad5  float,
             @WDate Char(10)


 AS
	INSERT INTO  Precios
			(
		             Clave,
		             Cliente,
		             Terminado,   
		             Precio,	
		             Descripcion,
		             Fecha1,
		             Factura1,
		             Precio1,
		             Cantidad1,
		             Fecha2,
		             Factura2,
		             Precio2,
		             Cantidad2,
		             Fecha3,
		             Factura3,
		             Precio3,
		             Cantidad3,
		             Fecha4,
		             Factura4,
		             Precio4,
		             Cantidad4,
		             Fecha5,
		             Factura5,
		             Precio5,
		             Cantidad5,
		             WDate 
			)

VALUES
			(
		             @Clave,
		             @Cliente,
		             @Terminado,   
		             @Precio,	
		             @Descripcion,
		             @Fecha1,
		             @Factura1,
		             @Precio1,
		             @Cantidad1,
		             @Fecha2,
		             @Factura2,
		             @Precio2,
		             @Cantidad2,
		             @Fecha3,
		             @Factura3,
		             @Precio3,
		             @Cantidad3,
		             @Fecha4,
		             @Factura4,
		             @Precio4,
		             @Cantidad4,
		             @Fecha5,
		             @Factura5,
		             @Precio5,
		             @Cantidad5,
		             @WDate 
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaPedidoDevolII]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPedidoDevolII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaPedidoDevolII]
             @Clave char(8),
	@Pedido int,
	@Renglon int,
	@Cliente char(6),
	@Fecha char(10),
	@Observaciones char(100),
	@Terminado char(12),
	@Cantidad float,
	@Precio float,
	@Linea int,
	@Facturado float,
	@Importe int,
	@Autorizo char(1),
	@Impresion char(1) ,
	@TipoPro  char(1)  ,
	@Articulo  char(10) ,
	@FechaOrd char(8) ,
	@TipoPedido char(2)


 AS
	INSERT INTO  PedidoDevol
			(
		             Clave ,
			Pedido ,
			Renglon ,
			Cliente ,
			Fecha ,
			Observaciones ,
			Terminado ,
			Cantidad ,
			Precio ,
			Linea ,
			Facturado ,
			Importe ,
			Autorizo,
			Impresion ,
			TipoPro   ,
			Articulo  ,
			FechaOrd  ,
			TipoPedido
			)

VALUES
			(
		             @Clave ,
			@Pedido ,
			@Renglon ,
			@Cliente ,
			@Fecha ,
			@Observaciones ,
			@Terminado ,
			@Cantidad ,
			@Precio ,
			@Linea ,
			@Facturado ,
			@Importe ,
			@Autorizo ,
			@Impresion ,
			@TipoPro  ,
			@Articulo  ,
			@FechaOrd  ,
			@TipoPedido
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaPedidoDevol]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPedidoDevol]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaPedidoDevol]
             @Clave char(8),
	@Pedido int,
	@Renglon int,
	@Cliente char(6),
	@Fecha char(10),
	@Observaciones char(100),
	@Terminado char(12),
	@Cantidad float,
	@Precio float,
	@Linea int,
	@Facturado float,
	@Importe int,
	@Autorizo char(1),
	@Impresion char(1) ,
	@TipoPro  char(1)  ,
	@Articulo  char(10) ,
	@FechaOrd char(8)


 AS
	INSERT INTO  PedidoDevol
			(
		             Clave ,
			Pedido ,
			Renglon ,
			Cliente ,
			Fecha ,
			Observaciones ,
			Terminado ,
			Cantidad ,
			Precio ,
			Linea ,
			Facturado ,
			Importe ,
			Autorizo,
			Impresion ,
			TipoPro   ,
			Articulo  ,
			FechaOrd 
			)

VALUES
			(
		             @Clave ,
			@Pedido ,
			@Renglon ,
			@Cliente ,
			@Fecha ,
			@Observaciones ,
			@Terminado ,
			@Cantidad ,
			@Precio ,
			@Linea ,
			@Facturado ,
			@Importe ,
			@Autorizo ,
			@Impresion ,
			@TipoPro  ,
			@Articulo  ,
			@FechaOrd
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaPedido]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPedido]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaPedido]
             @Clave char(8),
	@Pedido int,
	@Renglon int,
	@Cliente char(6),
	@Fecha char(10),
	@Fecentrega char(10),
	@Hora char(5),
	@Observaciones char(100),
	@Terminado char(12),
	@Cantidad float,
	@Envase1 int,
	@Canti1 int,
	@Envase2 int,
	@Canti2 int,
	@Envase3 int,
	@Canti3 int,
	@Envase4 int,
	@Canti4 int,
	@Fechaord char(8),
	@Precio float,
	@Linea int,
	@Facturado float,
	@Importe int,
	@Marca1 nvarchar(20),
	@Marca2 nvarchar(20),
	@Marca3 nvarchar(20),
	@Destino nvarchar(15),
	@Autorizo char(1),
	@Impresion char(1) ,
	@Tipoped int,
	@Cantidad1 float  ,
	@Cantidad2   float  ,
	@Lote1 int  ,
	@CantiLote1  float  ,
	@Lote2 int ,
	@CantiLote2  float  ,
	@Lote3 int  ,
	@CantiLote3  float  ,
	@Lote4 int  ,
	@CantiLote4  float  ,
	@Lote5 int  ,
	@CantiLote5  float  ,
	@Env1 int  ,
	@CantiEnv1  float  ,
	@Env2 int  ,
	@CantiEnv2  float  ,
	@Env3 int  ,
	@CantiEnv3  float  ,
	@Env4 int  ,
	@CantiEnv4  float  ,
	@Env5 int  ,
	@CantiEnv5  float  ,
	@Version int,
	@OrdFecEntrega char(8) ,
	@OrdenCpa char(10)  ,
	  @ttime varchar(10) ,
             @TipoPro  char(1)  ,
	@Articulo  char(10)


 AS
	INSERT INTO  Pedido
			(
		             Clave ,
			Pedido ,
			Renglon ,
			Cliente ,
			Fecha ,
			Fecentrega ,
			Hora ,
			Observaciones ,
			Terminado ,
			Cantidad ,
			Envase1 ,
			Canti1 ,
			Envase2 ,
			Canti2 ,
			Envase3 ,
			Canti3 ,
			Envase4 ,
			Canti4 ,
			Fechaord ,
			Precio ,
			Linea ,
			Facturado ,
			Importe ,
			Marca1 ,
			Marca2 ,
			Marca3 ,
			Destino ,
			Autorizo,
			Impresion ,
			Tipoped ,
			Cantidad1   ,
			Cantidad2   ,
			Lote1  ,
			CantiLote1,
			Lote2  ,
			CantiLote2,
			Lote3  ,
			CantiLote3,
			Lote4  ,
			CantiLote4,
			Lote5  ,
			CantiLote5,
			Env1  ,
			CantiEnv1,
			Env2  ,
			CantiEnv2,
			Env3  ,
			CantiEnv3,
			Env4  ,
			CantiEnv4,
			Env5  ,
			CantiEnv5   ,
			Version   ,
			OrdFecEntrega  ,
			OrdenCpa     ,
			  ttime  ,
                                        TipoPro   ,
			Articulo
			)

VALUES
			(
		             @Clave ,
			@Pedido ,
			@Renglon ,
			@Cliente ,
			@Fecha ,
			@Fecentrega ,
			@Hora ,
			@Observaciones ,
			@Terminado ,
			@Cantidad ,
			@Envase1 ,
			@Canti1 ,
			@Envase2 ,
			@Canti2 ,
			@Envase3 ,
			@Canti3 ,
			@Envase4 ,
			@Canti4 ,
			@Fechaord ,
			@Precio ,
			@Linea ,
			@Facturado ,
			@Importe ,
			@Marca1 ,
			@Marca2 ,
			@Marca3 ,
			@Destino ,
			@Autorizo ,
			@Impresion ,
			@Tipoped  ,
			@Cantidad1   ,
			@Cantidad2   ,
			@Lote1  ,
			@CantiLote1  ,
			@Lote2  ,
			@CantiLote2  ,
			@Lote3  ,
			@CantiLote3  ,
			@Lote4  ,
			@CantiLote4  ,
			@Lote5  ,
			@CantiLote5  ,
			@Env1  ,
			@CantiEnv1  ,
			@Env2  ,
			@CantiEnv2  ,
			@Env3  ,
			@CantiEnv3  ,
			@Env4  ,
			@CantiEnv4  ,
			@Env5  ,
			@CantiEnv5   ,
			@Version   ,
			@OrdFecEntrega  ,
			@OrdenCpa   ,
			 @ttime  ,
                                        @TipoPro  ,
			@Articulo
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaPagos]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPagos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaPagos]
	@Clave varchar(8),                     
	@Orden varchar(6),   
	@Renglon varchar(2), 
	@Proveedor varchar(11),
	@Fecha varchar(10),
	@Fechaord varchar(8),  
	@Tipoord  varchar(1),   
	@RetGanancias real,
	@RetIva real, 
	@RetOtra real,                                       
	@Retencion real,                                       
	@Tiporeg   real ,                                          
	@Tipo1  varchar(2)    ,                                          
	@Letra1  varchar(1),
	@Punto1 varchar(4), 
	@Numero1 varchar(8),
	@Importe1 real,
	@Tipo2 varchar(2), 
	@Numero2 varchar(8),                                             
	@Fecha2 varchar(10),  
	@Banco2 smallint,
	@Importe2 real,
	@Observaciones2 varchar(30),
	@Empresa smallint,
	@Concepto smallint,
	@Observaciones varchar(50),
	@Importe float,
	@Fechaord2 varchar(8),
	@Consecionaria smallint,
	@Impolist float,
	@Cuenta varchar(10)
AS
INSERT INTO Pagos
	(
		Clave ,                     
		Orden ,   
		Renglon , 
		Proveedor ,
		Fecha ,
		Fechaord ,  
		Tipoord ,   
		RetGanancias ,
		RetIva , 
		RetOtra ,                                       
		Retencion ,                                       
		Tiporeg  ,                                          
		Tipo1      ,                                          
		Letra1  ,
		Punto1 , 
		Numero1 ,
		Importe1 ,
		Tipo2 , 
		Numero2 ,                                             
		Fecha2 ,  
		Banco2 ,
		Importe2 ,
		Observaciones2 ,
		Empresa ,
		Concepto ,
		Observaciones ,
		Importe ,
		Fechaord2 ,
		Consecionaria ,
		Impolist ,
		Cuenta 
	)	
VALUES
	(
		@Clave ,                     
		@Orden ,   
		@Renglon , 
		@Proveedor ,
		@Fecha ,
		@Fechaord ,  
		@Tipoord ,   
		@RetGanancias ,
		@RetIva , 
		@RetOtra ,                                       
		@Retencion ,                                       
		@Tiporeg  ,                                          
		@Tipo1      ,                                          
		@Letra1  ,
		@Punto1 , 
		@Numero1 ,
		@Importe1 ,
		@Tipo2 , 
		@Numero2 ,                                             
		@Fecha2 ,  
		@Banco2 ,
		@Importe2 ,
		@Observaciones2 ,
		@Empresa ,
		@Concepto ,
		@Observaciones ,
		@Importe ,
		@Fechaord2 ,
		@Consecionaria ,
		@Impolist ,
		@Cuenta 
	)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaPago]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaPago]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaPago]
             @Pago int,
	@Nombre char(50),
	@Dias real,
	@Plazo real,
	@Tasa real,
	@Descuento real

 AS
	INSERT INTO  Pago
			(
			Pago  ,
			Nombre   ,
			Dias  ,
			Plazo   ,
			Tasa   ,
		             Descuento
			)

VALUES
			(
			@Pago  ,
			@Nombre   ,
			@Dias  ,
			@Plazo   ,
			@Tasa   ,
		             @Descuento
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaOt]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaOt]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaOt]
             @Codigo int,
	@Fecha char(10),
	@Cliente char(6),
	@Razon char(50),
	@Preparacion char(50) ,
	@Solidez char(50) ,
	@Observaciones1 char(50)   ,
	@Observaciones2 char(50)   ,
	@Observaciones3  char(50)   ,
	@Solicitante  char(50)   ,
	@Compo char(50),
	@Compo1 int,
	@Compo2 int,
	@Compo3 int,
	@Compo4 int,
	@Compo5 int,
	@Compo6 int,
	@Compo7 int,
	@Compo8 int,
	@Compo9 int,
	@Compo10 int,
	@Compo11 int,
	@Compo12 int,
	@Compo13 int,
	@Compo14 int,
	@Traba char(50),
	@Trabajo1 int,
	@Trabajo2 int,
	@Trabajo3 int,
	@Trabajo4 int,
	@Trabajo5 int,
	@Trabajo6 int,
	@Trabajo7 int,
	@Trabajo8 int,
	@Trabajo9 int,
	@Trabajo10 int,
	@Trabajo11 int,
	@Trabajo12 int,
	@Trabajo13 int,
	@Trabajo14 int,
	@Color char(50),
	@Color1 int,
	@Color2 int,
	@Color3 int,
	@Color4 int,
	@Color5 int,
	@Color6 int,
	@Color7 int,
	@Color8 int,
	@Color9 int,
	@Color10 int,
	@Color11 int,
	@Color12 int,
	@Color13 int,
	@Color14 int,
	@Color15 int,
	@Color16 int,
	@Color17 int,
	@Color18 int,
	@Color19 int,
	@Color20 int,
	@Color21 int,
	@Maqui char(50),
	@Maquina1 int,
	@Maquina2 int,
	@Maquina3 int,
	@Maquina4 int,
	@Maquina5 int,
	@Maquina6 int,
	@Maquina7 int,
	@Maquina8 int,
	@Maquina9 int,
	@Maquina10 int,
	@Maquina11 int,
	@Maquina12 int,
	@Maquina13 int,
	@Maquina14 int,
	@FechaCompro char(10),
	@FechaSalida char(10),
	@OrdFecha char(8),
	@OrdFechaCompro char(8),
	@OrdFechaSalida char(8)  ,
	@Clave int

 AS
	INSERT INTO  Ot
			(
		             Codigo ,
			Fecha ,
			Cliente ,
			Razon ,
			Preparacion ,
			Solidez ,
			Observaciones1 ,
			Observaciones2 ,
			Observaciones3  ,
			Solicitante  ,
			Compo ,
			Compo1 ,
			Compo2 ,
			Compo3 ,
			Compo4 ,
			Compo5 ,
			Compo6 ,
			Compo7 ,
			Compo8 ,
			Compo9 ,
			Compo10 ,
			Compo11 ,
			Compo12 ,
			Compo13 ,
			Compo14 ,
			Traba ,
			Trabajo1 ,
			Trabajo2 ,
			Trabajo3 ,
			Trabajo4 ,
			Trabajo5 ,
			Trabajo6 ,
			Trabajo7 ,
			Trabajo8 ,
			Trabajo9 ,
			Trabajo10 ,
			Trabajo11 ,
			Trabajo12 ,
			Trabajo13 ,
			Trabajo14 ,
			Color ,
			Color1 ,
			Color2 ,
			Color3 ,
			Color4 ,
			Color5 ,
			Color6 ,
			Color7 ,
			Color8 ,
			Color9 ,
			Color10 ,
			Color11 ,
			Color12 ,
			Color13 ,
			Color14 ,
			Color15 ,
			Color16 ,
			Color17 ,
			Color18 ,
			Color19 ,
			Color20 ,
			Color21 ,
			Maqui,
			Maquina1 ,
			Maquina2 ,
			Maquina3 ,
			Maquina4 ,
			Maquina5 ,
			Maquina6 ,
			Maquina7 ,
			Maquina8 ,
			Maquina9 ,
			Maquina10 ,
			Maquina11 ,
			Maquina12 ,
			Maquina13 ,
			Maquina14 ,
			FechaCompro ,
			FechaSalida ,
			OrdFecha ,
			OrdFechaCompro ,
			OrdFechaSalida  ,
			Clave
			)

VALUES
			(
		             @Codigo ,
			@Fecha ,
			@Cliente ,
			@Razon ,
			@Preparacion ,
			@Solidez ,
			@Observaciones1 ,
			@Observaciones2 ,
			@Observaciones3  ,
			@Solicitante  ,
			@Compo ,
			@Compo1 ,
			@Compo2 ,
			@Compo3 ,
			@Compo4 ,
			@Compo5 ,
			@Compo6 ,
			@Compo7 ,
			@Compo8 ,
			@Compo9 ,
			@Compo10 ,
			@Compo11 ,
			@Compo12 ,
			@Compo13 ,
			@Compo14 ,
			@Traba ,
			@Trabajo1 ,
			@Trabajo2 ,
			@Trabajo3 ,
			@Trabajo4 ,
			@Trabajo5 ,
			@Trabajo6 ,
			@Trabajo7 ,
			@Trabajo8 ,
			@Trabajo9 ,
			@Trabajo10 ,
			@Trabajo11 ,
			@Trabajo12 ,
			@Trabajo13 ,
			@Trabajo14 ,
			@Color ,
			@Color1 ,
			@Color2 ,
			@Color3 ,
			@Color4 ,
			@Color5 ,
			@Color6 ,
			@Color7 ,
			@Color8 ,
			@Color9 ,
			@Color10 ,
			@Color11 ,
			@Color12 ,
			@Color13 ,
			@Color14 ,
			@Color15 ,
			@Color16 ,
			@Color17 ,
			@Color18 ,
			@Color19 ,
			@Color20 ,
			@Color21 ,
			@Maqui,
			@Maquina1 ,
			@Maquina2 ,
			@Maquina3 ,
			@Maquina4 ,
			@Maquina5 ,
			@Maquina6 ,
			@Maquina7 ,
			@Maquina8 ,
			@Maquina9 ,
			@Maquina10 ,
			@Maquina11 ,
			@Maquina12 ,
			@Maquina13 ,
			@Maquina14 ,
			@FechaCompro ,
			@FechaSalida ,
			@OrdFecha ,
			@OrdFechaCompro ,
			@OrdFechaSalida   ,
			@Clave
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaOrdenIII]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaOrdenIII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaOrdenIII]
             @Clave char(8),
	@Orden int,
	@Renglon int,
	@Fecha char(10),
	@Proveedor char(11),
	@Articulo char(10),
	@Cantidad float,
	@Precio float,
	@Fecha1 char(10), 
	@Fecha2 char(10),
	@Condicion char(40),
	@Recibida float,
	@Saldo float,
	@Fechaord char(8), 
	@Liberada float,
	@Devuelta float,
	@Fechaentrega char(10),
	@WDate char(10)  ,
	@Moneda  int  ,
	@Tipo  int,
	@Carpeta int ,
	@Derechos  float  ,
	@Origen  char(50)


 AS
	INSERT INTO  Orden
			(
		             Clave ,
			Orden ,
			Renglon ,
			Fecha ,
			Proveedor ,
			Articulo ,
			Cantidad ,
			Precio ,
			Fecha1 ,
			Fecha2 ,
			Condicion ,
			Recibida ,
			Saldo ,
			Fechaord ,
			Liberada ,
			Devuelta ,
			Fechaentrega ,
			WDate   ,
			Moneda   ,
			Tipo  ,
			Carpeta  ,
			Derechos  ,
			Origen
			)

VALUES
			(
		             @Clave ,
			@Orden ,
			@Renglon ,
			@Fecha ,
			@Proveedor ,
			@Articulo ,
			@Cantidad ,
			@Precio ,
			@Fecha1 ,
			@Fecha2 ,
			@Condicion ,
			@Recibida ,
			@Saldo ,
			@Fechaord ,
			@Liberada ,
			@Devuelta ,
			@Fechaentrega ,
			@WDate   ,
			@Moneda   ,
			@Tipo   ,
			@Carpeta  ,
			@Derechos   ,
			@Origen
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaOrdenII]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaOrdenII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaOrdenII]
             @Clave char(8),
	@Orden int,
	@Renglon int,
	@Fecha char(10),
	@Proveedor char(11),
	@Articulo char(10),
	@Cantidad float,
	@Precio float,
	@Fecha1 char(10), 
	@Fecha2 char(10),
	@Condicion char(40),
	@Recibida float,
	@Saldo float,
	@Fechaord char(8), 
	@Liberada float,
	@Devuelta float,
	@Fechaentrega char(10),
	@WDate char(10)  ,
	@Moneda  int  ,
	@Tipo  int,
	@Carpeta int 


 AS
	INSERT INTO  Orden
			(
		             Clave ,
			Orden ,
			Renglon ,
			Fecha ,
			Proveedor ,
			Articulo ,
			Cantidad ,
			Precio ,
			Fecha1 ,
			Fecha2 ,
			Condicion ,
			Recibida ,
			Saldo ,
			Fechaord ,
			Liberada ,
			Devuelta ,
			Fechaentrega ,
			WDate   ,
			Moneda   ,
			Tipo  ,
			Carpeta 
			)

VALUES
			(
		             @Clave ,
			@Orden ,
			@Renglon ,
			@Fecha ,
			@Proveedor ,
			@Articulo ,
			@Cantidad ,
			@Precio ,
			@Fecha1 ,
			@Fecha2 ,
			@Condicion ,
			@Recibida ,
			@Saldo ,
			@Fechaord ,
			@Liberada ,
			@Devuelta ,
			@Fechaentrega ,
			@WDate   ,
			@Moneda   ,
			@Tipo   ,
			@Carpeta 
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaOrden]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaOrden]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaOrden]
             @Clave char(8),
	@Orden int,
	@Renglon int,
	@Fecha char(10),
	@Proveedor char(11),
	@Articulo char(10),
	@Cantidad float,
	@Precio float,
	@Fecha1 char(10), 
	@Fecha2 char(10),
	@Condicion char(40),
	@Recibida float,
	@Saldo float,
	@Fechaord char(8), 
	@Liberada float,
	@Devuelta float,
	@Fechaentrega char(10),
	@WDate char(10)


 AS
	INSERT INTO  Orden
			(
		             Clave ,
			Orden ,
			Renglon ,
			Fecha ,
			Proveedor ,
			Articulo ,
			Cantidad ,
			Precio ,
			Fecha1 ,
			Fecha2 ,
			Condicion ,
			Recibida ,
			Saldo ,
			Fechaord ,
			Liberada ,
			Devuelta ,
			Fechaentrega ,
			WDate 
			)

VALUES
			(
		             @Clave ,
			@Orden ,
			@Renglon ,
			@Fecha ,
			@Proveedor ,
			@Articulo ,
			@Cantidad ,
			@Precio ,
			@Fecha1 ,
			@Fecha2 ,
			@Condicion ,
			@Recibida ,
			@Saldo ,
			@Fechaord ,
			@Liberada ,
			@Devuelta ,
			@Fechaentrega ,
			@WDate 
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaMuestraImpre]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMuestraImpre]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaMuestraImpre]
             @Numero int,
	@Fecha char(10) ,
             @Codigo char(15),
             @Descripcion char(50),
             @Cantidad char(10),
	@DescriCliente char(50),
	@Cliente char(50),
	@Observaciones char(50),
	@Fecha2 char(10),
	@Codigo2 char(15),
	@Descripcion2 char(50) ,
	@Lote char(10) ,
	@Observaciones2 char(50),
	@Cantidad2 char(10),
	@Actualiza char(1)

 AS
	INSERT INTO  MuestraImpre
			(
		             Numero,
			Fecha ,
			Codigo ,
		             Descripcion,
		             Cantidad,
		             DescriCliente,
		             Cliente,
		             Observaciones,
		             Fecha2,
		             Codigo2,
			Descripcion2  ,
			Lote  ,
			Observaciones2   ,
			Cantidad2   ,
			Actualiza
			)

VALUES
			(
		             @Numero,
			@Fecha ,
			@Codigo ,
		             @Descripcion,
		             @Cantidad,
		             @DescriCliente,
		             @Cliente,
		             @Observaciones,
		             @Fecha2,
		             @Codigo2,
			@Descripcion2  ,
			@Lote  ,
			@Observaciones2   ,
			@Cantidad2   ,
			@Actualiza
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaMuestra]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMuestra]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaMuestra]
             @Codigo int,
	@Producto char(12),
	@Articulo  char(10),
	@Ensayo char(15),
	@Nombre char(50),
	@Fecha char(10),
	@OrdFecha char(8),
	@Cantidad char(15),
	@Cliente char(6),
	@Razon char(50),
	@DescriCliente char(50),
	@Vendedor int,
	@DesVendedor char(50) ,
	@Observaciones  char(50) 

 AS
	INSERT INTO  Muestra
			(
		             Codigo ,
			Producto ,
			Articulo ,
			Ensayo ,
			Nombre ,
			Fecha ,
			OrdFecha ,
			Cantidad ,
			Cliente ,
			Razon ,
			DescriCliente ,
			Vendedor ,
			DesVendedor ,
			Observaciones
			)

VALUES
			(
		             @Codigo ,
			@Producto ,
			@Articulo  ,
			@Ensayo , 
			@Nombre  ,
			@Fecha ,
			@OrdFecha ,
			@Cantidad ,
			@Cliente ,
			@Razon  ,
			@DescriCliente ,
			@Vendedor ,
			@DesVendedor  ,
			@Observaciones
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaMovvar]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMovvar]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaMovvar]
             @Clave char(8),
	@Codigo int,
	@Renglon int,
	@Fecha char(10),
	@Tipo char(1),
	@Articulo char(10),
	@Terminado char(12),
	@Cantidad float,
	@Fechaord char(8),
	@Movi char(1),
	@Tipomov char(1),
	@Observaciones char(50),
	@WDate char(10),
	@Marca char(1) ,
	@Lote int


 AS
	INSERT INTO  Movvar
			(
		             Clave ,
			Codigo ,
			Renglon ,
			Fecha ,
			Tipo ,
			Articulo ,
			Terminado ,
			Cantidad ,
			Fechaord ,
			Movi ,
			Tipomov ,
			Observaciones ,
			WDate ,
			Marca  ,
			Lote
			)

VALUES
			(
		             @Clave ,
			@Codigo ,
			@Renglon ,
			@Fecha ,
			@Tipo ,
			@Articulo ,
			@Terminado ,
			@Cantidad ,
			@Fechaord ,
			@Movi ,
			@Tipomov ,
			@Observaciones ,
			@WDate ,
			@Marca  ,
			@Lote
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaMovlab]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMovlab]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaMovlab]
             @Clave char(8),
	@Codigo int,
	@Renglon int,
	@Fecha char(10),
	@Tipo char(1),
	@Articulo char(10),
	@Terminado char(12),
	@Cantidad float,
	@Fechaord char(8),
	@Movi char(1),
	@Tipomov char(1),
	@Observaciones char(50),
	@WDate char(10),
	@Marca char(1) ,
	@Lote int

 AS
	INSERT INTO  Movlab
			(
             			Clave ,
			Codigo ,
			Renglon ,
			Fecha ,
			Tipo ,
			Articulo ,
			Terminado ,
			Cantidad ,
			Fechaord ,
			Movi ,
			Tipomov ,
			Observaciones ,
			WDate ,
			Marca  ,
			Lote
			)

VALUES
			(
             			@Clave ,
			@Codigo ,
			@Renglon ,
			@Fecha ,
			@Tipo ,
			@Articulo ,
			@Terminado ,
			@Cantidad ,
			@Fechaord ,
			@Movi ,
			@Tipomov ,
			@Observaciones ,
			@WDate ,
			@Marca  ,
			@Lote
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaMovguia]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMovguia]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaMovguia]
             @Clave char(9),
	@Tipomov int,
	@Codigo int,
	@Renglon int,
	@Fecha char(10),
	@Tipo char(1),
	@Articulo char(10),
	@Terminado char(12),
	@Cantidad float,
	@Fechaord char(8),
	@Movi char(1),
	@Observaciones char(50),
	@WDate char(10),
	@Marca char(1),
	@Destino int ,
	@Lote int ,
	@Saldo float ,
	@Partida Int


 AS
	INSERT INTO  Guia
			(
		             Clave ,
			Tipomov  ,
			Codigo ,
			Renglon ,
			Fecha ,
			Tipo ,
			Articulo ,
			Terminado ,
			Cantidad ,
			Fechaord ,
			Movi ,
			Observaciones ,
			WDate ,
			Marca ,
			Destino ,
			Lote ,
			Saldo , 
			Partida
			)

VALUES
			(
		             @Clave ,
			@Tipomov, 
			@Codigo ,
			@Renglon ,
			@Fecha ,
			@Tipo ,
			@Articulo ,
			@Terminado ,
			@Cantidad ,
			@Fechaord ,
			@Movi ,
			@Observaciones ,
			@WDate ,
			@Marca ,
			@Destino  ,
			@Lote ,
			@Saldo  ,
			@Partida
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaMovgasCon]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMovgasCon]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaMovgasCon]
             @Clave char(10),
             @Empresa int,
             @Carpeta int,
             @Fecha char(10),
	@Derechos float,
	@Orden int,
	@Concepto int,
	@Importe float ,
	@Auxiliar float ,
	@OrdFecha char(8) ,
	@Proveedor char(11)  ,
	@Origen   char(30)  ,
	@Moneda  int  ,
	@Marca int  ,
	@ImpoDerechos  float  ,
	@FechaLLegada char(10)   ,
	@OrdFechaLLegada char(8)   ,
	@CostoFlete  float  ,
	@Gastos  Float   ,
	@Pagado  float

 AS
	INSERT INTO  MovgasCon
			(
		             Clave,
			Empresa ,
		             Carpeta,
		             Fecha,
		             Derechos,
		             Orden,
		             Concepto,
		             Importe,
		             Auxiliar,
			OrdFecha  ,
			Proveedor  ,
			Origen   ,
			Moneda   ,
			Marca    ,
			ImpoDerechos   ,
			FechaLLegada   ,
			OrdFechaLLegada   ,
			CostoFlete   ,
			Gastos   ,
			Pagado
			)

VALUES
			(
		             @Clave,
			@Empresa   ,
		             @Carpeta,
		             @Fecha,
		             @Derechos,
		             @Orden,
		             @Concepto,
		             @Importe,
		             @Auxiliar,
			@OrdFecha  ,
			@Proveedor  ,
			@Origen   ,
			@Moneda,
			@Marca   ,
			@ImpoDerechos  ,
			@FechaLLegada   ,
			@OrdFechaLLegada  ,
			@CostoFlete  ,
			@Gastos   ,
			@Pagado
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaMovgas]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMovgas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaMovgas]
             @Clave char(8),
             @Carpeta int,
             @Renglon int,
             @Fecha char(10),
	@Derechos float,
	@Orden int,
	@Concepto int,
	@Importe float ,
	@Auxiliar float ,
	@OrdFecha char(8) ,
	@Proveedor char(11)  ,
	@Origen   char(30)  ,
	@Moneda  int   ,
	@Marca  char(1),
	@ImpoDerechos float  ,
	@FechaLLegada  char(10) ,
	@OrdFechaLLegada char(8)  ,
	@CostoFlete  float  ,
	@Gastos   float  ,
	@Pagado float  ,
	@Empresa Int

 AS
	INSERT INTO  Movgas
			(
		             Clave,
		             Carpeta,
		             Renglon,
		             Fecha,
		             Derechos,
		             Orden,
		             Concepto,
		             Importe,
		             Auxiliar,
			OrdFecha  ,
			Proveedor  ,
			Origen   ,
			Moneda  ,
			Marca   ,
			ImpoDerechos  ,
			FechaLLegada  ,
			OrdFechaLLegada  ,
			CostoFlete  ,
			Gastos  ,
			Pagado  ,
			Empresa
			)

VALUES
			(
		             @Clave,
		             @Carpeta,
		             @Renglon,
		             @Fecha,
		             @Derechos,
		             @Orden,
		             @Concepto,
		             @Importe,
		             @Auxiliar,
			@OrdFecha  ,
			@Proveedor  ,
			@Origen   ,
			@Moneda  ,
			@Marca  ,
			@ImpoDerechos  ,
			@FechaLlegada  ,
			@OrdFechaLLegada  ,
			@CostoFlete   ,
			@Gastos  ,
			@Pagado  ,
			@Empresa
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaMovenv]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMovenv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaMovenv]
             @Clave char(9),
             @Tipo char(1),
             @Codigo int,
             @Renglon int,
             @Fecha char(10),
             @Fechaord char(8),
             @Cliente char(6),
             @Envase int,
             @Movimiento char(1),
             @Cantidad float


 AS
	INSERT INTO  Movenv
			(
		             Clave,
		             Tipo,
		             Codigo,
		             Renglon,
		             Fecha,
		             Fechaord,
		             Cliente,
		             Envase,
		             Movimiento,
		             Cantidad
			)

VALUES
			(
		             @Clave,
		             @Tipo,
		             @Codigo,
		             @Renglon,
		             @Fecha,
		             @Fechaord,
		             @Cliente,
		             @Envase,
		             @Movimiento,
		             @Cantidad
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaMinimoPlanta]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMinimoPlanta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaMinimoPlanta]
             @Codigo char(12),
	@Articulo char(10),
	@Terminado char(12),
	@Descripcion char(50),
	@Stock1 float,
	@Stock2 float,
	@Stock3 float,
	@Stock4 float,
	@Stock5 float,
	@Minimo1 float ,
	@Minimo2 float ,
	@Minimo3 float ,
	@Minimo4 float ,
	@Minimo5 float 

 AS
	INSERT INTO  Minimo
			(
		             Codigo ,
			Articulo ,
			Terminado ,
			Descripcion ,
			Stock1 ,
			Stock2 ,
			Stock3 ,
			Stock4 ,
			Stock5 ,
			Minimo1  ,
			Minimo2,
			Minimo3,
			Minimo4,
			Minimo5
			)

VALUES
			(
		             @Codigo ,
			@Articulo ,
			@Terminado ,
			@Descripcion  ,
			@Stock1 ,
			@Stock2 ,
			@Stock3 ,
			@Stock4 ,
			@Stock5 ,
			@Minimo1  ,
			@Minimo2,
			@Minimo3,
			@Minimo4,
			@Minimo5
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaMinimo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMinimo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaMinimo]
             @Codigo char(12),
	@Articulo char(10),
	@Terminado char(12),
	@Descripcion char(50),
	@Stock1 float,
	@Stock2 float,
	@Stock3 float,
	@Stock4 float,
	@Stock5 float,
	@Minimo float 

 AS
	INSERT INTO  Minimo
			(
		             Codigo ,
			Articulo ,
			Terminado ,
			Descripcion ,
			Stock1 ,
			Stock2 ,
			Stock3 ,
			Stock4 ,
			Stock5 ,
			Minimo
			)

VALUES
			(
		             @Codigo ,
			@Articulo ,
			@Terminado ,
			@Descripcion  ,
			@Stock1 ,
			@Stock2 ,
			@Stock3 ,
			@Stock4 ,
			@Stock5 ,
			@Minimo
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaMarcas]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaMarcas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaMarcas]
             @Clave char(21),
	@Articulo char(10),
	@Proveedor char(11),
	@Descripcion char(50)
 AS
	INSERT INTO  Marcas
			(
			Clave     ,
			Articulo ,                                       
			Proveedor  ,
			Descripcion 
			)
VALUES
			(
          	 		@Clave,
			@Articulo,
			@Proveedor,
			@Descripcion
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaLineaMp]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaLineaMp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaLineaMp]
             @Linea int,
	@Nombre char(50)


 AS
	INSERT INTO  LineasMp
			(
			Linea   ,
		             Nombre
			)

VALUES
			(
		             @Linea,
			@Nombre
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaLinea]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaLinea]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaLinea]
             @Linea int,
	@Nombre char(50)


 AS
	INSERT INTO  Lineas
			(
			Linea   ,
		             Nombre
			)

VALUES
			(
		             @Linea,
			@Nombre
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaLaudo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaLaudo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaLaudo]
             @Clave char(8),
	@Laudo int,
	@Renglon int,
	@Fecha char(10),
	@Articulo char(10),
	@Liberada float,
	@Devuelta float,
	@Orden int,
	@Marca char(1),
	@Lote int,
	@Rechazo int,
	@Informe int,
	@Actualiza char(1),
	@WDate char(10),
	@Saldo float  ,
	@Origen  char(50)   ,
	@PartiOri  char(20)   ,
	@Envase   int


 AS
	INSERT INTO  Laudo
			(
		             Clave ,
			Laudo ,
			Renglon ,
			Fecha ,
			Articulo ,
			Liberada ,
			Devuelta ,
			Orden ,
			Marca ,
			Lote ,
			Rechazo ,
			Informe ,
			Actualiza,
			WDate  ,
			Saldo   ,
			Origen   ,
			PartiOri   ,
			Envase
			)

VALUES
			(
		             @Clave ,
			@Laudo ,
			@Renglon ,
			@Fecha ,
			@Articulo ,
			@Liberada ,
			@Devuelta ,
			@Orden ,
			@Marca ,
			@Lote ,
			@Rechazo ,
			@Informe ,
			@Actualiza,
			@WDate ,
			@Saldo   ,
			@Origen  ,
			@PartiOri   ,
			@Envase
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaIvaCompras]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaIvaCompras]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaIvaCompras]
	@NroInterno  int,
	@Proveedor char(11),
	@Tipo char(2),
	@Letra char(1),
	@Punto char(4),
	@Numero char(8),  
	@Fecha  char(10),    
	@Vencimiento char(10),
	@Vencimiento1 char(10),
	@Periodo    char(10),
	@Neto   float,                                               
	@Iva21  float,                                               
	@Iva5   float,                                               
	@Iva27  float,                                               
	@Ib     float,                                               
	@Exento float,                                               
	@Contado char(1),
	@Impre char(2),
	@Ordfecha char(8),
	@Empresa smallint,
	@Netolist    float,                                          
	@ExentoList  float    ,
	@Paridad float  ,
	@Pago  int                                     
AS
	INSERT INTO IvaComp
			(
			NroInterno  ,
			Proveedor   ,
			Tipo ,
			Letra ,
			Punto ,
			Numero ,  
			Fecha   ,   
			Vencimiento ,
			Vencimiento1 ,
			Periodo    ,
			Neto        ,                                          
			Iva21        ,                                         
			Iva5          ,                                        
			Iva27          ,                                       
			Ib              ,                                      
			Exento           ,                                     
			Contado ,
			Impre ,
			Ordfecha ,
			Empresa ,
			Netolist ,                                             
			ExentoList  ,
			Paridad   ,
			Pago                                          
			)
		VALUES
			(	
			@NroInterno,
			@Proveedor,
			@Tipo,
			@Letra,
			@Punto,
			@Numero,  
			@Fecha,    
			@Vencimiento,
			@Vencimiento1,
			@Periodo,
			@Neto,                                               
			@Iva21,                                               
			@Iva5,                                               
			@Iva27,                                               
			@Ib,                                               
			@Exento,                                               
			@Contado,
			@Impre,
			@Ordfecha,
			@Empresa,
			@Netolist,                                          
			@ExentoList    ,
			@Paridad    ,
			@Pago
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaInventario]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaInventario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaInventario]
             @Clave char(8),
	@Numero int,
	@Renglon int,
	@Tipo char(1),
	@Articulo char(10),
	@Terminado char(12),
	@Talon int  ,
	@Cantidad float,
	@Lote int  ,
	@Ubicacion  char(20)  ,
	@Observaciones char(30)


 AS
	INSERT INTO  Inventario
			(
		             Clave ,
			Numero ,
			Renglon ,
			Tipo ,
			Articulo ,
			Terminado ,
			Talon  ,
			Cantidad ,
			Lote ,
			Ubicacion  ,
			Observaciones 
			)

VALUES
			(
		             @Clave ,
			@Numero ,
			@Renglon ,
			@Tipo ,
			@Articulo ,
			@Terminado ,
			@Talon ,
			@Cantidad ,
			@Lote ,
			@Ubicacion  ,
			@Observaciones
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaInforme]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaInforme]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaInforme]
             @Clave char(8),
	@Informe int,
	@Renglon int,
	@Fecha char(10),
	@Remito int,
	@Proveedor char(11),
	@Orden int,
	@Articulo char(10),
	@Cantidad int,
	@Resta int,
	@Fechaord char(8),
	@WDate char(10),
	@Envase int


 AS
	INSERT INTO  Informe
			(
		             Clave ,
			Informe ,
			Renglon ,
			Fecha ,
			Remito ,
			Proveedor ,
			Orden ,
			Articulo ,
			Cantidad ,
			Resta ,
			Fechaord ,
			WDate ,
			Envase
			)

VALUES
			(
		             @Clave ,
			@Informe ,
			@Renglon ,
			@Fecha ,
			@Remito ,
			@Proveedor ,
			@Orden ,
			@Articulo ,
			@Cantidad ,
			@Resta ,
			@Fechaord ,
			@WDate ,
			@Envase
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaImputacion]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaImputacion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaImputacion]
	@Clave varchar(24),                    
	@TipoMovi varchar(1), 
	@Proveedor varchar(11),   
	@TipoComp varchar(2),
	@LetraComp varchar(1), 
	@PuntoComp varchar(4), 
	@NroComp varchar(8), 
	@Renglon varchar(2),
	@Fecha   varchar(10),   
	@Observaciones varchar(30),                  
	@Cuenta   varchar(10)  ,
	@Debito   float  ,                                           
	@Credito  float  ,                                           
	@FechaOrd varchar(8),
	@Titulo   varchar(30),                      
	@Empresa smallint,
	@DebitoList float,                                            
	@CreditoList float,                                           
	@NroInterno int
AS
INSERT INTO Imputac
	(
		Clave,                    
		TipoMovi, 
		Proveedor,   
		TipoComp, 
		LetraComp, 
		PuntoComp, 
		NroComp , 
		Renglon, 
		Fecha  ,    
		Observaciones,                  
		Cuenta     ,
		Debito     ,                                           
		Credito    ,                                           
		FechaOrd ,
		Titulo   ,                      
		Empresa ,
		DebitoList,                                            
		CreditoList,                                           
		NroInterno  
	)	
VALUES
	(
		@Clave,
		@TipoMovi,
		@Proveedor,
		@TipoComp,
		@LetraComp,
		@PuntoComp,
		@NroComp,
		@Renglon,
		@Fecha,
		@Observaciones,
		@Cuenta,
		@Debito,
		@Credito,
		@FechaOrd,
		@Titulo,
		@Empresa,
		@DebitoList,
		@CreditoList,
		@NroInterno
	)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaHoja]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaHoja]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaHoja]
             @Clave char(8),
	@Hoja int,
	@Renglon int,
	@Fecha char(10),
	@Producto char(12),
	@Cantidad float,
	@Tipo char(1),
	@Lote int,
	@Articulo char(10),
	@Terminado char(12),
	@Teorico float,
	@Real float,
	@Fechaing char(10),
	@Fechaingord char(8),
	@WDate char(10),
	@WImporte float,
	@Marca char(1) ,
	@Saldo float  ,
	@Lote1 int  ,
	@Canti1 float  ,
	@Lote2 int   ,
	@Canti2 float  ,
	@Lote3 int  ,
	@Canti3 float  ,
	@Costo1 float  ,
	@Costo2  float  ,
	@Costo3 float

 AS
	INSERT INTO  Hoja
			(
		             Clave ,
			Hoja ,
			Renglon ,
			Fecha ,
			Producto ,
			Cantidad ,
			Tipo,
			Lote ,
			Articulo,
			Terminado ,
			Teorico ,
			Real ,
			Fechaing ,
			Fechaingord ,
			WDate ,
			WImporte ,
			Marca ,
			Saldo   ,
			Lote1  ,
			Canti1  ,
			Lote2  ,
			Canti2  ,
			Lote3  ,
			Canti3  ,
			Costo1  ,
			Costo2  ,
			Costo3
			)

VALUES
			(
		             @Clave ,
			@Hoja ,
			@Renglon ,
			@Fecha ,
			@Producto ,
			@Cantidad ,
			@Tipo,
			@Lote ,
			@Articulo,
			@Terminado ,
			@Teorico ,
			@Real ,
			@Fechaing ,
			@Fechaingord ,
			@WDate ,
			@WImporte ,
			@Marca  ,
			@Saldo   ,
			@Lote1  ,
			@Canti1  ,
			@Lote2  ,
			@Canti2  ,
			@Lote3  ,
			@Canti3  ,
			@Costo1  ,
			@Costo2  ,
			@Costo3
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaGasimpo]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaGasimpo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaGasimpo]
             @Codigo int,
	@Nombre char(50)


 AS
	INSERT INTO  Gasimpo
			(
			Codigo   ,
		             Nombre
			)

VALUES
			(
		             @Codigo,
			@Nombre
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaEstadisticaDev]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEstadisticaDev]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaEstadisticaDev]
             @Clave char(12),
             @Tipo int,
             @Numero int,
             @Renglon int,
             @Articulo char(12),
             @Cantidad float,
             @Precio float,
             @PrecioUS float,
             @Importe float,
             @ImporteUS float,
             @Cliente char(6),
             @Paridad float,
             @Vendedor int,
             @Rubro int,
             @Linea int,
             @Costo1 float,
             @Costo2 float,
             @Coeficiente float,
             @Pedido int,
             @Fecha char(10),
             @Importe1 float,
             @Importe2 float,
             @Importe3 float,
             @Importe4 float,
             @Ordfecha char(8),
             @WArticulo char(8),
             @Remito char(10),
             @WDate char(10),
             @WCantidad float,
             @WImporte float,
             @WImporteUs float,
             @Marca char(1) ,
             @Lote1 int ,
             @Canti1 float ,
             @Lote2 int ,
             @Canti2 float, 
             @Lote3 int,
             @Canti3 float,
             @Lote4 int ,
             @Canti4 float, 
             @Lote5 int,
             @Canti5 float  ,
	@Entrada float  ,
	@Tipopro char(2)  ,
	@Hoja  int,
	@TipoProDy char(1)  ,
	@ArticuloDy char(10)


 AS
	INSERT INTO  Estadistica
			(
		             Clave,
		             Tipo,
		             Numero,
		             Renglon,
		             Articulo,
		             Cantidad,
		             Precio,
		             PrecioUS,
		             Importe,
		             ImporteUs,
		             Cliente,
		             Paridad,
		             Vendedor,
		             Rubro,
		             Linea,
		             Costo1,
		             Costo2,
		             Coeficiente,
		             Pedido,
		             Fecha,
		             Importe1,
		             Importe2,
		             Importe3,
		             Importe4,
		             Ordfecha,
		             WArticulo,
		             Remito,
		             WDate,
		             WCantidad,
		             WImporte,
		             WImporteUs,
		             Marca ,
		             Lote1,
		             Canti1,
		             Lote2,
		             Canti2,
		             Lote3,
		             Canti3,
		             Lote4,
		             Canti4,
		             Lote5,
		             Canti5  ,
			Entrada   ,
			Tipopro  ,
			Hoja,
			TipoProDy   ,
			ArticuloDy
			)

VALUES
			(
		             @Clave,
		             @Tipo,
		             @Numero,
		             @Renglon,
		             @Articulo,
		             @Cantidad,
		             @Precio,
		             @PrecioUS,
		             @Importe,
		             @ImporteUS,
		             @Cliente,
		             @Paridad,
		             @Vendedor,
		             @Rubro,
		             @Linea,
		             @Costo1,
		             @Costo2,
		             @Coeficiente,
		             @Pedido,
		             @Fecha,
		             @Importe1,
		             @Importe2,
		             @Importe3,
		             @Importe4,
		             @Ordfecha,
		             @WArticulo,
		             @Remito,
		             @WDate,
		             @WCantidad,
		             @WImporte,
		             @WImporteUs,
		             @Marca,
		             @Lote1,
		             @Canti1,
		             @Lote2,
		             @Canti2,
		             @Lote3,
		             @Canti3,
		             @Lote4,
		             @Canti4,
		             @Lote5,
		             @Canti5  ,
			@Entrada  ,
			@Tipopro   ,
			@Hoja  ,
			@TipoProDy   ,
			@ArticuloDy			
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaEstadistica]    Script Date: 05/01/2016 17:47:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEstadistica]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaEstadistica]
             @Clave char(12),
             @Tipo int,
             @Numero int,
             @Renglon int,
             @Articulo char(12),
             @Cantidad float,
             @Precio float,
             @PrecioUS float,
             @Importe float,
             @ImporteUS float,
             @Cliente char(6),
             @Paridad float,
             @Vendedor int,
             @Rubro int,
             @Linea int,
             @Costo1 float,
             @Costo2 float,
             @Coeficiente float,
             @Pedido int,
             @Fecha char(10),
             @Importe1 float,
             @Importe2 float,
             @Importe3 float,
             @Importe4 float,
             @Ordfecha char(8),
             @WArticulo char(8),
             @Remito char(10),
             @WDate char(10),
             @WCantidad float,
             @WImporte float,
             @WImporteUs float,
             @Marca char(1) ,
             @Lote1 int ,
             @Canti1 float ,
             @Lote2 int ,
             @Canti2 float, 
             @Lote3 int,
             @Canti3 float,
             @Lote4 int ,
             @Canti4 float, 
             @Lote5 int,
             @Canti5 float   ,
	@TipoProDy char(1)  ,
	@ArticuloDy char(10)


 AS
	INSERT INTO  Estadistica
			(
		             Clave,
		             Tipo,
		             Numero,
		             Renglon,
		             Articulo,
		             Cantidad,
		             Precio,
		             PrecioUS,
		             Importe,
		             ImporteUs,
		             Cliente,
		             Paridad,
		             Vendedor,
		             Rubro,
		             Linea,
		             Costo1,
		             Costo2,
		             Coeficiente,
		             Pedido,
		             Fecha,
		             Importe1,
		             Importe2,
		             Importe3,
		             Importe4,
		             Ordfecha,
		             WArticulo,
		             Remito,
		             WDate,
		             WCantidad,
		             WImporte,
		             WImporteUs,
		             Marca ,
		             Lote1,
		             Canti1,
		             Lote2,
		             Canti2,
		             Lote3,
		             Canti3,
		             Lote4,
		             Canti4,
		             Lote5,
		             Canti5  ,
			TipoProDy   ,
			ArticuloDy
			)

VALUES
			(
		             @Clave,
		             @Tipo,
		             @Numero,
		             @Renglon,
		             @Articulo,
		             @Cantidad,
		             @Precio,
		             @PrecioUS,
		             @Importe,
		             @ImporteUS,
		             @Cliente,
		             @Paridad,
		             @Vendedor,
		             @Rubro,
		             @Linea,
		             @Costo1,
		             @Costo2,
		             @Coeficiente,
		             @Pedido,
		             @Fecha,
		             @Importe1,
		             @Importe2,
		             @Importe3,
		             @Importe4,
		             @Ordfecha,
		             @WArticulo,
		             @Remito,
		             @WDate,
		             @WCantidad,
		             @WImporte,
		             @WImporteUs,
		             @Marca,
		             @Lote1,
		             @Canti1,
		             @Lote2,
		             @Canti2,
		             @Lote3,
		             @Canti3,
		             @Lote4,
		             @Canti4,
		             @Lote5,
		             @Canti5  ,
			@TipoProDy  ,
			@ArticuloDy
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaEspeCli]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEspeCli]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaEspeCli]
             @Cliente char(6),
	@Terminado char(12),
	@Especificaciones char(30)

 AS
	INSERT INTO  EspeCli
			(
			Cliente   ,
		             Terminado       ,                                     
			Especificaciones
			)

VALUES
			(
			@Cliente   ,
		             @Terminado       ,                                     
			@Especificaciones
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaEspecificaciones]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEspecificaciones]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaEspecificaciones]
             @Producto char(10),
	@Ensayo1 int,
	@Valor1 char(50),
	@Ensayo2 int,
	@Valor2 char(50),
	@Ensayo3 int,
	@Valor3 char(50),
	@Ensayo4 int,
	@Valor4 char(50),
	@Ensayo5 int,
	@Valor5 char(50),
	@Ensayo6 int,
	@Valor6 char(50),
	@Ensayo7 int,
	@Valor7 char(50),
	@Ensayo8 int,
	@Valor8 char(50),
	@Ensayo9 int,
	@Valor9 char(50),
	@Ensayo10 int,
	@Valor10 char(50),
	@WDate char(10)

 AS
	INSERT INTO  Especificaciones
			(
			Producto   ,
		             Ensayo1        ,                                     
		             Valor1        ,                                     
		             Ensayo2        ,                                     
		             Valor2        ,                                     
		             Ensayo3        ,                                     
		             Valor3        ,                                     
		             Ensayo4        ,                                     
		             Valor4        ,                                     
		             Ensayo5        ,                                     
		             Valor5        ,                                     
		             Ensayo6        ,                                     
		             Valor6        ,                                     
		             Ensayo7        ,                                     
		             Valor7        ,                                     
		             Ensayo8        ,                                     
		             Valor8        ,                                     
		             Ensayo9        ,                                     
		             Valor9       ,                                     
		             Ensayo10        ,                                     
		             Valor10        ,                                     
			WDate
			)

VALUES
			(
		             @Producto,
			@Ensayo1,
			@Valor1,
			@Ensayo2,
			@Valor2,
			@Ensayo3,
			@Valor3,
			@Ensayo4,
			@Valor4,
			@Ensayo5,
			@Valor5,
			@Ensayo6,
			@Valor6,
			@Ensayo7,
			@Valor7,
			@Ensayo8,
			@Valor8,
			@Ensayo9,
			@Valor9,
			@Ensayo10,
			@Valor10,
			@WDate
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaEspecif]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEspecif]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaEspecif]
             @Producto char(12),
	@Ensayo1 int,
	@Valor1 char(100),
	@Ensayo2 int,
	@Valor2 char(100),
	@Ensayo3 int,
	@Valor3 char(100),
	@Ensayo4 int,
	@Valor4 char(100),
	@Ensayo5 int,
	@Valor5 char(100),
	@Ensayo6 int,
	@Valor6 char(100),
	@Ensayo7 int,
	@Valor7 char(100),
	@Ensayo8 int,
	@Valor8 char(100),
	@Ensayo9 int,
	@Valor9 char(100),
	@Ensayo10 int,
	@Valor10 char(100),
	@WDate char(10)

 AS
	INSERT INTO  Especif
			(
			Producto   ,
		             Ensayo1        ,                                     
		             Valor1        ,                                     
		             Ensayo2        ,                                     
		             Valor2        ,                                     
		             Ensayo3        ,                                     
		             Valor3        ,                                     
		             Ensayo4        ,                                     
		             Valor4        ,                                     
		             Ensayo5        ,                                     
		             Valor5        ,                                     
		             Ensayo6        ,                                     
		             Valor6        ,                                     
		             Ensayo7        ,                                     
		             Valor7        ,                                     
		             Ensayo8        ,                                     
		             Valor8        ,                                     
		             Ensayo9        ,                                     
		             Valor9       ,                                     
		             Ensayo10        ,                                     
		             Valor10        ,                                     
			WDate
			)

VALUES
			(
		             @Producto,
			@Ensayo1,
			@Valor1,
			@Ensayo2,
			@Valor2,
			@Ensayo3,
			@Valor3,
			@Ensayo4,
			@Valor4,
			@Ensayo5,
			@Valor5,
			@Ensayo6,
			@Valor6,
			@Ensayo7,
			@Valor7,
			@Ensayo8,
			@Valor8,
			@Ensayo9,
			@Valor9,
			@Ensayo10,
			@Valor10,
			@WDate
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaEnvase]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEnvase]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaEnvase]
             @Envases int,
	@Descripcion char(50),
	@Abreviatura char(10),
	@Kilos int

 AS
	INSERT INTO  Envases
			(
			Envases   ,
		             Descripcion        ,                                     
			Abreviatura      ,                                    
			Kilos
			)

VALUES
			(
		             @Envases,
			@Descripcion,
			@Abreviatura,
			@Kilos
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaEntdev]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEntdev]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaEntdev]
             @Clave char(8),
	@Codigo int,
	@Renglon int,
	@Fecha char(10),
	@Terminado char(12),
	@Cantidad float,
	@Fechaord char(8),
	@Observaciones char(50),
	@Marca char(1) ,
	@Lote int  ,
	@Cliente char(6)  ,
	@Saldo  float  ,
	@Laboratorio   float


 AS
	INSERT INTO  Entdev
			(
		             Clave ,
			Codigo ,
			Renglon ,
			Fecha ,
			Terminado ,
			Cantidad ,
			Fechaord ,
			Observaciones ,
			Marca  ,
			Lote  ,
			Cliente   ,
			Saldo   ,
			Laboratorio
			)

VALUES
			(
		             @Clave ,
			@Codigo ,
			@Renglon ,
			@Fecha ,
			@Terminado ,
			@Cantidad ,
			@Fechaord ,
			@Observaciones ,
			@Marca  ,
			@Lote   ,
			@Cliente   ,
			@Saldo   ,
			@Laboratorio
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaEnsayos]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaEnsayos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaEnsayos]
             @Codigo int,
	@Descripcion char(50),
	@WDate char(10)

 AS
	INSERT INTO  Ensayos
			(
			Codigo   ,
		             Descripcion        ,                                     
			WDate
			)

VALUES
			(
		             @Codigo,
			@Descripcion,
			@WDate
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaDevcon]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaDevcon]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaDevcon]
             @Clave char(8),
	@Numero int,
	@Renglon int,
	@Cliente char(6),
	@Fecha char(10),
	@Observaciones char(50),
	@Terminado char(12),
	@Cantidad float,
	@Fechaord char(8),
	@Precio float,
	@Linea int,
	@Importe int,
	@Remito int,
	@Lote int


 AS
	INSERT INTO  Devcon
			(
		             Clave ,
			Numero ,
			Renglon ,
			Cliente ,
			Fecha ,
			Observaciones ,
			Terminado ,
			Cantidad ,
			Fechaord ,
			Precio ,
			Linea ,
			Importe ,
			Remito ,
			Lote
			)

VALUES
			(
		             @Clave ,
			@Numero ,
			@Renglon ,
			@Cliente ,
			@Fecha ,
			@Observaciones ,
			@Terminado ,
			@Cantidad ,
			@Fechaord ,
			@Precio ,
			@Linea ,
			@Importe ,
			@Remito ,
			@Lote
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaDesccomp]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaDesccomp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaDesccomp]
             @Clave char(12),
             @Tipo char(2),
             @Numero char(8),
             @Renglon char(2),
             @Descripcion char(50),
             @Importe float,
             @Empresa smallint,
             @WDate char(10)


 AS
	INSERT INTO  Desccomp
			(
		             Clave,
		             Tipo,
		             Numero,
		             Renglon,
		             Descripcion,
		             Importe,
		             Empresa,
		             WDate
			)

VALUES
			(
		             @Clave,
		             @Tipo,
		             @Numero,
		             @Renglon,
		             @Descripcion,
		             @Importe,
		             @Empresa,
		             @WDate
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaDepositos]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaDepositos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE procedure [dbo].[AltaDepositos]
	@Clave    VarChar(8),
	@Deposito VarChar(6),
	@Renglon  varchar(2),
	@Banco    smallint,
	@Fecha    VarChar(10),  
	@FechaOrd VarChar(8),
	@Importe  Float,                                             
	@Acredita VarChar(10),  
	@AcreditaOrd VarChar(8),
	@Tipo2    VarChar(2),
	@Numero2  VarChar(8),
	@Fecha2   VarChar(10),  
	@Importe2 real,                
	@Observaciones2 VarChar(20),       
	@Empresa  smallint,
	@Impolista float
AS
INSERT INTO
	Depositos
		(
		Clave    ,
		Deposito ,
		Renglon ,
		Banco  ,
		Fecha   ,  
		FechaOrd ,
		Importe   ,                                           
		Acredita   ,
		AcreditaOrd ,
		Tipo2 ,
		Numero2, 
		Fecha2  ,  
		Importe2 ,               
		Observaciones2,      
		Empresa ,
		Impolista
		)
VALUES
		(
		@Clave    ,
		@Deposito ,
		@Renglon ,
		@Banco  ,
		@Fecha   ,   
		@FechaOrd ,
		@Importe   ,                                            
		@Acredita   ,
		@AcreditaOrd ,
		@Tipo2 ,
		@Numero2,  
		@Fecha2  ,   
		@Importe2 ,                
		@Observaciones2,       
		@Empresa ,
		@Impolista
		)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaCuenta]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCuenta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaCuenta]
             @Cuenta varchar(20),
	@Descripcion varchar(100),
	@Nivel int,
	@Empresa int
 AS
	INSERT INTO  Cuenta
			(
			Cuenta     ,
			Descripcion ,                                       
			Nivel  ,
			Empresa 
			)
VALUES
			(
		        @Cuenta,
			@Descripcion,
			@Nivel,
			@Empresa
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaCtaPrv]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCtaPrv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaCtaPrv]
	@Clave varchar(26),                     
	@Proveedor varchar(11),   
	@Letra varchar(1), 
	@Tipo varchar(2),
	@Punto varchar(4),
	@Numero varchar(8),  
	@fecha   varchar(10),   
	@Estado varchar(1),
	@Vencimiento varchar(10), 
	@Vencimiento1 varchar(50),                                       
	@Total   float    ,                                          
	@Saldo   float    ,                                          
	@OrdFecha  varchar(8),
	@OrdVencimiento varchar(8), 
	@Impre varchar(2),
	@Empresa smallint, 
	@SaldoList float,                                             
	@NroInterno int,  
	@Lista  varchar(1),
	@Acumulado float   ,
	@Paridad   float   ,
	@Pago  int
AS
INSERT INTO CtaCtePrv
	(
		Clave ,                     
		Proveedor,   
		Letra, 
		Tipo ,
		Punto ,
		Numero ,  
		fecha   ,   
		Estado ,
		Vencimiento, 
		Vencimiento1,                                       
		Total       ,                                          
		Saldo       ,                                          
		OrdFecha ,
		OrdVencimiento, 
		Impre ,
		Empresa, 
		SaldoList,                                             
		NroInterno,  
		Lista ,
		Acumulado   ,
		Paridad   ,
		Pago
	)	
VALUES
	(
		@Clave ,                     
		@Proveedor ,   
		@Letra , 
		@Tipo,
		@Punto,
		@Numero,  
		@fecha,   
		@Estado,
		@Vencimiento, 
		@Vencimiento1,                                       
		@Total,                                          
		@Saldo,                                          
		@OrdFecha,
		@OrdVencimiento, 
		@Impre,
		@Empresa, 
		@SaldoList,                                             
		@NroInterno,  
		@Lista,
		@Acumulado   ,
		@Paridad   ,
		@Pago
	)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaCtacteVarios]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCtacteVarios]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaCtacteVarios]
             @Clave char(12),
             @Tipo char(2),
             @Numero int,
             @Renglon char(2),
             @Cliente char(6),
             @Fecha char(10),
             @Estado char(1),
             @Vencimiento char(10),
             @Vencimiento1 char(10),
             @Total float,
             @TotalUs float,
             @Saldo float,
             @SaldoUs float,
             @Ordfecha char(8),
             @Ordvencimiento char(8),
             @Ordvencimiento1 char(8),
             @Impre char(2),
             @Empresa smallint,
             @Neto float,
             @Iva1 float,
             @Iva2 float,
             @Pedido char(6),
             @Remito char(10),
             @Orden char(10),
             @Paridad Float,
             @Provincia char(2),
             @Vendedor int,
             @Rubro int,
             @Comprobante char(8),
             @Aceptada char(1),
             @Costo float,
             @Importe1 float,
             @Importe2 float,
             @Importe3 float,
             @Importe4 float,
             @Importe5 float,
             @Importe6 float,
             @Importe7 float,
             @WDate char(10) ,
	@Seguro float  ,
	@Flete float  ,
	@ImpoIb  float  ,
	@NroFactura  int ,
	@NroRecibo  int



 AS
	INSERT INTO  Ctacte
			(
		             Clave,
		             Tipo,
		             Numero,
		             Renglon,
		             Cliente,
		             Fecha,
		             Estado,
		             Vencimiento,
		             Vencimiento1,
		             Total,
		             TotalUs,
		             Saldo,
		             SaldoUs,
		             Ordfecha,
		             Ordvencimiento,
		             Ordvencimiento1,
		             Impre,
		             Empresa,
		             Neto,
		             Iva1,
		             Iva2,
		             Pedido,
		             Remito,
		             Orden,
		             Paridad,
		             Provincia,
		             Vendedor,
		             Rubro,
		             Comprobante,
		             Aceptada,
		             Costo,
		             Importe1,
		             Importe2,
		             Importe3,
		             Importe4,
		             Importe5,
		             Importe6,
		             Importe7,
		             WDate  ,
			Seguro  ,
			Flete   ,
			ImpoIb ,
			NroFactura  ,
			NroRecibo
			)			

VALUES
			(
		             @Clave,
		             @Tipo,
		             @Numero,
		             @Renglon,
		             @Cliente,
		             @Fecha,
		             @Estado,
		             @Vencimiento,
		             @Vencimiento1,
		             @Total,
		             @TotalUs,
		             @Saldo,
		             @SaldoUs,
		             @Ordfecha,
		             @Ordvencimiento,
		             @Ordvencimiento1,
		             @Impre,
		             @Empresa,
		             @Neto,
		             @Iva1,
		             @Iva2,
		             @Pedido,
		             @Remito,
		             @Orden,
		             @Paridad,
		             @Provincia,
		             @Vendedor,
		             @Rubro,
		             @Comprobante,
		             @Aceptada,
		             @Costo,
		             @Importe1,
		             @Importe2,
		             @Importe3,
		             @Importe4,
		             @Importe5,
		             @Importe6,
		             @Importe7,
		             @WDate  ,
			@Seguro  ,
			@Flete  ,
			@ImpoIb  ,
			@NroFactura  ,
			@NroRecibo
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaCtaCtePrv]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCtaCtePrv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaCtaCtePrv]
	@Clave varchar(26),                     
	@Proveedor varchar(11),   
	@Letra varchar(1), 
	@Tipo varchar(2),
	@Punto varchar(4),
	@Numero varchar(8),  
	@fecha   varchar(10),   
	@Estado varchar(1),
	@Vencimiento varchar(10), 
	@Vencimiento1 varchar(50),                                       
	@Total   float    ,                                          
	@Saldo   float    ,                                          
	@OrdFecha  varchar(8),
	@OrdVencimiento varchar(8), 
	@Impre varchar(2),
	@Empresa smallint, 
	@SaldoList float,                                             
	@NroInterno int,  
	@Lista  varchar(1),
	@Acumulado float
AS
INSERT INTO CtaCtePrv
	(
		Clave ,                     
		Proveedor,   
		Letra, 
		Tipo ,
		Punto ,
		Numero ,  
		fecha   ,   
		Estado ,
		Vencimiento, 
		Vencimiento1,                                       
		Total       ,                                          
		Saldo       ,                                          
		OrdFecha ,
		OrdVencimiento, 
		Impre ,
		Empresa, 
		SaldoList,                                             
		NroInterno,  
		Lista ,
		Acumulado
	)	
VALUES
	(
		@Clave ,                     
		@Proveedor ,   
		@Letra , 
		@Tipo,
		@Punto,
		@Numero,  
		@fecha,   
		@Estado,
		@Vencimiento, 
		@Vencimiento1,                                       
		@Total,                                          
		@Saldo,                                          
		@OrdFecha,
		@OrdVencimiento, 
		@Impre,
		@Empresa, 
		@SaldoList,                                             
		@NroInterno,  
		@Lista,
		@Acumulado
	)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaCtaCteCli]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCtaCteCli]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[AltaCtaCteCli]
	@Clave 		varchar(12),       
	@Tipo 		varchar(2),
	@Numero      	varchar(4),
	@Renglon 	varchar(2),
	@Cliente 	varchar(6),
	@fecha      	varchar(10),
	@Estado 	varchar(1),
	@Vencimiento 	varchar(10),
	@Vencimiento1 	varchar(10),
	@Total          float,                           
	@TotalUs        float,                                       
	@Saldo          float,                                       
	@SaldoUs        float,                                       
	@OrdFecha 	varchar(8),
	@OrdVencimiento varchar(8),
	@OrdVencimiento1 varchar(8),
	@Impre 		varchar(2),
	@Empresa 	integer,
	@Neto           float,                                       
	@Iva1           float,                                       
	@Iva2           float,                                       
	@Pedido 	varchar(6),
	@Remito     	varchar(10),
	@Orden      	varchar(10),
	@Paridad        float,                           
	@Provincia 	varchar(2),
	@Vendedor    	integer,
	@Rubro       	integer,
	@Comprobante 	varchar(8),
	@Aceptada 	varchar(1),
	@Costo          float,                            
	@Importe1       float,                                       
	@Importe2       float,                                       
	@Importe3       float,                                       
	@Importe4       float,                                       
	@Importe5       float,                                       
	@Importe6       float,                                       
	@Importe7       float,                                       
	@WDate     	varchar(10) ,
	@Seguro float  ,
	@Flete float  ,
	@ImpoIb  float




AS
INSERT INTO
	CtaCte
		(
		Clave        ,
		Tipo ,
		Numero,      
		Renglon, 
		Cliente ,
		fecha    ,  
		Estado ,
		Vencimiento,
		Vencimiento1,
		Total        ,                                        
		TotalUs       ,                                       
		Saldo          ,                                      
		SaldoUs         ,                                     
		OrdFecha ,
		OrdVencimiento,
		OrdVencimiento1,
		Impre ,
		Empresa,
		Neto    ,                                             
		Iva1     ,                                            
		Iva2      ,                                           
		Pedido ,
		Remito  ,  
		Orden    , 
		Paridad   ,                                           
		Provincia ,
		Vendedor   ,
		Rubro       ,
		Comprobante ,
		Aceptada ,
		Costo     ,                                           
		Importe1   ,                                          
		Importe2    ,                                         
		Importe3     ,                                        
		Importe4      ,                                       
		Importe5       ,                                      
		Importe6        ,                                     
		Importe7         ,                                    
		WDate        ,
		Seguro  ,
		Flete   ,
		ImpoIb
		)
VALUES
		(
		@Clave        ,
		@Tipo ,
		@Numero,      
		@Renglon, 
		@Cliente ,
		@fecha    ,  
		@Estado ,
		@Vencimiento ,
		@Vencimiento1 ,
		@Total         ,                                        
		@TotalUs        ,                                       
		@Saldo           ,                                      
		@SaldoUs          ,                                     
		@OrdFecha ,
		@OrdVencimiento ,
		@OrdVencimiento1 ,
		@Impre ,
		@Empresa, 
		@Neto    ,                                              
		@Iva1     ,                                             
		@Iva2      ,                                            
		@Pedido ,
		@Remito  ,   
		@Orden    ,  
		@Paridad   ,                                            
		@Provincia ,
		@Vendedor   , 
		@Rubro       ,
		@Comprobante ,
		@Aceptada ,
		@Costo     ,                                            
		@Importe1   ,                                           
		@Importe2    ,                                          
		@Importe3     ,                                         
		@Importe4      ,                                        
		@Importe5       ,                                       
		@Importe6        ,                                      
		@Importe7         ,                                     
		@WDate        ,
		@Seguro    ,
		@Flete    ,
		@ImpoIb
		)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaCtacte1]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCtacte1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaCtacte1]
             @Clave char(12),
             @Tipo char(2),
             @Numero int,
             @Renglon char(2),
             @Cliente char(6),
             @Fecha char(10),
             @Estado char(1),
             @Vencimiento char(10),
             @Vencimiento1 char(10),
             @Total float,
             @TotalUs float,
             @Saldo float,
             @SaldoUs float,
             @Ordfecha char(8),
             @Ordvencimiento char(8),
             @Ordvencimiento1 char(8),
             @Impre char(2),
             @Empresa smallint,
             @Neto float,
             @Iva1 float,
             @Iva2 float,
             @Pedido char(6),
             @Remito char(10),
             @Orden char(10),
             @Paridad Float,
             @Provincia char(2),
             @Vendedor int,
             @Rubro int,
             @Comprobante char(8),
             @Aceptada char(1),
             @Costo float,
             @Importe1 float,
             @Importe2 float,
             @Importe3 float,
             @Importe4 float,
             @Importe5 float,
             @Importe6 float,
             @Importe7 float,
             @WDate char(10) ,
	@Seguro float  ,
	@Flete float  ,
	@ImpoIb  float


 AS
	INSERT INTO  Ctacte
			(
		             Clave,
		             Tipo,
		             Numero,
		             Renglon,
		             Cliente,
		             Fecha,
		             Estado,
		             Vencimiento,
		             Vencimiento1,
		             Total,
		             TotalUs,
		             Saldo,
		             SaldoUs,
		             Ordfecha,
		             Ordvencimiento,
		             Ordvencimiento1,
		             Impre,
		             Empresa,
		             Neto,
		             Iva1,
		             Iva2,
		             Pedido,
		             Remito,
		             Orden,
		             Paridad,
		             Provincia,
		             Vendedor,
		             Rubro,
		             Comprobante,
		             Aceptada,
		             Costo,
		             Importe1,
		             Importe2,
		             Importe3,
		             Importe4,
		             Importe5,
		             Importe6,
		             Importe7,
		             WDate  ,
			Seguro  ,
			Flete  ,
			ImpoIb
			)			

VALUES
			(
		             @Clave,
		             @Tipo,
		             @Numero,
		             @Renglon,
		             @Cliente,
		             @Fecha,
		             @Estado,
		             @Vencimiento,
		             @Vencimiento1,
		             @Total,
		             @TotalUs,
		             @Saldo,
		             @SaldoUs,
		             @Ordfecha,
		             @Ordvencimiento,
		             @Ordvencimiento1,
		             @Impre,
		             @Empresa,
		             @Neto,
		             @Iva1,
		             @Iva2,
		             @Pedido,
		             @Remito,
		             @Orden,
		             @Paridad,
		             @Provincia,
		             @Vendedor,
		             @Rubro,
		             @Comprobante,
		             @Aceptada,
		             @Costo,
		             @Importe1,
		             @Importe2,
		             @Importe3,
		             @Importe4,
		             @Importe5,
		             @Importe6,
		             @Importe7,
		             @WDate   ,
			@Seguro  ,
			@Flete   ,
			@ImpoIb
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaCtacte]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCtacte]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaCtacte]
             @Clave char(12),
             @Tipo char(2),
             @Numero int,
             @Renglon char(2),
             @Cliente char(6),
             @Fecha char(10),
             @Estado char(1),
             @Vencimiento char(10),
             @Vencimiento1 char(10),
             @Total float,
             @TotalUs float,
             @Saldo float,
             @SaldoUs float,
             @Ordfecha char(8),
             @Ordvencimiento char(8),
             @Ordvencimiento1 char(8),
             @Impre char(2),
             @Empresa smallint,
             @Neto float,
             @Iva1 float,
             @Iva2 float,
             @Pedido char(6),
             @Remito char(10),
             @Orden char(10),
             @Paridad Float,
             @Provincia char(2),
             @Vendedor int,
             @Rubro int,
             @Comprobante char(8),
             @Aceptada char(1),
             @Costo float,
             @Importe1 float,
             @Importe2 float,
             @Importe3 float,
             @Importe4 float,
             @Importe5 float,
             @Importe6 float,
             @Importe7 float,
             @WDate char(10) ,
	@Seguro float  ,
	@Flete float  ,
	@ImpoIb  float



 AS
	INSERT INTO  Ctacte
			(
		             Clave,
		             Tipo,
		             Numero,
		             Renglon,
		             Cliente,
		             Fecha,
		             Estado,
		             Vencimiento,
		             Vencimiento1,
		             Total,
		             TotalUs,
		             Saldo,
		             SaldoUs,
		             Ordfecha,
		             Ordvencimiento,
		             Ordvencimiento1,
		             Impre,
		             Empresa,
		             Neto,
		             Iva1,
		             Iva2,
		             Pedido,
		             Remito,
		             Orden,
		             Paridad,
		             Provincia,
		             Vendedor,
		             Rubro,
		             Comprobante,
		             Aceptada,
		             Costo,
		             Importe1,
		             Importe2,
		             Importe3,
		             Importe4,
		             Importe5,
		             Importe6,
		             Importe7,
		             WDate  ,
			Seguro  ,
			Flete   ,
			ImpoIb  
			)			

VALUES
			(
		             @Clave,
		             @Tipo,
		             @Numero,
		             @Renglon,
		             @Cliente,
		             @Fecha,
		             @Estado,
		             @Vencimiento,
		             @Vencimiento1,
		             @Total,
		             @TotalUs,
		             @Saldo,
		             @SaldoUs,
		             @Ordfecha,
		             @Ordvencimiento,
		             @Ordvencimiento1,
		             @Impre,
		             @Empresa,
		             @Neto,
		             @Iva1,
		             @Iva2,
		             @Pedido,
		             @Remito,
		             @Orden,
		             @Paridad,
		             @Provincia,
		             @Vendedor,
		             @Rubro,
		             @Comprobante,
		             @Aceptada,
		             @Costo,
		             @Importe1,
		             @Importe2,
		             @Importe3,
		             @Importe4,
		             @Importe5,
		             @Importe6,
		             @Importe7,
		             @WDate  ,
			@Seguro  ,
			@Flete  ,
			@ImpoIb
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaCotizaII]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCotizaII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaCotizaII]
             @Clave char(8),
	@Cotiza int,
	@Renglon int,
	@Fecha char(10),
	@Proveedor char(11),
	@Articulo char(10),
	@Precio float,
	@Condicion char(40),
	@Observaciones char(40),
	@Fechaord char(8),
	@WDate char(10),
	@Moneda int


 AS
	INSERT INTO  Cotiza
			(
			Clave   ,
		             Cotiza        ,                                     
			Renglon      ,                                    
		             Fecha       ,                                     
			Proveedor      ,                                    
		             Articulo        ,                                     
			Precio      ,                                    
		             Condicion        ,                                     
			Observaciones      ,                                    
		             Fechaord        ,                                     
			WDate      ,
			Moneda
			)

VALUES
			(
			@Clave   ,
		             @Cotiza        ,                                     
			@Renglon      ,                                    
		             @Fecha       ,                                     
			@Proveedor      ,                                    
		             @Articulo        ,                                     
			@Precio      ,                                    
		             @Condicion        ,                                     
			@Observaciones      ,                                    
		             @Fechaord        ,                                     
			@WDate      ,
			@Moneda
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaCotiza]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCotiza]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaCotiza]
             @Clave char(8),
	@Cotiza int,
	@Renglon int,
	@Fecha char(10),
	@Proveedor char(11),
	@Articulo char(10),
	@Precio float,
	@Condicion char(40),
	@Observaciones char(40),
	@Fechaord char(8),
	@WDate char(10)


 AS
	INSERT INTO  Cotiza
			(
			Clave   ,
		             Cotiza        ,                                     
			Renglon      ,                                    
		             Fecha       ,                                     
			Proveedor      ,                                    
		             Articulo        ,                                     
			Precio      ,                                    
		             Condicion        ,                                     
			Observaciones      ,                                    
		             Fechaord        ,                                     
			WDate      
			)

VALUES
			(
			@Clave   ,
		             @Cotiza        ,                                     
			@Renglon      ,                                    
		             @Fecha       ,                                     
			@Proveedor      ,                                    
		             @Articulo        ,                                     
			@Precio      ,                                    
		             @Condicion        ,                                     
			@Observaciones      ,                                    
		             @Fechaord        ,                                     
			@WDate      
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaConsig]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaConsig]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaConsig]
             @Clave char(8),
	@Numero int,
	@Renglon int,
	@Cliente char(6),
	@Fecha char(10),
	@Observaciones char(50),
	@Terminado char(12),
	@Cantidad float,
	@Envase1 int,
	@Canti1 int,
	@Envase2 int,
	@Canti2 int,
	@Envase3 int,
	@Canti3 int,
	@Envase4 int,
	@Canti4 int,
	@Fechaord char(8),
	@Precio float,
	@Linea int,
	@Facturado int,
	@Importe int,
	@Marca char(1),
	@Lote int


 AS
	INSERT INTO  Consig
			(
		             Clave ,
			Numero ,
			Renglon ,
			Cliente ,
			Fecha ,
			Observaciones ,
			Terminado ,
			Cantidad ,
			Envase1 ,
			Canti1 ,
			Envase2 ,
			Canti2 ,
			Envase3 ,
			Canti3 ,
			Envase4 ,
			Canti4 ,
			Fechaord ,
			Precio ,
			Linea ,
			Facturado ,
			Importe ,
			Marca,
			Lote
			)

VALUES
			(
		             @Clave ,
			@Numero ,
			@Renglon ,
			@Cliente ,
			@Fecha ,
			@Observaciones ,
			@Terminado ,
			@Cantidad ,
			@Envase1 ,
			@Canti1 ,
			@Envase2 ,
			@Canti2 ,
			@Envase3 ,
			@Canti3 ,
			@Envase4 ,
			@Canti4 ,
			@Fechaord ,
			@Precio ,
			@Linea ,
			@Facturado ,
			@Importe ,
			@Marca,
			@Lote
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaComposicion]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaComposicion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaComposicion]
             @Clave char(14),
	@Terminado char(12),
	@Renglon char(2),
	@Tipo char(1),
	@Articulo1 char(10),
	@Articulo2 char(12),
	@Cantidad float,
	@WDate char(10),
	@Costo1 float,
	@Costo2 float


 AS
	INSERT INTO  Composicion
			(
		             Clave,
			Terminado,
			Renglon,
			Tipo,
			Articulo1,
			Articulo2,
			Cantidad,
			WDate,
			Costo1,
			Costo2
			)

VALUES
			(
		             @Clave,
			@Terminado,
			@Renglon,
			@Tipo,
			@Articulo1,
			@Articulo2,
			@Cantidad,
			@WDate,
			@Costo1,
			@Costo2
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaCliente1]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCliente1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaCliente1]
             @Cliente char(6),
	@Razon char(50),
	@Direccion char(50),
	@Localidad char(50),
	@Provincia char(2),	
	@Postal char(10),	
	@EMail char(40),	
	@Fax char(20),
	@Telefono char(40),
	@Cuit char(15),
	@Contacto char(50),
	@Observaciones char(100),	
	@Vendedor int,
	@Iva char(1),
	@Rubro int,
	@Horario char(20),
	@Pago1 int,
	@Pago2 int,
	@Limite float,
	@Minimo float,
	@Direntrega char(50),
	@Importe1 float,
	@Importe2 float,
	@Importe3 float,
	@Importe4 float,
	@Importe5 float,
	@Importe6 float,
	@WDate char(10)  ,
	@Precio char(2)  ,
	@Ib  int


 AS
	INSERT INTO  Cliente
			(
		             Cliente,
			Razon,
			Direccion,
			Localidad,
			Provincia ,	
			Postal,	
			EMail,	
			Fax ,
			Telefono ,
			Cuit ,
			Contacto ,
			Observaciones ,	
			Vendedor ,
			Iva ,
			Rubro ,
			Horario ,
			Pago1 ,
			Pago2 ,
			Limite ,
			Minimo ,
			Direntrega ,
			Importe1 ,
			Importe2 ,
			Importe3 ,
			Importe4 ,
			Importe5 ,
			Importe6 ,
			WDate    ,
			Precio  ,
			Ib
			)

VALUES
			(
		             @Cliente,
			@Razon,
			@Direccion,
			@Localidad,
			@Provincia ,	
			@Postal,	
			@EMail,	
			@Fax ,
			@Telefono ,
			@Cuit ,
			@Contacto ,
			@Observaciones ,	
			@Vendedor ,
			@Iva ,
			@Rubro ,
			@Horario ,
			@Pago1 ,
			@Pago2 ,
			@Limite ,
			@Minimo ,
			@Direntrega ,
			@Importe1 ,
			@Importe2 ,
			@Importe3 ,
			@Importe4 ,
			@Importe5 ,
			@Importe6 ,
			@WDate   ,
			@Precio  ,
			@Ib
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaCliente]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCliente]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaCliente]
             @Cliente char(6),
	@Razon char(50),
	@Direccion char(50),
	@Localidad char(50),
	@Provincia char(2),	
	@Postal char(10),	
	@EMail char(40),	
	@Fax char(20),
	@Telefono char(40),
	@Cuit char(15),
	@Contacto char(50),
	@Observaciones char(100),	
	@Vendedor int,
	@Iva char(1),
	@Rubro int,
	@Horario char(20),
	@Pago1 int,
	@Pago2 int,
	@Limite float,
	@Minimo float,
	@Direntrega char(50),
	@Importe1 float,
	@Importe2 float,
	@Importe3 float,
	@Importe4 float,
	@Importe5 float,
	@Importe6 float,
	@WDate char(10) ,
	@Precio char(2)  ,
	@Ib  int


 AS
	INSERT INTO  Cliente
			(
		             Cliente,
			Razon,
			Direccion,
			Localidad,
			Provincia ,	
			Postal,	
			EMail,	
			Fax ,
			Telefono ,
			Cuit ,
			Contacto ,
			Observaciones ,	
			Vendedor ,
			Iva ,
			Rubro ,
			Horario ,
			Pago1 ,
			Pago2 ,
			Limite ,
			Minimo ,
			Direntrega ,
			Importe1 ,
			Importe2 ,
			Importe3 ,
			Importe4 ,
			Importe5 ,
			Importe6 ,
			WDate   , 
			Precio  ,
			Ib
			)

VALUES
			(
		             @Cliente,
			@Razon,
			@Direccion,
			@Localidad,
			@Provincia ,	
			@Postal,	
			@EMail,	
			@Fax ,
			@Telefono ,
			@Cuit ,
			@Contacto ,
			@Observaciones ,	
			@Vendedor ,
			@Iva ,
			@Rubro ,
			@Horario ,
			@Pago1 ,
			@Pago2 ,
			@Limite ,
			@Minimo ,
			@Direntrega ,
			@Importe1 ,
			@Importe2 ,
			@Importe3 ,
			@Importe4 ,
			@Importe5 ,
			@Importe6 ,
			@WDate   ,
			@Precio  ,
			@Ib
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaCarpeta]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCarpeta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaCarpeta]
             @Carpeta int,
	@Articulo char(10) ,
	@Cantidad float ,
	@CostoFlete float ,
	@Importe float  ,
	@Arancel  float  ,
	@Costo float  ,
	@Gastos  float  ,
	@Precio float  ,
	@Coeficiente   float  ,
	@Clave char(8)  ,
	@Leyenda int

 AS
	INSERT INTO  Carpeta
			(
			Carpeta   ,
		             Articulo       ,                                     
			Cantidad      ,                                    
		             CostoFlete      ,                                     
			Importe      ,                                    
		             Arancel        ,                                     
			Costo      ,                                    
		             Gastos        ,                                     
			Precio     ,
			Coeficiente  ,
			Clave  ,
			Leyenda
			)

VALUES
			(
			@Carpeta   ,
		             @Articulo       ,                                     
			@Cantidad      ,                                    
		             @CostoFlete      ,                                     
			@Importe      ,                                    
		             @Arancel        ,                                     
			@Costo      ,                                    
		             @Gastos       ,                                     
			@Precio   ,
			@Coeficiente  ,
			@Clave  ,
			@Leyenda
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaCambioAdm]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCambioAdm]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaCambioAdm]

             @Fecha char(10),
             	@Cambio float,
	@Ordfecha char(10)


 AS
	INSERT INTO  CambioAdm
			(
			Fecha   ,
			Cambio  ,
		             OrdFecha
			)

VALUES
			(
		             @Fecha,
		             @Cambio,
			@OrdFecha
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaCambio]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaCambio]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaCambio]

             @Fecha char(10),
             	@Cambio float,
	@Ordfecha char(10)


 AS
	INSERT INTO  Cambios
			(
			Fecha   ,
			Cambio  ,
		             OrdFecha
			)

VALUES
			(
		             @Fecha,
		             @Cambio,
			@OrdFecha
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[Altabanco]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Altabanco]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[Altabanco]
             @Banco smallint,
	@Nombre varchar(50),
	@Cuenta varchar(10),
	@Empresa smallint
 AS
	INSERT INTO  Banco
			(
			Banco     ,
			Nombre ,                                       
			Cuenta  ,
			Empresa 
			)
VALUES
			(
		        @Banco,
			@Nombre,
			@Cuenta,
			@Empresa
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaAtributos]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaAtributos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaAtributos]
             @Operador  int,
	@Proceso int,
	@Atributo1 char(50),
	@Atributo2 char(50),
	@Atributo3 char(50),
	@Atributo4 char(50),
	@Atributo5 char(50),
	@Atributo6 char(50),
	@Atributo7 char(50),
	@Atributo8 char(50),
	@Atributo9 char(50),
	@Atributo10 char(50)


 AS
	INSERT INTO  Atributos
			(
			Operador     ,
			Proceso ,                                       
			Atributo1 ,                                       
			Atributo2 ,                                       
			Atributo3 ,                                       
			Atributo4 ,                                       
			Atributo5 ,                                       
			Atributo6 ,                                       
			Atributo7 ,                                       
			Atributo8 ,                                       
			Atributo9 ,                                       
			Atributo10
			)
VALUES
			(
			@Operador,
			@Proceso ,
			@Atributo1 ,
			@Atributo2 ,
			@Atributo3 ,
			@Atributo4 ,
			@Atributo5 ,
			@Atributo6 ,
			@Atributo7 ,
			@Atributo8 ,
			@Atributo9 ,
			@Atributo10 
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaArticuloII]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaArticuloII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaArticuloII]
             @Codigo char(10),
	@Descripcion char(50),
	@Costo1 float,
	@Costo2 float,
	@Inicial float,
	@Entradas float,
	@Salidas float,
	@Minimo float,
	@Laboratorio float,
	@Unidad char(10),
	@Pedido float,
	@Deposito char(20),
	@Envase int,
	@Rs char(1),
	@Fecha char(10),
	@Orden int,
	@Dife float,
	@Proveedor char(11),
	@WDate char(10),
	@Flete float,
	@Moneda char(3),
	@Controla int, 
	@Densidad char(20) ,
	@Costo3 float  ,
	@WCosto1 float  ,
	@WCosto2 float  ,
	@WCosto3 float  ,
	@Venta   float

 AS
	INSERT INTO  Articulo
			(
			Codigo   ,
		             Descripcion        ,                                     
			Costo1      ,                                    
		             Costo2        ,                                     
			Inicial      ,                                    
		             Entradas        ,                                     
			Salidas      ,                                    
		             Minimo        ,                                     
			Laboratorio      ,                                    
		             Unidad        ,                                     
			Pedido      ,                                    
			Deposito      ,                                    
		             Envase        ,                                     
			Rs      ,                                    
		             Fecha        ,                                     
			Orden      ,                                    
		             Dife        ,                                     
			Proveedor      ,                                    
		             WDate        ,                                     
			Flete      ,                                    
			Moneda   ,
			Controla  ,
			Densidad  ,
			Costo3  ,
			WCosto1  ,
			WCosto2   ,
			WCosto3   ,
			Venta
			)

VALUES
			(
			@Codigo   ,
		             @Descripcion        ,                                     
			@Costo1      ,                                    
		             @Costo2        ,                                     
			@Inicial      ,                                     
		             @Entradas        ,                                     
			@Salidas      ,                                    
		             @Minimo        ,                                     
			@Laboratorio      ,                                    
		             @Unidad        ,                                     
			@Pedido      ,                                    
			@Deposito      ,                                    
		             @Envase        ,                                     
			@Rs      ,                                    
		             @Fecha        ,                                     
			@Orden      ,                                    
		             @Dife        ,                                     
			@Proveedor      ,                                    
		             @WDate        ,                                     
			@Flete      ,                                    
			@Moneda   ,
			@Controla  ,
			@Densidad  ,
			@Costo3  ,
			@WCosto1   ,
			@WCosto2   ,
			@WCosto3   ,
			@Venta
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[AltaArticulo]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AltaArticulo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[AltaArticulo]
             @Codigo char(10),
	@Descripcion char(50),
	@Costo1 float,
	@Costo2 float,
	@Inicial float,
	@Entradas float,
	@Salidas float,
	@Minimo float,
	@Laboratorio float,
	@Unidad char(10),
	@Pedido float,
	@Deposito char(20),
	@Envase int,
	@Rs char(1),
	@Fecha char(10),
	@Orden int,
	@Dife float,
	@Proveedor char(11),
	@WDate char(10),
	@Flete float,
	@Moneda char(3),
	@Controla int, 
	@Densidad char(20) ,
	@Costo3 float

 AS
	INSERT INTO  Articulo
			(
			Codigo   ,
		             Descripcion        ,                                     
			Costo1      ,                                    
		             Costo2        ,                                     
			Inicial      ,                                    
		             Entradas        ,                                     
			Salidas      ,                                    
		             Minimo        ,                                     
			Laboratorio      ,                                    
		             Unidad        ,                                     
			Pedido      ,                                    
			Deposito      ,                                    
		             Envase        ,                                     
			Rs      ,                                    
		             Fecha        ,                                     
			Orden      ,                                    
		             Dife        ,                                     
			Proveedor      ,                                    
		             WDate        ,                                     
			Flete      ,                                    
			Moneda   ,
			Controla  ,
			Densidad  ,
			Costo3
			)

VALUES
			(
			@Codigo   ,
		             @Descripcion        ,                                     
			@Costo1      ,                                    
		             @Costo2        ,                                     
			@Inicial      ,                                     
		             @Entradas        ,                                     
			@Salidas      ,                                    
		             @Minimo        ,                                     
			@Laboratorio      ,                                    
		             @Unidad        ,                                     
			@Pedido      ,                                    
			@Deposito      ,                                    
		             @Envase        ,                                     
			@Rs      ,                                    
		             @Fecha        ,                                     
			@Orden      ,                                    
		             @Dife        ,                                     
			@Proveedor      ,                                    
		             @WDate        ,                                     
			@Flete      ,                                    
			@Moneda   ,
			@Controla  ,
			@Densidad  ,
			@Costo3
			)' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaSaldoCtaCtePrv]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaSaldoCtaCtePrv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ActualizaSaldoCtaCtePrv]
	@Clave VarChar(26),
	@Saldo Float
AS
UPDATE	CtaCtePrv
SET
	Saldo = @Saldo
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaSaldoCtaCteCli]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaSaldoCtaCteCli]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ActualizaSaldoCtaCteCli]
	@Clave varchar(12),
	@Saldo float,
	@SaldoUs float,
	@Wdate varchar(10)
AS
UPDATE 	CtaCte
SET
	Saldo = @Saldo,
	SaldoUs = @SaldoUS,
	Wdate = @Wdate
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRetencionPagos]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRetencionPagos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ActualizaRetencionPagos]
	@Clave     varchar(15),
	@Neto      float,                                            
	@Retenido  float                                            
AS      
UPDATE
		Retencion
SET 
	Neto    	= @Neto      ,                                                                                    
	Retenido	= @Retenido
WHERE
	Clave 		= @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRetencion]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRetencion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ActualizaRetencion]
	@Clave     varchar(15),
	@Neto      float,                                            
	@Anticipo  float,                                            
	@Bruto     float,                                            
	@Iva       float,
	@Retenido  float                                            
AS      
UPDATE
		Retencion
SET
	Neto    	= @Neto      ,                                                                                    
	Retenido	= @Retenido  ,                                                                                    
	Anticipo	= @Anticipo  ,                                                                                        
	Bruto   	= @Bruto     ,                                                                                      
	Iva     	= @Iva          
WHERE
	Clave 		= @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRecibosSalvaMarcaII]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRecibosSalvaMarcaII]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE procedure [dbo].[ActualizaRecibosSalvaMarcaII]
AS
UPDATE
	Recibos
SET
	Marca = ""

WHERE
	Impolist = 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRecibosSalvaMarca]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRecibosSalvaMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE procedure [dbo].[ActualizaRecibosSalvaMarca]
AS
UPDATE
	Recibos
SET
	Impolist = 1

WHERE
	Marca = "X"' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRecibosOtroVI]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRecibosOtroVI]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE procedure [dbo].[ActualizaRecibosOtroVI]
	@Recibo VarChar(6),
	@Fecha2 char(10)  ,
	@FechaOrd2 char(8)
AS
UPDATE
	Recibos
SET
	Fechadepo = @Fecha2 ,
	FechadepoOrd = @Fechaord2

WHERE
	Recibo = @Recibo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRecibosOtro]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRecibosOtro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE procedure [dbo].[ActualizaRecibosOtro]
	@Clave VarChar(8)
AS
UPDATE
	Recibos
SET
	Estado2 = "X"

WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRecibosMarca]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRecibosMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE procedure [dbo].[ActualizaRecibosMarca]
	@Recibo VarChar(6) ,
	@FechaDepo char(10)  ,
	@FechaDepoOrd char(8),
	@Marca char(1)
AS
UPDATE
	Recibos
SET
	Marca = @Marca  ,
	FechaDepo = @FechaDepo   ,
	FechaDepoOrd =  @FechaDepoOrd

WHERE
	Recibo = @Recibo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaRecibos]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaRecibos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE procedure [dbo].[ActualizaRecibos]
	@Clave VarChar(8),
	@Estado2 VarChar(1),
	@Destino VarChar(50)
AS
UPDATE
	Recibos
SET
	Estado2 = @Estado2,
	Destino = @Destino
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaMovGasMarca]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaMovGasMarca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE procedure [dbo].[ActualizaMovGasMarca]
	@Carpeta  int,
	@Marca char(1)
AS
UPDATE
	MovGas
SET
	Marca = @Marca  

WHERE
	Carpeta = @Carpeta' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaIvaComprasCai]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaIvaComprasCai]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ActualizaIvaComprasCai]
	@NroInterno  int,
	@Cai char(14),
	@VtoCai char(10)
AS
	UPDATE IvaComp
	SET
		Cai = @Cai,
		VtoCai = @VtoCai
WHERE
	NroInterno = @NroInterno' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaIvaCompras]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaIvaCompras]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ActualizaIvaCompras]
	@NroInterno  int,
	@Proveedor char(11),
	@Tipo char(2),
	@Letra char(1),
	@Punto char(4),
	@Numero char(8),  
	@Fecha  char(10),    
	@Vencimiento char(10),
	@Vencimiento1 char(10),
	@Periodo    char(10),
	@Neto   float,                                               
	@Iva21  float,                                               
	@Iva5   float,                                               
	@Iva27  float,                                               
	@Ib     float,                                               
	@Exento float,                                               
	@Contado char(1),
	@Impre char(2),
	@Ordfecha char(8),
	@Empresa smallint,
	@Netolist    float,                                          
	@ExentoList  float     ,
	@Paridad  float  ,
	@Pago int                                   
AS
	UPDATE IvaComp
	SET
		Proveedor = @Proveedor,
		Tipo = @Tipo,
		Letra = @Letra,
		Punto = @Punto,
		Numero = @Numero,    
		Fecha = @Fecha,   
		Vencimiento = @Vencimiento,
		Vencimiento1 = @Vencimiento1,
		Periodo = @Periodo,
		Neto = @Neto,
		Iva21 = @Iva21,
		Iva5 = @Iva5,
		Iva27 = @Iva27,                                               
		Ib = @Ib,                                               
		Exento = @Exento,                                               
		Contado = @Contado,
		Impre = @Impre,
		Ordfecha = @Ordfecha,
		Empresa = @Empresa,
		Netolist = @Netolist,                                                                                       
		ExentoList = @ExentoList      ,
		Paridad = @Paridad  ,
		Pago = @Pago                                      
WHERE
	NroInterno = @NroInterno' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaEstadistica]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaEstadistica]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ActualizaEstadistica]
	AS
UPDATE 	Estadistica
SET
	Empresa = 1' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaCtaPrvSaldo]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaCtaPrvSaldo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ActualizaCtaPrvSaldo]
	@Clave varchar(26),                     
	@Saldo   float 
AS
UPDATE CtaCtePrv
SET	
	Saldo = @Saldo     
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaCtaCteUs2]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaCtaCteUs2]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ActualizaCtaCteUs2]
AS
UPDATE 	CtaCte
SET
	Paridad = Total / TotalUs

where

	TotalUs <> 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaCtaCteUs1]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaCtaCteUs1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ActualizaCtaCteUs1]
AS
UPDATE 	CtaCte
SET
	SaldoUs = 0

where

	TotalUs = 0' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaCtaCteUs]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaCtaCteUs]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ActualizaCtaCteUs]
AS
UPDATE 	CtaCte
SET
	SaldoUs = Saldo' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaCtaCtePrv]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaCtaCtePrv]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ActualizaCtaCtePrv]
	@Clave varchar(26),                     
	@Proveedor varchar(11),   
	@Letra varchar(1), 
	@Tipo varchar(2),
	@Punto varchar(4),
	@Numero varchar(8),  
	@fecha   varchar(10),   
	@Estado varchar(1),
	@Vencimiento varchar(10), 
	@Vencimiento1 varchar(50),                                       
	@Total   float    ,                                          
	@Saldo   float    ,                                          
	@OrdFecha  varchar(8),
	@OrdVencimiento varchar(8), 
	@Impre varchar(2),
	@Empresa smallint, 
	@SaldoList float,                                             
	@NroInterno int,  
	@Lista  varchar(1),
	@Acumulado float
AS
UPDATE CtaCtePrv
SET	
	Proveedor = @Proveedor,   
	Letra = @Letra, 
	Tipo = @Tipo,
	Punto = @Punto,
	Numero = @Punto,  
	Fecha = @Fecha   ,   
	Estado = @Estado,
	Vencimiento = @Vencimiento, 
	Vencimiento1 = @Vencimiento1,                                       
	Total = @Total      ,                                          
	Saldo = @Saldo      ,                                          
	OrdFecha = @OrdFecha ,
	OrdVencimiento = @OrdVencimiento, 
	Impre = @Impre,
	Empresa = @Empresa, 
	SaldoList = @SaldoList,                                             
	NroInterno = @NroInterno,  
	Lista = @Lista,
	Acumulado = @Acumulado
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaCtaCte]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaCtaCte]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ActualizaCtaCte]
	@Clave varchar(12),
	@Saldo float,
	@SaldoUs float,
	@WEstado varchar(8),
	@Wdate varchar(10)
AS
UPDATE 	CtaCte
SET
	Saldo = @Saldo,
	SaldoUs = @SaldoUS,
	Estado = @WEstado,
	Wdate = @Wdate
WHERE
	Clave = @Clave' 
END
GO
/****** Object:  StoredProcedure [dbo].[ActualizaAuxiliar]    Script Date: 05/01/2016 17:47:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ActualizaAuxiliar]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[ActualizaAuxiliar]
	@Nombre char(50), 
	@Direccion char(50)
AS
UPDATE 	Auxiliar
SET
	Nombre = @Nombre,
	Direccion = @Direccion' 
END
GO
