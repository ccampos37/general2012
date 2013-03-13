if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ct_detcomprob2007_ct_cabcomprob2007]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_detcomprob2007] DROP CONSTRAINT FK_ct_detcomprob2007_ct_cabcomprob2007
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ct_detcomprob2007_ct_centrocosto]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_detcomprob2007] DROP CONSTRAINT FK_ct_detcomprob2007_ct_centrocosto
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ct_detcomprob2007_ct_operacion]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_detcomprob2007] DROP CONSTRAINT FK_ct_detcomprob2007_ct_operacion
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ct_detcomprob2007_gr_documento]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_detcomprob2007] DROP CONSTRAINT FK_ct_detcomprob2007_gr_documento
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ct_detcomprob2007_gr_moneda]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_detcomprob2007] DROP CONSTRAINT FK_ct_detcomprob2007_gr_moneda
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tri_insertaranalitico2007]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[tri_insertaranalitico2007]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[fn_xx]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[fn_xx]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_analiticoentidad]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_analiticoentidad]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_detallegastos]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_detallegastos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_te_conciliacionCaja]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_te_conciliacionCaja]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bancomoneda]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bancomoneda]
GO


if exists (select * from dbo.systypes where name = N'fechaact')
exec sp_droptype N'fechaact'
GO

if exists (select * from dbo.systypes where name = N'numvalor')
exec sp_droptype N'numvalor'
GO

if exists (select * from dbo.systypes where name = N'usuariocodigo')
exec sp_droptype N'usuariocodigo'
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UW_ZeroDefault]') and OBJECTPROPERTY(id, N'IsDefault') = 1)
drop default [dbo].[UW_ZeroDefault]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ceros]') and OBJECTPROPERTY(id, N'IsDefault') = 1)
drop default [dbo].[ceros]
GO

if not exists (select * from dbo.sysusers where name = N'guest' and hasdbaccess = 1)
	EXEC sp_grantdbaccess N'guest'
GO


CREATE DEFAULT UW_ZeroDefault AS 0

GO
create default [ceros] as 0
GO
setuser
GO

EXEC sp_addtype N'fechaact', N'datetime', N'not null'
GO

setuser
GO

setuser
GO

EXEC sp_addtype N'numvalor', N'decimal(20,4)', N'not null'
GO

setuser
GO

setuser
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[numvalor]'
GO

setuser
GO

setuser
GO

EXEC sp_addtype N'usuariocodigo', N'nchar (8)', N'not null'
GO

setuser
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE FUNCTION [dbo].[fn_xx] ( @a float  )
RETURNS float
 WITH SCHEMABINDING
 AS  
BEGIN  
  return (@a * 1)	
END



GO
SET QUOTED_IDENTIFIER OFF 
GO


ALTER TABLE [dbo].[co_detprovi2007] WITH NOCHECK ADD 
	CONSTRAINT [PK_co_detprovi2007] PRIMARY KEY  CLUSTERED 
	(
		[cabprovinumero],
		[detproviitem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[co_detprovi2007] ADD 
	CONSTRAINT [FK_co_detprovi2007_ct_centrocosto] FOREIGN KEY 
	(
		[centrocostocodigo]
	) REFERENCES [dbo].[ct_centrocosto] (
		[centrocostocodigo]
	) ON UPDATE CASCADE  NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[ct_estcomprob] WITH NOCHECK ADD 
	CONSTRAINT [PK_ct_estcomprob] PRIMARY KEY  CLUSTERED 
	(
		[estcomprobcodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_cabcomprob2007] WITH NOCHECK ADD 
	CONSTRAINT [PK_ct_cabcomprob2007] PRIMARY KEY  CLUSTERED 
	(
		[cabcomprobmes],
		[asientocodigo],
		[subasientocodigo],
		[cabcomprobnumero]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_centrocosto] WITH NOCHECK ADD 
	CONSTRAINT [PK_ct_centrocosto] PRIMARY KEY  CLUSTERED 
	(
		[centrocostocodigo]
	)  ON [PRIMARY] 
GO


ALTER TABLE [dbo].[ct_detcomprob2007] WITH NOCHECK ADD 
	CONSTRAINT [PK__ct_detcomprob2007] PRIMARY KEY  CLUSTERED 
	(
		[cabcomprobmes],
		[cabcomprobnumero],
		[subasientocodigo],
		[asientocodigo],
		[detcomprobitem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_operacion] WITH NOCHECK ADD 
	CONSTRAINT [PK_ct_operacion] PRIMARY KEY  CLUSTERED 
	(
		[operacioncodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[gr_documento] WITH NOCHECK ADD 
	CONSTRAINT [PK_gr_documento] PRIMARY KEY  CLUSTERED 
	(
		[documentocodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[gr_moneda] WITH NOCHECK ADD 
	CONSTRAINT [PK_gr_moneda] PRIMARY KEY  CLUSTERED 
	(
		[monedacodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[gr_usuario] WITH NOCHECK ADD 
	CONSTRAINT [PK_gr_usuario] PRIMARY KEY  CLUSTERED 
	(
		[usuariocodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[servicios] WITH NOCHECK ADD 
	CONSTRAINT [PK_servicios] PRIMARY KEY  CLUSTERED 
	(
		[ser_codigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_cuenta] WITH NOCHECK ADD 
	CONSTRAINT [PK_ct_cuenta] PRIMARY KEY  CLUSTERED 
	(
		[cuentacodigo]
	)  ON [PRIMARY] 
GO


ALTER TABLE [dbo].[te_FormadePago] WITH NOCHECK ADD 
	CONSTRAINT [PK_te_FormadePago] PRIMARY KEY  CLUSTERED 
	(
		[FormadePagocodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[te_cabecerarecibos] WITH NOCHECK ADD 
	CONSTRAINT [PK_te_cabecerarecibos] PRIMARY KEY  CLUSTERED 
	(
		[cabrec_numrecibo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[te_detallerecibos] WITH NOCHECK ADD 
	CONSTRAINT [PK_te_detallerecibos] PRIMARY KEY  CLUSTERED 
	(
		[cabrec_numrecibo],
		[detrec_item]
	)  ON [PRIMARY] 
GO


ALTER TABLE [dbo].[te_detallerecibos] ADD 
	CONSTRAINT [FK_te_detallerecibos_te_cabecerarecibos] FOREIGN KEY 
	(
		[cabrec_numrecibo]
	) REFERENCES [dbo].[te_cabecerarecibos] (
		[cabrec_numrecibo]
	) ON DELETE CASCADE  ON UPDATE CASCADE  NOT FOR REPLICATION 
GO



ALTER TABLE [dbo].[te_rendiciones] WITH NOCHECK ADD 
	CONSTRAINT [PK_te_rendiciones] PRIMARY KEY  CLUSTERED 
	(
		[oficinacodigo],
		[monedacodigo],
		[rendicionnumero]
	)  ON [PRIMARY] 
GO


ALTER TABLE [dbo].[ct_analitico] WITH NOCHECK ADD 
	CONSTRAINT [PK_ct_analitico] PRIMARY KEY  CLUSTERED 
	(
		[analiticocodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_entidad] WITH NOCHECK ADD 
	CONSTRAINT [PK_ct_entidad] PRIMARY KEY  CLUSTERED 
	(
		[entidadcodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_tipoanalitico] WITH NOCHECK ADD 
	CONSTRAINT [PK_ct_tipoanalitico] PRIMARY KEY  CLUSTERED 
	(
		[tipoanaliticocodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_analitico] ADD 
	CONSTRAINT [FK_ct_analitico_ct_entidad] FOREIGN KEY 
	(
		[entidadcodigo]
	) REFERENCES [dbo].[ct_entidad] (
		[entidadcodigo]
	),
	CONSTRAINT [FK_ct_analitico_ct_tipoanalitico] FOREIGN KEY 
	(
		[tipoanaliticocodigo]
	) REFERENCES [dbo].[ct_tipoanalitico] (
		[tipoanaliticocodigo]
	)
GO

ALTER TABLE [dbo].[vt_cliente] WITH NOCHECK ADD 
	CONSTRAINT [PK_vt_cliente] PRIMARY KEY  CLUSTERED 
	(
		[clientecodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[vt_detallepedido] WITH NOCHECK ADD 
	CONSTRAINT [PK_vt_detallepedido] PRIMARY KEY  CLUSTERED 
	(
		[pedidonumero],
		[detpeditem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[vt_modoventa] WITH NOCHECK ADD 
	CONSTRAINT [PK_vt_modoventa] PRIMARY KEY  CLUSTERED 
	(
		[modovtacodigo]
	)  ON [PRIMARY] 
GO

alter table [dbo].[PK_co_detprovi2007] 	with nocheck add
	CONSTRAINT [PK_co_detprovi2007] PRIMARY KEY  CLUSTERED 
	(
		[cabprovinumero],
		[detproviitem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[vt_pedido] WITH NOCHECK ADD 
	CONSTRAINT [PK_vt_pedido] PRIMARY KEY  CLUSTERED 
	(
		[pedidonumero]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[vt_puntoventa] WITH NOCHECK ADD 
	CONSTRAINT [PK_vt_puntoventa] PRIMARY KEY  CLUSTERED 
	(
		[puntovtacodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_detcomprob2007] ADD 
	CONSTRAINT [DF_ct_detcomprob2007_analiticocodigo] DEFAULT ('00') FOR [analiticocodigo],
	CONSTRAINT [DF_ct_detcomprob2007_monedacodigo] DEFAULT ('00') FOR [monedacodigo],
	CONSTRAINT [DF_ct_detcomprob2007_centrocostocodigo] DEFAULT ('00') FOR [centrocostocodigo],
	CONSTRAINT [DF_ct_detcomprob2007_documentocodigo] DEFAULT ('00') FOR [documentocodigo],
	CONSTRAINT [DF_ct_detcomprob2007_operacioncodigo] DEFAULT ('00') FOR [operacioncodigo],
	CONSTRAINT [DF_ct_detcomprob2007_cuentacodigo] DEFAULT ('00') FOR [cuentacodigo],
	CONSTRAINT [DF_ct_detcomprob2007_detcomprobnumdocumento] DEFAULT ('') FOR [detcomprobnumdocumento],
	CONSTRAINT [DF_ct_detcomprob2007_detcomprobglosa] DEFAULT ('') FOR [detcomprobglosa],
	CONSTRAINT [DF_ct_detcomprob2007_detcomprobruc] DEFAULT ('') FOR [detcomprobruc],
	CONSTRAINT [DF_ct_detcomprob2007_detcomprobauto] DEFAULT (0) FOR [detcomprobauto],
	CONSTRAINT [DF_ct_detcomprob2007_detcomprobformacambio] DEFAULT (1) FOR [detcomprobformacambio],
	CONSTRAINT [DF_ct_detcomprob2007_detcomprobajusteuser] DEFAULT (0) FOR [detcomprobajusteuser],
	CONSTRAINT [DF_ct_detcomprob2007_plantillaasientoinafecto_1] DEFAULT (0) FOR [plantillaasientoinafecto],
	CONSTRAINT [DF_ct_detcomprob2007_plantillaasientoinafecto] DEFAULT ('00') FOR [tipdocref],
	CONSTRAINT [DF_ct_detcomprob2007_detcomprobnumref] DEFAULT ('') FOR [detcomprobnumref],
	CONSTRAINT [DF_ct_detcomprob2007_detcomprobconci] DEFAULT (0) FOR [detcomprobconci],
	CONSTRAINT [DF_ct_detcomprob2007_detcomprobnlibro] DEFAULT ('') FOR [detcomprobnlibro]
GO

ALTER TABLE [dbo].[stkart] ADD 
	CONSTRAINT [DF_stkart_STALMA] DEFAULT (0) FOR [STALMA],
	CONSTRAINT [DF_stkart_STSKREF] DEFAULT (0) FOR [STSKREF],
	CONSTRAINT [DF_stkart_STSKMIN] DEFAULT (0) FOR [STSKMIN],
	CONSTRAINT [DF_stkart_STSKMAX] DEFAULT (0) FOR [STSKMAX],
	CONSTRAINT [DF_stkart_STPUNREP] DEFAULT (0) FOR [STPUNREP],
	CONSTRAINT [DF_stkart_STSEMREP] DEFAULT (0) FOR [STSEMREP],
	CONSTRAINT [DF_stkart_STLOTCOM] DEFAULT (0) FOR [STLOTCOM],
	CONSTRAINT [DF_stkart_STSKCOM] DEFAULT (0) FOR [STSKCOM],
	CONSTRAINT [DF_stkart_STKPREPRO] DEFAULT (0) FOR [STKPREPRO],
	CONSTRAINT [DF_stkart_STKPREULT] DEFAULT (0) FOR [STKPREULT]
GO


EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_detcomprob2007].[detcomprobdebe]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_detcomprob2007].[detcomprobhaber]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_detcomprob2007].[detcomprobtipocambio]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_detcomprob2007].[detcomprobussdebe]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_detcomprob2007].[detcomprobusshaber]'
GO

setuser
GO

setuser
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumdebe01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumdebe02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumdebe03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumdebe04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumdebe05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumdebe06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumdebe07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumdebe08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumdebe09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumdebe10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumdebe11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumdebe12]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumhaber01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumhaber02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumhaber03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumhaber04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumhaber05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumhaber06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumhaber07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumhaber08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumhaber09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumhaber10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumhaber11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumhaber12]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumussdebe01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumussdebe02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumussdebe03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumussdebe04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumussdebe05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumussdebe06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumussdebe07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumussdebe08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumussdebe09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumussdebe10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumussdebe11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumussdebe12]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumusshaber01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumusshaber02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumusshaber03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumusshaber04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumusshaber05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumusshaber06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumusshaber07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumusshaber08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumusshaber09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumusshaber10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumusshaber11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoacumusshaber12]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldodebe00]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldodebe01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldodebe02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldodebe03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldodebe04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldodebe05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldodebe06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldodebe07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldodebe08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldodebe09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldodebe10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldodebe11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldodebe12]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldohaber00]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldohaber01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldohaber02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldohaber03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldohaber04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldohaber05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldohaber06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldohaber07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldohaber08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldohaber09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldohaber10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldohaber11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldohaber12]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoussdebe00]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoussdebe01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoussdebe02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoussdebe03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoussdebe04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoussdebe05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoussdebe06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoussdebe07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoussdebe08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoussdebe09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoussdebe10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoussdebe11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldoussdebe12]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldousshaber00]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldousshaber01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldousshaber02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldousshaber03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldousshaber04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldousshaber05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldousshaber06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldousshaber07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldousshaber08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldousshaber09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldousshaber10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldousshaber11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_saldos2007].[saldousshaber12]'
GO

setuser
GO
------

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos12]'
GO
------

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss12]'
GO

------
EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum12]'
GO

------

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss12]'
GO


setuser
GO






ALTER TABLE [dbo].[Listapre1] ADD 
	CONSTRAINT [FK_Listapre1_MAEART] FOREIGN KEY 
	(
		[productocodigo]
	) REFERENCES [dbo].[MAEART] (
		[ACODIGO]
	) ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[MOVALMDET] ADD 
	CONSTRAINT [FK_MOVALMDET_MAEART] FOREIGN KEY 
	(
		[DECODIGO]
	) REFERENCES [dbo].[MAEART] (
		[ACODIGO]
	) ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[STKLOTE] ADD 
	CONSTRAINT [FK_STKLOTE_MAEART] FOREIGN KEY 
	(
		[STSCODIGO]
	) REFERENCES [dbo].[MAEART] (
		[ACODIGO]
	) ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[ct_detcomprob2007] ADD 
	CONSTRAINT [FK_ct_detcomprob2007_ct_cabcomprob2007] FOREIGN KEY 
	(
		[cabcomprobmes],
		[asientocodigo],
		[subasientocodigo],
		[cabcomprobnumero]
	) REFERENCES [dbo].[ct_cabcomprob2007] (
		[cabcomprobmes],
		[asientocodigo],
		[subasientocodigo],
		[cabcomprobnumero]
	) ON DELETE CASCADE  ON UPDATE CASCADE ,
	CONSTRAINT [FK_ct_detcomprob2007_ct_centrocosto] FOREIGN KEY 
	(
		[centrocostocodigo]
	) REFERENCES [dbo].[ct_centrocosto] (
		[centrocostocodigo]
	) ON UPDATE CASCADE ,
	CONSTRAINT [FK_ct_detcomprob2007_ct_operacion] FOREIGN KEY 
	(
		[operacioncodigo]
	) REFERENCES [dbo].[ct_operacion] (
		[operacioncodigo]
	) ON UPDATE CASCADE ,
	CONSTRAINT [FK_ct_detcomprob2007_gr_documento] FOREIGN KEY 
	(
		[tipdocref]
	) REFERENCES [dbo].[gr_documento] (
		[documentocodigo]
	) NOT FOR REPLICATION ,
	CONSTRAINT [FK_ct_detcomprob2007_gr_moneda] FOREIGN KEY 
	(
		[monedacodigo]
	) REFERENCES [dbo].[gr_moneda] (
		[monedacodigo]
	)
GO

ALTER TABLE [dbo].[listapre2] ADD 
	CONSTRAINT [FK_listapre2_MAEART] FOREIGN KEY 
	(
		[productocodigo]
	) REFERENCES [dbo].[MAEART] (
		[ACODIGO]
	) ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[te_detallerecibos] ADD 
	CONSTRAINT [FK_te_detallerecibos_te_cabecerarecibos] FOREIGN KEY 
	(
		[cabrec_numrecibo]
	) REFERENCES [dbo].[te_cabecerarecibos] (
		[cabrec_numrecibo]
	) ON UPDATE CASCADE  NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[vt_detallepedido] ADD 
	CONSTRAINT [FK_vt_detallepedido_MAEART] FOREIGN KEY 
	(
		[productocodigo]
	) REFERENCES [dbo].[MAEART] (
		[ACODIGO]
	) ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[vt_param_impoexpo] ADD 
	CONSTRAINT [FK_vt_param_impoexpo_vt_cliente] FOREIGN KEY 
	(
		[clientecodigo]
	) REFERENCES [dbo].[vt_cliente] (
		[clientecodigo]
	)
GO


alter table [dbo].[co_detprovi2007] ADD 
	CONSTRAINT [FK_co_detprovi2007_ct_centrocosto] FOREIGN KEY 
	(
		[centrocostocodigo]
	) REFERENCES [ct_centrocosto] (
		[centrocostocodigo]
	) ON UPDATE CASCADE  NOT FOR REPLICATION 
) ON [PRIMARY]
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  VIEW dbo.v_bancomoneda
AS
SELECT     a.cbanco_codigo, a.cbanco_numero, b.monedasimbolo, a.cbanco_referenciacta, a.cbanco_nrocheque, a.monedacodigo
FROM         dbo.te_cuentabancos a INNER JOIN
                      dbo.gr_moneda b ON a.monedacodigo = b.monedacodigo




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE VIEW dbo.v_analiticoentidad
AS
SELECT     a.analiticocodigo, a.tipoanaliticocodigo, b.*
FROM         dbo.ct_analitico a INNER JOIN
                      dbo.ct_entidad b ON a.entidadcodigo = b.entidadcodigo



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE   VIEW dbo.v_detallegastos

as 
select z.*,
soles =case when z.monedacodigo='02' then
            z.cabprovitipcambio*z.detprovitotal  
         else  z.detprovitotal end,
dolares =case when z.monedacodigo='01' then
            z.detprovitotal/z.cabprovitipcambio  
         else  z.detprovitotal end
from 
(
   Select  a.cabprovioficina,j.vendedornombres,
       b.centrocostocodigo,i.centrocostodescripcion,
       A.cabprovinumero,A.cabprovinumaux,
       proveedorcodigo,cabprovirznsoc =a.cabprovirznsoc,
       cabentidadrznsoc= h.entidadrazonsocial,
       b.entidadcodigo,A.modoprovicod,D.modoprovidesc,
       A.documetocodigo,A.cabprovinumdoc,A.cabprovifchdoc,
       A.cabprovifchconta,A.monedacodigo,A.cabprovimes,
       cabprovitipcambio=case when A.cabprovitipcambio=0 then 1 else A.cabprovitipcambio end ,
       B.gastoscodigo,
       gastosdescripcion=B.gastoscodigo+' '+C.gastosdescripcion,
          
       detproviimpbru=Case When G.tdocumentotipo='A' then  B.detproviimpbru * -1 else B.detproviimpbru end,
       detproviimpigv=Case When G.tdocumentotipo='A' then  B.detproviimpigv * -1 else B.detproviimpigv end, 
       detproviimpina=Case When G.tdocumentotipo='A' then  B.detproviimpina * -1 else B.detproviimpina end, 
       detprovitotal= Case When G.tdocumentotipo='A' then  B.detprovitotal * -1 else B.detprovitotal end
   from dbo.co_cabprovi2006 A 
       inner join dbo.co_detprovi2006 B
                on A.cabprovinumero=B.cabprovinumero
       LEFT join dbo.co_gastos C 
         on  B.gastoscodigo =c.gastoscodigo 
       inner join dbo.co_modoprovi D     
         on  A.modoprovicod=D.modoprovicod       
       left join  dbo.cp_proveedor E 
          on  A.proveedorcodigo=e.clientecodigo  
       inner join dbo.cp_tipodocumento G 
          on  a.documetocodigo=g.tdocumentocodigo  
       left join  dbo.ct_entidad h 
          on  b.entidadcodigo=h.entidadcodigo 
       left join  dbo.ct_centrocosto i 
          on  b.centrocostocodigo=i.centrocostocodigo 
       inner join  dbo.cp_oficina j 
          on  a.cabprovioficina=j.vendedorcodigo 
--       inner join dbo.co_sistema k
--          on a.cabprovioficina=k.sistemaoficina
) as z

--WHere Z.cabprovioficina in (select sistemaoficina dbo.co_sistema) 
 
--- select * from v_detallegastos order by cabprovioficina


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE           view v_te_conciliacionCaja
--Declare 

as 
select 
cajadescripcion=case yy.tipocajabase when 'C' then
                     g.cajadescripcion else h.bancodescripcion end,
empresaresumen= case yy.cabrec_transferenciaautomatico when 1 then
                     g.cajadescripcion else yy.empresadescripcion end,
yy.* from 
(
select zz.tipo,zz.chkconcil,zz.rendicionnumero,
b.cabcomprobnumero,zz.cabrec_numrecibo,
zz.detrec_fechacancela,zz.detrec_cajabanco1,b.monedacodigo,
d.monedadescripcion,b.cabrec_ingsal,
zz.detrec_tipodoc_concepto,zz.detrec_numdocumento,

tipoingreso=case  B.cabrec_ingsal when  'I' then
             'INGRESOS ' else 'EGRESOS ' end,
zz.detrec_tipocajabanco,
e.centrocostonivel,e.centrocostodescripcion,
zz.empresacodigo,f.empresadescripcion ,
empresacodigodescripcion = zz.empresacodigo+' '+f.empresadescripcion ,
Td_Concep=Isnull(
    case upper(isnull(rtrim(ltrim(detrec_adicionactacte)),'X')) 
       	when 'P' then (Select X.tdocumentodescripcion from dbo.cp_tipodocumento X 
                        Where X.tdocumentocodigo=zz.detrec_tipodoc_concepto)
        When 'C' then (Select Y.tdocumentodescripcion from  dbo.cc_tipodocumento Y 
                        Where Y.tdocumentocodigo=zz.detrec_tipodoc_concepto)           
       	Else  (Select G.conceptodescripcion  from dbo.te_conceptocaja G  
                where G.conceptocodigo=zz.detrec_tipodoc_concepto) End,''),
b.cabrec_transferenciaautomatico,
b.cabrec_numreciboegreso,
tipocajabase = case b.cabrec_transferenciaautomatico when 1 then
                  (select top 1 j.detrec_tipocajabanco from te_detallerecibos j
                   inner join te_cabecerarecibos k 
                    on j.cabrec_numrecibo=k.cabrec_numrecibo
                    where k.cabrec_numreciboegreso=b.cabrec_numreciboegreso
                     and k.cabrec_numrecibo<>b.cabrec_numrecibo
                   )
              else zz.detrec_tipocajabanco end,
cajabase = case b.cabrec_transferenciaautomatico when 1 then
                  (select top 1 j.detrec_cajabanco1 from te_detallerecibos j
                   inner join te_cabecerarecibos k 
                    on j.cabrec_numrecibo=k.cabrec_numrecibo
                    where k.cabrec_numreciboegreso=b.cabrec_numreciboegreso
                     and k.cabrec_numrecibo<>b.cabrec_numrecibo
                   )
              else zz.detrec_cajabanco1 end,
ruc=Isnull( case upper(isnull(rtrim(ltrim(detrec_adicionactacte)),'X')) 
       	When 'P' then (Select Top 1 P.clienteruc  from dbo.cp_proveedor P 
                        Where P.clientecodigo=b.clientecodigo)
        When 'C' then (Select Top 1 Cl.clienteruc  from  dbo.vt_cliente Cl 
                        Where Cl.clientecodigo=b.clientecodigo)           
        Else  '' End,''),
ProveCliConc=Isnull(case upper(isnull(rtrim(ltrim(detrec_adicionactacte)),'X')) 
       	When 'P' then (Select Top 1 P.clienterazonsocial  from dbo.cp_proveedor P 
                         Where P.clientecodigo=b.clientecodigo)
        When 'C' then (Select Top 1 Cl.clienterazonsocial  from  dbo.vt_cliente Cl
			 Where Cl.clientecodigo=b.clientecodigo)           
	Else  b.cabrec_descripcion End,''),
zz.MONTO,zz.gastos,zz.costos,zz.provision,zz.detrec_monedacancela,
B.cabrec_estadoreg,B.cabrec_fechadocumento,zz.fechconcil,
c.gastosdescripcion,c.gastosequivalente 
from dbo.te_cabecerarecibos  B INNER JOIN
(
select z.*,monto =z.importe/c.cargoapeimpape*d.detprovitotal,
         gastos=d.gastoscodigo ,costos=d.centrocostocodigo,
         provision=d.cabprovinumero,TIPO='P',w.empresacodigo
           from te_cabecerarecibos y WITH (NOLOCK) INNER join 
   ( select 
        a.cabrec_numrecibo,b.clientecodigo,a.detrec_tipodoc_concepto,
        a.detrec_numdocumento,a.detrec_tipocajabanco,a.detrec_cajabanco1,
        a.detrec_monedacancela,a.detrec_fechacancela,
        a.chkconcil,a.fechconcil,a.detrec_adicionactacte,a.rendicionnumero,
        importe=sum(a.detrec_importesoles) 
     from te_detallerecibos a WITH (NOLOCK)
          inner join te_cabecerarecibos b WITH (NOLOCK) on a.cabrec_numrecibo=B.cabrec_numrecibo  
     where detrec_adicionactacte='P' and b.cabcomprobnumero = 0
           AND ISNULL(detalle_no_saldos,1)=0 AND DETREC_ESTADOREG=0
--        and b.cabrec_numrecibo='212999' 
           and b.cabcomprobnumero= 0 
     group by a.cabrec_numrecibo,b.clientecodigo,
           a.detrec_tipodoc_concepto,a.detrec_numdocumento,
           a.detrec_tipocajabanco,a.detrec_cajabanco1,
           a.detrec_monedacancela,a.detrec_fechacancela,
           a.chkconcil,a.fechconcil,a.detrec_adicionactacte,a.rendicionnumero
     ) as z
      on z.cabrec_numrecibo=y.cabrec_numrecibo
     inner join cp_cargo c 
           on z.clientecodigo+z.detrec_tipodoc_concepto+z.detrec_numdocumento=
              c.clientecodigo+c.documentocargo+c.cargonumdoc
     inner join co_detprovi2006 d WITH (NOLOCK) on c.abononumplanilla=d.cabprovinumero
     inner join co_cabprovi2006 w WITH (NOLOCK) on d.cabprovinumero=w.cabprovinumero
union ALL

select 
   a.cabrec_numrecibo,b.clientecodigo,a.detrec_tipodoc_concepto,
   a.detrec_numdocumento,a.detrec_tipocajabanco,a.detrec_cajabanco1,
   a.detrec_monedacancela,a.detrec_fechacancela,
   a.chkconcil,a.fechconcil,a.detrec_adicionactacte,a.rendicionnumero,
   a.detrec_importesoles,
   MONTO =a.detrec_importesoles,
   gastos=a.detrec_gastos,
   costos =a.centrocostocodigo,
   provision =b.cabcomprobnumero, tipo=' ',b.empresacodigo
   from te_detallerecibos a WITH (NOLOCK)
     inner join te_cabecerarecibos b WITH (NOLOCK) on a.cabrec_numrecibo=B.cabrec_numrecibo
   WHERE (detrec_adicionactacte<>'P' or 
          detrec_adicionactacte='P'  AND b.cabcomprobnumero  >  0 ) AND
     DETREC_ESTADOREG=0 and ISNULL(detalle_no_saldos,1)=0 
---and a.cabrec_numrecibo='212706'
) as zz
 on  zz.cabrec_numrecibo=B.cabrec_numrecibo 
 left join dbo.co_gastos c WITH (NOLOCK) on  zz.gastos=c.gastoscodigo 
 left join dbo.gr_moneda  d WITH (NOLOCK) on  b.monedacodigo=d.monedacodigo 
 left join dbo.ct_centrocosto  e WITH (NOLOCK) on  zz.costos=e.centrocostocodigo 
 left join dbo.co_multiempresas f WITH (NOLOCK) on  zz.empresacodigo=f.empresacodigo 
) as yy
left join dbo.te_codigocaja g WITH (NOLOCK) on  yy.cajabase=g.cajacodigo 
left join dbo.gr_banco h WITH (NOLOCK) on  yy.cajabase=h.bancocodigo 



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create     TRIGGER  tri_insertaranalitico2007  ON dbo.ct_detcomprob2007
FOR INSERT
AS


Declare @Fecha datetime
Select @Fecha=cabcomprobfeccontable 
From inserted a inner join ct_cabcomprob2007 b
                  on A.cabcomprobmes=B.cabcomprobmes and A.asientocodigo=B.asientocodigo and 
                     A.subasientocodigo=B.subasientocodigo and A.cabcomprobnumero=B.cabcomprobnumero 
--dateadd(day,-1,cast('01/'+ cast(cabcomprobmes+1 as varchar(2))+'/2002' as datetime))
Insert dbo.ct_ctacteanalitico2007
(cabcomprobmes, detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo, documentocodigo, operacioncodigo, cuentacodigo, 
 ctacteanaliticofechaconta, analiticocodigo, ctacteanaliticonumdocumento, ctacteanaliticofechadoc, ctacteanaliticoglosa, ctacteanaliticodebe, 
 ctacteanaliticoussdebe, ctacteanaliticohaber, ctacteanaliticousshaber, ctacteanaliticocancel, ctacteanaliticofechaven,monedacodigo)
select 
   cabcomprobmes,detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo,documentocodigo, operacioncodigo,d.cuentacodigo, 
   @Fecha, analiticocodigo, detcomprobnumdocumento,detcomprobfechaemision,detcomprobglosa, detcomprobdebe,
   detcomprobussdebe, detcomprobhaber,detcomprobusshaber,0,detcomprobfechavencimiento,monedacodigo
from
     inserted d inner join ct_cuenta C on d.cuentacodigo=c.cuentacodigo
where  not (analiticocodigo ='00' or analiticocodigo is null or rtrim(analiticocodigo)='' ) and 
            not (documentocodigo ='00' or  documentocodigo is null or rtrim(documentocodigo)='' )  and  
            not      (rtrim(detcomprobnumdocumento)=''  or detcomprobnumdocumento is null) 
       and c.cuentaestadoanalitico=1  


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


INSERT INTO gr_moneda
    (monedacodigo, monedadescripcion, monedaabreviatura, 
    monedasimbolo, usuariocodigo, fechaact)
SELECT '00', '(Ninguno)', '(Ninguno)', '(Ninguno)','Sistema', '20/08/2002' UNION ALL
SELECT '01', 'MONEDA NACIONAL', 'SOLES', 'S/.','Sistema', '20/08/2002' UNION ALL
SELECT '02', 'MONEDA EXTRANJERA', 'DOLAR', 'USD','Sistema', '20/08/2002' 


INSERT INTO ct_operacion
    (operacioncodigo, operaciondescripcion, usuariocodigo, 
    fechaact)
VALUES ('00', '(Ninguno)', 'Sistema', '20/08/2002')

go

INSERT INTO ct_centrocosto
    (centrocostocodigo, centrocostodescripcion, 
    centrocostodescrcorta, centrocostotipo, usuariocodigo, 
    fechaact)
VALUES ('00', '(Ninguno)', '(Ninguno)', 'N', 'Sistema', '20/08/2002')

go

INSERT INTO ct_tipoanalitico
    (tipoanaliticocodigo,tipoanaliticodescripcion,usuariocodigo,fechaact)
VALUES ('00','(Ninguno)','Sistema','20/08/2002')

go

INSERT INTO ct_entidad
    (entidadcodigo,entidadrazonsocial,entidaddireccion,entidadtelefono,entidadruc,usuariocodigo,fechaact)
VALUES ('00','(Ninguno)','(Ninguno)','(Ninguno)','(Ninguno)','Sistema','20/08/2002')

go

INSERT INTO ct_analitico
    (analiticocodigo,entidadcodigo,tipoanaliticocodigo,usuariocodigo,fechaact)
VALUES ('00','00','00','Sistema','20/08/2002')

go



INSERT INTO ct_estcomprob
   (estcomprobcodigo,estcomprobdescripcion,usuariocodigo,fechaact)
SELECT '01','ERRADO','Sistema','20/08/2002' UNION ALL
SELECT '02','REGISTRADO','Sistema','20/08/2002' UNION ALL
SELECT '03','PROCESADO','Sistema','20/08/2002'

go

INSERT INTO gr_documento
    (documentocodigo, documentodescripcion, documentoregcompras,documentoregventas,
     documentoregletrasxcobrar, documentoregletrasxpagar, documentonotacredito, 
     usuariocodigo,fechaact)
VALUES ('00', '(Ninguno)',0,0,0,0,0,'Sistema','20/08/2002')


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_gastos2007]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_gastos2007]
GO

CREATE TABLE [dbo].[ct_gastos2007] (
	[cuentacodigo] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[centrocostocodigo] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[gastos00] [numeric](20, 4) NOT NULL ,
	[gastosuss00] [numeric](20, 4) NOT NULL ,
	[gastos01] [numeric](20, 4) NOT NULL ,
	[gastosacum01] [numeric](20, 4) NOT NULL ,
	[gastosuss01] [numeric](20, 4) NOT NULL ,
	[gastosacumuss01] [numeric](20, 4) NOT NULL ,
	[gastos02] [numeric](20, 4) NOT NULL ,
	[gastosacum02] [numeric](20, 4) NOT NULL ,
	[gastosuss02] [numeric](20, 4) NOT NULL ,
	[gastosacumuss02] [numeric](20, 4) NOT NULL ,
	[gastos03] [numeric](20, 4) NOT NULL ,
	[gastosacum03] [numeric](20, 4) NOT NULL ,
	[gastosuss03] [numeric](20, 4) NOT NULL ,
	[gastosacumuss03] [numeric](20, 4) NOT NULL ,
	[gastos04] [numeric](20, 4) NOT NULL ,
	[gastosacum04] [numeric](20, 4) NOT NULL ,
	[gastosuss04] [numeric](20, 4) NOT NULL ,
	[gastosacumuss04] [numeric](20, 4) NOT NULL ,
	[gastos05] [numeric](20, 4) NOT NULL ,
	[gastosacum05] [numeric](20, 4) NOT NULL ,
	[gastosuss05] [numeric](20, 4) NOT NULL ,
	[gastosacumuss05] [numeric](20, 4) NOT NULL ,
	[gastos06] [numeric](20, 4) NOT NULL ,
	[gastosacum06] [numeric](20, 4) NOT NULL ,
	[gastosuss06] [numeric](20, 4) NOT NULL ,
	[gastosacumuss06] [numeric](20, 4) NOT NULL ,
	[gastos07] [numeric](20, 4) NOT NULL ,
	[gastosacum07] [numeric](20, 4) NOT NULL ,
	[gastosuss07] [numeric](20, 4) NOT NULL ,
	[gastosacumuss07] [numeric](20, 4) NOT NULL ,
	[gastos08] [numeric](20, 4) NOT NULL ,
	[gastosacum08] [numeric](20, 4) NOT NULL ,
	[gastosuss08] [numeric](20, 4) NOT NULL ,
	[gastosacumuss08] [numeric](20, 4) NOT NULL ,
	[gastos09] [numeric](20, 4) NOT NULL ,
	[gastosacum09] [numeric](20, 4) NOT NULL ,
	[gastosuss09] [numeric](20, 4) NOT NULL ,
	[gastosacumuss09] [numeric](20, 4) NOT NULL ,
	[gastos10] [numeric](20, 4) NOT NULL ,
	[gastosacum10] [numeric](20, 4) NOT NULL ,
	[gastosuss10] [numeric](20, 4) NOT NULL ,
	[gastosacumuss10] [numeric](20, 4) NOT NULL ,
	[gastos11] [numeric](20, 4) NOT NULL ,
	[gastosacum11] [numeric](20, 4) NOT NULL ,
	[gastosuss11] [numeric](20, 4) NOT NULL ,
	[gastosacumuss11] [numeric](20, 4) NOT NULL ,
	[gastos12] [numeric](20, 4) NOT NULL ,
	[gastosacum12] [numeric](20, 4) NOT NULL ,
	[gastosuss12] [numeric](20, 4) NOT NULL ,
	[gastosacumuss12] [numeric](20, 4) NOT NULL ,
	[usuariocodigo] [nchar] (8) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fechaact] [datetime] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ct_gastos2007] WITH NOCHECK ADD 
	CONSTRAINT [PK_ct_gastos2007] PRIMARY KEY  CLUSTERED 
	(
		[cuentacodigo],
		[centrocostocodigo]
	)  ON [PRIMARY] 
GO

setuser
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos00]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastos12]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacum12]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosacumuss12]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss00]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss01]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss02]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss03]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss04]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss05]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss06]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss07]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss08]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss09]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss10]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss11]'
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[ct_gastos2007].[gastosuss12]'
GO

setuser
GO

ALTER TABLE [dbo].[ct_gastos2007] ADD 
	CONSTRAINT [FK_ct_gastos2007_ct_centrocosto] FOREIGN KEY 
	(
		[centrocostocodigo]
	) REFERENCES [dbo].[ct_centrocosto] (
		[centrocostocodigo]
	) ON UPDATE CASCADE ,
	CONSTRAINT [FK_ct_gastos2007_ct_cuenta] FOREIGN KEY 
	(
		[cuentacodigo]
	) REFERENCES [dbo].[ct_cuenta] (
		[cuentacodigo]
	) ON UPDATE CASCADE 
GO




--delete from ct_subasiento
--delete from ct_asiento
--delete from ct_operacion
--delete from ct_analitico
--delete from ct_tipoanalitico
--delete from ct_cuenta