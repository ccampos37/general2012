if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ct_detcomprobxxxx_ct_cabcomprobxxxx]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_detcomprobxxxx] DROP CONSTRAINT FK_ct_detcomprobxxxx_ct_cabcomprobxxxx
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_ctacteanalitixxxx__09746778]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_ctacteanaliticoxxxx] DROP CONSTRAINT FK__ct_ctacteanalitixxxx__09746778
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tri_insertaranaliticoxxxx]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[tri_insertaranaliticoxxxx]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_cabcomprobxxxx]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_cabcomprobxxxx]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_ctacteanaliticoxxxx]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_ctacteanaliticoxxxx]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_detcomprobxxxx]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_detcomprobxxxx]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_saldosxxxx]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_saldosxxxx]
GO

CREATE TABLE [dbo].[ct_cabcomprobxxxx] (
	[cabcomprobmes] [int] NOT NULL ,
	[cabcomprobnumero] [char] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cabcomprobfeccontable] [datetime] NOT NULL ,
	[subasientocodigo] [char] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[estcomprobcodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[asientocodigo] [char] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cabcomprobobservaciones] [varchar] (150) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fechaact] [datetime] NULL ,
	[cabcomprobglosa] [varchar] (30) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cabcomprobtotdebe] [numvalor] NOT NULL ,
	[cabcomprobtothaber] [numvalor] NOT NULL ,
	[cabcomprobtotussdebe] [numvalor] NOT NULL ,
	[cabcomprobtotusshaber] [numvalor] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_ctacteanaliticoxxxx] (
	[cabcomprobmes] [int] NOT NULL ,
	[detcomprobitem] [char] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cabcomprobnumero] [char] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[subasientocodigo] [char] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[asientocodigo] [char] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[documentocodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[operacioncodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[cuentacodigo] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[ctacteanaliticofechaconta] [datetime] NOT NULL ,
	[analiticocodigo] [char] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[ctacteanaliticonumdocumento] [varchar] (23) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[ctacteanaliticofechadoc] [datetime] NOT NULL ,
	[ctacteanaliticoglosa] [varchar] (60) COLLATE Modern_Spanish_CI_AS NULL ,
	[ctacteanaliticodebe] [numvalor] NOT NULL ,
	[ctacteanaliticoussdebe] [numvalor] NOT NULL ,
	[ctacteanaliticohaber] [numvalor] NOT NULL ,
	[ctacteanaliticousshaber] [numvalor] NOT NULL ,
	[ctacteanaliticocancel] [bit] NOT NULL ,
	[ctacteanaliticofechaven] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_detcomprobxxxx] (
	[cabcomprobmes] [int] NOT NULL ,
	[cabcomprobnumero] [char] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[subasientocodigo] [char] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[analiticocodigo] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[asientocodigo] [char] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[detcomprobitem] [char] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[monedacodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[centrocostocodigo] [char] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[documentocodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[operacioncodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cuentacodigo] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[detcomprobnumdocumento] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[detcomprobfechaemision] [datetime] NOT NULL ,
	[detcomprobfechavencimiento] [datetime] NULL ,
	[detcomprobglosa] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[detcomprobdebe] [numvalor] NOT NULL ,
	[detcomprobhaber] [numvalor] NOT NULL ,
	[detcomprobusshaber] [numvalor] NOT NULL ,
	[detcomprobussdebe] [numvalor] NOT NULL ,
	[detcomprobtipocambio] [numvalor] NOT NULL ,
	[detcomprobruc] [char] (13) COLLATE Modern_Spanish_CI_AS NULL ,
	[detcomprobauto] [bit] NOT NULL ,
	[detcomprobformacambio] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[detcomprobajusteuser] [bit] NOT NULL ,
	[plantillaasientoinafecto] [bit] NOT NULL ,
	[tipdocref] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[detcomprobnumref] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[detcomprobconci] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_saldosxxxx] (
	[cuentacodigo] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[saldodebe00] [numvalor] NOT NULL ,
	[saldohaber00] [numvalor] NOT NULL ,
	[saldoussdebe00] [numvalor] NOT NULL ,
	[saldousshaber00] [numvalor] NOT NULL ,
	[saldodebe01] [numvalor] NOT NULL ,
	[saldohaber01] [numvalor] NOT NULL ,
	[saldoacumhaber01] [numvalor] NOT NULL ,
	[saldoacumdebe01] [numvalor] NOT NULL ,
	[saldoussdebe01] [numvalor] NOT NULL ,
	[saldousshaber01] [numvalor] NOT NULL ,
	[saldoacumusshaber01] [numvalor] NOT NULL ,
	[saldoacumussdebe01] [numvalor] NOT NULL ,
	[saldodebe02] [numvalor] NOT NULL ,
	[saldohaber02] [numvalor] NOT NULL ,
	[saldoacumdebe02] [numvalor] NOT NULL ,
	[saldoacumhaber02] [numvalor] NOT NULL ,
	[saldoussdebe02] [numvalor] NOT NULL ,
	[saldousshaber02] [numvalor] NOT NULL ,
	[saldoacumussdebe02] [numvalor] NOT NULL ,
	[saldoacumusshaber02] [numvalor] NOT NULL ,
	[saldodebe03] [numvalor] NOT NULL ,
	[saldohaber03] [numvalor] NOT NULL ,
	[saldoacumdebe03] [numvalor] NOT NULL ,
	[saldoacumhaber03] [numvalor] NOT NULL ,
	[saldoussdebe03] [numvalor] NOT NULL ,
	[saldousshaber03] [numvalor] NOT NULL ,
	[saldoacumussdebe03] [numvalor] NOT NULL ,
	[saldoacumusshaber03] [numvalor] NOT NULL ,
	[saldodebe04] [numvalor] NOT NULL ,
	[saldohaber04] [numvalor] NOT NULL ,
	[saldoacumdebe04] [numvalor] NOT NULL ,
	[saldoacumhaber04] [numvalor] NOT NULL ,
	[saldoussdebe04] [numvalor] NOT NULL ,
	[saldousshaber04] [numvalor] NOT NULL ,
	[saldoacumussdebe04] [numvalor] NOT NULL ,
	[saldoacumusshaber04] [numvalor] NOT NULL ,
	[saldodebe05] [numvalor] NOT NULL ,
	[saldohaber05] [numvalor] NOT NULL ,
	[saldoacumdebe05] [numvalor] NOT NULL ,
	[saldoacumhaber05] [numvalor] NOT NULL ,
	[saldoussdebe05] [numvalor] NOT NULL ,
	[saldousshaber05] [numvalor] NOT NULL ,
	[saldoacumussdebe05] [numvalor] NOT NULL ,
	[saldoacumusshaber05] [numvalor] NOT NULL ,
	[saldodebe06] [numvalor] NOT NULL ,
	[saldohaber06] [numvalor] NOT NULL ,
	[saldoacumdebe06] [numvalor] NOT NULL ,
	[saldoacumhaber06] [numvalor] NOT NULL ,
	[saldoussdebe06] [numvalor] NOT NULL ,
	[saldousshaber06] [numvalor] NOT NULL ,
	[saldoacumussdebe06] [numvalor] NOT NULL ,
	[saldoacumusshaber06] [numvalor] NOT NULL ,
	[saldodebe07] [numvalor] NOT NULL ,
	[saldohaber07] [numvalor] NOT NULL ,
	[saldoacumdebe07] [numvalor] NOT NULL ,
	[saldoacumhaber07] [numvalor] NOT NULL ,
	[saldoussdebe07] [numvalor] NOT NULL ,
	[saldousshaber07] [numvalor] NOT NULL ,
	[saldoacumussdebe07] [numvalor] NOT NULL ,
	[saldoacumusshaber07] [numvalor] NOT NULL ,
	[saldodebe08] [numvalor] NOT NULL ,
	[saldohaber08] [numvalor] NOT NULL ,
	[saldoacumdebe08] [numvalor] NOT NULL ,
	[saldoacumhaber08] [numvalor] NOT NULL ,
	[saldoussdebe08] [numvalor] NOT NULL ,
	[saldousshaber08] [numvalor] NOT NULL ,
	[saldoacumussdebe08] [numvalor] NOT NULL ,
	[saldoacumusshaber08] [numvalor] NOT NULL ,
	[saldodebe09] [numvalor] NOT NULL ,
	[saldohaber09] [numvalor] NOT NULL ,
	[saldoacumdebe09] [numvalor] NOT NULL ,
	[saldoacumhaber09] [numvalor] NOT NULL ,
	[saldoussdebe09] [numvalor] NOT NULL ,
	[saldousshaber09] [numvalor] NOT NULL ,
	[saldoacumussdebe09] [numvalor] NOT NULL ,
	[saldoacumusshaber09] [numvalor] NOT NULL ,
	[saldodebe10] [numvalor] NOT NULL ,
	[saldohaber10] [numvalor] NOT NULL ,
	[saldoacumdebe10] [numvalor] NOT NULL ,
	[saldoacumhaber10] [numvalor] NOT NULL ,
	[saldoussdebe10] [numvalor] NOT NULL ,
	[saldousshaber10] [numvalor] NOT NULL ,
	[saldoacumussdebe10] [numvalor] NOT NULL ,
	[saldoacumusshaber10] [numvalor] NOT NULL ,
	[saldodebe11] [numvalor] NOT NULL ,
	[saldohaber11] [numvalor] NOT NULL ,
	[saldoacumdebe11] [numvalor] NOT NULL ,
	[saldoacumhaber11] [numvalor] NOT NULL ,
	[saldoussdebe11] [numvalor] NOT NULL ,
	[saldousshaber11] [numvalor] NOT NULL ,
	[saldoacumussdebe11] [numvalor] NOT NULL ,
	[saldoacumusshaber11] [numvalor] NOT NULL ,
	[saldodebe12] [numvalor] NOT NULL ,
	[saldohaber12] [numvalor] NOT NULL ,
	[saldoacumdebe12] [numvalor] NOT NULL ,
	[saldoacumhaber12] [numvalor] NOT NULL ,
	[saldoussdebe12] [numvalor] NOT NULL ,
	[saldousshaber12] [numvalor] NOT NULL ,
	[saldoacumussdebe12] [numvalor] NOT NULL ,
	[saldoacumusshaber12] [numvalor] NOT NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ct_cabcomprobxxxx] WITH NOCHECK ADD 
	CONSTRAINT [PK__ct_cabcomprobxxxx__6CD828CA] PRIMARY KEY  CLUSTERED 
	(
		[cabcomprobnumero],
		[cabcomprobmes],
		[subasientocodigo],
		[asientocodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_detcomprobxxxx] WITH NOCHECK ADD 
	CONSTRAINT [PK__ct_detcomprobxxxx__7EF6D905] PRIMARY KEY  CLUSTERED 
	(
		[cabcomprobmes],
		[cabcomprobnumero],
		[subasientocodigo],
		[asientocodigo],
		[detcomprobitem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_saldosxxxx] WITH NOCHECK ADD 
	CONSTRAINT [PK_ct_saldosxxxx] PRIMARY KEY  CLUSTERED 
	(
		[cuentacodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_cabcomprobxxxx] WITH NOCHECK ADD 
	CONSTRAINT [DF_ct_cabcomprobxxxx_cabcomprobglosa] DEFAULT (0) FOR [cabcomprobglosa],
	CONSTRAINT [DF_ct_cabcomprobxxxx_cabcomprobtotdebe] DEFAULT (0) FOR [cabcomprobtotdebe],
	CONSTRAINT [DF_ct_cabcomprobxxxx_cabcomprobtothaber] DEFAULT (0) FOR [cabcomprobtothaber],
	CONSTRAINT [DF_ct_cabcomprobxxxx_cabcomprobtotussdebe] DEFAULT (0) FOR [cabcomprobtotussdebe],
	CONSTRAINT [DF_ct_cabcomprobxxxx_cabcomprobtotusshaber] DEFAULT (0) FOR [cabcomprobtotusshaber]
GO

 CREATE  INDEX [XIF21ct_cabcomprobxxxx] ON [dbo].[ct_cabcomprobxxxx]([subasientocodigo], [asientocodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF36ct_cabcomprobxxxx] ON [dbo].[ct_cabcomprobxxxx]([estcomprobcodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [fecha] ON [dbo].[ct_cabcomprobxxxx]([cabcomprobfeccontable]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ct_ctacteanaliticoxxxx] WITH NOCHECK ADD 
	CONSTRAINT [DF_ct_ctacteanaliticoxxxx_cabcomprobmes] DEFAULT (0) FOR [cabcomprobmes],
	CONSTRAINT [DF_ct_ctacteanaliticoxxxx_ctacteanaliticodebe] DEFAULT (0) FOR [ctacteanaliticodebe],
	CONSTRAINT [DF_ct_ctacteanaliticoxxxx_ctacteanaliticoussdebe] DEFAULT (0) FOR [ctacteanaliticoussdebe],
	CONSTRAINT [DF_ct_ctacteanaliticoxxxx_ctacteanaliticohaber] DEFAULT (0) FOR [ctacteanaliticohaber],
	CONSTRAINT [DF_ct_ctacteanaliticoxxxx_ctacteanaliticousshaber] DEFAULT (0) FOR [ctacteanaliticousshaber]
GO

 CREATE  INDEX [XIF28ct_ctacteanaliticoxxxx] ON [dbo].[ct_ctacteanaliticoxxxx]([analiticocodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF31ct_ctacteanaliticoxxxx] ON [dbo].[ct_ctacteanaliticoxxxx]([cuentacodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF32ct_ctacteanaliticoxxxx] ON [dbo].[ct_ctacteanaliticoxxxx]([documentocodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF33ct_ctacteanaliticoxxxx] ON [dbo].[ct_ctacteanaliticoxxxx]([cabcomprobmes], [detcomprobitem], [cabcomprobnumero], [subasientocodigo], [asientocodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF38ct_ctacteanaliticoxxxx] ON [dbo].[ct_ctacteanaliticoxxxx]([operacioncodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF39ct_ctacteanaliticoxxxx] ON [dbo].[ct_ctacteanaliticoxxxx]([operacioncodigo]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ct_detcomprobxxxx] WITH NOCHECK ADD 
	CONSTRAINT [DF_ct_detcomprobxxxx_analiticocodigo] DEFAULT ('00') FOR [analiticocodigo],
	CONSTRAINT [DF_ct_detcomprobxxxx_monedacodigo] DEFAULT ('00') FOR [monedacodigo],
	CONSTRAINT [DF_ct_detcomprobxxxx_centrocostocodigo] DEFAULT ('00') FOR [centrocostocodigo],
	CONSTRAINT [DF_ct_detcomprobxxxx_documentocodigo] DEFAULT ('00') FOR [documentocodigo],
	CONSTRAINT [DF_ct_detcomprobxxxx_operacioncodigo] DEFAULT ('00') FOR [operacioncodigo],
	CONSTRAINT [DF_ct_detcomprobxxxx_cuentacodigo] DEFAULT ('00') FOR [cuentacodigo],
	CONSTRAINT [DF_ct_detcomprobxxxx_detcomprobnumdocumento] DEFAULT ('') FOR [detcomprobnumdocumento],
	CONSTRAINT [DF_ct_detcomprobxxxx_detcomprobglosa] DEFAULT ('') FOR [detcomprobglosa],
	CONSTRAINT [DF_ct_detcomprobxxxx_detcomprobdebe] DEFAULT (0) FOR [detcomprobdebe],
	CONSTRAINT [DF_ct_detcomprobxxxx_detcomprobhaber] DEFAULT (0) FOR [detcomprobhaber],
	CONSTRAINT [DF_ct_detcomprobxxxx_detcomprobusshaber] DEFAULT (0) FOR [detcomprobusshaber],
	CONSTRAINT [DF_ct_detcomprobxxxx_detcomprobussdebe] DEFAULT (0) FOR [detcomprobussdebe],
	CONSTRAINT [DF_ct_detcomprobxxxx_detcomprobtipocambio] DEFAULT (0) FOR [detcomprobtipocambio],
	CONSTRAINT [DF_ct_detcomprobxxxx_detcomprobruc] DEFAULT ('') FOR [detcomprobruc],
	CONSTRAINT [DF_ct_detcomprobxxxx_detcomprobauto] DEFAULT (0) FOR [detcomprobauto],
	CONSTRAINT [DF_ct_detcomprobxxxx_detcomprobformacambio] DEFAULT (1) FOR [detcomprobformacambio],
	CONSTRAINT [DF_ct_detcomprobxxxx_detcomprobajusteuser] DEFAULT (0) FOR [detcomprobajusteuser],
	CONSTRAINT [DF_ct_detcomprobxxxx_plantillaasientoinafecto_1] DEFAULT (0) FOR [plantillaasientoinafecto],
	CONSTRAINT [DF_ct_detcomprobxxxx_plantillaasientoinafecto] DEFAULT ('00') FOR [tipdocref],
	CONSTRAINT [DF_ct_detcomprobxxxx_detcomprobnumref] DEFAULT ('') FOR [detcomprobnumref],
	CONSTRAINT [DF_ct_detcomprobxxxx_detcomprobconci] DEFAULT (0) FOR [detcomprobconci]
GO

 CREATE  INDEX [XIF16ct_detcomprobxxxx] ON [dbo].[ct_detcomprobxxxx]([cabcomprobnumero], [cabcomprobmes], [subasientocodigo], [asientocodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF24ct_detcomprobxxxx] ON [dbo].[ct_detcomprobxxxx]([operacioncodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF25ct_detcomprobxxxx] ON [dbo].[ct_detcomprobxxxx]([documentocodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF35ct_detcomprobxxxx] ON [dbo].[ct_detcomprobxxxx]([centrocostocodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF37ct_detcomprobxxxx] ON [dbo].[ct_detcomprobxxxx]([monedacodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [analitico] ON [dbo].[ct_detcomprobxxxx]([detcomprobnumdocumento]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ct_saldosxxxx] WITH NOCHECK ADD 
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe00] DEFAULT (0) FOR [saldodebe00],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber00] DEFAULT (0) FOR [saldohaber00],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe001] DEFAULT (0) FOR [saldoussdebe00],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber001] DEFAULT (0) FOR [saldousshaber00],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe01] DEFAULT (0) FOR [saldodebe01],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber01] DEFAULT (0) FOR [saldohaber01],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber01] DEFAULT (0) FOR [saldoacumhaber01],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe01] DEFAULT (0) FOR [saldoacumdebe01],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe011] DEFAULT (0) FOR [saldoussdebe01],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber011] DEFAULT (0) FOR [saldousshaber01],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber011] DEFAULT (0) FOR [saldoacumusshaber01],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe011] DEFAULT (0) FOR [saldoacumussdebe01],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe02] DEFAULT (0) FOR [saldodebe02],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber02] DEFAULT (0) FOR [saldohaber02],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe02] DEFAULT (0) FOR [saldoacumdebe02],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber02] DEFAULT (0) FOR [saldoacumhaber02],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe021] DEFAULT (0) FOR [saldoussdebe02],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber021] DEFAULT (0) FOR [saldousshaber02],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe021] DEFAULT (0) FOR [saldoacumussdebe02],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber021] DEFAULT (0) FOR [saldoacumusshaber02],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe03] DEFAULT (0) FOR [saldodebe03],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber03] DEFAULT (0) FOR [saldohaber03],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe03] DEFAULT (0) FOR [saldoacumdebe03],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber03] DEFAULT (0) FOR [saldoacumhaber03],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe031] DEFAULT (0) FOR [saldoussdebe03],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber031] DEFAULT (0) FOR [saldousshaber03],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe031] DEFAULT (0) FOR [saldoacumussdebe03],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber031] DEFAULT (0) FOR [saldoacumusshaber03],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe04] DEFAULT (0) FOR [saldodebe04],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber04] DEFAULT (0) FOR [saldohaber04],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe04] DEFAULT (0) FOR [saldoacumdebe04],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber04] DEFAULT (0) FOR [saldoacumhaber04],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe041] DEFAULT (0) FOR [saldoussdebe04],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber041] DEFAULT (0) FOR [saldousshaber04],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe041] DEFAULT (0) FOR [saldoacumussdebe04],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber041] DEFAULT (0) FOR [saldoacumusshaber04],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe05] DEFAULT (0) FOR [saldodebe05],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber05] DEFAULT (0) FOR [saldohaber05],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe05] DEFAULT (0) FOR [saldoacumdebe05],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber05] DEFAULT (0) FOR [saldoacumhaber05],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe051] DEFAULT (0) FOR [saldoussdebe05],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber051] DEFAULT (0) FOR [saldousshaber05],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe051] DEFAULT (0) FOR [saldoacumussdebe05],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber051] DEFAULT (0) FOR [saldoacumusshaber05],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe06] DEFAULT (0) FOR [saldodebe06],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber06] DEFAULT (0) FOR [saldohaber06],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe06] DEFAULT (0) FOR [saldoacumdebe06],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber06] DEFAULT (0) FOR [saldoacumhaber06],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe061] DEFAULT (0) FOR [saldoussdebe06],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber061] DEFAULT (0) FOR [saldousshaber06],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe061] DEFAULT (0) FOR [saldoacumussdebe06],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber061] DEFAULT (0) FOR [saldoacumusshaber06],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe07] DEFAULT (0) FOR [saldodebe07],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber07] DEFAULT (0) FOR [saldohaber07],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe07] DEFAULT (0) FOR [saldoacumdebe07],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber07] DEFAULT (0) FOR [saldoacumhaber07],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe071] DEFAULT (0) FOR [saldoussdebe07],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber071] DEFAULT (0) FOR [saldousshaber07],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe071] DEFAULT (0) FOR [saldoacumussdebe07],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber071] DEFAULT (0) FOR [saldoacumusshaber07],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe08] DEFAULT (0) FOR [saldodebe08],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber08] DEFAULT (0) FOR [saldohaber08],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe08] DEFAULT (0) FOR [saldoacumdebe08],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber08] DEFAULT (0) FOR [saldoacumhaber08],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe081] DEFAULT (0) FOR [saldoussdebe08],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber081] DEFAULT (0) FOR [saldousshaber08],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe081] DEFAULT (0) FOR [saldoacumussdebe08],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber081] DEFAULT (0) FOR [saldoacumusshaber08],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe09] DEFAULT (0) FOR [saldodebe09],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber09] DEFAULT (0) FOR [saldohaber09],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe09] DEFAULT (0) FOR [saldoacumdebe09],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber09] DEFAULT (0) FOR [saldoacumhaber09],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe091] DEFAULT (0) FOR [saldoussdebe09],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber091] DEFAULT (0) FOR [saldousshaber09],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe091] DEFAULT (0) FOR [saldoacumussdebe09],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber091] DEFAULT (0) FOR [saldoacumusshaber09],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe10] DEFAULT (0) FOR [saldodebe10],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber10] DEFAULT (0) FOR [saldohaber10],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe10] DEFAULT (0) FOR [saldoacumdebe10],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber10] DEFAULT (0) FOR [saldoacumhaber10],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe101] DEFAULT (0) FOR [saldoussdebe10],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber101] DEFAULT (0) FOR [saldousshaber10],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe101] DEFAULT (0) FOR [saldoacumussdebe10],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber101] DEFAULT (0) FOR [saldoacumusshaber10],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe11] DEFAULT (0) FOR [saldodebe11],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber11] DEFAULT (0) FOR [saldohaber11],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe11] DEFAULT (0) FOR [saldoacumdebe11],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber11] DEFAULT (0) FOR [saldoacumhaber11],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe111] DEFAULT (0) FOR [saldoussdebe11],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber111] DEFAULT (0) FOR [saldousshaber11],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe111] DEFAULT (0) FOR [saldoacumussdebe11],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber111] DEFAULT (0) FOR [saldoacumusshaber11],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe12] DEFAULT (0) FOR [saldodebe12],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber12] DEFAULT (0) FOR [saldohaber12],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe12] DEFAULT (0) FOR [saldoacumdebe12],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber12] DEFAULT (0) FOR [saldoacumhaber12],
	CONSTRAINT [DF_ct_saldosxxxx_saldodebe121] DEFAULT (0) FOR [saldoussdebe12],
	CONSTRAINT [DF_ct_saldosxxxx_saldohaber121] DEFAULT (0) FOR [saldousshaber12],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumdebe121] DEFAULT (0) FOR [saldoacumussdebe12],
	CONSTRAINT [DF_ct_saldosxxxx_saldoacumhaber121] DEFAULT (0) FOR [saldoacumusshaber12]
GO

ALTER TABLE [dbo].[ct_cabcomprobxxxx] ADD 
	CONSTRAINT [FK__ct_cabcom__estcoxxxx__6DCC4D03] FOREIGN KEY 
	(
		[estcomprobcodigo]
	) REFERENCES [dbo].[ct_estcomprob] (
		[estcomprobcodigo]
	),
	CONSTRAINT [FK__ct_cabcomprobxxxx__6EC0713C] FOREIGN KEY 
	(
		[subasientocodigo],
		[asientocodigo]
	) REFERENCES [dbo].[ct_subasiento] (
		[subasientocodigo],
		[asientocodigo]
	)
GO

ALTER TABLE [dbo].[ct_ctacteanaliticoxxxx] ADD 
	CONSTRAINT [FK__ct_ctacte__analixxxx__0C50D423] FOREIGN KEY 
	(
		[analiticocodigo]
	) REFERENCES [dbo].[ct_analitico] (
		[analiticocodigo]
	),
	CONSTRAINT [FK__ct_ctacte__cuentxxxx__0B5CAFEA] FOREIGN KEY 
	(
		[cuentacodigo]
	) REFERENCES [dbo].[ct_cuenta] (
		[cuentacodigo]
	),
	CONSTRAINT [FK__ct_ctacte__documxxxx__0A688BB1] FOREIGN KEY 
	(
		[documentocodigo]
	) REFERENCES [dbo].[gr_documento] (
		[documentocodigo]
	),
	CONSTRAINT [FK__ct_ctacte__operaxxxx__0880433F] FOREIGN KEY 
	(
		[operacioncodigo]
	) REFERENCES [dbo].[ct_operacion] (
		[operacioncodigo]
	),
	CONSTRAINT [FK__ct_ctacteanalitixxxx__09746778] FOREIGN KEY 
	(
		[cabcomprobmes],
		[cabcomprobnumero],
		[subasientocodigo],
		[asientocodigo],
		[detcomprobitem]
	) REFERENCES [dbo].[ct_detcomprobxxxx] (
		[cabcomprobmes],
		[cabcomprobnumero],
		[subasientocodigo],
		[asientocodigo],
		[detcomprobitem]
	) ON DELETE CASCADE  ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[ct_detcomprobxxxx] ADD 
	CONSTRAINT [FK_ct_detcomprobxxxx_ct_cabcomprobxxxx] FOREIGN KEY 
	(
		[cabcomprobnumero],
		[cabcomprobmes],
		[subasientocodigo],
		[asientocodigo]
	) REFERENCES [dbo].[ct_cabcomprobxxxx] (
		[cabcomprobnumero],
		[cabcomprobmes],
		[subasientocodigo],
		[asientocodigo]
	) ON DELETE CASCADE ,
	CONSTRAINT [FK_ct_detcomprobxxxx_ct_centrocosto] FOREIGN KEY 
	(
		[centrocostocodigo]
	) REFERENCES [dbo].[ct_centrocosto] (
		[centrocostocodigo]
	),
	CONSTRAINT [FK_ct_detcomprobxxxx_ct_operacion] FOREIGN KEY 
	(
		[operacioncodigo]
	) REFERENCES [dbo].[ct_operacion] (
		[operacioncodigo]
	),
	CONSTRAINT [FK_ct_detcomprobxxxx_gr_documento] FOREIGN KEY 
	(
		[documentocodigo]
	) REFERENCES [dbo].[gr_documento] (
		[documentocodigo]
	),
	CONSTRAINT [FK_ct_detcomprobxxxx_gr_documento1] FOREIGN KEY 
	(
		[tipdocref]
	) REFERENCES [dbo].[gr_documento] (
		[documentocodigo]
	),
	CONSTRAINT [FK_ct_detcomprobxxxx_gr_moneda] FOREIGN KEY 
	(
		[monedacodigo]
	) REFERENCES [dbo].[gr_moneda] (
		[monedacodigo]
	)
GO

ALTER TABLE [dbo].[ct_saldosxxxx] ADD 
	CONSTRAINT [FK__ct_saldos__cuentxxxx__1B9317B3] FOREIGN KEY 
	(
		[cuentacodigo]
	) REFERENCES [dbo].[ct_cuenta] (
		[cuentacodigo]
	)
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER  tri_insertaranaliticoxxxx  ON [dbo].[ct_detcomprobxxxx] 
FOR INSERT
AS
Insert dbo.ct_ctacteanaliticoxxxx
(cabcomprobmes, detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo, documentocodigo, operacioncodigo, cuentacodigo, 
 ctacteanaliticofechaconta, analiticocodigo, ctacteanaliticonumdocumento, ctacteanaliticofechadoc, ctacteanaliticoglosa, ctacteanaliticodebe, 
 ctacteanaliticoussdebe, ctacteanaliticohaber, ctacteanaliticousshaber, ctacteanaliticocancel, ctacteanaliticofechaven)
select 
   cabcomprobmes,detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo,documentocodigo, operacioncodigo, cuentacodigo, 
   detcomprobfechaemision, analiticocodigo, detcomprobnumdocumento,detcomprobfechaemision,detcomprobglosa, detcomprobdebe,
   detcomprobussdebe, detcomprobhaber,detcomprobusshaber,0,detcomprobfechavencimiento
from
     inserted
where  not (analiticocodigo ='00' or analiticocodigo is null or rtrim(analiticocodigo)='' ) and 
            not (documentocodigo ='00' or  documentocodigo is null or rtrim(documentocodigo)='' )  and  
            not      (rtrim(detcomprobnumdocumento)=''  or detcomprobnumdocumento is null)



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

