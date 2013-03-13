
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_cabcomprob2007]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_cabcomprob2007]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_ctacteanalitico2007]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_ctacteanalitico2007]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_detcomprob2007]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_detcomprob2007]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_detcomprobxxxx]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_detcomprobxxxx]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_gastos2007]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_gastos2007]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_gastosxxxx]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_gastosxxxx]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_saldos2007]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_saldos2007]
GO


DELETE ct_analitico
delete ct_entidad
delete ct_estcomprob

SET ANSI_NULLS ON 
GO


CREATE TABLE [dbo].[ct_cabcomprob2007] (
	[cabcomprobmes] [int] NOT NULL ,
	[asientocodigo] [char] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[subasientocodigo] [char] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cabcomprobnumero] [char] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cabcomprobfeccontable] [datetime] NOT NULL ,
	[usuariocodigo] [nchar] (8) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[estcomprobcodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[cabcomprobobservaciones] [varchar] (150) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fechaact] [datetime] NULL ,
	[cabcomprobglosa] [varchar] (30) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cabcomprobtotdebe] [numeric](20, 4) NOT NULL ,
	[cabcomprobtothaber] [numeric](20, 4) NOT NULL ,
	[cabcomprobtotussdebe] [numeric](20, 4) NOT NULL ,
	[cabcomprobtotusshaber] [numeric](20, 4) NOT NULL ,
	[cabcomprobgrabada] [bit] NULL ,
	[cabcomprobnref] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[cabcomprobnlibro] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[cabcomprobnprovi] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO


CREATE TABLE [dbo].[ct_ctacteanalitico2007] (
	[cabcomprobmes] [int] NOT NULL ,
	[asientocodigo] [char] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[subasientocodigo] [char] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cabcomprobnumero] [char] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[detcomprobitem] [char] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[documentocodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[operacioncodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[cuentacodigo] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[ctacteanaliticofechaconta] [datetime] NOT NULL ,
	[analiticocodigo] [char] (15) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[ctacteanaliticonumdocumento] [varchar] (23) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[ctacteanaliticofechadoc] [datetime] NOT NULL ,
	[ctacteanaliticoglosa] [varchar] (60) COLLATE Modern_Spanish_CI_AS NULL ,
	[ctacteanaliticodebe] [numeric](20, 4) NOT NULL ,
	[ctacteanaliticoussdebe] [numeric](20, 4) NOT NULL ,
	[ctacteanaliticohaber] [numeric](20, 4) NOT NULL ,
	[ctacteanaliticousshaber] [numeric](20, 4) NOT NULL ,
	[ctacteanaliticocancel] [bit] NOT NULL ,
	[ctacteanaliticofechaven] [datetime] NULL ,
	[monedacodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO


CREATE TABLE [dbo].[ct_detcomprob2007] (
	[cabcomprobmes] [int] NOT NULL ,
	[asientocodigo] [char] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[subasientocodigo] [char] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cabcomprobnumero] [char] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[detcomprobitem] [char] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[analiticocodigo] [char] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[monedacodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[centrocostocodigo] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
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
	[detcomprobconci] [int] NULL ,
	[detcomprobnlibro] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[detcomprobfecharef] [datetime] NULL 
) ON [PRIMARY]
GO


CREATE TABLE [dbo].[ct_gastos2007] (
	[cuentacodigo] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[centrocostocodigo] [varchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
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
	[gastosacumusshaber04] [numeric](20, 4) NOT NULL ,
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
	[gastoshaber12] [numeric](20, 4) NOT NULL ,
	[gastosacum12] [numeric](20, 4) NOT NULL ,
	[gastosuss12] [numeric](20, 4) NOT NULL ,
	[gastosacumuss12] [numeric](20, 4) NOT NULL ,
	[usuariocodigo] [nchar] (8) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fechaact] [datetime] NOT NULL 
) ON [PRIMARY]
GO


CREATE TABLE [dbo].[ct_saldos2007] (
	[cuentacodigo] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[saldodebe00] [numvalor] NULL ,
	[saldohaber00] [numvalor] NULL ,
	[saldoussdebe00] [numvalor] NULL ,
	[saldousshaber00] [numvalor] NULL ,
	[saldodebe01] [numvalor] NULL ,
	[saldohaber01] [numvalor] NULL ,
	[saldoacumhaber01] [numvalor] NULL ,
	[saldoacumdebe01] [numvalor] NULL ,
	[saldoussdebe01] [numvalor] NULL ,
	[saldousshaber01] [numvalor] NULL ,
	[saldoacumusshaber01] [numvalor] NULL ,
	[saldoacumussdebe01] [numvalor] NULL ,
	[saldodebe02] [numvalor] NULL ,
	[saldohaber02] [numvalor] NULL ,
	[saldoacumdebe02] [numvalor] NULL ,
	[saldoacumhaber02] [numvalor] NULL ,
	[saldoussdebe02] [numvalor] NULL ,
	[saldousshaber02] [numvalor] NULL ,
	[saldoacumussdebe02] [numvalor] NULL ,
	[saldoacumusshaber02] [numvalor] NULL ,
	[saldodebe03] [numvalor] NULL ,
	[saldohaber03] [numvalor] NULL ,
	[saldoacumdebe03] [numvalor] NULL ,
	[saldoacumhaber03] [numvalor] NULL ,
	[saldoussdebe03] [numvalor] NULL ,
	[saldousshaber03] [numvalor] NULL ,
	[saldoacumussdebe03] [numvalor] NULL ,
	[saldoacumusshaber03] [numvalor] NULL ,
	[saldodebe04] [numvalor] NULL ,
	[saldohaber04] [numvalor] NULL ,
	[saldoacumdebe04] [numvalor] NULL ,
	[saldoacumhaber04] [numvalor] NULL ,
	[saldoussdebe04] [numvalor] NULL ,
	[saldousshaber04] [numvalor] NULL ,
	[saldoacumussdebe04] [numvalor] NULL ,
	[saldoacumusshaber04] [numvalor] NULL ,
	[saldodebe05] [numvalor] NULL ,
	[saldohaber05] [numvalor] NULL ,
	[saldoacumdebe05] [numvalor] NULL ,
	[saldoacumhaber05] [numvalor] NULL ,
	[saldoussdebe05] [numvalor] NULL ,
	[saldousshaber05] [numvalor] NULL ,
	[saldoacumussdebe05] [numvalor] NULL ,
	[saldoacumusshaber05] [numvalor] NULL ,
	[saldodebe06] [numvalor] NULL ,
	[saldohaber06] [numvalor] NULL ,
	[saldoacumdebe06] [numvalor] NULL ,
	[saldoacumhaber06] [numvalor] NULL ,
	[saldoussdebe06] [numvalor] NULL ,
	[saldousshaber06] [numvalor] NULL ,
	[saldoacumussdebe06] [numvalor] NULL ,
	[saldoacumusshaber06] [numvalor] NULL ,
	[saldodebe07] [numvalor] NULL ,
	[saldohaber07] [numvalor] NULL ,
	[saldoacumdebe07] [numvalor] NULL ,
	[saldoacumhaber07] [numvalor] NULL ,
	[saldoussdebe07] [numvalor] NULL ,
	[saldousshaber07] [numvalor] NULL ,
	[saldoacumussdebe07] [numvalor] NULL ,
	[saldoacumusshaber07] [numvalor] NULL ,
	[saldodebe08] [numvalor] NULL ,
	[saldohaber08] [numvalor] NULL ,
	[saldoacumdebe08] [numvalor] NULL ,
	[saldoacumhaber08] [numvalor] NULL ,
	[saldoussdebe08] [numvalor] NULL ,
	[saldousshaber08] [numvalor] NULL ,
	[saldoacumussdebe08] [numvalor] NULL ,
	[saldoacumusshaber08] [numvalor] NULL ,
	[saldodebe09] [numvalor] NULL ,
	[saldohaber09] [numvalor] NULL ,
	[saldoacumdebe09] [numvalor] NULL ,
	[saldoacumhaber09] [numvalor] NULL ,
	[saldoussdebe09] [numvalor] NULL ,
	[saldousshaber09] [numvalor] NULL ,
	[saldoacumussdebe09] [numvalor] NULL ,
	[saldoacumusshaber09] [numvalor] NULL ,
	[saldodebe10] [numvalor] NULL ,
	[saldohaber10] [numvalor] NULL ,
	[saldoacumdebe10] [numvalor] NULL ,
	[saldoacumhaber10] [numvalor] NULL ,
	[saldoussdebe10] [numvalor] NULL ,
	[saldousshaber10] [numvalor] NULL ,
	[saldoacumussdebe10] [numvalor] NULL ,
	[saldoacumusshaber10] [numvalor] NULL ,
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

