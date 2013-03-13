if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_ctacte__analixxxx__0C50D423]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_ctacteanaliticoxxxx] DROP CONSTRAINT FK__ct_ctacte__analixxxx__0C50D423
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_contab__aplic__7755B73D]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_contab] DROP CONSTRAINT FK__ct_contab__aplic__7755B73D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_ctasxa__asien__0E391C95]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_ctasxasiento] DROP CONSTRAINT FK__ct_ctasxa__asien__0E391C95
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_subasi__asien__6AEFE058]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_subasiento] DROP CONSTRAINT FK__ct_subasi__asien__6AEFE058
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_contab__centr__7B264821]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_contab] DROP CONSTRAINT FK__ct_contab__centr__7B264821
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_contab__cuent__7A3223E8]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_contab] DROP CONSTRAINT FK__ct_contab__cuent__7A3223E8
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_ctacte__cuentxxxx__0B5CAFEA]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_ctacteanaliticoxxxx] DROP CONSTRAINT FK__ct_ctacte__cuentxxxx__0B5CAFEA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_distri__cuent__57DD0BE4]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_distribucion] DROP CONSTRAINT FK__ct_distri__cuent__57DD0BE4
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_planti__cuent__14E61A24]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_plantillaasiento] DROP CONSTRAINT FK__ct_planti__cuent__14E61A24
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_saldos__cuentxxxx__1B9317B3]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_saldosxxxx] DROP CONSTRAINT FK__ct_saldos__cuentxxxx__1B9317B3
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_ctacteanalitixxxx__09746778]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_ctacteanaliticoxxxx] DROP CONSTRAINT FK__ct_ctacteanalitixxxx__09746778
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_analit__entid__5E8A0973]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_analitico] DROP CONSTRAINT FK__ct_analit__entid__5E8A0973
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_cabcom__estcoxxxx__6DCC4D03]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_cabcomprobxxxx] DROP CONSTRAINT FK__ct_cabcom__estcoxxxx__6DCC4D03
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_contab__opera__7C1A6C5A]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_contab] DROP CONSTRAINT FK__ct_contab__opera__7C1A6C5A
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_ctacte__operaxxxx__078C1F06]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_ctacteanaliticoxxxx] DROP CONSTRAINT FK__ct_ctacte__operaxxxx__078C1F06
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_ctacte__operaxxxx__0880433F]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_ctacteanaliticoxxxx] DROP CONSTRAINT FK__ct_ctacte__operaxxxx__0880433F
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_planti__opera__15DA3E5D]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_plantillaasiento] DROP CONSTRAINT FK__ct_planti__opera__15DA3E5D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_cabcomprobxxxx__6EC0713C]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_cabcomprobxxxx] DROP CONSTRAINT FK__ct_cabcomprobxxxx__6EC0713C
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_contab__7849DB76]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_contab] DROP CONSTRAINT FK__ct_contab__7849DB76
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_plantillaasie__16CE6296]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_plantillaasiento] DROP CONSTRAINT FK__ct_plantillaasie__16CE6296
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_analit__tipoa__5F7E2DAC]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_analitico] DROP CONSTRAINT FK__ct_analit__tipoa__5F7E2DAC
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_cuenta__tipoa__55F4C372]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_cuenta] DROP CONSTRAINT FK__ct_cuenta__tipoa__55F4C372
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_contab__docum__7D0E9093]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_contab] DROP CONSTRAINT FK__ct_contab__docum__7D0E9093
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_ctacte__documxxxx__0A688BB1]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_ctacteanaliticoxxxx] DROP CONSTRAINT FK__ct_ctacte__documxxxx__0A688BB1
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_contab__moned__793DFFAF]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_contab] DROP CONSTRAINT FK__ct_contab__moned__793DFFAF
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK__ct_subasi__monedxxxx__69FBBC1F]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ct_subasiento] DROP CONSTRAINT FK__ct_subasi__monedxxxx__69FBBC1F
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tri_asientoauto]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[tri_asientoauto]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tri_insertaranalitico]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[tri_insertaranalitico]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_analiticoentidad]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_analiticoentidad]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_analitico]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_analitico]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_aplicacion]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_aplicacion]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_asiento]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_asiento]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_cabcomprobxxxx]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_cabcomprobxxxx]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_centrocosto]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_centrocosto]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_contab]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_contab]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_ctacteanaliticoxxxx]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_ctacteanaliticoxxxx]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_ctasxasiento]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_ctasxasiento]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_cuenta]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_cuenta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_detcomprobxxxx]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_detcomprobxxxx]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_distribucion]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_distribucion]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_entidad]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_entidad]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_estcomprob]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_estcomprob]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_general]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_general]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_operacion]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_operacion]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_paramgastos]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_paramgastos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_paramlibaux]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_paramlibaux]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_plantillaasiento]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_plantillaasiento]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_ratiosfinan]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_ratiosfinan]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_saldosxxxx]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_saldosxxxx]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_sistema]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_sistema]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_strucbalance]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_strucbalance]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_strucganper]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_strucganper]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_subasiento]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_subasiento]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_tipoanalitico]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_tipoanalitico]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_tipocambio]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_tipocambio]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ct_totalizadorgp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ct_totalizadorgp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[gr_documento]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[gr_documento]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[gr_moneda]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[gr_moneda]
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

setuser
GO

EXEC sp_addtype N'fechaact', N'datetime', N'not null'
GO

setuser
GO

setuser
GO

EXEC sp_addtype N'numvalor', N'numeric(20,4)', N'null'
GO

setuser
GO

setuser
GO

EXEC sp_addtype N'usuariocodigo', N'nchar (8)', N'not null'
GO

setuser
GO

CREATE TABLE [dbo].[ct_analitico] (
	[analiticocodigo] [char] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[entidadcodigo] [char] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[tipoanaliticocodigo] [char] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_aplicacion] (
	[aplicacioncodigo] [char] (18) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[aplicaciondescripcion] [char] (18) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_asiento] (
	[asientocodigo] [char] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[asientodescripcion] [varchar] (30) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
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

CREATE TABLE [dbo].[ct_centrocosto] (
	[centrocostocodigo] [char] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[centrocostodescripcion] [char] (30) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[centrocostodescrcorta] [char] (15) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[centrocostotipo] [char] (1) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_contab] (
	[aplicacioncodigo] [char] (18) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cuentacodigo] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[centrocostocodigo] [char] (5) COLLATE Modern_Spanish_CI_AS NULL ,
	[operacioncodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[monedacodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[documentocodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[subasientocodigo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[asientocodigo] [char] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[contabnivel01] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[contabnivel02] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[contabitem] [char] (1) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tabglosa] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[contabnivel03] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tabcargoabono] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[contabasientoaux] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[contabsubasientoaux] [char] (2) COLLATE Modern_Spanish_CI_AS NULL 
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

CREATE TABLE [dbo].[ct_ctasxasiento] (
	[asientocodigo] [char] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[ctasxasientodescrip] [varchar] (40) COLLATE Modern_Spanish_CI_AS NULL ,
	[ctasxasientovalorbruto] [varchar] (60) COLLATE Modern_Spanish_CI_AS NULL ,
	[ctasxasientoigv1] [varchar] (12) COLLATE Modern_Spanish_CI_AS NULL ,
	[ctasxasientoigv2] [varchar] (12) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_cuenta] (
	[cuentacodigo] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cuentadescripcion] [char] (35) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tipoanaliticocodigo] [char] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cuentalineaactivo] [float] NOT NULL ,
	[cuentalineapasivo] [float] NOT NULL ,
	[cuentalineaegp] [float] NOT NULL ,
	[cuentanatu] [char] (1) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cuentanategp] [char] (1) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cuentagrupo] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[fechaact] [fechaact] NOT NULL ,
	[cuentaestadoccostos] [bit] NOT NULL ,
	[cuentaestadoanalitico] [bit] NOT NULL ,
	[cuentadocumento] [bit] NOT NULL ,
	[cuentanivel] [int] NOT NULL ,
	[cuentaestadodistribucion] [bit] NOT NULL ,
	[cuentaestado] [bit] NOT NULL 
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
	[detcomprobnumdocumento] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
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
	[plantillaasientoinafecto] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_distribucion] (
	[cuentacodigo] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[distribucioncargo] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[distribucionabono] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[distribucionporcen] [float] NOT NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_entidad] (
	[entidadcodigo] [char] (6) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[entidadrazonsocial] [char] (40) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[entidaddireccion] [char] (25) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[entidadruc] [char] (11) COLLATE Modern_Spanish_CI_AS NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[entidadtelefono] [char] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_estcomprob] (
	[estcomprobcodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[estcomprobdescripcion] [varchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_general] (
	[generalanno] [char] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[generalmes01] [bit] NOT NULL ,
	[generalmes02] [bit] NOT NULL ,
	[generalmes03] [bit] NOT NULL ,
	[generalmes04] [bit] NOT NULL ,
	[generalmes05] [bit] NOT NULL ,
	[generalmes06] [bit] NOT NULL ,
	[generalmes07] [bit] NOT NULL ,
	[generalmes08] [bit] NOT NULL ,
	[generalmes09] [bit] NOT NULL ,
	[generalmes10] [bit] NOT NULL ,
	[generalmes11] [bit] NOT NULL ,
	[generalmes12] [bit] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_operacion] (
	[operacioncodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[operaciondescripcion] [char] (25) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_paramgastos] (
	[paramgastoslinutil] [varchar] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[paramgastosctautil] [varchar] (12) COLLATE Modern_Spanish_CI_AS NULL ,
	[paramgastoslinventa] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[paramgastoslinadmin] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[paramgastoslindiv] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[paramgastosactivo] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[paramgastosgastoadmin] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[paramgastosgastoventa] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[paramgastosgastoprod] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[paramgastosgastofinan] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[paramgastosgastodiv] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_paramlibaux] (
	[paramlibauxtipo] [varchar] (6) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[paramlibauxdescripcion] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[paramlibauxvalbru] [varchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[paramlibauxigv1] [varchar] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[paramlibauxigv2] [varchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[paramlibauxtiporeg] [char] (4) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_plantillaasiento] (
	[subasientocodigo] [char] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[asientocodigo] [char] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[plantillaasientocorrela] [int] NOT NULL ,
	[operacioncodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[cuentacodigo] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[iddebeohaber] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[plantillaasientoinafecto] [bit] NULL ,
	[usuariocodigo] [usuariocodigo] NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_ratiosfinan] (
	[ratiosfinanlinea] [varchar] (6) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[ratiosfinannivel1] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[ratiosfinandescrip1] [varchar] (110) COLLATE Modern_Spanish_CI_AS NULL ,
	[ratiosfinandescrip2] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[ratiosfinanformula] [varchar] (120) COLLATE Modern_Spanish_CI_AS NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
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

CREATE TABLE [dbo].[ct_sistema] (
	[sistemadescripcionempresa] [char] (40) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[sistemadescrcortaempresa] [char] (15) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[sistemaesttipodescrempresa] [bit] NOT NULL ,
	[sistemadireccionempresa] [char] (40) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[cia_perdefa] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[sistemadirdataanoactivo] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[perant] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[sistemaempresaruc] [char] (11) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[sistemadirdataanoanterior] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[sistemaestcuadreasiento] [bit] NOT NULL ,
	[sistemaestimpresionasiento] [bit] NOT NULL ,
	[sistemaconfiguracuenta] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[sistemaultimonivel] [int] NULL ,
	[monedacodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_strucbalance] (
	[strucbalancelinea] [varchar] (6) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[strucbalancenivel1] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[strucbalancedescrip1] [varchar] (60) COLLATE Modern_Spanish_CI_AS NULL ,
	[strucbalancesigno1] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[strucbalancenivel2] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[strucbalancedescrip2] [varchar] (60) COLLATE Modern_Spanish_CI_AS NULL ,
	[strucbalancesigno2] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_strucganper] (
	[strucganperlinea] [varchar] (6) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[strucganpernivel] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[strucganperacumes] [float] NULL ,
	[strucganpersigno1] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[strucganperdescrip] [varchar] (60) COLLATE Modern_Spanish_CI_AS NULL ,
	[strucganperacumanno] [float] NULL ,
	[strucganpersigno2] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[strucganpersaldo] [bit] NOT NULL ,
	[strucganpergrupo] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_subasiento] (
	[subasientocodigo] [char] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[asientocodigo] [char] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[monedacodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[subasientodescripcion] [varchar] (30) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[subasientonumcorr01] [float] NULL ,
	[subasientonumcorr02] [float] NULL ,
	[subasientonumcorr03] [float] NULL ,
	[subasientonumcorr04] [float] NULL ,
	[subasientonumcorr05] [float] NULL ,
	[subasientonumcorr06] [float] NULL ,
	[subasientonumcorr07] [float] NULL ,
	[subasientonumcorr08] [float] NULL ,
	[subasientonumcorr09] [float] NULL ,
	[subasientonumcorr10] [float] NULL ,
	[subasientonumcorr11] [float] NULL ,
	[subasientonumcorr12] [float] NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[subasientoglosa] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[fechaact] [fechaact] NOT NULL ,
	[subasientorepitedoc] [bit] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_tipoanalitico] (
	[tipoanaliticocodigo] [char] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[tipoanaliticodescripcion] [varchar] (40) COLLATE Modern_Spanish_CI_AS NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_tipocambio] (
	[tipocambiofecha] [datetime] NOT NULL ,
	[tipocambiocompra] [numvalor] NOT NULL ,
	[tipocambioventa] [numvalor] NOT NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[tipocambiopromedio] [numvalor] NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ct_totalizadorgp] (
	[totalizadorgplinea] [varchar] (6) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[totalizadorgpformulames] [varchar] (90) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[totalizadorgpformulaacum] [varchar] (90) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[totalizadorgpantig] [varchar] (6) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[gr_documento] (
	[documentocodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[documentodescripcion] [char] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[documentoregcompras] [bit] NOT NULL ,
	[documentoregventas] [bit] NOT NULL ,
	[documentoregletrasxcobrar] [bit] NOT NULL ,
	[documentoregletrasxpagar] [bit] NOT NULL ,
	[documentonotacredito] [bit] NOT NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[gr_moneda] (
	[monedacodigo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[monedadescripcion] [char] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[monedaabreviatura] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[monedasimbolo] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[usuariocodigo] [usuariocodigo] NOT NULL ,
	[fechaact] [fechaact] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ct_analitico] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[analiticocodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_aplicacion] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[aplicacioncodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_asiento] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[asientocodigo]
	)  ON [PRIMARY] 
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

ALTER TABLE [dbo].[ct_centrocosto] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[centrocostocodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_contab] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[aplicacioncodigo],
		[contabnivel01],
		[contabnivel02],
		[monedacodigo],
		[contabitem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_cuenta] WITH NOCHECK ADD 
	CONSTRAINT [PK__ct_cuenta__55009F39] PRIMARY KEY  CLUSTERED 
	(
		[cuentacodigo]
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

ALTER TABLE [dbo].[ct_entidad] WITH NOCHECK ADD 
	CONSTRAINT [PK__ct_entidad__5BAD9CC8] PRIMARY KEY  CLUSTERED 
	(
		[entidadcodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_estcomprob] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[estcomprobcodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_general] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[generalanno]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_operacion] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[operacioncodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_paramgastos] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[paramgastoslinutil]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_paramlibaux] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[paramlibauxtipo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_plantillaasiento] WITH NOCHECK ADD 
	CONSTRAINT [PK__ct_plantillaasie__13F1F5EB] PRIMARY KEY  CLUSTERED 
	(
		[subasientocodigo],
		[asientocodigo],
		[plantillaasientocorrela],
		[cuentacodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_ratiosfinan] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[ratiosfinanlinea]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_saldosxxxx] WITH NOCHECK ADD 
	CONSTRAINT [PK_ct_saldosxxxx] PRIMARY KEY  CLUSTERED 
	(
		[cuentacodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_strucbalance] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[strucbalancelinea]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_strucganper] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[strucganperlinea]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_subasiento] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[subasientocodigo],
		[asientocodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_tipoanalitico] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[tipoanaliticocodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_tipocambio] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[tipocambiofecha]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ct_totalizadorgp] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[totalizadorgplinea]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[gr_documento] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[documentocodigo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[gr_moneda] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[monedacodigo]
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

 CREATE  INDEX [XIF45ct_contab] ON [dbo].[ct_contab]([documentocodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF47ct_contab] ON [dbo].[ct_contab]([operacioncodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF48ct_contab] ON [dbo].[ct_contab]([centrocostocodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF49ct_contab] ON [dbo].[ct_contab]([cuentacodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF50ct_contab] ON [dbo].[ct_contab]([monedacodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF52ct_contab] ON [dbo].[ct_contab]([subasientocodigo], [asientocodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF53ct_contab] ON [dbo].[ct_contab]([aplicacioncodigo]) ON [PRIMARY]
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

 CREATE  INDEX [XIF34ct_ctasxasiento] ON [dbo].[ct_ctasxasiento]([asientocodigo]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ct_cuenta] WITH NOCHECK ADD 
	CONSTRAINT [DF_ct_cuenta_cuentacodigo] DEFAULT (0) FOR [cuentacodigo],
	CONSTRAINT [DF_ct_cuenta_tipoanaliticocodigo] DEFAULT ('00') FOR [tipoanaliticocodigo],
	CONSTRAINT [DF_ct_cuenta_cuentaestadoccostos] DEFAULT (0) FOR [cuentaestadoccostos],
	CONSTRAINT [DF_ct_cuenta_cuentaestadoanalitico] DEFAULT (0) FOR [cuentaestadoanalitico],
	CONSTRAINT [DF_ct_cuenta_cuentadocumento] DEFAULT (0) FOR [cuentadocumento],
	CONSTRAINT [DF_ct_cuenta_cuentanivel] DEFAULT (0) FOR [cuentanivel],
	CONSTRAINT [DF_ct_cuenta_cuentaestadodistribucion] DEFAULT (0) FOR [cuentaestadodistribucion],
	CONSTRAINT [DF_ct_cuenta_cuentaestado] DEFAULT (1) FOR [cuentaestado]
GO

 CREATE  INDEX [XIF54ct_cuenta] ON [dbo].[ct_cuenta]([tipoanaliticocodigo]) ON [PRIMARY]
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
	CONSTRAINT [DF_ct_detcomprobxxxx_plantillaasientoinafecto] DEFAULT (0) FOR [plantillaasientoinafecto]
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

ALTER TABLE [dbo].[ct_distribucion] WITH NOCHECK ADD 
	CONSTRAINT [DF_ct_distribucion_distribucionporcen] DEFAULT (0) FOR [distribucionporcen]
GO

 CREATE  INDEX [XIF59ct_distribucion] ON [dbo].[ct_distribucion]([cuentacodigo]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ct_entidad] WITH NOCHECK ADD 
	CONSTRAINT [DF_ct_entidad_entidadruc] DEFAULT ('') FOR [entidadruc],
	CONSTRAINT [DF_ct_entidad_entidadtelefono] DEFAULT ('') FOR [entidadtelefono]
GO

ALTER TABLE [dbo].[ct_general] WITH NOCHECK ADD 
	CONSTRAINT [DF_ct_general_generalmes01] DEFAULT (0) FOR [generalmes01],
	CONSTRAINT [DF_ct_general_generalmes02] DEFAULT (0) FOR [generalmes02],
	CONSTRAINT [DF_ct_general_generalmes03] DEFAULT (0) FOR [generalmes03],
	CONSTRAINT [DF_ct_general_generalmes04] DEFAULT (0) FOR [generalmes04],
	CONSTRAINT [DF_ct_general_generalmes05] DEFAULT (0) FOR [generalmes05],
	CONSTRAINT [DF_ct_general_generalmes06] DEFAULT (0) FOR [generalmes06],
	CONSTRAINT [DF_ct_general_generalmes07] DEFAULT (0) FOR [generalmes07],
	CONSTRAINT [DF_ct_general_generalmes08] DEFAULT (0) FOR [generalmes08],
	CONSTRAINT [DF_ct_general_generalmes09] DEFAULT (0) FOR [generalmes09],
	CONSTRAINT [DF_ct_general_generalmes10] DEFAULT (0) FOR [generalmes10],
	CONSTRAINT [DF_ct_general_generalmes11] DEFAULT (0) FOR [generalmes11],
	CONSTRAINT [DF_ct_general_generalmes12] DEFAULT (0) FOR [generalmes12]
GO

ALTER TABLE [dbo].[ct_plantillaasiento] WITH NOCHECK ADD 
	CONSTRAINT [DF_ct_plantillaasiento_usuariocodigo] DEFAULT (0) FOR [usuariocodigo]
GO

 CREATE  INDEX [XIF20ct_plantillaasiento] ON [dbo].[ct_plantillaasiento]([subasientocodigo], [asientocodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF23ct_plantillaasiento] ON [dbo].[ct_plantillaasiento]([operacioncodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF60ct_plantillaasiento] ON [dbo].[ct_plantillaasiento]([cuentacodigo]) ON [PRIMARY]
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

ALTER TABLE [dbo].[ct_sistema] WITH NOCHECK ADD 
	CONSTRAINT [DF_ct_sistema_sistemaesttipodescrempresa] DEFAULT (0) FOR [sistemaesttipodescrempresa],
	CONSTRAINT [DF_ct_sistema_sistemaestcuadreasiento] DEFAULT (0) FOR [sistemaestcuadreasiento],
	CONSTRAINT [DF_ct_sistema_sistemaestimpresionasiento] DEFAULT (0) FOR [sistemaestimpresionasiento]
GO

ALTER TABLE [dbo].[ct_subasiento] WITH NOCHECK ADD 
	CONSTRAINT [DF_ct_subasiento_subasientorepitedoc] DEFAULT (0) FOR [subasientorepitedoc]
GO

 CREATE  INDEX [XIF18ct_subasiento] ON [dbo].[ct_subasiento]([asientocodigo]) ON [PRIMARY]
GO

 CREATE  INDEX [XIF26ct_subasiento] ON [dbo].[ct_subasiento]([monedacodigo]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ct_analitico] ADD 
	CONSTRAINT [FK__ct_analit__entid__5E8A0973] FOREIGN KEY 
	(
		[entidadcodigo]
	) REFERENCES [dbo].[ct_entidad] (
		[entidadcodigo]
	),
	 FOREIGN KEY 
	(
		[tipoanaliticocodigo]
	) REFERENCES [dbo].[ct_tipoanalitico] (
		[tipoanaliticocodigo]
	)
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

ALTER TABLE [dbo].[ct_contab] ADD 
	 FOREIGN KEY 
	(
		[subasientocodigo],
		[asientocodigo]
	) REFERENCES [dbo].[ct_subasiento] (
		[subasientocodigo],
		[asientocodigo]
	),
	 FOREIGN KEY 
	(
		[aplicacioncodigo]
	) REFERENCES [dbo].[ct_aplicacion] (
		[aplicacioncodigo]
	),
	 FOREIGN KEY 
	(
		[centrocostocodigo]
	) REFERENCES [dbo].[ct_centrocosto] (
		[centrocostocodigo]
	),
	CONSTRAINT [FK__ct_contab__cuent__7A3223E8] FOREIGN KEY 
	(
		[cuentacodigo]
	) REFERENCES [dbo].[ct_cuenta] (
		[cuentacodigo]
	),
	 FOREIGN KEY 
	(
		[documentocodigo]
	) REFERENCES [dbo].[gr_documento] (
		[documentocodigo]
	),
	 FOREIGN KEY 
	(
		[monedacodigo]
	) REFERENCES [dbo].[gr_moneda] (
		[monedacodigo]
	),
	 FOREIGN KEY 
	(
		[operacioncodigo]
	) REFERENCES [dbo].[ct_operacion] (
		[operacioncodigo]
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
	CONSTRAINT [FK__ct_ctacte__operaxxxx__078C1F06] FOREIGN KEY 
	(
		[operacioncodigo]
	) REFERENCES [dbo].[ct_operacion] (
		[operacioncodigo]
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

ALTER TABLE [dbo].[ct_ctasxasiento] ADD 
	 FOREIGN KEY 
	(
		[asientocodigo]
	) REFERENCES [dbo].[ct_asiento] (
		[asientocodigo]
	)
GO

ALTER TABLE [dbo].[ct_cuenta] ADD 
	CONSTRAINT [FK__ct_cuenta__tipoa__55F4C372] FOREIGN KEY 
	(
		[tipoanaliticocodigo]
	) REFERENCES [dbo].[ct_tipoanalitico] (
		[tipoanaliticocodigo]
	)
GO

ALTER TABLE [dbo].[ct_distribucion] ADD 
	CONSTRAINT [FK__ct_distri__cuent__57DD0BE4] FOREIGN KEY 
	(
		[cuentacodigo]
	) REFERENCES [dbo].[ct_cuenta] (
		[cuentacodigo]
	)
GO

ALTER TABLE [dbo].[ct_plantillaasiento] ADD 
	CONSTRAINT [FK__ct_planti__cuent__14E61A24] FOREIGN KEY 
	(
		[cuentacodigo]
	) REFERENCES [dbo].[ct_cuenta] (
		[cuentacodigo]
	),
	CONSTRAINT [FK__ct_planti__opera__15DA3E5D] FOREIGN KEY 
	(
		[operacioncodigo]
	) REFERENCES [dbo].[ct_operacion] (
		[operacioncodigo]
	),
	CONSTRAINT [FK__ct_plantillaasie__16CE6296] FOREIGN KEY 
	(
		[subasientocodigo],
		[asientocodigo]
	) REFERENCES [dbo].[ct_subasiento] (
		[subasientocodigo],
		[asientocodigo]
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

ALTER TABLE [dbo].[ct_subasiento] ADD 
	 FOREIGN KEY 
	(
		[asientocodigo]
	) REFERENCES [dbo].[ct_asiento] (
		[asientocodigo]
	),
	CONSTRAINT [FK__ct_subasi__monedxxxx__69FBBC1F] FOREIGN KEY 
	(
		[monedacodigo]
	) REFERENCES [dbo].[gr_moneda] (
		[monedacodigo]
	)
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

CREATE TRIGGER tri_asientoauto  ON dbo.ct_detcomprobxxxx 
FOR INSERT
AS

--Actualizar la Cabecera
Declare 
@sub varchar(4) ,@asi varchar(3),@mes float,@comp varchar(10) ,
@ad numvalor,@ah numvalor,@ads numvalor,@ahs numvalor     	

Select distinct @mes=cabcomprobmes, @comp=cabcomprobnumero, @sub=subasientocodigo, 
                      @asi=asientocodigo from ct_detcomprobxxxx

--Calculando el total de las cabeceras
select  @ad=isnull(sum(detcomprobdebe),0), @ah =isnull(sum(detcomprobhaber),0),@ahs= isnull(sum(detcomprobusshaber),0), @ads=isnull(sum(detcomprobussdebe),0) 
from dbo.ct_detcomprobxxxx 
where  cabcomprobmes=@mes and  cabcomprobnumero=@comp and  
            asientocodigo=@asi and  subasientocodigo=@sub 


Update dbo.ct_cabcomprobxxxx
set  cabcomprobtotdebe=@ad,
       cabcomprobtothaber=@ah, 
       cabcomprobtotussdebe=@ads,
       cabcomprobtotusshaber=@ahs,
      estcomprobcodigo='03'
where  cabcomprobmes=@mes and  cabcomprobnumero=@comp and  
            asientocodigo=@asi and  subasientocodigo=@sub









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER  tri_insertaranalitico  ON [dbo].[ct_detcomprobxxxx] 
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

