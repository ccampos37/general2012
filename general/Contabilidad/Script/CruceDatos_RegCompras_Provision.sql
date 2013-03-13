--select * from ##tmpregcompraconta where cabcomprobnumero not in 
--	 (select Compconta from dbo.##tmpregcompraDesa3)


alter procedure ct_ListarDiferenciasCompras
(
	@Base				varchar(50),
	@BaseCompra 	varchar(50),
	@Server			varchar(50),
	@anno				varchar(4),
	@mes				varchar(2)	
)

as
Declare	@cadsql varchar(4000)

set nocount on

/*Ejecuta los Stores de Reg. Compras Contabilidad*/
exec ct_ValidaRegistroCompras_rpt @Base,@BaseCompra,@Server,@anno,@mes,'060,061,062,063,064,','60%33%34%46%63%64%65%9%28%38%73%77%403%','401110,401113,','403140,','401174,'


/*Ejecuta los Stores de Reg. Compras Provisión*/
exec co_registrocompra_rpt @BaseCompra,@Base,@BaseCompra,@Server,@Server,@Server,@Anno,@Mes,'Desa3'



set @cadsql='
/*Base Imponible*/
select ''Base_Imponible'' as concepto,A.Compconta,A.numaux,Prov=Round(A.BaseImpCD,2),Conta=Round(B.baseimpgrab,2)
--select a.*,b.*
from 
	##tmpregcompraDesa3 A,
	##tmpregcompraconta B
where 
	A.Compconta collate  Modern_Spanish_CI_AI= B.cabcomprobnumero collate  Modern_Spanish_CI_AI and
	Round(A.BaseImpCD,2)<>Round(B.baseimpgrab,2)	
order by 1

/*Base Referencial*/
select ''Base_Referencial'' as concepto,A.Compconta,A.numaux,Prov=Round(A.BaseRef,2),Conta=Round(B.MontoReferencia,2)
--select a.*,b.*
from 
	##tmpregcompraDesa3 A,
	##tmpregcompraconta B
where 
	A.Compconta collate  Modern_Spanish_CI_AI= B.cabcomprobnumero collate  Modern_Spanish_CI_AI and
	Round(A.BaseRef,2)<>Round(B.MontoReferencia,2)	
order by 1


/*Cruce de los Montos Inafectos*/
select ''Inafecto'' as concepto,A.Compconta,A.numaux,Prov=Round(A.Inafecto,2),Conta=Round(B.montoinafecto,2)
from 
	##tmpregcompraDesa3 A,
	##tmpregcompraconta B
where 
	A.Compconta collate  Modern_Spanish_CI_AI= B.cabcomprobnumero collate  Modern_Spanish_CI_AI
	and Round(A.Inafecto,2)<>Round(B.montoinafecto,2)
order by 1

/*Cruce de los Montos IGV*/
select ''IGV'' as concepto,A.Compconta,A.numaux,Prov=Round(A.IGVCD,2),Conta=Round(B.igvimpgrab,2)
from 
	##tmpregcompraDesa3 A,
	##tmpregcompraconta B
where 
	A.Compconta collate  Modern_Spanish_CI_AI= B.cabcomprobnumero collate  Modern_Spanish_CI_AI
	and Round(A.IGVCD,2)<>Round(B.igvimpgrab,2)

/*Cruce de las IES*/
select ''IES'' as concepto,A.compconta,A.numaux,Prov=Round(A.IES,2),Conta=Round(B.IES,2)
from 
	##tmpregcompraDesa3 A,
	##tmpregcompraconta B
where 
	A.Compconta collate  Modern_Spanish_CI_AI= B.cabcomprobnumero collate  Modern_Spanish_CI_AI
	and Round(A.IES,2)<>Round(B.IES,2)

/*Cruce del Impuesto Renta*/
select ''Imp_Renta'' as concepto,A.compconta,A.numaux,Prov=Round(A.ImpRta,2),Conta=Round(B.RENTA,2)
from 
	##tmpregcompraDesa3 A,
	##tmpregcompraconta B
where 
	A.Compconta collate  Modern_Spanish_CI_AI= B.cabcomprobnumero collate  Modern_Spanish_CI_AI
	and Round(A.ImpRta,2)<>Round(B.RENTA,2)'

set nocount off

exec(@cadsql)


--exec ct_ListarDiferenciasCompras 'Contaprueba','camtex_tj','server_tc','2003','05'
