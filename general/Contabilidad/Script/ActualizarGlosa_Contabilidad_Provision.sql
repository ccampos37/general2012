select * from ct_cabcomprob2003 where asientocodigo like '06%' and cabcomprobmes=1
select * from ct_detcomprob2003 where asientocodigo like '06%' and cabcomprobmes=1


/*Actualizar Glosa Detalle Tinto*/
update contaprueba.dbo.ct_detcomprob2003 set detcomprobglosa=isnull(left(ltrim(rtrim(zz.detproviglosa)),50),'')
--select a.*,zz.*
from 
	contaprueba.dbo.ct_detcomprob2003 a,
	(select b.cabcomprobnumero,a.cabprovinumero,a.detproviglosa
	from 
		[camtex_tinto_compras].dbo.co_detprovi2003 a,
		[contaprueba].dbo.ct_cabcomprob2003 b
	where 
		a.cabprovimes=4 and b.cabcomprobmes=4 and
		cast(a.cabprovinumero as varchar(20))=cast(b.cabcomprobnprovi as varchar(20)) ) as ZZ
where 
	a.cabcomprobnumero=zz.cabcomprobnumero and 
	asientocodigo like '06%' and cabcomprobmes=4


/*Actualizar Glosa Detalle Teje*/
update contaprueba.dbo.ct_detcomprob2003 set detcomprobglosa=isnull(left(ltrim(rtrim(zz.detproviglosa)),50),'')
--select a.*,zz.*
from 
	contaprueba.dbo.ct_detcomprob2003 a,
	(select b.cabcomprobnumero,a.cabprovinumero,a.detproviglosa
	from 
		[camtex_teje_compras].dbo.co_detprovi2003 a,
		[contaprueba].dbo.ct_cabcomprob2003 b
	where 
		a.cabprovimes=4 and b.cabcomprobmes=4 and
		cast(a.cabprovinumero as varchar(20))=cast(b.cabcomprobnprovi as varchar(20)) ) as ZZ
where 
	a.cabcomprobnumero=zz.cabcomprobnumero and 
	asientocodigo like '06%' and cabcomprobmes=4


select b.cabprovinumaux,b.cabprovinumero,b.cabprovinconta,
left('Prov.Compra ' + cast(b.cabprovinumero as varchar(5)) + ' NºAux. ' + cast(b.cabprovinumaux as varchar),30),
a.cabcomprobglosa

/*Actualizar Glosa Cabecera*/
update [contaprueba].dbo.ct_cabcomprob2003 
set cabcomprobglosa=left('Prov.Compra ' + cast(b.cabprovinumero as varchar(5)) + ' NºAux. ' + cast(b.cabprovinumaux as varchar),30)
--select a.*,b.*
from 
	[contaprueba].dbo.ct_cabcomprob2003 a,
	[camtex_teje_compras].dbo.co_cabprovi2003 b
where a.cabcomprobmes=4 and
		a.cabcomprobnumero collate  Modern_Spanish_CI_AI = b.cabprovinconta collate  Modern_Spanish_CI_AI


/*Actualizar el NºComprobante Contable para Tinto*/
update [server_tc].[camtex_tinto].dbo.co_cabprovi2003 
	set cabprovinconta=b.cabprovinconta
--select a.*,b.*
from 
	[server_tc].[camtex_tinto].dbo.co_cabprovi2003  a,
	[pc06].[camtex_tinto].dbo.co_cabprovi2003 b
where 
	a.cabprovinumero=b.cabprovinumero and a.cabprovimes=4

/*Actualizar el NºComprobante Contable para Teje*/
update [server_tc].[camtex_tj].dbo.co_cabprovi2003 
	set cabprovinconta=b.cabprovinconta
--select a.*,b.*
from 
	[server_tc].[camtex_tj].dbo.co_cabprovi2003  a,
	[pc06].[camtex_teje_compras].dbo.co_cabprovi2003 b
where 
	a.cabprovinumero=b.cabprovinumero and a.cabprovimes=4

/*Actualizar Nº Comprobante Contable SalIndefonso*/
update [server_tc].[sanildefonso].dbo.co_cabprovi2003 
	set cabprovinconta=b.cabprovinconta
--select a.*,b.*
from 
	[server_tc].[sanildefonso].dbo.co_cabprovi2003  a,
	[pc06].[sanildefonso].dbo.co_cabprovi2003 b
where 
	a.cabprovinumero=b.cabprovinumero and a.cabprovimes=4




select * from comprasprueba.dbo.co_cabprovi2003 where cabprovimes=1
select * from comprasprueba.dbo.co_detprovi2003 where cabprovimes=1

select * from contaprueba.dbo.ct_cabcomprob2003 
	where asientocodigo like '06%' and cabcomprobmes=4
order by cabcomprobnprovi


/*Para Tintorería*/
update contaprueba.dbo.ct_cabcomprob2003
set cabcomprobnprovi=left(a.cabprovinumaux,2)+'01'+substring(a.cabprovinumaux,3,5)
--select a.cabprovinumaux,left(a.cabprovinumaux,2)+'01'+substring(a.cabprovinumaux,3,5)   
from
	camtex_tinto.dbo.co_cabprovi2003 a,
	contaprueba.dbo.ct_cabcomprob2003 b
where 	a.cabprovimes=5 and b.cabcomprobmes=5 and
			a.cabprovinconta=b.cabcomprobnumero

/*Para Tejeduria*/
update contaprueba.dbo.ct_cabcomprob2003
set cabcomprobnprovi=left(a.cabprovinumaux,2)+'02'+substring(a.cabprovinumaux,3,5)
--select a.cabprovinumaux,left(a.cabprovinumaux,2)+'02'+substring(a.cabprovinumaux,3,5)   
from
	camtex_teje_compras.dbo.co_cabprovi2003 a,
	contaprueba.dbo.ct_cabcomprob2003 b
where 	a.cabprovimes=4 and 
			a.cabprovinconta=b.cabcomprobnumero


select * from [pc06].contaprueba.dbo.ct_cabcomprob2003 
	where asientocodigo like '06%' and cabcomprobmes=4
order by cabcomprobnprovi desc



select * from camtex_teje_compras.dbo.co_cabprovi2003 
	where cabprovinconta='0406100553'


/*Actualizar Glosa Detalle San Ildefonso*/
update prueba_contaprueba_sanil.dbo.ct_detcomprob2003 set detcomprobglosa=isnull(left(ltrim(rtrim(zz.detproviglosa)),50),'')
--select a.*,zz.*
from 
	prueba_contaprueba_sanil.dbo.ct_detcomprob2003 a,
	(select b.cabcomprobnumero,a.cabprovinumero,a.detproviglosa
	from 
		[sanildefonso].dbo.co_detprovi2003 a,
		[prueba_contaprueba_sanil].dbo.ct_cabcomprob2003 b
	where 
		a.cabprovimes=4 and b.cabcomprobmes=4 and
		cast(a.cabprovinumero as varchar(20))=cast(b.cabcomprobnprovi as varchar(20)) ) as ZZ
where 
	a.cabcomprobnumero=zz.cabcomprobnumero and 
	asientocodigo like '06%' and cabcomprobmes=4

/*Actualizar Glosa Cabecera San Ildefonso*/
update [prueba_contaprueba_sanil].dbo.ct_cabcomprob2003 
set cabcomprobglosa=left('Prov.Compra ' + cast(b.cabprovinumero as varchar(5)) + ' NºAux. ' + cast(b.cabprovinumaux as varchar),30)
--select a.*,b.*
from 
	[prueba_contaprueba_sanil].dbo.ct_cabcomprob2003 a,
	[server_tc].[sanildefonso].dbo.co_cabprovi2003 b
where a.cabcomprobmes=4 and
		a.cabcomprobnumero collate  Modern_Spanish_CI_AI = b.cabprovinconta collate  Modern_Spanish_CI_AI