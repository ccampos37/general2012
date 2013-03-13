/*
select * from contaprueba.dbo.ct_cabcomprob2002 where asientocodigo='012' and cabcomprobmes=12
select Recibo=substring(cabcomprobnprovi,4,6) from 
	contaprueba.dbo.ct_detcomprob2002 a,
	contaprueba.dbo.ct_cabcomprob2002 b
where a.asientocodigo='012' and a.cuentacodigo like '104%' and a.cabcomprobmes=12 and
	a.cabcomprobnumero=b.cabcomprobnumero
order by 1

/*tipdocref,detcomprobnumref */
select *
from [contaprueba].dbo.ct_detcomprob2002
where 


select detrec_tdqc,detrec_ndqc from ventas_prueba.dbo.te_detallerecibos 
where cabrec_numrecibo='201360'

select * from ventas_prueba.dbo.te_cabecerarecibos 
where cabrec_numrecibo='201384'

select * from ventas_prueba.dbo.te_detallerecibos 
where cabrec_numrecibo='201384'

select * from ventas_prueba.dbo.te_detallerecibos
*/

select * from [pc06].[contaprueba].dbo.ct_detcomprob2002


--update [contaprueba].dbo.ct_detcomprob2002 set 
--	tipdocref=YY.detrec_tdqc, detcomprobnumref=YY.detrec_ndqc

/*Actualizar los Recibos de Egresos*/
update [pc06].[contaprueba].dbo.ct_detcomprob2002 set 
	documentocodigo=YY.detrec_tdqc, detcomprobnumdocumento=YY.detrec_ndqc
from
	[contaprueba].dbo.ct_detcomprob2002 H,
(select a.cabcomprobmes,a.cabcomprobnumero,zz.cabrec_numrecibo,zz.detrec_tdqc,zz.detrec_ndqc 
from contaprueba.dbo.ct_cabcomprob2002 a, 
(select distinct cabrec_numrecibo,detrec_tdqc,detrec_ndqc from ventas_prueba.dbo.te_detallerecibos  
where cabrec_numrecibo in
(select cabrec_numrecibo=substring(cabcomprobnprovi,4,6) from 
	contaprueba.dbo.ct_detcomprob2002 a,
	contaprueba.dbo.ct_cabcomprob2002 b
	where a.asientocodigo='012' and a.cuentacodigo like '104%' and a.cabcomprobmes=12 and
		a.cabcomprobnumero=b.cabcomprobnumero) 
and detrec_tdqc='59') as ZZ
where substring(a.cabcomprobnprovi,4,6)=zz.cabrec_numrecibo ) as YY
where H.cabcomprobnumero=YY.cabcomprobnumero and H.cuentacodigo like '104%' 


/*Actualizar los Transferencias*/
--update [contaprueba].dbo.ct_detcomprob2002 set 
--	documentocodigo=YY.detrec_tdqc, detcomprobnumdocumento=YY.detrec_ndqc


select * from camtex_tinto.dbo.te_detallerecibos 
select * from camtex_tinto.dbo.te_cabecerarecibos


--select ww.* from
select numcomprob=h.cabcomprobnumero,ww.*,h.cuentacodigo
--update [pc06].[contaprueba].dbo.ct_detcomprob2002 
--set 
--	detcomprobnumdocumento=WW.detrec_numdocumento
from
	[pc06].[contaprueba].dbo.ct_detcomprob2002 H,
(select zz.*,yy.*
from
(select a.cabrec_numrecibo,cabrec_numreciboegreso,b.detrec_numdocumento
from 
	camtex_tinto.dbo.te_cabecerarecibos a,
	camtex_tinto.dbo.te_detallerecibos b
where 
	a.cabrec_numrecibo=b.cabrec_numrecibo and
	b.detrec_tipodoc_concepto='90') as ZZ,
(select cabrec_recibo=substring(cabcomprobnprovi,4,6),a.cabcomprobnumero from 
	[pc06].[contaprueba].dbo.ct_detcomprob2002 a,
	[pc06].[contaprueba].dbo.ct_cabcomprob2002 b
	where a.asientocodigo='013' and a.cuentacodigo like '104%' and a.cabcomprobmes=12 and
		a.cabcomprobnumero=b.cabcomprobnumero) as YY
where zz.cabrec_numreciboegreso=yy.cabrec_recibo) as ww
where H.cabcomprobnumero=ww.cabcomprobnumero and H.cuentacodigo like '104%' and detcomprobdebe=0



--select * from [pc06].contaprueba.dbo.ct_detcomprob2002 where cabcomprobnumero='1201300098'
--detrec_tipodoc_concepto 
--detrec_numdocumento
