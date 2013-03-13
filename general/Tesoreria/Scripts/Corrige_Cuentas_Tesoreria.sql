select detrec_cajabanco1,detrec_numctacte,detrec_monedacancela,count(*) from te_detallerecibos
where detrec_tipocajabanco='B'
group by detrec_cajabanco1,detrec_numctacte,detrec_monedacancela


select 	a.cabrec_numrecibo,
			b.clientecodigo,
			c.clienterazonsocial,
			b.operacioncodigo,
			b.monedacodigo,
			a.detrec_numdocumento,
			a.detrec_fechacancela,
			a.detrec_monedacancela,
			a.detrec_importesoles,
			a.detrec_importedolares
from 
	te_detallerecibos a,
	te_cabecerarecibos b,
	cp_proveedor c
where 
	detrec_tipocajabanco='B' and 
	detrec_cajabanco1='' and
	detrec_estadoreg<>1 and
	a.cabrec_numrecibo=b.cabrec_numrecibo and
	b.clientecodigo=c.clientecodigo
order by a.cabrec_numrecibo

select * from 	cp_proveedor c




select cbanco_codigo,monedacodigo,cbanco_numero from te_cuentabancos

select * from te_detallerecibos
where detrec_cajabanco1='01' and detrec_numctacte='191-1148049-1-63' and detrec_monedacancela='01'
update te_detallerecibos set detrec_monedacancela='02' 
where  detrec_cajabanco1='01' and detrec_numctacte='191-1148049-1-63' and detrec_monedacancela='01'



select * from te_detallerecibos


/*
200552
200553
200555
200558
200559
100077
*/

delete from te_cabecerarecibos where cabrec_numrecibo='100077'
select * from cp_abono where abononumdoc like 'H00362631'
--delete from cp_abono where abononumdoc like 'H00362631'
select * from cp_cargo where cargonumdoc like 'H00362631'
update cp_cargo set cargoapeimppag=0, cargoapeflgcan=0 where cargonumdoc like 'H00362631'

--00100000034
--00100024646
--00100004966
--00100004975
--00100007015
--12100093466
--H00362631