
select count(*) as Num_Recibos,ZZ.detrec_tipodoc_concepto,c.conceptodescripcion from 
(select * from te_detallerecibos
where cabrec_numrecibo in 
(select a.cabrec_numrecibo 
	from te_cabecerarecibos a, te_operaciongeneral b
where 
   a.operacioncodigo=b.operacioncodigo and
	b.operacioncontrolaclienteprov='X')
and 
 detrec_tipodoc_concepto not in ('80','90') ) as ZZ,
 te_conceptocaja c
where ZZ.detrec_tipodoc_concepto=c.conceptocodigo
group by ZZ.detrec_tipodoc_concepto,c.conceptocodigo,c.conceptodescripcion 
order by 2

