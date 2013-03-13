select * from  te_cabecerarecibos where cabrec_numreciboegreso='000003'
Select cabrec_numrecibo from te_cabecerarecibos where cabrec_numreciboegreso='000003'


select * from te_cabecerarecibos

--201401
--201508
--201536

delete from te_cabecerarecibos where cabrec_numrecibo='201536'


select cabrec_numrecibo,cabrec_numreciboegreso,MontoSoles=cast(Round(cabrec_totsoles,2) as numeric(15,2)),
  MontoDolar=cast(Round(cabrec_totdolares,2) as numeric(15,2)) 
	from te_cabecerarecibos where cabrec_numreciboegreso in
(select cabrec_numreciboegreso
from 
(select count(*) as Valor,cabrec_numreciboegreso from te_cabecerarecibos 
	where cabrec_numreciboegreso<>''
	group by cabrec_numreciboegreso
	having count(*)<>2) as ZZ)
order by 2


Select cabrec_numrecibo from te_cabecerarecibos where cabrec_numreciboegreso='200629'
select a.monedacodigo,b.detrec_monedacancela,sum(b.detrec_importesoles) as detrec_importesoles, sum(b.detrec_importedolares) as detrec_importedolares, a.cabrec_tipocambio  FROM te_cabecerarecibos a, te_detallerecibos b WHERE a.cabrec_numrecibo=b.cabrec_numrecibo AND a.cabrec_numrecibo='200629' Group by a.monedacodigo,b.detrec_monedacancela,a.cabrec_tipocambio


--select * into bkte_detallerecibos from te_detallerecibos
delete from te_cabecerarecibos where cabrec_numrecibo in
(select cabrec_numrecibo
	from te_cabecerarecibos where cabrec_numreciboegreso in
(select cabrec_numreciboegreso
from 
(select count(*) as Valor,cabrec_numreciboegreso from te_cabecerarecibos 
	where cabrec_numreciboegreso<>''
	group by cabrec_numreciboegreso
	having count(*)=1) as ZZ))
order by cabrec_numreciboegreso
