select zz.*,b.entidadrazonsocial from
(select cabcomprobnumero, documentocodigo, operacioncodigo, 
		ctacteanaliticonumdocumento,a.analiticocodigo,ctacteanaliticofechadoc,
		ctacteanaliticodebe,ctacteanaliticohaber,
		ctacteanaliticoussdebe,ctacteanaliticousshaber,
		ctacteanaliticocancel,cuentacodigo
from ct_ctacteanalitico2002 a 
where documentocodigo in ('02','03') and 
		ctacteanaliticocancel is null and 
		left(cuentacodigo,3) like '421%') as ZZ,
v_analiticoentidad b
where
	ZZ.analiticocodigo=b.analiticocodigo
order by zz.analiticocodigo,zz.cuentacodigo,zz.ctacteanaliticonumdocumento,zz.operacioncodigo


--select * from ct_ctacteanalitico2002
/*
select * from ct_ctacteanalitico2002 where cabcomprobnumero='0201000669' 
				and ctacteanaliticonumdocumento='000-0195508'
select * from v_analiticoentidad
*/