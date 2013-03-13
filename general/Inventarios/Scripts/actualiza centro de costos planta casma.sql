
update planta_casma.dbo.movalmdet 
set decencos=z.equivalencia
from planta_casma.dbo.movalmdet,
(
select d.equivalencia,empresa=right(rtrim(b.id_empresa),2),a=right('00000000000'+rtrim(cast(a.registrosalida_id as char(10))),11),
a.producto_id,precio_unitario
from detregistrosalida a 
left join registrosalida b on  a.registrosalida_id=b.registrosalida_id
left join centro_costo d on a.id_centro_costo=d.id_centro_costo
where a.cantidadretirada > 0 
) as z

where deALMA=right(rtrim(z.empresa),2) and deTD='NS' and deNUMDOC=a and decodigo=z.producto_id
