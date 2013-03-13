select cabcomprobmes,cabcomprobnumero,cuentacodigo,documentocodigo,detcomprobnumdocumento,
	detcomprobfechaemision,detcomprobdebe,detcomprobhaber,detcomprobtipocambio
 from ct_detcomprob2002 where analiticocodigo='20264474702002'
order by cabcomprobmes,cabcomprobnumero




select  *
from ct_detcomprob2002

select * from ct_entidad where entidadrazonsocial like '%LyCR%'