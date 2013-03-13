select from ct_cabcomprob2003 where asientocodigo='014' and cabcomprobmes=1


select '1' AS X,documentocargo,cargonumdoc,monedacodigo,cargoapeimpape,cargoapefecemi 
FROM vt_cargo
WHERE cargonumdoc='00600000672'
UNION

--abonocanmoneda
SELECT '2' AS X,documentoabono,abononumdoc,abonocanmoncan,abonocanimcan,abonocanfecan 
from vt_abono  
WHERE ABONOnumdoc='00600000672'

