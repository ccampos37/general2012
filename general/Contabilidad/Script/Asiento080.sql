/*Borra Registros del Comprobante*/
--delete ct_cabcomprob2002 where cabcomprobnumero='0108000001'

/*Genera Cabezera del Comprobante*/
INSERT ct_cabcomprob2002 (cabcomprobmes,cabcomprobnumero,cabcomprobfeccontable,subasientocodigo,usuariocodigo,estcomprobcodigo,asientocodigo,cabcomprobobservaciones,fechaact,cabcomprobglosa,cabcomprobtotdebe,cabcomprobtothaber,cabcomprobtotussdebe,cabcomprobtotusshaber,cabcomprobgrabada,cabcomprobnref,cabcomprobnlibro)
select month(fecha) as cabcomprobmes, 
  '0108000001' as cabcomprobnumero ,
   fecha as cabcomprobfeccontable,
  '0003' as subasientocodigo,'SISTEMA' as usuariocodigo,'01' as estcomprobcodigo, 
  '080' as asientocodigo , 'Asiento Apertura' as cabcomprobobservaciones   ,'01/01/2002' as fechaact, 
  'Asiento Apertura' as cabcomprobglosa, ROUND(sum(isnull(debe,0) ),2) as cabcomprobtotdebe, ROUND(sum(isnull(haber,0)),2) as cabcomprobtothaber, 
  ROUND(sum(isnull(debeD,0)),2) as cabcomprobtotussdebe, ROUND(sum(isnull(haberD,0)),2) as cabcomprobtotusshaber, 
  '0' as cabcomprobgrabada, '0' as cabcomprobnref, '0' as cabcomprobnlibro
  from dbo.mov80 
  group by fecha,mon


/*Inserta Detalle del Comprobante*/
INSERT ct_detcomprob2002 (cabcomprobmes, cabcomprobnumero, subasientocodigo, analiticocodigo, asientocodigo, detcomprobitem, monedacodigo, centrocostocodigo, documentocodigo, operacioncodigo, cuentacodigo, detcomprobnumdocumento, detcomprobfechaemision, detcomprobfechavencimiento, detcomprobglosa, detcomprobdebe, detcomprobhaber, detcomprobusshaber, detcomprobussdebe, detcomprobtipocambio, detcomprobruc, detcomprobauto, detcomprobformacambio, detcomprobajusteuser, plantillaasientoinafecto)
SELECT 
  month(fecha) as cabcomprobmes,
  	'0108000001'  as cabcomprobnumero,
  	'0003' as subasientocodigo, '00'  as analiticocodigo, 
	'080' as asientocodigo, item as detcomprobitem, '01' as monedacodigo, 
  	'00' as centrocostocodigo, '00' as documentocodigo, '01' as operacioncodigo, cuenta as cuentacodigo, '' as detcomprobnumdocumento, fecha as detcomprobfechaemision, fecha as detcomprobfechavencimiento,
  	isnull(GLOSA,'') as detcomprobglosa, ROUND(isnull(debe,0) ,2) as detcomprobdebe, ROUND(isnull(haber,0),2) as detcomprobhaber, ROUND(isnull(haber,0)/tc,2) as detcomprobusshaber, ROUND(isnull(debe,0)/tc,2) as detcomprobussdebe,
  	round(tc,4) as detcomprobtipocambio, '' as detcomprobruc, '0' as detcomprobauto, '0' as detcomprobformacambio, '0' as detcomprobajusteuser , '0' as plantillaasientoinafecto 
  FROM mov80