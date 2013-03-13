use contaprueba

/*Actualizar el Flag del asiento Inafecto*/
--select * from ct_detcomprob2003 
--update ct_detcomprob2003 set plantillaasientoinafecto=1
	where cabcomprobnumero in 
		(select cabcomprobnumero from ct_detcomprob2003 
			where asientocodigo like '07%' and cabcomprobmes=5 and detcomprobnumdocumento like '001%' 
				and cuentacodigo='401110' and detcomprobhaber=0)
	and cuentacodigo='701101' and detcomprobhaber>0
order by detcomprobnumdocumento


--select * from ct_detcomprob2003 
--update ct_detcomprob2003 set analiticocodigo='00000003462002'
	where cabcomprobnumero in 
		(select cabcomprobnumero from ct_detcomprob2003 
			where asientocodigo like '07%' and cabcomprobmes=5 and detcomprobnumdocumento like '001%' 
				and cuentacodigo='401110' and detcomprobhaber=0)
	and cuentacodigo='121101' and detcomprobdebe>0


--select * from ct_detcomprob2003 
--update ct_detcomprob2003 set plantillaasientoinafecto=1
	where cabcomprobnumero in 
		(select cabcomprobnumero from ct_detcomprob2003 
			where asientocodigo like '07%' and cabcomprobmes=5 
				and cuentacodigo='401110' and detcomprobhaber=0)
	and cuentacodigo='701101' and detcomprobhaber>0
order by detcomprobnumdocumento