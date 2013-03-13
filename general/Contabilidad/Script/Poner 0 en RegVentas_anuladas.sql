use contaprueba
use camtex_tinto


select * from ct_cabcomprob2002 
--update ct_detcomprob2002 
--	set detcomprobdebe=0,detcomprobhaber=0,
--		detcomprobusshaber=0,detcomprobussdebe=0
where cabcomprobnumero in
	(select cabcomprobnumero from ct_detcomprob2002 
		where asientocodigo like '07%' and
				detcomprobhaber=0.01 and 
				cuentacodigo like '70%')
--order by 2,detcomprobitem



select * from ct_detcomprob2002 where cabcomprobnumero='1007000085'

select * from ct_cabcomprob2002 
update ct_detcomprob2002 
	set detcomprobdebe=0,detcomprobhaber=0,
		detcomprobusshaber=0,detcomprobussdebe=0
where cabcomprobnumero in
	(select cabcomprobnumero from ct_detcomprob2002 
		where asientocodigo like '07%' and
				detcomprobdebe=0.01 and 
				cuentacodigo like '70%')

