select * from ct_cabcomprob2002 where cabcomprobnumero='0906200772'
select * from ct_detcomprob2002 where cabcomprobnumero='0906200772'

select * from ct_cabcomprob2002 where cabcomprobnumero='0906200819'
select * from ct_detcomprob2002 where cabcomprobnumero='0906200819'


select * from ct_cabcomprob2002 where cabcomprobnumero like '__064%' and isnull(cabcomprobgrabada,0)<>0
select * from ct_cabcomprob2002 where cabcomprobnumero like '__064%' and cabcomprobgrabada is null

select * from ct_detcomprob2002 where cabcomprobnumero like '__062%' and asientocodigo='062' 
		and documentocodigo='03' and 


select * from ct_cabcomprob2002 a, ct_detcomprob2002 b
where a.cabcomprobnumero like '__062%' and a.asientocodigo='062' and
		a.cabcomprobnumero=b.cabcomprobnumero


select * from ct_detcomprob2002 a
--update ct_detcomprob2002 set plantillaasientoinafecto='1'
where cabcomprobnumero like '__062%' and
		asientocodigo='062' and documentocodigo='03' and isnull(plantillaasientoinafecto,0)=0 and
		detcomprobauto=0 and cuentacodigo<>'421101' and 
      cabcomprobmes<>12

order by 1 




--update ct_cabcomprob2002 set cabcomprobgrabada='0' 
--  where cabcomprobnumero like '__064%' and isnull(cabcomprobgrabada,0)<>0