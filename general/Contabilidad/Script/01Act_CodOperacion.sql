--select * from ct_detcomprob2002 where
--asientocodigo in ('060','061','062','064') and operacioncodigo='00'


--update ct_ctacteanalitico2002 set ctacteanaliticocancel=null
--select * from ct_ctacteanalitico2002

/*Actualiza los Códigos de Operación*/
update ct_detcomprob2002 set operacioncodigo='01' where
   asientocodigo in ('060','061','062','064') and operacioncodigo='00'

update ct_detcomprob2002 set operacioncodigo='01' where
   asientocodigo in ('070','071','072','073','074') and operacioncodigo='00'

update ct_detcomprob2002 set operacioncodigo='03' where
   asientocodigo in ('010') and operacioncodigo='00'
