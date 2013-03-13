insert ct_entidad
(entidadcodigo,entidadruc,entidadrazonsocial,usuariocodigo,fechaact)
select distinct left(analiticocodigo,11),left(analiticocodigo,11),' ','sa',getdate() from ct_detcomprob2008 
where cuentacodigo='421100' 
and asientocodigo='019' and left(analiticocodigo,11) not in ( select entidadcodigo from ct_entidad )

insert ct_analitico
(analiticocodigo,entidadcodigo,tipoanaliticocodigo)
select distinct analiticocodigo,left(analiticocodigo,11),'001' from ct_detcomprob2008 
where cuentacodigo='421100' 
and asientocodigo='019' and analiticocodigo not in ( select analiticocodigo from ct_analitico )
