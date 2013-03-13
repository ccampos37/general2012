Use dbContabsanil

--select * from dbo.ct_ctacteanalitico2003
--delete from dbo.ct_ctacteanalitico2003
declare @anno varchar(4)
declare @sqlcad varchar(4000)

set @anno='2003'
set @sqlcad='
	Insert dbo.ct_ctacteanalitico' +@anno+ '
		(cabcomprobmes, detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo, documentocodigo, operacioncodigo, cuentacodigo, 
 		ctacteanaliticofechaconta, analiticocodigo, ctacteanaliticonumdocumento, ctacteanaliticofechadoc, ctacteanaliticoglosa, ctacteanaliticodebe, 
 		ctacteanaliticoussdebe, ctacteanaliticohaber, ctacteanaliticousshaber, ctacteanaliticocancel, ctacteanaliticofechaven,monedacodigo)
	select 
   	A.cabcomprobmes,detcomprobitem, A.cabcomprobnumero, A.subasientocodigo, A.asientocodigo,documentocodigo, operacioncodigo, cuentacodigo, 
   	B.cabcomprobfeccontable, analiticocodigo, detcomprobnumdocumento,detcomprobfechaemision,detcomprobglosa, detcomprobdebe,
   	detcomprobussdebe, detcomprobhaber,detcomprobusshaber,null,detcomprobfechavencimiento,monedacodigo
	from
   	ct_detcomprob' +@anno+ ' A ,ct_cabcomprob' +@anno+ ' B  
	where 
		 A.cabcomprobnumero=B.cabcomprobnumero and 
     	not (A.analiticocodigo =''00'' or A.analiticocodigo is null or rtrim(A.analiticocodigo)='''' ) and 
     	not (A.documentocodigo =''00'' or  A.documentocodigo is null or rtrim(A.documentocodigo)='''' )  and  
     	not (rtrim(A.detcomprobnumdocumento)=''''  or A.detcomprobnumdocumento is null)'

exec(@sqlcad)


--select count(*) from dbo.ct_ctacteanalitico2002