Insert dbo.ct_ctacteanalitico2003
(cabcomprobmes, detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo, documentocodigo, operacioncodigo, cuentacodigo, 
 ctacteanaliticofechaconta, analiticocodigo, ctacteanaliticonumdocumento, ctacteanaliticofechadoc, ctacteanaliticoglosa, ctacteanaliticodebe, 
 ctacteanaliticoussdebe, ctacteanaliticohaber, ctacteanaliticousshaber, ctacteanaliticocancel, ctacteanaliticofechaven,monedacodigo,ctacteanaliticosaldo)
select 
	cabcomprobmes, detcomprobitem, cabcomprobnumero='00'+left(cabcomprobnumero,2)+substring(cabcomprobnumero,4,2)+right(cabcomprobnumero,4), 
	subasientocodigo, asientocodigo, documentocodigo, operacioncodigo, cuentacodigo, 
 	ctacteanaliticofechaconta, analiticocodigo, ctacteanaliticonumdocumento, 
	ctacteanaliticofechadoc,ctacteanaliticoglosa, ctacteanaliticodebe, 
 	ctacteanaliticoussdebe, ctacteanaliticohaber, ctacteanaliticousshaber, ctacteanaliticocancel, ctacteanaliticofechaven,monedacodigo,ctacteanaliticosaldo='2002'
	from
		dbo.ct_ctacteanalitico2002
	where  
		ctacteanaliticocancel is null and
		cabcomprobmes=12
		

select '00'+left(cabcomprobnumero,2)+substring(cabcomprobnumero,4,2)+right(cabcomprobnumero,4)   from dbo.ct_ctacteanalitico2002 where cabcomprobnumero='1207203108'




select 
   cabcomprobmes,detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo,documentocodigo, operacioncodigo, cuentacodigo, 
   '31/12/2002', analiticocodigo, detcomprobnumdocumento,detcomprobfechaemision,detcomprobglosa, detcomprobdebe,
   detcomprobussdebe, detcomprobhaber,detcomprobusshaber,null,detcomprobfechavencimiento,monedacodigo
from
     ct_detcomprob2002
where  not (analiticocodigo ='00' or analiticocodigo is null or rtrim(analiticocodigo)='' ) and 
            not (documentocodigo ='00' or  documentocodigo is null or rtrim(documentocodigo)='' )  and  
            not      (rtrim(detcomprobnumdocumento)=''  or detcomprobnumdocumento is null) and


--select * from 	dbo.ct_ctacteanalitico2003 where ctacteanaliticosaldo=2002