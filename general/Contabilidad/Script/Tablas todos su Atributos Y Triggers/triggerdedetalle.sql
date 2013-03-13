CREATE TRIGGER tri_asientoauto2002  ON dbo.ct_detcomprob2002
FOR INSERT
AS

--Actualizar la Cabecera
Declare 
@sub varchar(4) ,@asi varchar(3),@mes float,@comp varchar(10) ,
@ad numvalor,@ah numvalor,@ads numvalor,@ahs numvalor     	

Select distinct @mes=cabcomprobmes, @comp=cabcomprobnumero, @sub=subasientocodigo, 
                      @asi=asientocodigo from ct_detcomprob2002

--Calculando el total de las cabeceras
select  @ad=isnull(sum(detcomprobdebe),0), @ah =isnull(sum(detcomprobhaber),0),@ahs= isnull(sum(detcomprobusshaber),0), @ads=isnull(sum(detcomprobussdebe),0) 
from dbo.ct_detcomprob2002 
where  cabcomprobmes=@mes and  cabcomprobnumero=@comp and  
            asientocodigo=@asi and  subasientocodigo=@sub 


Update dbo.ct_cabcomprob2002
set  cabcomprobtotdebe=@ad,
       cabcomprobtothaber=@ah, 
       cabcomprobtotussdebe=@ads,
       cabcomprobtotusshaber=@ahs,
      estcomprobcodigo='03'
where  cabcomprobmes=@mes and  cabcomprobnumero=@comp and  
            asientocodigo=@asi and  subasientocodigo=@sub
go

CREATE TRIGGER  tri_insertaranalitico2002  ON [dbo].[ct_detcomprob2002] 
FOR INSERT
AS
Insert dbo.ct_ctacteanalitico2002
(cabcomprobmes, detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo, documentocodigo, operacioncodigo, cuentacodigo, 
 ctacteanaliticofechaconta, analiticocodigo, ctacteanaliticonumdocumento, ctacteanaliticofechadoc, ctacteanaliticoglosa, ctacteanaliticodebe, 
 ctacteanaliticoussdebe, ctacteanaliticohaber, ctacteanaliticousshaber, ctacteanaliticocancel, ctacteanaliticofechaven)
select 
   cabcomprobmes,detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo,documentocodigo, operacioncodigo, cuentacodigo, 
   detcomprobfechaemision, analiticocodigo, detcomprobnumdocumento,detcomprobfechaemision,detcomprobglosa, detcomprobdebe,
   detcomprobussdebe, detcomprobhaber,detcomprobusshaber,0,detcomprobfechavencimiento
from
     inserted
where  not (analiticocodigo ='00' or analiticocodigo is null or rtrim(analiticocodigo)='' ) and 
            not (documentocodigo ='00' or  documentocodigo is null or rtrim(documentocodigo)='' )  and  
            not      (rtrim(detcomprobnumdocumento)=''  or detcomprobnumdocumento is null)


