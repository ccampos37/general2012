
ALTER  TRIGGER  tri_insertaranalitico2002  ON dbo.ct_detcomprob2002
FOR INSERT
AS

Declare @Fecha datetime
Select @Fecha=cabcomprobfeccontable 
From inserted A,ct_cabcomprob2002 B 
Where A.cabcomprobmes=B.cabcomprobmes and 
      A.asientocodigo=B.asientocodigo and 
      A.subasientocodigo=B.subasientocodigo and 
      A.cabcomprobnumero=B.cabcomprobnumero


Insert dbo.ct_ctacteanalitico2002
(cabcomprobmes, detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo, documentocodigo, operacioncodigo, cuentacodigo, 
 ctacteanaliticofechaconta, analiticocodigo, ctacteanaliticonumdocumento, ctacteanaliticofechadoc, ctacteanaliticoglosa, ctacteanaliticodebe, 
 ctacteanaliticoussdebe, ctacteanaliticohaber, ctacteanaliticousshaber, ctacteanaliticocancel, ctacteanaliticofechaven,monedacodigo)
select 
   cabcomprobmes,detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo,documentocodigo, operacioncodigo, cuentacodigo, 
   @Fecha, analiticocodigo, detcomprobnumdocumento,detcomprobfechaemision,detcomprobglosa, detcomprobdebe,
   detcomprobussdebe, detcomprobhaber,detcomprobusshaber,0,detcomprobfechavencimiento,monedacodigo
from
     inserted
where  not (analiticocodigo ='00' or analiticocodigo is null or rtrim(analiticocodigo)='' ) and 
            not (documentocodigo ='00' or  documentocodigo is null or rtrim(documentocodigo)='' )  and  
            not      (rtrim(detcomprobnumdocumento)=''  or detcomprobnumdocumento is null)
GO

ALTER  TRIGGER  tri_insertaranaliticoxxxx  ON dbo.ct_detcomprobxxxx
FOR INSERT
AS

Declare @Fecha datetime
Select @Fecha=cabcomprobfeccontable 
From inserted A,ct_cabcomprobxxxx B 
Where A.cabcomprobmes=B.cabcomprobmes and 
      A.asientocodigo=B.asientocodigo and 
      A.subasientocodigo=B.subasientocodigo and 
      A.cabcomprobnumero=B.cabcomprobnumero

Insert dbo.ct_ctacteanaliticoxxxx
(cabcomprobmes, detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo, documentocodigo, operacioncodigo, cuentacodigo, 
 ctacteanaliticofechaconta, analiticocodigo, ctacteanaliticonumdocumento, ctacteanaliticofechadoc, ctacteanaliticoglosa, ctacteanaliticodebe, 
 ctacteanaliticoussdebe, ctacteanaliticohaber, ctacteanaliticousshaber, ctacteanaliticocancel, ctacteanaliticofechaven,monedacodigo)
select 
   cabcomprobmes,detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo,documentocodigo, operacioncodigo, cuentacodigo, 
   @Fecha, analiticocodigo, detcomprobnumdocumento,detcomprobfechaemision,detcomprobglosa, detcomprobdebe,
   detcomprobussdebe, detcomprobhaber,detcomprobusshaber,0,detcomprobfechavencimiento,monedacodigo
from
     inserted
where  not (analiticocodigo ='00' or analiticocodigo is null or rtrim(analiticocodigo)='' ) and 
            not (documentocodigo ='00' or  documentocodigo is null or rtrim(documentocodigo)='' )  and  
            not      (rtrim(detcomprobnumdocumento)=''  or detcomprobnumdocumento is null)

GO

Update ct_ctacteanalitico2002
Set ctacteanaliticofechaconta=B.cabcomprobfeccontable 
From ct_ctacteanalitico2002 A,dbo.ct_cabcomprob2002 B
Where
    A.cabcomprobmes=B.cabcomprobmes and
    A.asientocodigo=B.asientocodigo and 
    A.subasientocodigo=B.subasientocodigo and 
    A.cabcomprobnumero=B.cabcomprobnumero   




