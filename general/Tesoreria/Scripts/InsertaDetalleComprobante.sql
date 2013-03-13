--select * from contaprueba.dbo.ct_cabcomprob2002 where cabcomprobnumero='1201204824'


Insert Into [Contaprueba].dbo.ct_detcomprob2002
(cabcomprobmes, cabcomprobnumero, subasientocodigo, analiticocodigo, asientocodigo,
 detcomprobitem, monedacodigo, centrocostocodigo, documentocodigo, operacioncodigo,
 cuentacodigo, detcomprobnumdocumento, detcomprobfechaemision, detcomprobfechavencimiento,
 detcomprobglosa, detcomprobdebe, detcomprobhaber, detcomprobusshaber, detcomprobussdebe,
 detcomprobtipocambio, detcomprobruc, detcomprobauto, detcomprobformacambio,
 detcomprobajusteuser, plantillaasientoinafecto, tipdocref,
 detcomprobnumref, detcomprobconci, detcomprobnlibro, detcomprobfecharef)


--select len(ZZ.detcomprobglosa) from 

Select Distinct
       A.cabcomprobmes, 
       comprobnumero='12'+A.asientocodigo+replicate('0',5-len(B.cabcomprobnumero))+ltrim(rtrim(cast(B.cabcomprobnumero as varchar(20)))),
       A.subasientocodigo,
       analitico=case when left(A.cuentacodigo,2)='42' then 
                         case Rtrim(ltrim(isnull(cli.clienteruc,''))) 
                         when '00000000000' then cli.clientecodigo 
                         When '' Then cli.clientecodigo
                         else  cli.clienteruc end + '001'
                   Else '00' end,
       A.asientocodigo,
       A.detcomprobitem as detcomprobitem, A.monedacodigo, A.centrocostocodigo, A.documentocodigo, A.operacioncodigo,
       A.cuentacodigo as cuentacodigo, A.detcomprobnumdocumento, A.detcomprobfechaemision, A.detcomprobfechavencimiento,
       left(A.detcomprobglosa,50) as detcomprobglosa, A.detcomprobdebe, A.detcomprobhaber, A.detcomprobusshaber, A.detcomprobussdebe,
       A.detcomprobtipocambio,
		 detcomprobruc=isnull(cli.clienteruc,'')  , A.detcomprobauto, A.detcomprobformacambio,
       A.detcomprobajusteuser, A.plantillaasientoinafecto,
       tipdocref= case when rtrim(isnull(A.tipdocref,'00'))='' then '00' else isnull(A.tipdocref,'00') end,
       A.detcomprobnumref, A.detcomprobconci, 
       comprobnlibro='12012'+replicate('0',5-len(1+B.correlibro))+ltrim(rtrim(cast(1+B.correlibro as varchar(10))))
       , A.detcomprobfecharef
from [##tmpgenasientodetDesarrollo3] A, 
     [##tmpgenasientocabDesarrollo3] B,[Ventas_Prueba].dbo.cp_proveedor cli      
Where ltrim(rtrim(A.numprovi))=rtrim(ltrim(B.numprovi)) 
      and A.analiticocodigo=Cli.clientecodigo /*) as ZZ*/
---order by 2

--delete  from ##tmpgenasientodetDesarrollo3 where cuentacodigo=''



--delete from ##tmpgenasientodetDesarrollo3 where cuentacodigo='471110'