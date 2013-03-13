--USE MARFICE_VENTAS
create  proc cp_antigudeuda_rpt
--Declare 
@Base varchar(100),
@BaseConta varchar(100),
@op   varchar(1),
@cliente varchar(20),
@tipdoc varchar(2),
@dias varchar(5),
@simbo varchar(3),
@fecharef varchar(20)

as
/*  
	Op=1 Documentos Vencidos
    Y diferentes de 1 documentos por vencer
*/

/*Set @Base='ventas_prueba'
Set @BaseConta='contaprueba'
Set @op='1'
Set @cliente='%%'
Set @tipdoc='%%'
Set @dias='17'
Set @simbo='<'
Set @fecharef='37698'*/

Declare @SqlCad varchar(4000),
        @SqlVar varchar(4000) 

Set @SqlCad=' 
select A.clientecodigo,
       A.cargoapefecvct,A.monedacodigo,
       A.documentocargo,DescDoc=isnull(D.tdocumentodesccorta,''No existe Descripcion''),	  	
       cargoapeimpape=isnull(A.cargoapeimpape,0),
       cargoapeimppag=isnull(A.cargoapeimppag,0)  ,C.tipocambioventa,
       totalMdolar=case when A.monedacodigo=''02'' then  isnull(A.cargoapeimpape,0) else 0 end, 
       totalMSoles=case when A.monedacodigo=''01'' then  isnull(A.cargoapeimpape,0) else 0 end,
	   totalDdolar=case when A.monedacodigo=''02'' then  isnull(A.cargoapeimppag,0) else 0 end, 
       totalDSoles=case when A.monedacodigo=''01'' then  isnull(A.cargoapeimppag,0) else 0 end,	
       B.clienteruc,B.clienterazonsocial,
       Dias=datediff(day,'+@fecharef+',floor(cast(A.cargoapefecvct as real)))
        
from 
	 ['+@base+'].dbo.cp_cargo A
     Inner join ['+@base+'].dbo.cp_proveedor B   
     on A.clientecodigo=B.clientecodigo
     left outer join ['+@BaseConta+'].dbo.ct_tipocambio C  
     on A.cargoapefecemi=C.tipocambiofecha 
     left outer join ['+@base+'].dbo.cp_tipodocumento D
     on A.documentocargo=D.tdocumentocodigo 
       
     
Where      	 
    A.cargoapeflgreg is null and 
    isnull(A.cargoapeflgcan,0) <>1   and   
    A.clientecodigo like '''+@cliente+''' and 
    A.documentocargo like '''+@tipdoc+'''' 
    --Para documentos vencidos 

if @op='1' 
    Set @SqlVar=' and floor(cast(A.cargoapefecvct as real)) <'+@fecharef
Else
Begin  
	if ltrim(rtrim(isnull(@dias,'')))='' 
		Set @SqlVar=' and floor(cast(A.cargoapefecvct as real)) >='+@fecharef
    Else
		Set @SqlVar=' and datediff(day,'+@fecharef+',floor(cast(A.cargoapefecvct as real)))>=0 and 
                          datediff(day,'+@fecharef+',floor(cast(A.cargoapefecvct as real))) '+@simbo+' '+@dias
End 
set @SqlCad=@SqlCad+@SqlVar
EXEC (@SqlCad+@SqlVar)



