SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






--Declare 
--Author: Fernando Cossio

ALTER        Proc co_listagastos_rpt
   @BaseCompra varchar(50),
   @BaseConta  varchar(50),
   @BaseVenta  varchar(50),
   @Prove      varchar(2),
   @Ano        varchar(4), 
   @flagfecha  varchar(1),
   @Fechaini   Float,
   @fechafin   Float,
   @cuenta     varchar(20), 
   @flagtipo   varchar(1)='0'

As

Declare 
       @XFechaini  varchar(12),
       @XFechaFin  varchar(12),
       @SqlCad varchar(4000)

If @flagfecha=0 
  Begin
     Set    @XFechaini=cast (@Fechaini as varchar(12))  
     Set    @XFechaFin=Cast(@fechafin as varchar(12))
  End 
Else
  Begin
  Set  @XFechaini=' null '
  Set  @XFechaFin=' null '	
End 

If @flagtipo=0 
   Set  @SqlCad='
     Select A.cabprovinumero,A.cabprovinumaux,
       A.cabprovimes,A.proveedorcodigo,A.cabprovirznsoc,
       A.modoprovicod,
       D.modoprovidesc,
       A.documetocodigo,A.cabprovinumdoc,A.cabprovifchdoc,
       A.cabprovifchconta,A.monedacodigo,A.cabprovitipcambio,
       B.gastoscodigo,C.gastosdescripcion,   
       detproviimpbru=Case When G.tdocumentotipo=''A'' then  B.detproviimpbru * -1 else B.detproviimpbru end,
       detproviimpigv=Case When G.tdocumentotipo=''A'' then  B.detproviimpigv * -1 else B.detproviimpigv end, 
       detproviimpina=Case When G.tdocumentotipo=''A'' then  B.detproviimpina * -1 else B.detproviimpina end, 
       detprovitotal= Case When G.tdocumentotipo=''A'' then  B.detprovitotal * -1 else B.detprovitotal end
     from 
       ['+@BaseCompra+'].dbo.co_cabprovi'+@Ano+' A 
       inner join ['+@BaseCompra+'].dbo.co_detprovi'+@Ano+' B
         on A.cabprovinumero=B.cabprovinumero
       inner join ['+@BaseCompra+'].dbo.co_gastos C 
         on  B.gastoscodigo =c.gastoscodigo    
       inner join ['+@BaseCompra+'].dbo.co_modoprovi D     
         on  A.modoprovicod=D.modoprovicod       
       inner join  ['+@BaseVenta+'].dbo.cp_tipodocumento G 
         on  a.documetocodigo=g.tdocumentocodigo
 Where   
       A.modoprovicod like '''+@Prove+''' and 
       B.gastoscodigo like '''+@cuenta+''' '
--- and  (1='+@flagfecha+' or floor(cast(A.cabprovifchconta as float))  between '+@XFechaini+' and '+@XFechaFin+')'
execute(@SqlCad)

If @flagtipo=1 
   Set  @SqlCad='
     Select A.cabprovinumero,A.cabprovinumaux,
       A.cabprovimes,A.proveedorcodigo,A.cabprovirznsoc,A.modoprovicod,
       D.modoprovidesc,
       A.documetocodigo,A.cabprovinumdoc,A.cabprovifchdoc,
       A.cabprovifchconta,A.monedacodigo,A.cabprovitipcambio,
       B.gastoscodigo,C.gastosdescripcion,   
       detproviimpbru=Case When G.tdocumentotipo=''A'' then  B.detproviimpbru * -1 else B.detproviimpbru end,
       detproviimpigv=Case When G.tdocumentotipo=''A'' then  B.detproviimpigv * -1 else B.detproviimpigv end, 
       detproviimpina=Case When G.tdocumentotipo=''A'' then  B.detproviimpina * -1 else B.detproviimpina end, 
       detprovitotal= Case When G.tdocumentotipo=''A'' then  B.detprovitotal * -1 else B.detprovitotal end
     from 
       ['+@BaseCompra+'].dbo.co_cabprovi'+@Ano+' A 
       inner join ['+@BaseCompra+'].dbo.co_detprovi'+@Ano+' B
         on A.cabprovinumero=B.cabprovinumero
       inner join ['+@BaseCompra+'].dbo.co_gastos C 
         on  B.gastoscodigo =c.gastoscodigo 
       inner join ['+@BaseCompra+'].dbo.co_modoprovi D     
         on  A.modoprovicod=D.modoprovicod       
       inner join  ['+@BaseVenta+'].dbo.cp_tipodocumento G 
          on  a.documetocodigo=g.tdocumentocodigo  
   Where   
       A.modoprovicod like '''+@Prove+''' and 
       B.gastoscodigo like '''+@cuenta+''' '
--      and 
--      (1='+@flagfecha+' or floor(cast(A.cabprovifchconta as float))  between '+@XFechaini+' and '+@XFechaFin+')'
   print(@SqlCad)


--EXECUTE co_listagastos_rpt 'FOX','FOX','FOX','%%','2005','0',38700,38800,'%%',1







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

