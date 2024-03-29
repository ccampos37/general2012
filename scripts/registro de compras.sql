USE [marfice]
GO
/****** Object:  StoredProcedure [dbo].[ct_LibroRegistroCompras_rpt]    Script Date: 10/28/2011 09:13:28 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


/*
drop  Proc ct_registrocompras_rpt
exec ct_LibroRegistroCompras_rpt 'planta_casma','03','2011','09','060,009','33%34%60%62%63%64%65%67%','40111,40112','401902,40116,','401700,'
*/

ALTER Proc [dbo].[ct_LibroRegistroCompras_rpt]

--Declare 
@BASE         VARCHAR(100),
@empresa varchar(2),
@ANNO         VARCHAR(4),
@MES          VARCHAR(2),  
@ASIENTOSPLAN   VARCHAR(500),
@CTASPLANCOMP VARCHAR(500),
@CTASIGV      VARCHAR(200),
@CTASIES      VARCHAR(100),
@CTASRENTA    VARCHAR(100)
AS
Declare 
@sqlvar varchar(8000),@sqlvar1 varchar(8000),
@CADASIENTOSPLAN   VARCHAR(1000),
@CADCTASPLANCOMP VARCHAR(2000),
@CADCTASIGV      VARCHAR(1000),
@CADCTASIES      VARCHAR(500),
@CADCTASRENTA    VARCHAR(500)
Set @CADASIENTOSPLAN='('+dbo.fn_ArmaCriterio(@ASIENTOSPLAN,',','')+')'
Set @CADCTASPLANCOMP='('+dbo.fn_ArmaCriterio(@CTASPLANCOMP,'%','F.cuentacodigo')+')'
Set @CADCTASIGV ='('+dbo.fn_ArmaCriterio(@CTASIGV,',','F.cuentacodigo')+')'
Set @CADCTASIES='('+dbo.fn_ArmaCriterio(@CTASIES,',','F.cuentacodigo')+')'
Set @CADCTASRENTA='('+dbo.fn_ArmaCriterio(@CTASRENTA,',','F.cuentacodigo')+')'
Set @sqlvar=''+
'SELECT A.cabcomprobmes,A.cabcomprobnumero,A.subasientocodigo,A.asientocodigo,H.asientodescripcion, 
        A.analiticocodigo,G.entidadruc,G.entidadrazonsocial,A.monedacodigo,
        documentocodigo=isnull(A.documentocodigo,''''),T.documentodescripcion,
        serie=left(isnull(a.detcomprobnumdocumento,''''),4),
        detcomprobnumdocumento=Right(isnull(A.detcomprobnumdocumento,''''),10),
        A.detcomprobfechaemision,a.detcomprobfechavencimiento,
	baseimpgrab=Isnull((Select isnull((case when isnull(Z.cabcomprobgrabada,0)=1 
                   then case when sum(F.detcomprobdebe)>0 then sum(F.detcomprobdebe) 
                             else sum(F.detcomprobhaber) * -1 end end),0)
              FROM ['+@base+'].dbo.ct_detcomprob'+@anno+' F
               Where a.empresacodigo=f.empresacodigo and A.cabcomprobmes=F.cabcomprobmes and f.detcomprobauto=0 and
	             A.cabcomprobnumero=F.cabcomprobnumero and A.asientocodigo=F.asientocodigo and 
	             A.subasientocodigo=F.subasientocodigo and A.documentocodigo=F.documentocodigo and 
	             A.detcomprobnumdocumento=F.detcomprobnumdocumento And F.asientocodigo IN '+@CADASIENTOSPLAN+' AND 
                     F.plantillaasientoinafecto=0 and '+@CADCTASPLANCOMP + 
                    ' Group By F.cabcomprobnumero,F.asientocodigo,F.subasientocodigo,F.documentocodigo,F.detcomprobnumdocumento),0) , 
        baseimpnograb=Isnull((Select isnull((case when isnull(Z.cabcomprobgrabada,0)=0 
                      then case when sum(F.detcomprobdebe)>0 then sum(F.detcomprobdebe) 
                             else sum(F.detcomprobhaber) * -1 end end),0)                
		      FROM ['+@base+'].dbo.ct_detcomprob'+@anno+' F
                      Where a.empresacodigo=f.empresacodigo and A.cabcomprobmes=F.cabcomprobmes and f.detcomprobauto=0 and 
				      A.cabcomprobnumero=F.cabcomprobnumero and A.asientocodigo=F.asientocodigo and 
				      A.subasientocodigo=F.subasientocodigo and A.documentocodigo=F.documentocodigo and 
			          A.detcomprobnumdocumento=F.detcomprobnumdocumento And F.asientocodigo IN '+@CADASIENTOSPLAN+' AND 
        F.plantillaasientoinafecto=0 and '+                  
		@CADCTASPLANCOMP +
        ' Group By F.cabcomprobnumero,F.asientocodigo,F.subasientocodigo,F.documentocodigo,F.detcomprobnumdocumento),0),	            
        montoinafecto=isnull((SELECT TOP 1
                      Case when F.plantillaasientoinafecto=1 then 
                                isnull((case when sum(F.detcomprobdebe)>=0 then sum(F.detcomprobdebe) 
                                       else sum(F.detcomprobhaber) * -1 end ),0)
                      Else 0 end
       			FROM ['+@BASE+'].dbo.ct_detcomprob'+@ANNO+' F
        			WHERE F.plantillaasientoinafecto=1 and  not('+@CADCTASIES+')
					and F.detcomprobauto=0 and a.empresacodigo=f.empresacodigo and A.cabcomprobmes=F.cabcomprobmes and 
				      A.cabcomprobnumero=F.cabcomprobnumero and A.asientocodigo=F.asientocodigo and 
				      A.subasientocodigo=F.subasientocodigo and A.documentocodigo=F.documentocodigo and 
			          A.detcomprobnumdocumento=F.detcomprobnumdocumento   
         Group By F.cabcomprobnumero,F.asientocodigo,F.subasientocodigo,F.documentocodigo,F.detcomprobnumdocumento,F.plantillaasientoinafecto),0), 
       	igvimpgrab=ISNULL(( SELECT top 1 isnull((case when isnull(Z.cabcomprobgrabada,0)=1 
                   			then case when F.detcomprobdebe>0 then F.detcomprobdebe 
                            	 else F.detcomprobhaber * -1 end end),0)         	
       				FROM ['+@BASE+'].dbo.ct_detcomprob'+@ANNO+' F
       	 			WHERE a.empresacodigo=f.empresacodigo and A.cabcomprobmes=F.cabcomprobmes and 
				        A.cabcomprobnumero=F.cabcomprobnumero and A.asientocodigo=F.asientocodigo and 
				        A.subasientocodigo=F.subasientocodigo and A.documentocodigo=F.documentocodigo and 
			             '+@CADCTASIGV+'),0),
        igvimpnograb=ISNULL(( SELECT top 1 isnull((case when isnull(Z.cabcomprobgrabada,0)=0 
                                    then case when F.detcomprobdebe>0 then F.detcomprobdebe 
                                    else F.detcomprobhaber * -1 end end),0)
       	  			FROM ['+@BASE+'].dbo.ct_detcomprob'+@ANNO+' F
       	  			WHERE a.empresacodigo=f.empresacodigo and  A.cabcomprobmes=F.cabcomprobmes and 
				        A.cabcomprobnumero=F.cabcomprobnumero and A.asientocodigo=F.asientocodigo and 
				        A.subasientocodigo=F.subasientocodigo and A.documentocodigo=F.documentocodigo and             
			            '+@CADCTASIGV+'),0),
        IES=ISNULL(( SELECT top 1
         	          isnull(case when F.detcomprobdebe>0 then F.detcomprobdebe 
                             else F.detcomprobhaber * -1 end,0)         	       		
       	  			FROM ['+@BASE+'].dbo.ct_detcomprob'+@ANNO+' F
       	  			WHERE A.cabcomprobmes=F.cabcomprobmes and A.cabcomprobnumero=F.cabcomprobnumero and 
				      	A.asientocodigo=F.asientocodigo and A.subasientocodigo=F.subasientocodigo and 
			          	A.documentocodigo=F.documentocodigo and A.detcomprobnumdocumento=F.detcomprobnumdocumento and              
			            '+@CADCTASIES+'),0),
		RENTA=ISNULL(( SELECT top 1
         	          isnull(case when F.detcomprobdebe>0 then F.detcomprobdebe 
                             else F.detcomprobhaber * -1 end,0)         	       		
       	  			FROM ['+@BASE+'].dbo.ct_detcomprob'+@ANNO+' F
       	  			WHERE a.empresacodigo=f.empresacodigo and A.cabcomprobmes=F.cabcomprobmes and 
				      	A.cabcomprobnumero=F.cabcomprobnumero and A.asientocodigo=F.asientocodigo and 
				      	A.subasientocodigo=F.subasientocodigo and A.documentocodigo=F.documentocodigo and 
			          	A.detcomprobnumdocumento=F.detcomprobnumdocumento and '+@CADCTASRENTA+'),0),
		  tipdocref=isnull(A.tipdocref,''''),detcomprobnumref=isnull(A.detcomprobnumref,''''),A.detcomprobtipocambio,		
        MontoReferencia=case when A.monedacodigo=''02'' then  
                           isnull((case when A.detcomprobussdebe>0 then A.detcomprobussdebe * -1  
                                   else A.detcomprobusshaber  end),0) 
                           else 0 end,A.detcomprobnlibro,
		  NumAuxiliar=cabcomprobnprovi,A.centrocostocodigo, g.identidadcodigo, A.detcomprobfecharef,
		  a.detcomprobnumerodetraccion , a.detcomprobfechadetraccion ,aaaa='''+@anno+''' '
Set @sqlvar1=' FROM ['+@base+'].dbo.ct_detcomprob'+@anno+' A inner join ['+@base+'].dbo.ct_cabcomprob'+@anno+' Z 
  on a.empresacodigo=z.empresacodigo and A.cabcomprobmes=Z.cabcomprobmes and A.cabcomprobnumero=Z.cabcomprobnumero and A.asientocodigo=Z.asientocodigo and 
	           A.subasientocodigo=Z.subasientocodigo 		
  	    inner join ['+@base+'].dbo.v_analiticoentidad G on  A.analiticocodigo=G.analiticocodigo  
            inner join ['+@base+'].dbo.ct_asiento H on  A.asientocodigo=H.asientocodigo  
            inner join ['+@base+'].dbo.gr_documento T on  A.documentocodigo=T.documentocodigo          
        WHERE a.empresacodigo='''+@empresa+''' and A.asientocodigo IN '+@CADASIENTOSPLAN+' AND t.documentoregcompras=1 and
              A.cuentacodigo like ''42%'' and A.cabcomprobmes='''+@mes+''' order by a.cabcomprobnumero '
execute(@sqlvar+@sqlvar1)
--print(@sqlvar)
--   exec ct_LibroRegistroCompras_rpt 'conta2010','02','2010','12','060,009','33%34%60%62%63%64%65%67%','40111,40112','401902,40116,','401700,'









