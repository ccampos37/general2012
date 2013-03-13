/*
SELECT A.cabcomprobmes,A.cabcomprobnumero,A.subasientocodigo,A.asientocodigo,H.asientodescripcion, 
        A.analiticocodigo,
        entidadruc=case when left(A.analiticocodigo,8)='88888888' then A.detcomprobruc else G.entidadruc end,
        entidadrazonsocial=case when left(A.analiticocodigo,8)='88888888' then A.detcomprobglosa else G.entidadrazonsocial end,        
        A.monedacodigo,
        documentocodigo=isnull(A.documentocodigo,''),
        tdserie=isnull(A.documentocodigo,'')+left(isnull(A.detcomprobnumdocumento,''),3),
        T.documentodescripcion,
        detcomprobnumdocumento=isnull(A.detcomprobnumdocumento,''),
        A.detcomprobfechaemision,

		baseimpgrab=Isnull((Select isnull((case when isnull(Z.cabcomprobgrabada,0)=1 
                   then case when sum(F.detcomprobhaber)>0 then sum(F.detcomprobhaber) 
                             else sum(F.detcomprobdebe) * -1 end end),0)
        FROM [CONTAPRUEBA].dbo.ct_detcomprob2003 F
        Where  
		A.cabcomprobmes=F.cabcomprobmes and 
				      A.cabcomprobnumero=F.cabcomprobnumero and 
				      A.asientocodigo=F.asientocodigo and 
				      A.subasientocodigo=F.subasientocodigo and 
			          A.documentocodigo=F.documentocodigo and 
			          A.detcomprobnumdocumento=F.detcomprobnumdocumento And
        F.asientocodigo IN ('070','071','072','073','074') AND 
        F.detcomprobauto=0 and  F.plantillaasientoinafecto=0 and (F.cuentacodigo like '70%' or F.cuentacodigo like '74%' or F.cuentacodigo like '75%' or F.cuentacodigo like '76%' or F.cuentacodigo like '77%') Group By F.cabcomprobnumero,F.asientocodigo,F.subasientocodigo,F.documentocodigo,F.detcomprobnumdocumento),0) , 

        baseimpnograb=Isnull((Select isnull((case when isnull(Z.cabcomprobgrabada,0)=0 
                   then case when sum(F.detcomprobhaber)>0 then sum(F.detcomprobhaber) 
                             else sum(F.detcomprobdebe) * -1 end end),0)                
		FROM [CONTAPRUEBA].dbo.ct_detcomprob2003 F
        Where  
		A.cabcomprobmes=F.cabcomprobmes and 
				      A.cabcomprobnumero=F.cabcomprobnumero and 
				      A.asientocodigo=F.asientocodigo and 
				      A.subasientocodigo=F.subasientocodigo and 
			          A.documentocodigo=F.documentocodigo and 
			          A.detcomprobnumdocumento=F.detcomprobnumdocumento And
        F.asientocodigo IN ('070','071','072','073','074') AND 
        F.detcomprobauto=0 and  F.plantillaasientoinafecto=0 and (F.cuentacodigo like '70%' or F.cuentacodigo like '74%' or F.cuentacodigo like '75%' or F.cuentacodigo like '76%' or F.cuentacodigo like '77%') Group By F.cabcomprobnumero,F.asientocodigo,F.subasientocodigo,F.documentocodigo,F.detcomprobnumdocumento),0),	            

        montoinafecto=isnull((SELECT TOP 1
                      Case when F.plantillaasientoinafecto=1 then 
                                isnull((case when F.detcomprobhaber>=0 then F.detcomprobhaber 
                                       else F.detcomprobdebe * -1 end ),0)
                      Else 0 end
          
        			FROM [CONTAPRUEBA].dbo.ct_detcomprob2003 F
        			WHERE
			          	F.plantillaasientoinafecto=1 and 	F.detcomprobauto=0 and 
			          	A.cabcomprobmes=F.cabcomprobmes and 
				      	A.cabcomprobnumero=F.cabcomprobnumero and 
				      	A.asientocodigo=F.asientocodigo and 
				      	A.subasientocodigo=F.subasientocodigo and 
			          	A.documentocodigo=F.documentocodigo and (F.cuentacodigo like '70%' or F.cuentacodigo like '74%' or F.cuentacodigo like '75%' or F.cuentacodigo like '76%' or F.cuentacodigo like '77%') AND
			          	A.detcomprobnumdocumento=F.detcomprobnumdocumento),0), 
       	igvimpgrab=ISNULL(( SELECT top 1 
         	       		isnull((case when isnull(Z.cabcomprobgrabada,0)=1 
                   			then case when F.detcomprobhaber>0 then F.detcomprobhaber 
                            	 else F.detcomprobdebe * -1 end end),0)         	
       		
       	 			FROM [CONTAPRUEBA].dbo.ct_detcomprob2003 F
       	 			WHERE	   		
			            A.cabcomprobmes=F.cabcomprobmes and 
				        A.cabcomprobnumero=F.cabcomprobnumero and 
				        A.asientocodigo=F.asientocodigo and 
				        A.subasientocodigo=F.subasientocodigo and 
			            A.documentocodigo=F.documentocodigo and 
			            F.detcomprobauto=0 and (F.cuentacodigo like '401110')),0),

        igvimpnograb=ISNULL(( SELECT top 1          	
         	                  isnull((case when isnull(Z.cabcomprobgrabada,0)=0 
                                    then case when F.detcomprobhaber>0 then F.detcomprobhaber 
                                    else F.detcomprobdebe * -1 end end),0)
       	  			FROM [CONTAPRUEBA].dbo.ct_detcomprob2003 F
       	  			WHERE	   		
			            A.cabcomprobmes=F.cabcomprobmes and 
				        A.cabcomprobnumero=F.cabcomprobnumero and 
				        A.asientocodigo=F.asientocodigo and 
				        A.subasientocodigo=F.subasientocodigo and 
			            A.documentocodigo=F.documentocodigo and             
			            F.detcomprobauto=0 and  (F.cuentacodigo like '401110')),0),
        flete=ISNULL(( SELECT top 1
         	          isnull(case when F.detcomprobhaber>0 then F.detcomprobhaber 
                             else F.detcomprobdebe * -1 end,0)         	       		
       	  			FROM [CONTAPRUEBA].dbo.ct_detcomprob2003 F
       	  			WHERE
				   		A.cabcomprobmes=F.cabcomprobmes and 
				      	A.cabcomprobnumero=F.cabcomprobnumero and 
				      	A.asientocodigo=F.asientocodigo and 
				      	A.subasientocodigo=F.subasientocodigo and 
			          	A.documentocodigo=F.documentocodigo and 
			          	A.detcomprobnumdocumento=F.detcomprobnumdocumento and              
			            F.detcomprobauto=0 and (F.cuentacodigo like '75%')),0),

		otros=ISNULL(( SELECT top 1
         	          isnull(case when F.detcomprobhaber>0 then F.detcomprobhaber 
                             else F.detcomprobdebe * -1 end,0)         	       		
       	  			FROM [CONTAPRUEBA].dbo.ct_detcomprob2003 F
       	  			WHERE
				   		A.cabcomprobmes=F.cabcomprobmes and 
				      	A.cabcomprobnumero=F.cabcomprobnumero and 
				      	A.asientocodigo=F.asientocodigo and 
				      	A.subasientocodigo=F.subasientocodigo and 
			          	A.documentocodigo=F.documentocodigo and 
			          	A.detcomprobnumdocumento=F.detcomprobnumdocumento and              
			            F.detcomprobauto=0 and (F.cuentacodigo like '76%')),0),

		  tipdocref=isnull(A.tipdocref,''),detcomprobnumref=isnull(A.detcomprobnumref,''),A.detcomprobtipocambio,		
          MontoReferencia=case when A.monedacodigo='02' then  
                           isnull((case when A.detcomprobusshaber>0 then A.detcomprobusshaber * -1  
                                   else A.detcomprobussdebe  end),0) 
                           else 0 end,A.detcomprobnlibro                 
		
			into RegVentas	

	      FROM [CONTAPRUEBA].dbo.ct_detcomprob2003 A,[CONTAPRUEBA].dbo.ct_cabcomprob2003 Z,   	          	
            [CONTAPRUEBA].dbo.v_analiticoentidad G, 
            [CONTAPRUEBA].dbo.ct_asiento H,
            [CONTAPRUEBA].dbo.gr_documento T         	        	
	WHERE
		A.cabcomprobmes=Z.cabcomprobmes and 
	   A.cabcomprobnumero=Z.cabcomprobnumero and 
	   A.asientocodigo=Z.asientocodigo and 
	   A.subasientocodigo=Z.subasientocodigo and 		
	   A.asientocodigo IN ('070','071','072','073','074') AND 
   	(A.cuentacodigo like '121%' ) and 
   	A.analiticocodigo=G.analiticocodigo and 
      A.asientocodigo=H.asientocodigo and 
      A.documentocodigo=T.documentocodigo and         
      A.detcomprobauto=0 and A.cabcomprobmes=1
*/

/*
select * from regventas 
order by detcomprobnumdocumento
select cabcomprobnumero,
		serie=left(detcomprobnumdocumento,3), 
		numero=substring(detcomprobnumdocumento,5,8) from regventas 
order by 2,3

*/

Declare @inicio as bigint
Declare @serieini as bigint
Declare @conta as bigint
DECLARE @cabcomprobnumero varchar(20),@serie varchar(3),@numero varchar(8)
DECLARE tablas CURSOR FOR 
		select cabcomprobnumero,
			serie=left(detcomprobnumdocumento,3), 
			numero=substring(detcomprobnumdocumento,5,8) from regventas 
		order by 2,3   

	OPEN tablas
	/* Leer cada registro del cursor  */
	
	FETCH NEXT FROM tablas INTO @cabcomprobnumero, @serie, @numero
	set @inicio=cast(@numero as bigint)	
	set @serieini=cast(@serie as bigint)
	FETCH NEXT FROM tablas INTO @cabcomprobnumero, @serie, @numero
	WHILE @@FETCH_STATUS = 0
	BEGIN
      --@inicio+@conta

		if  cast(@serie as bigint) <>@serieini
		 begin
			set @serieini=cast(@serie as bigint)
			set @inicio=cast(@numero as bigint)
			set @conta=1
		end
	
		--set @cadsql='12'
		print @cabcomprobnumero+@inicio
      --exec(@cadsql)
 	FETCH NEXT FROM tablas INTO @cabcomprobnumero, @serie, @numero
   END
	CLOSE tablas
	DEALLOCATE tablas

select * into contaprueba.dbo.regventas from regventas