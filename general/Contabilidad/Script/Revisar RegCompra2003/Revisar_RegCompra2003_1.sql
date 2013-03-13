SELECT A.cabcomprobmes,A.cabcomprobnumero,A.subasientocodigo,A.asientocodigo,H.asientodescripcion, 
        A.analiticocodigo,G.entidadruc,G.entidadrazonsocial,A.monedacodigo,
        documentocodigo=isnull(A.documentocodigo,''),
        T.documentodescripcion,
        detcomprobnumdocumento=isnull(A.detcomprobnumdocumento,''),
        A.detcomprobfechaemision,

/*
		baseimpgrab=Isnull((Select isnull((case when isnull(Z.cabcomprobgrabada,0)=1 
                   then case when sum(F.detcomprobdebe)>0 then sum(F.detcomprobdebe) 
                             else sum(F.detcomprobhaber) * -1 end end),0)
        FROM [Contaprueba].dbo.ct_detcomprob2003 F
        Where  
		A.cabcomprobmes=F.cabcomprobmes and 
				      A.cabcomprobnumero=F.cabcomprobnumero and 
				      A.asientocodigo=F.asientocodigo and 
				      A.subasientocodigo=F.subasientocodigo and 
			          A.documentocodigo=F.documentocodigo and 
			          A.detcomprobnumdocumento=F.detcomprobnumdocumento And
        F.asientocodigo IN ('060','061','062','063','064') AND 
        F.detcomprobauto=0 and  F.plantillaasientoinafecto=0 and (F.cuentacodigo like '60%' or F.cuentacodigo like '33%' or F.cuentacodigo like '34%' or F.cuentacodigo like '46%' or F.cuentacodigo like '63%' or F.cuentacodigo like '64%' or F.cuentacodigo like '65%' or F.cuentacodigo like '9%' or F.cuentacodigo like '28%' or F.cuentacodigo like '38%' or F.cuentacodigo like '73%' or F.cuentacodigo like '77%') Group By F.cabcomprobnumero,F.asientocodigo,F.subasientocodigo,F.documentocodigo,F.detcomprobnumdocumento),0) , 

        baseimpnograb=Isnull((Select isnull((case when isnull(Z.cabcomprobgrabada,0)=0 
                   then case when sum(F.detcomprobdebe)>0 then sum(F.detcomprobdebe) 
                             else sum(F.detcomprobhaber) * -1 end end),0)                
		FROM [Contaprueba].dbo.ct_detcomprob2003 F
        Where  
		A.cabcomprobmes=F.cabcomprobmes and 
				      A.cabcomprobnumero=F.cabcomprobnumero and 
				      A.asientocodigo=F.asientocodigo and 
				      A.subasientocodigo=F.subasientocodigo and 
			          A.documentocodigo=F.documentocodigo and 
			          A.detcomprobnumdocumento=F.detcomprobnumdocumento And
        F.asientocodigo IN ('060','061','062','063','064') AND 
        F.detcomprobauto=0 and  F.plantillaasientoinafecto=0 and (F.cuentacodigo like '60%' or F.cuentacodigo like '33%' or F.cuentacodigo like '34%' or F.cuentacodigo like '46%' or F.cuentacodigo like '63%' or F.cuentacodigo like '64%' or F.cuentacodigo like '65%' or F.cuentacodigo like '9%' or F.cuentacodigo like '28%' or F.cuentacodigo like '38%' or F.cuentacodigo like '73%' or F.cuentacodigo like '77%') Group By F.cabcomprobnumero,F.asientocodigo,F.subasientocodigo,F.documentocodigo,F.detcomprobnumdocumento),0),	            
*/

        montoinafecto=isnull((SELECT TOP 1
                      Case when F.plantillaasientoinafecto=1 then 
                                isnull((case when sum(F.detcomprobdebe)>=0 then sum(F.detcomprobdebe) 
                                       else sum(F.detcomprobhaber) * -1 end ),0)
                      Else 0 end
          
        			FROM [Contaprueba].dbo.ct_detcomprob2003 F
        			WHERE
			          F.plantillaasientoinafecto=1 and 	F.detcomprobauto=0 and 
			          A.cabcomprobmes=F.cabcomprobmes and 
				      A.cabcomprobnumero=F.cabcomprobnumero and 
				      A.asientocodigo=F.asientocodigo and 
				      A.subasientocodigo=F.subasientocodigo and 
			          A.documentocodigo=F.documentocodigo and 
			          A.detcomprobnumdocumento=F.detcomprobnumdocumento   
         Group By F.cabcomprobnumero,F.asientocodigo,F.subasientocodigo,F.documentocodigo,F.detcomprobnumdocumento,F.plantillaasientoinafecto),0), 

/*
       	igvimpgrab=ISNULL(( SELECT top 1 
         	       		isnull((case when isnull(Z.cabcomprobgrabada,0)=1 
                   			then case when F.detcomprobdebe>0 then F.detcomprobdebe 
                            	 else F.detcomprobhaber * -1 end end),0)         	
       		
       	 			FROM [Contaprueba].dbo.ct_detcomprob2003 F
       	 			WHERE	   		
			            A.cabcomprobmes=F.cabcomprobmes and 
				        A.cabcomprobnumero=F.cabcomprobnumero and 
				        A.asientocodigo=F.asientocodigo and 
				        A.subasientocodigo=F.subasientocodigo and 
			            A.documentocodigo=F.documentocodigo and 
			            F.detcomprobauto=0 and (F.cuentacodigo like '401110' or F.cuentacodigo like '401113')),0),

        igvimpnograb=ISNULL(( SELECT top 1          	
         	                  isnull((case when isnull(Z.cabcomprobgrabada,0)=0 
                                    then case when F.detcomprobdebe>0 then F.detcomprobdebe 
                                    else F.detcomprobhaber * -1 end end),0)
       	  			FROM [Contaprueba].dbo.ct_detcomprob2003 F
       	  			WHERE	   		
			            A.cabcomprobmes=F.cabcomprobmes and 
				        A.cabcomprobnumero=F.cabcomprobnumero and 
				        A.asientocodigo=F.asientocodigo and 
				        A.subasientocodigo=F.subasientocodigo and 
			            A.documentocodigo=F.documentocodigo and             
			            F.detcomprobauto=0 and  (F.cuentacodigo like '401110' or F.cuentacodigo like '401113')),0),
        IES=ISNULL(( SELECT top 1
         	          isnull(case when F.detcomprobdebe>0 then F.detcomprobdebe 
                             else F.detcomprobhaber * -1 end,0)         	       		
       	  			FROM [Contaprueba].dbo.ct_detcomprob2003 F
       	  			WHERE
				   		A.cabcomprobmes=F.cabcomprobmes and 
				      	A.cabcomprobnumero=F.cabcomprobnumero and 
				      	A.asientocodigo=F.asientocodigo and 
				      	A.subasientocodigo=F.subasientocodigo and 
			          	A.documentocodigo=F.documentocodigo and 
			          	A.detcomprobnumdocumento=F.detcomprobnumdocumento and              
			            F.detcomprobauto=0 and (F.cuentacodigo like '401174')),0),

		RENTA=ISNULL(( SELECT top 1
         	          isnull(case when F.detcomprobdebe>0 then F.detcomprobdebe 
                             else F.detcomprobhaber * -1 end,0)         	       		
       	  			FROM [Contaprueba].dbo.ct_detcomprob2003 F
       	  			WHERE
				   		A.cabcomprobmes=F.cabcomprobmes and 
				      	A.cabcomprobnumero=F.cabcomprobnumero and 
				      	A.asientocodigo=F.asientocodigo and 
				      	A.subasientocodigo=F.subasientocodigo and 
			          	A.documentocodigo=F.documentocodigo and 
			          	A.detcomprobnumdocumento=F.detcomprobnumdocumento and              
			            F.detcomprobauto=0 and (F.cuentacodigo like '403140')),0),*/

		  tipdocref=isnull(A.tipdocref,''),detcomprobnumref=isnull(A.detcomprobnumref,''),A.detcomprobtipocambio,		
        MontoReferencia=case when A.monedacodigo='02' then  
                           isnull((case when A.detcomprobussdebe>0 then A.detcomprobussdebe * -1  
                                   else A.detcomprobusshaber  end),0) 
                           else 0 end,A.detcomprobnlibro,
		  NumAuxiliar=isnull('01'+ replicate('0',5-len(cabcomprobnprovi))+cabcomprobnprovi,'')				
       FROM [Contaprueba].dbo.ct_detcomprob2003 A,[Contaprueba].dbo.ct_cabcomprob2003 Z,   	          	
            [Contaprueba].dbo.v_analiticoentidad G, 
            [Contaprueba].dbo.ct_asiento H,
            [Contaprueba].dbo.gr_documento T         	        	
	WHERE
		A.cabcomprobmes=Z.cabcomprobmes and 
	   A.cabcomprobnumero=Z.cabcomprobnumero and 
	   A.asientocodigo=Z.asientocodigo and 
	   A.subasientocodigo=Z.subasientocodigo and 		
	   A.asientocodigo IN ('060','061','062','063','064') AND 
   	(A.cuentacodigo like '421%' ) and 
      A.analiticocodigo=G.analiticocodigo and 
      A.asientocodigo=H.asientocodigo and 
      A.documentocodigo=T.documentocodigo and         
      A.detcomprobauto=0 and A.cabcomprobmes=01 and
		A.cabcomprobnumero='0106100272'
