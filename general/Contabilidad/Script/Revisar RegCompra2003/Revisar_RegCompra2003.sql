SELECT A.cabcomprobmes,A.cabcomprobnumero,A.subasientocodigo,A.asientocodigo,H.asientodescripcion, 
        A.analiticocodigo,G.entidadruc,G.entidadrazonsocial,A.monedacodigo,
        documentocodigo=isnull(A.documentocodigo,''),
        T.documentodescripcion,
        detcomprobnumdocumento=isnull(A.detcomprobnumdocumento,''),
        A.detcomprobfechaemision,

		Inafecto=isnull((SELECT TOP 1
                      Case when F.plantillaasientoinafecto=1 then 
                                isnull((case when sum(F.detcomprobdebe)>=0 then sum(F.detcomprobdebe) 
                                       else sum(F.detcomprobhaber) * -1 end ),0)
                      Else 0 end
          
      		  			FROM [Contaprueba].dbo.ct_detcomprob2003 F
        					WHERE
			         		 	F.plantillaasientoinafecto=1 and F.detcomprobauto=0 and 
			          			F.cabcomprobmes=A.cabcomprobmes and 
				      			A.cabcomprobnumero=F.cabcomprobnumero and 
				      A.asientocodigo=F.asientocodigo and 
				      A.subasientocodigo=F.subasientocodigo and 
			          A.documentocodigo=F.documentocodigo and 
			          A.detcomprobnumdocumento=F.detcomprobnumdocumento   
         Group By F.cabcomprobnumero,F.asientocodigo,F.subasientocodigo,F.documentocodigo,F.detcomprobnumdocumento,F.plantillaasientoinafecto),0) 





        /*montoinafecto= isnull((SELECT TOP 1
                      Case when F.plantillaasientoinafecto=1 then 
                                isnull((case when sum(F.detcomprobdebe)>=0 then sum(F.detcomprobdebe) 
                                       else sum(F.detcomprobhaber) * -1 end ),0)
                      Else 0 end
          
        			FROM [Contaprueba].dbo.ct_detcomprob2003 F
        			WHERE
			          F.plantillaasientoinafecto=1 and F.detcomprobauto=0 and 
			          F.cabcomprobmes=A.cabcomprobmes and 
				      A.cabcomprobnumero=F.cabcomprobnumero and 
				      A.asientocodigo=F.asientocodigo and 
				      A.subasientocodigo=F.subasientocodigo and 
			          A.documentocodigo=F.documentocodigo and 
			          A.detcomprobnumdocumento=F.detcomprobnumdocumento   
         Group By F.cabcomprobnumero,F.asientocodigo,F.subasientocodigo,F.documentocodigo,F.detcomprobnumdocumento,F.plantillaasientoinafecto),0) */

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


--select * from contaprueba.dbo.ct_detcomprob2003 where cabcomprobnumero='0106100272'