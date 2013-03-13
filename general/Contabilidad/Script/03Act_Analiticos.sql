Use Contaprueba

if exists (select name from sysobjects where name='##DocCancela')
	drop table ##DocCancela

Select 
	AA.CuentaCodigo,AA.operacioncodigo,
	AA.analiticocodigo,AA.DocumentoCodigo,
	AA.detcomprobnumdocumento,
	AA.detcomprobfechaemision,AA.detcomprobformacambio,AA.detcomprobtipocambio,
	AA.monedacodigo,
	FechaCancela=BB.detcomprobfechaemision,
/*
	MontoProv= Round((AA.detcomprobdebe +  AA.detcomprobhaber),2),
   TotalPagado= Round(isnull((BB.Sdebe+BB.Shaber),0) ,2),
	Saldo= Round((AA.detcomprobdebe +  AA.detcomprobhaber) - isnull((BB.Sdebe+BB.Shaber),0) ,2)
*/

/*
	MontoProv=	Case When AA.monedacodigo='01' 
	              		Then Round((AA.detcomprobdebe +  AA.detcomprobhaber),2)
               		Else Round((AA.detcomprobussdebe+AA.detcomprobusshaber) ,2)
          		End, 
   TotalPagado=		Case When AA.monedacodigo='01' 
               		Then Round(isnull((BB.Sdebe+BB.Shaber),0) ,2)
               		Else Round(isnull((BB.Sdebeuss+BB.Shaberuss),0) ,2)
            	End,
	Saldo=		Case When AA.monedacodigo='01' 
               		Then 
								Round((AA.detcomprobdebe +  AA.detcomprobhaber)-
	                    	isnull((BB.Sdebe+BB.Shaber),0),2)
               		Else 
								Round((AA.detcomprobussdebe+AA.detcomprobusshaber) -
            		        isnull((BB.Sdebeuss+BB.Shaberuss),0) ,2)
          		End
*/
	MontoProvS=	Round((AA.detcomprobdebe +  AA.detcomprobhaber),2),
   TotalPagadoS= Round(isnull((BB.Sdebe+BB.Shaber),0) ,2),
	SaldoS=Round((AA.detcomprobdebe +  AA.detcomprobhaber)- isnull((BB.Sdebe+BB.Shaber),0),2),

	MontoProvD=	Round((AA.detcomprobussdebe+AA.detcomprobusshaber) ,2),
   TotalPagadoD= Round(isnull((BB.Sdebeuss+BB.Shaberuss),0) ,2),
	SaldoD=Round((AA.detcomprobussdebe+AA.detcomprobusshaber) -
            		        isnull((BB.Sdebeuss+BB.Shaberuss),0) ,2)

Into ##DocCancela
From [Contaprueba].dbo.ct_detcomprob2002 AA,
    (select A.CuentaCodigo,A.analiticocodigo,	   
		   	A.documentocodigo,A.detcomprobnumdocumento,
				A.detcomprobfechaemision, 	   		
		   	Sdebe=Round(Sum(A.detcomprobdebe),2)  ,
				Shaber=Round(sum(A.detcomprobhaber),2),
				Sdebeuss=Round(sum(A.detcomprobussdebe),2),
      		Shaberuss=Round(sum(A.detcomprobusshaber),2)       
	   from [Contaprueba].dbo.ct_detcomprob2002 A       
	   where		
    		A.operacioncodigo<>'01' and 
         A.analiticocodigo<>'00'
		Group by A.CuentaCodigo,A.analiticocodigo,	   
	    	A.documentocodigo,A.detcomprobnumdocumento,
			A.detcomprobfechaemision 
	 ) as BB
	Where 
		AA.operacioncodigo='01' and 
		AA.CuentaCodigo=BB.CuentaCodigo and 
    	AA.analiticocodigo=BB.analiticocodigo and 
    	AA.documentocodigo=BB.documentocodigo and   
    	AA.detcomprobnumdocumento=BB.detcomprobnumdocumento 


--select * from contaprueba.dbo.ct_detcomprob2002
--select count(*) from contaprueba.dbo.ct_ctacteanalitico2002 where cabcomprobmes=11
/*
select * from contaprueba.dbo.ct_ctacteanalitico2002 --where ctacteanaliticocancel='1'
  where analiticocodigo='20418885123001'
select * from doccanc where analiticocodigo='20418885123001'
*/

go
--select * from DocCancela where abs(SaldoS)>0  and abs(SaldoD)>0
delete  from ##DocCancela where abs(SaldoS)>0  and abs(SaldoD)>0

go

update [contaprueba].dbo.ct_ctacteanalitico2002 set ctacteanaliticocancel=ZZ.FechaCancela
  from 
	[contaprueba].dbo.ct_ctacteanalitico2002 Y,
	##DocCancela ZZ
where 
 	Y.analiticocodigo=ZZ.analiticocodigo and
	Y.ctacteanaliticonumdocumento=ZZ.detcomprobnumdocumento and
	Y.documentocodigo=ZZ.DocumentoCodigo 

/*
select * from contaprueba.dbo.ct_ctacteanalitico2002 where ctacteanaliticocancel='1' and
select * from doccancela where 
	analiticocodigo='2010016681001' and DocumentoCodigo='01' and detcomprobnumdocumento like '052-005572'
*/		
/*
select * from doccancela where 
	  detcomprobnumdocumento like '001-6609468'
select * from dbo.DocNoCuadran where detcomprobnumdocumento like '001-6609468'
select * from contaprueba.dbo.ct_ctacteanalitico2002 where ctacteanaliticonumdocumento='001-0002533'
select * from contaprueba.dbo.ct_detcomprob2002 where detcomprobnumdocumento='001-0002533'
*/