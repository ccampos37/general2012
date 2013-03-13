
--dbo.ct_detcomprob2002
--dbo.ct_cuenta
--dbo.ct_distribucion



/*Esta Consulta me Dice quien tiene asientos automaticos */

If Exists(Select name from tempdb.dbo.sysobjects Where name='##TempoAsiento')
	Drop Table ##TempoAsiento

Select Distinct A.cabcomprobmes, A.cabcomprobnumero,A.asientocodigo,A.subasientocodigo 
Into ##TempoAsiento

From 

(Select A.*      
 From  ct_detcomprob2002 A,          
(select A.cabcomprobnumero,A.cuentacodigo
 from ct_detcomprob2002 A,ct_cuenta B
 Where A.cuentacodigo=B.cuentacodigo  and 
      isnull(B.cuentaestadodistribucion,0) <>0
 ) as X
 Where A.cabcomprobnumero=X.cabcomprobnumero) as  A
 Where 
 A.cabcomprobnumero  not in(
Select Distinct A.cabcomprobnumero
       
From  ct_detcomprob2002 A,          
(select A.cabcomprobnumero,A.cuentacodigo
 from ct_detcomprob2002 A,ct_cuenta B
 Where A.cuentacodigo=B.cuentacodigo  and 
      isnull(B.cuentaestadodistribucion,0) <>0
 ) as X,ct_distribucion Z 

Where A.cabcomprobnumero=X.cabcomprobnumero and 
      X.cuentacodigo=Z.cuentacodigo and      
      (A.cuentacodigo=Z.distribucioncargo  
        or  A.cuentacodigo=Z.distribucionabono)  )

Order By A.cabcomprobmes, A.cabcomprobnumero,A.asientocodigo,A.subasientocodigo
select cabcomprobnumero,asientocodigo,subasientocodigo,cabcomprobmes
 from  ##TempoAsiento


/*Este el Cursor para actualizar los comprobantes */
/*
Declare @Xcabcomprobnumero varchar(10),@Xasientocodigo varchar(3),
        @Xsubasientocodigo varchar(4),@mes int


Declare GenAuto Cursor for 
select cabcomprobnumero,asientocodigo,subasientocodigo,cabcomprobmes
from  ##TempoAsiento

Open GenAuto

Fetch next from GenAuto into @Xcabcomprobnumero,@Xasientocodigo,@Xsubasientocodigo,@mes

While @@Fetch_status=0 
Begin
    Exec ct_grabaautomatico_pro 'Contaprueba','ct_cabcomprob2002',@mes,
    @Xcabcomprobnumero,@Xasientocodigo,@Xsubasientocodigo                                   
   	
    Exec ct_CalcComprob_pro '',Contaprueba,'2002',@mes,
    @Xasientocodigo,@Xsubasientocodigo,@Xcabcomprobnumero   
    Print  @Xasientocodigo +' '+@Xsubasientocodigo+' '+@Xcabcomprobnumero

    Fetch next from GenAuto into @Xcabcomprobnumero,@Xasientocodigo,@Xsubasientocodigo,@mes
End
Close GenAuto
Deallocate GenAuto
*/