/*
select * from planta_casma.dbo.ct_detcomprob2011
delete  planta_casma.dbo.ct_detcomprob2011
where detcomprobauto=1 
and ( left(cuentacodigo,1)='9' or  cuentacodigo='7911100' )


*/

Declare @Xcabcomprobnumero varchar(10),@Xasientocodigo varchar(3),
        @Xsubasientocodigo varchar(4),@Xtabla varchar(50)
declare @baseconta varchar(50), @empresa varchar(2) , @mes integer, @ano varchar(4)

set @ano = '2011'
Set @Xtabla='ct_detcomprob2011'
set @baseconta ='planta_casma'

Declare GenAuto Cursor for 
select B.cabcomprobnumero,B.asientocodigo,B.subasientocodigo, b.cabcomprobmes,b.empresacodigo 
from planta_casma.dbo.ct_cabcomprob2011 b

Open GenAuto
Fetch next from GenAuto into @Xcabcomprobnumero,@Xasientocodigo,@Xsubasientocodigo, @mes, @empresa
     While @@Fetch_status=0  
        Begin
           Exec ct_grabaautomatico_pro @baseconta,@Xtabla,@empresa,@mes,
                @Xcabcomprobnumero,@Xasientocodigo,@Xsubasientocodigo ,1                                 
 
           Exec ct_CalcComprob_pro '',@baseconta,@empresa,@Ano,@mes,
           @Xasientocodigo,@Xsubasientocodigo,@Xcabcomprobnumero   
           Fetch next from GenAuto into @Xcabcomprobnumero,@Xasientocodigo,@Xsubasientocodigo, @mes, @empresa
       End
    Close GenAuto
Deallocate GenAuto
