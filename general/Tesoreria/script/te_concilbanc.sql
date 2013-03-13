SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

ALTER    proc te_concilbanc
--Declare 
@Base   varchar(50), 
@cuenta varchar(50),
@concil varchar(1), 
@Fecharef  varchar(10),
@fecha    varchar(10)

as 
/*
 Set @Base='Ventas_Prueba'
 Set @cuenta='011-350-0100008495-62' 
 set @concil='2' 
 Set @Fecharef='01/01/2003'  
 Set @fecha='01/12/2002'  
*/

Declare @Sqlcad varchar(4000),@Sqlvar varchar(1000)

Set @Sqlcad='
select 
      A.chkconcil, 
      Anno=year(A.detrec_fechacancela),Mes=month(A.detrec_fechacancela),  A.cabrec_numrecibo,
      A.detrec_emisioncheque,A.detrec_tipodoc_concepto,  A.detrec_numdocumento,B.cabrec_ingsal,
      A.detrec_tipocajabanco,  A.detrec_numctacte,A.detrec_monedadocumento,
      A.detrec_importesoles,A.detrec_importedolares,A.detrec_monedacancela,
      A.detrec_tdqc,A.detrec_ndqc,A.detrec_fechacancela,B.cabrec_estadoreg,
      B.cabrec_fechadocumento,A.detrec_observacion,A.fechconcil 
from ['+@Base+'].dbo.te_detallerecibos A 
Inner join  ['+@Base+'].dbo.te_cabecerarecibos  B  on 
      A.cabrec_numrecibo=B.cabrec_numrecibo 
Where A.detrec_emisioncheque=''B'' and  A.detrec_tipocajabanco=''B'' and 
      ltrim(rtrim(Isnull(A.detrec_numctacte,'''')))  <>'''' and  B.cabrec_estadoreg <> 1 
      and ltrim(rtrim(Isnull(A.detrec_numctacte,'''')))='''+@cuenta+''' and A.detrec_fechacancela < '''+@Fecharef+''' and   
      ( fechconcil is null or fechconcil >='''+@fecha+''')'
If @concil='0' set @Sqlvar=''
If @concil='1' set @Sqlvar=' and (isnull(chkconcil,0)=1 and fechconcil <'''+@Fecharef+''')'
If @concil='2'
    set @Sqlvar=' and (case when fechconcil >='''+@Fecharef+''' and isnull(chkconcil,0)=1  then 0 else 1 end =0 or  
                  isnull(chkconcil,0)=0 ) '

exec(@Sqlcad+@Sqlvar+' order by A.detrec_fechacancela ')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

