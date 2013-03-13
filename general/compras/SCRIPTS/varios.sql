
--- SELECT * FROM xx_HOJA2$
--- SELECT * FROM CO_GASTOS where gastosnivel=1
--DELETE CO_GASTOS WHERE GASTOSNIVEL=2 

--- select * from invbrisa.dbo.co_gastos where gastosnivel=2


INSERT CO_GASTOS (gastoscodigo,gastosdescripcion,gastosestado,
gastosctrlcostos,CUENTACODIGO,
 gastosnivel, gastosestadodistribucion,usuariocodigo, fechaact)


update co_gastos
set  tipoanaliticocodigo=z.tipoanalisis
from co_gastos a,
(SELECT 
RIGHT('0'+RTRIM(GRUPO),2)+RIGHT('0'+RTRIM(CORRELATIVO),2) 
as codigo,
tipoanalisis= case when aa > 0 then 
                    right('001'+ltrim(str(aa)),3)
               else
                   '00' end  
 FROM xx_HOJA2$
WHERE NOT ISNULL(CORRELATIVO,'')='' ) as z
where a.gastoscodigo=z.codigo


SELECT RIGHT('0'+RTRIM(GRUPO),2)+RIGHT('0'+RTRIM(CORRELATIVO),2),right('000'+
,COUNT(*) FROM HOJA2$
WHERE NOT ISNULL(CORRELATIVO,'')=''
GROUP BY  RIGHT('0'+RTRIM(GRUPO),2)+RIGHT('0'+RTRIM(CORRELATIVO),2)