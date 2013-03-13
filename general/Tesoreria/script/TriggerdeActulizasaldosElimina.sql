CREATE TRIGGER ActualizaSaldoalElim  ON [dbo].[te_detallerecibos] 
FOR  DELETE 
AS

Select X.ano,X.mes,X.detrec_tipocajabanco,X.detrec_cajabanco1,X.detrec_numctacte,
       IngresoSoles=sum(x.IngresoSoles),EgresoSoles=sum(x.EgresoSoles),
       IngresoDolar=sum(x.IngresoDolar),EgresoDolar=sum(x.EgresoDolar)                         
       Into #tmp_saldo 
from 
(SELECT mes=month(A.detrec_fechacancela),ano=year(A.detrec_fechacancela),
       A.detrec_tipocajabanco, 
       A.detrec_cajabanco1,detrec_numctacte=rtrim(ltrim(isnull(A.detrec_numctacte,''))),        
       ImporteSoles=sum(A.detrec_importesoles),ImporteDolar=sum(A.detrec_importedolares), 
       IngresoSoles=case when  upper(B.cabrec_ingsal)='I' then sum(A.detrec_importesoles) else 0 end, 
       EgresoSoles=case when  upper(B.cabrec_ingsal)='E' then sum(A.detrec_importesoles) else 0 end, 
       IngresoDolar=case when  upper(B.cabrec_ingsal)='I' then sum(A.detrec_importedolares) else 0 end,  
       EgresoDolar=case when  upper(B.cabrec_ingsal)='E' then sum(A.detrec_importedolares) else 0 end 

FROM deleted  A,dbo.te_cabecerarecibos B
WHERE A.cabrec_numrecibo=B.cabrec_numrecibo and isnull(A.detrec_estadoreg,0) <> '1'
group by month(A.detrec_fechacancela),year(A.detrec_fechacancela) ,A.detrec_tipocajabanco,
               A.detrec_cajabanco1,rtrim(ltrim(isnull(A.detrec_numctacte,''))),B.cabrec_ingsal ) as X
group by X.ano,X.mes,X.detrec_tipocajabanco,X.detrec_tipocajabanco,
               X.detrec_cajabanco1,detrec_numctacte


/*Actualzo primero 
  Primero es la actualizacion por que se al actualizar no encuentra esos registron que se estan ingresando o actualizando 
  se insertan en la tabla control de saldos

 pero si se insertara primero entonces si no lo encuentra lo insertan tambien
 pero luego lo actulizaria y la cifra de los saldos no serian los correctos

*/

update  dbo.te_controlasaldos
set ctrlsaldo_saldodispoingre=B.ctrlsaldo_saldodispoingre+A.IngresoSoles,
    ctrlsaldo_saldodisposalida=B.ctrlsaldo_saldodisposalida+A.EgresoSoles
From #tmp_saldo A,dbo.te_controlasaldos B 
where A.ano=B.ctrlsaldo_año and 
      A.mes=B.ctrlsaldo_mes and   
      A.detrec_tipocajabanco=B.ctrlsaldo_tipobc and  
      A.detrec_cajabanco1=B.ctrlsaldo_bancocaja and        
      A.detrec_numctacte=rtrim(ltrim(isnull(B.ctrlsaldo_numectacte,''))) and  ctrlsaldo_mon='01'
      

update  dbo.te_controlasaldos
set ctrlsaldo_saldodispoingre=B.ctrlsaldo_saldodispoingre+A.IngresoDolar,
    ctrlsaldo_saldodisposalida=B.ctrlsaldo_saldodisposalida+A.EgresoDolar
From #tmp_saldo A,dbo.te_controlasaldos B 
where A.ano=B.ctrlsaldo_año and 
      A.mes=B.ctrlsaldo_mes and   
      A.detrec_tipocajabanco=B.ctrlsaldo_tipobc and  
      A.detrec_cajabanco1=B.ctrlsaldo_bancocaja and        
      A.detrec_numctacte=rtrim(ltrim(isnull(B.ctrlsaldo_numectacte,'')))  and  ctrlsaldo_mon='02'

/*Fin de la actualizacion*/




