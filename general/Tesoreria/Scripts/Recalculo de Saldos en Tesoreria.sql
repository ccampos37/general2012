Use VENTAS_PRUEBA


--Tabla A Ingresos soles
If exists(select * from tempdb.dbo.sysobjects where name='##te_recalcsaldos') 
	exec('Drop table ##te_recalcsaldos ')


Select IE=isnull(A.cabrec_ingsal,B.cabrec_ingsal),
       CajaoBanco=isnull(A.detrec_cajabanco1,B.detrec_cajabanco1), 
       numctacte=isnull(A.detrec_numctacte,B.detrec_numctacte),
       anno=isnull(A.Anno,B.Anno),
	   mes=isnull(A.mes,B.mes),
       mon=isnull(A.mon,B.mon),
       Impingresos=isnull(A.ImporteIngresos,0),
       ImpEgresos=isnull(B.ImporteEgresos,0),
       B.detrec_tipocajabanco
Into ##te_recalcsaldos

From 
(
Select 
	A.cabrec_ingsal,B.detrec_cajabanco1,B.detrec_numctacte,
    Anno=year(detrec_fechacancela),mes=month(detrec_fechacancela),
    mon='01',
    ImporteIngresos=sum(detrec_importesoles),B.detrec_tipocajabanco
From dbo.te_cabecerarecibos A  
Inner Join dbo.te_detallerecibos B
On A.cabrec_numrecibo=B.cabrec_numrecibo
Where A.cabrec_ingsal='I' 		
Group by A.cabrec_ingsal,B.detrec_cajabanco1,B.detrec_numctacte,
    year(b.detrec_fechacancela),month(b.detrec_fechacancela),B.detrec_tipocajabanco
Union all 
--Tabla A1 Ingresos Dolares
Select 
	A.cabrec_ingsal,B.detrec_cajabanco1,B.detrec_numctacte,
    Anno=year(detrec_fechacancela),mes=month(detrec_fechacancela),
    mon='02',
    ImporteIngresos=sum(detrec_importedolares),B.detrec_tipocajabanco
From dbo.te_cabecerarecibos A  
Inner Join dbo.te_detallerecibos B
On A.cabrec_numrecibo=B.cabrec_numrecibo
Where A.cabrec_ingsal='I' 		
Group by A.cabrec_ingsal,B.detrec_cajabanco1,B.detrec_numctacte,
    year(b.detrec_fechacancela),month(b.detrec_fechacancela),B.detrec_tipocajabanco
) as A 
Left outer join 
--Tabla B Egresos soles 
(
Select 
	A.cabrec_ingsal,B.detrec_cajabanco1,B.detrec_numctacte,
    Anno=year(detrec_fechacancela),mes=month(detrec_fechacancela),
    mon='01',
    ImporteEgresos=sum(detrec_importesoles),B.detrec_tipocajabanco
From dbo.te_cabecerarecibos A  
Inner Join dbo.te_detallerecibos B
On A.cabrec_numrecibo=B.cabrec_numrecibo
Where A.cabrec_ingsal='E' 		
Group by A.cabrec_ingsal,B.detrec_cajabanco1,B.detrec_numctacte,
    year(b.detrec_fechacancela),month(b.detrec_fechacancela),B.detrec_tipocajabanco   
Union all 
Select 
	A.cabrec_ingsal,B.detrec_cajabanco1,B.detrec_numctacte,
    Anno=year(detrec_fechacancela),mes=month(detrec_fechacancela),
    mon='02',
    ImporteEgresos=sum(detrec_importedolares),B.detrec_tipocajabanco
From dbo.te_cabecerarecibos A  
Inner Join dbo.te_detallerecibos B
On A.cabrec_numrecibo=B.cabrec_numrecibo
Where A.cabrec_ingsal='E' 		
Group by A.cabrec_ingsal,B.detrec_cajabanco1,B.detrec_numctacte,
    year(b.detrec_fechacancela),month(b.detrec_fechacancela),B.detrec_tipocajabanco
) as B

On A.detrec_cajabanco1=B.detrec_cajabanco1 and 
   A.detrec_numctacte=B.detrec_numctacte  and 
   A.mon=B.mon and A.anno=B.anno and A.mes=B.mes and A.detrec_tipocajabanco=B.detrec_tipocajabanco
order by A.anno, B.mes,isnull(A.detrec_cajabanco1,B.detrec_cajabanco1),
         isnull(A.detrec_numctacte,B.detrec_numctacte)


Select * From ##te_recalcsaldos

Update dbo.te_controlasaldos 
Set ctrlsaldo_saldocontaingre=B.Impingresos,
    ctrlsaldo_saldocontasalida=B.ImpEgresos  
from dbo.te_controlasaldos A,##te_recalcsaldos B 
where 
  A.ctrlsaldo_bancocaja=B.CajaoBanco and  
  A.ctrlsaldo_numectacte=B.numctacte and 
  A.ctrlsaldo_año=B.anno and 
  A.ctrlsaldo_mes=B.mes and 
  A.ctrlsaldo_mon=B.mon 


select * from dbo.te_controlasaldos

Alter table dbo.te_controlasaldos 
Alter Column ctrlsaldo_numectacte varchar(30)


Insert dbo.te_controlasaldos
(ctrlsaldo_bancocaja,ctrlsaldo_numectacte,ctrlsaldo_tipobc,
 ctrlsaldo_año, ctrlsaldo_mes,ctrlsaldo_mon, 
 ctrlsaldo_saldocontaingre,ctrlsaldo_saldocontasalida) 

Select CajaoBanco,numctacte,isnull(detrec_tipocajabanco,''),
       anno,mes,mon,Impingresos,ImpEgresos
From ##te_recalcsaldos

select * from dbo.te_controlasaldos
where ctrlsaldo_numectacte='011-350-0100008495-62'
   