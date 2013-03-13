Select * from ct_strucganper order by EGP_TIPO,EGP_LINEA
SELECT * from ct_strucganper_tot

DELETE FROM USUARIO

delete from  ct_strucganper where EGP_TIPO='03'


alter table ct_strucganper alter column EGP_DESCRI varchar(100)
 
INSERT INTO ct_strucganper
(EGP_TIPO,EGP_LINEA,EGP_NIVEL,EGP_SIGNO1,EGP_DESCRI,EGP_SIGNO2,
 EGP_CUENTA,EGP_SALDO)
 Select '03','1','C',null,'FLUJOS DE EFECTIVO DE ACTIVIDADES DE OPERACIÓN',null,null,0 union all
 Select '03','2','D',null,'Cobranzas a Clientes',null,'121',0 union all
 Select '03','3','D',null,'Otros cobros de operación',null,'129',0 union all
 Select '03','4','D',null,'Pagos a proveedores',null,'421',0 union all 
 Select '03','5','D',null,'Remuneraciones y beneficios sociales pagados',null,'41',0 union all 
 Select '03','6','D',null,'Pago de tributos',null,'40',0 union all 
 Select '03','7','D',null,'Otros pagos de operacion',null,null,0 union all 
 Select '03','8','C',null,null,null,null,0 union all
 Select '03','9','D',null,'Efectivo neto obtenido de actividades de operacion',null,null,0 union all
 Select '03','10','C',null,null,null,null,0 union all
 Select '03','11','C',null,'FLUJOS DE EFECTIVO DE ACTIVIDADES DE INVERSION:',null,null,0 union all
 Select '03','12','D',null,'Venta de maquinaria y equipo',null,'76',0 union all
 Select '03','13','D',null,'Compra de maquinaria y equipo',null,'33',0 union all
 Select '03','14','D',null,'Compra de inversiones en valores',null,'31',0 union all
 Select '03','15','C',null,null,null,null,0 union all
 Select '03','16','D',null,'Efectivo neto usado en actividades de inversión',null,null,0 union all
 Select '03','17','C',null,null,null,null,0 union all
 Select '03','18','C',null,'FLUJOS DE EFECTIVO DE ACTIVIDADES DE FINANCIAMIENTO:',null,null,0 union all
 Select '03','19','D',null,'Aumento (Disminución) en sobregiros y préstamos bancarios',null,null,0 union all
 Select '03','20','D',null,'Efectivo neto (usado) obtenido en actividades de financiamiento',null,null,0 union all
 Select '03','21','C',null,null,null,null,0 union all
 Select '03','22','D',null,'AUMENTO NETO EN EFECTIVO',null,null,0 union all
 Select '03','23','D',null,'EFECTIVO AL COMIENZO DEL AÑO',null,null,0 union all
 Select '03','24','C',null,null,null,null,0 union all
 Select '03','25','D',null,'EFECTIVO AL FIN DEL AÑO',null,null,0 