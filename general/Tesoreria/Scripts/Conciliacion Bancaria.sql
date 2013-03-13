use ventas_prueba
/*Se crea un campo para marcar cuando esta conciliado*/
Alter table dbo.te_detallerecibos 
add chkconcil bit

/*La fecha en que se conciliacion*/
Alter table dbo.te_detallerecibos 
add fechconcil datetime
/**/

Alter table  dbo.te_controlasaldos
add ctrlsaldo_saldobanco numeric(20,3)

Alter table  dbo.te_controlasaldos
add ctrlsaldo_ingresobanco numeric(20,3)

Alter table  dbo.te_controlasaldos
add ctrlsaldo_egresobanco numeric(20,3)

Alter table dbo.te_controlasaldos 
add ctrlsaldo_año varchar(4)

Alter table dbo.te_controlasaldos 
add ctrlsaldo_mes varchar(2)

Alter table dbo.te_controlasaldos
add ctrlsaldo_mon varchar(2)

SELECT * FROM SYSFILES

ctrlsaldo_bancocaja
ctrlsaldo_numectacte



80 - Transferencias 
90 - Transferencias

select 
	A.chkconcil,
	Anno=year(A.detrec_fechacancela),Mes=month(A.detrec_fechacancela), 	
 	A.cabrec_numrecibo,A.detrec_emisioncheque,A.detrec_tipodoc_concepto,
        A.detrec_numdocumento,B.cabrec_ingsal,A.detrec_tipocajabanco, 
        A.detrec_numctacte,A.detrec_monedadocumento,
        A.detrec_importesoles,A.detrec_importedolares,A.detrec_monedacancela,
        A.detrec_tdqc,A.detrec_ndqc,A.detrec_fechacancela,B.cabrec_estadoreg,
        B.cabrec_fechadocumento    
       
from te_detallerecibos A
Inner join te_cabecerarecibos  B
      on A.cabrec_numrecibo=B.cabrec_numrecibo 
where 
     A.detrec_emisioncheque='B' and 
     A.detrec_tipocajabanco='B'	and 
     ltrim(rtrim(Isnull(A.detrec_numctacte,'')))  <>'' and 
     B.cabrec_estadoreg <> 1 and isnull(chkconcil,0)=0 
order by detrec_fechacancela

select * from dbo.te_cuentabancos
cbanco_numero,monedacodigo,cbanco_codigo 

select * from dbo.gr_banco
bancocodigo,bancodescripcion 

select * from dbo.te_detallerecibos
select  A.chkconcil,  Anno=year(A.detrec_fechacancela),Mes=month(A.detrec_fechacancela),  A.cabrec_numrecibo,A.detrec_emisioncheque,A.detrec_tipodoc_concepto,  A.detrec_numdocumento,B.cabrec_ingsal,A.detrec_tipocajabanco,  A.detrec_numctacte,A.detrec_monedadocumento,  A.detrec_importesoles,A.detrec_importedolares,A.detrec_monedacancela,  A.detrec_tdqc,A.detrec_ndqc,A.detrec_fechacancela,B.cabrec_estadoreg,  B.cabrec_fechadocumento,A.detrec_observacion  from te_detallerecibos A  Inner join te_cabecerarecibos  B  on A.cabrec_numrecibo=B.cabrec_numrecibo  Where    ltrim(rtrim(Isnull(A.detrec_numctacte,'')))  <>'' and  B.cabrec_estadoreg <> 1 and ltrim(rtrim(Isnull(A.detrec_numctacte,'')))='191-1098578-0-45' order by A.detrec_fechacancela


ctrlsaldo_saldocontaingre, ctrlsaldo_saldocontasalida   

dbo.Te_SaldoIni

select * from dbo.te_controlasaldos

/*Store para el reporte */

Create proc te_concilbanc
--Declare 
@Base   varchar(50), 
@cuenta varchar(50),
@concil varchar(1)

as 
/* Set @Base='Ventas_Prueba'
 Set @cuenta='191-1109347-1-34' 
 set @concil=1 */

Declare @Sqlcad varchar(4000),@Sqlvar varchar(1000)


Set @Sqlcad='
select 
      A.chkconcil, 
      Anno=year(A.detrec_fechacancela),Mes=month(A.detrec_fechacancela),  A.cabrec_numrecibo,
      A.detrec_emisioncheque,A.detrec_tipodoc_concepto,  A.detrec_numdocumento,B.cabrec_ingsal,
      A.detrec_tipocajabanco,  A.detrec_numctacte,A.detrec_monedadocumento,
      A.detrec_importesoles,A.detrec_importedolares,A.detrec_monedacancela,
      A.detrec_tdqc,A.detrec_ndqc,A.detrec_fechacancela,B.cabrec_estadoreg,
      B.cabrec_fechadocumento,A.detrec_observacion 
from ['+@Base+'].dbo.te_detallerecibos A 
Inner join  ['+@Base+'].dbo.te_cabecerarecibos  B  on 
      A.cabrec_numrecibo=B.cabrec_numrecibo 
Where A.detrec_emisioncheque=''B'' and  A.detrec_tipocajabanco=''B'' and 
      ltrim(rtrim(Isnull(A.detrec_numctacte,'''')))  <>'''' and  B.cabrec_estadoreg <> 1 
      and ltrim(rtrim(Isnull(A.detrec_numctacte,'''')))='''+@cuenta+''' ' 

If @concil='1' set @Sqlvar=' and isnull(chkconcil,0)=1 '
If @concil='0' set @Sqlvar=''
 
Exec(@Sqlcad+@Sqlvar+' order by A.detrec_fechacancela ')



ctrlsaldo_tipobc
ctrlsaldo_saldobanco + 
(ctrlsaldo_ingresobanco - ctrlsaldo_egresobanco)





