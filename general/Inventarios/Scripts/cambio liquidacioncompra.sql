select * into al_tempfile from tempfile
select * into al_gtempfile from gtempfile

select * into al_jtempo  from jtempo
select * into al_jdetatempo  from jdetatempo

---***

select * into al_liquidacioncompra from vt_pedido
select * into al_detalleliquidacioncompra from vt_detallepedido

select * into al_tempoliquidacioncompra01 from tempopedido01
delete al_tempoliquidacioncompra01
select * into al_tempoliquidacioncompra02 from tempopedido02
delete al_tempoliquidacioncompra02
select * into al_tempoliquidacioncompra03 from tempopedido03
delete al_tempoliquidacioncompra03

select * into al_tempodetalleliquidacioncompra01 from tempodetallepedido01
delete al_tempodetalleliquidacioncompra01
select * into al_tempodetalleliquidacioncompra02 from tempodetallepedido02
delete al_tempodetalleliquidacioncompra02
select * into al_tempodetalleliquidacioncompra03 from tempodetallepedido03
delete al_tempodetalleliquidacioncompra03
