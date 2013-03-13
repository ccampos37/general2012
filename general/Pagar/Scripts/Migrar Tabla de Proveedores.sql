--select * from ct070400 where ana_grupo='P'
--select * from cp_proveedor

select * from xxx_vt_cliente

insert cp_proveedor (clientecodigo,clienteruc,clientedireccion,clienterazonsocial,
	clientetipopais,clientetipopersona,clientefechaactivacion,usuariocodigo,fechaact)
--ANA_ANALIS,ANA_RUC,ANA_DIRCOD,ANA_NOMCOD
select ANA_ANALIS,ANA_RUC,ANA_DIRCOD,ANA_NOMCOD,
 case when len(rtrim(isnull(ANA_RUC,'0')))=11 then '1' else '2' end ,
 case when len(rtrim(isnull(ANA_RUC,'0')))=11 then '2' else '1' end ,
 '01/11/2002',
'SISTEMA','05/11/2002' from ct070400 where ana_grupo='P'

update cp_proveedor set negociocodigo='01', estadoreg='0',clienteaval='0',clientediasmaxpagocont=0
update cp_proveedor set clientecodigo=clienteruc where clienteruc is not null and len(clienteruc)=11

select * from cp_cargo