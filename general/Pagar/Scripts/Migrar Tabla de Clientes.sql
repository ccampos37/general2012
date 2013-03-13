select * from  Maestro_Clientes

select *  into bk_vt_cliente from vt_cliente

select * from vt_cliente
delete from vt_cliente

insert vt_cliente (clientecodigo,clienteruc,clientedireccion,clienterazonsocial,
    clientetelefono,clientefax,clientemail,
	clientetipopais,clientetipopersona,clientefechaactivacion,usuariocodigo,fechaact)
select ContactoNro,ContactoRUC,rtrim(ContactoDireccion),ltrim(ContactoNombres),
   cast(ContactoTelefonos as varchar(25)),cast(ContactoFaxes as varchar(10)),ContactoEMail,
   case when ltrim(ContactoPais)='PERU' then '1' else '2' end,
   case when len(rtrim(isnull(ContactoRUC,'0')))=11 then '2' else '1' end ,
   '01/11/2002',
   'SISTEMA','05/11/2002' from MaestroClientes

update vt_cliente set negociocodigo='01', estadoreg='0',clienteaval='0',clientediasmaxpagocont=0