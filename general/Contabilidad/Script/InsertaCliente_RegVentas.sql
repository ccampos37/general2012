/*Primera Parte*/
INSERT INTO [CONTAPRUEBA].dbo.ct_entidad
(entidadcodigo, entidadrazonsocial, entidaddireccion, 
 entidadruc, entidadtelefono, entidadtipocontri,
 usuariocodigo, fechaact)
--Verificar luego los proveedores que no tenga ruc y que el sistema
--de un mensaje de cuales son
SELECT DISTINCT 
	entidadcodigo=Left(case when A.pedidotipofac='01' 
         then case when isnull(rtrim(ltrim(A.clienteruc)),'') ='' then A.clientecodigo else A.clienteruc end 
         else A.clientecodigo End,11), 
    entidadrazonsocial=isnull((select top 1 clienterazonsocial from [TRANSFER].dbo.Vt_cliente Cli where Cli.clientecodigo=A.clientecodigo),'No Tiene'),
    entidaddireccion=' ',
    entidadruc=Left(case when A.pedidotipofac='01' 
         then case when isnull(rtrim(ltrim(A.clienteruc)),'') ='' then A.clientecodigo else A.clienteruc end 
         else A.clientecodigo End,11),
    entidadtelefono=' ',
    entidadtipocontri='00',usuariocodigo='Sys',fechaact=Getdate(),
	 B.fechaact

FROM [TRANSFER].dbo.Vt_Pedido A 
     Left Outer Join [CONTAPRUEBA].dbo.ct_entidad B
     On rtrim(ltrim(Left(case when A.pedidotipofac='01' 
         then case when isnull(rtrim(ltrim(A.clienteruc)),'') ='' then A.clientecodigo else A.clienteruc end 
         else A.clientecodigo End,11)))
        =rtrim(ltrim(B.entidadcodigo)) 
WHERE Month(A.pedidofecha)=03 and year(A.pedidofecha)=2003 and 
		B.fechaact is null and A.pedidotipofac <> '80'
order by 1

/*Segunda Parte*/
INSERT INTO [CONTAPRUEBA].dbo.ct_analitico
(analiticocodigo, entidadcodigo, tipoanaliticocodigo, usuariocodigo, fechaact)

SELECT DISTINCT
      analiticocodigo=Left(case when A.pedidotipofac='01' 
         then case when isnull(rtrim(ltrim(A.clienteruc)),'') ='' then A.clientecodigo else A.clienteruc end 
         else A.clientecodigo End,11)
                +'002',
      entidadcodigo=Left(case when A.pedidotipofac='01' 
       then case when isnull(rtrim(ltrim(A.clienteruc)),'') ='' then A.clientecodigo else A.clienteruc end 
       else A.clientecodigo End,11),
       tipoanaliticocodigo='002',usuariocodigo='Sys',fechaact=getdate()
FROM   [TRANSFER].dbo.Vt_Pedido A 
       Left Outer Join  [CONTAPRUEBA].dbo.ct_analitico B
       On Left(case when A.pedidotipofac='01' 
       then case when isnull(rtrim(ltrim(A.clienteruc)),'') ='' then A.clientecodigo else A.clienteruc end 
       else A.clientecodigo End,11)+'002'
          =B.analiticocodigo 
WHERE Month(A.pedidofecha)=03 and year(A.pedidofecha)=2003 and 
      B.fechaact is null AND  A.pedidotipofac <> '80'



--select * from [CONTAPRUEBA].dbo.ct_entidad where entidadcodigo='10292773183'
select * from transfer.dbo.vt_cliente where clienteruc='10081253531'
select * from transfer.dbo.vt_pedido where clientecodigo='2110'
delete from transfer.dbo.vt_cliente where clienteruc='10081253531' and clientecodigo='2114'