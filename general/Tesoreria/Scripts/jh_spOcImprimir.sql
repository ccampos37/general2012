CREATE  Procedure jh_spOcImprimir
@Orden		int,
@Cond			int
As
if @Cond=1	--Hilos
	Select a.*, b.*, d.HiloDescripcion as Descripcion, ContactoNombres=c.clienterazonsocial,
          ContactoDireccion=c.clientedireccion,ContactoTelefonos=c.clientetelefono,ContactoRuc=c.clienteruc 
          ,contactoFaxes=c.clientefax , 
	e.UnidadMedida as UAlma, e.UnidadMedida2 as UCompr,	Factor = case UnidadOperador when 'm' then unidaddimension else 1/unidaddimension end,
	f.PagoCondDescripcion, 	MonedaNomb = case a.moneda when 1 then 'Soles' else 'Dolares' end, g.Fecha as CronoFech, g.Cant as CronCant
	from  OrdenCompra a INNER JOIN
                      OrdenCompraDetalle b ON a.OrdenNro = b.OrdenNro INNER JOIN
                      dbo.[cp_proveedor] c ON c.clientecodigo = A.ContactoNro INNER JOIN
                      dbo.Unidades e ON b.Um = e.UnidadMedidaID INNER JOIN
                      dbo.[Maestro Hilos] d ON b.Codigo = d.HiloCodigo INNER JOIN
                      dbo.PagoCondicion f ON a.PagoCondId = f.PagoCondId LEFT OUTER JOIN
                      OrdenCompraCronograma g ON b.OrdenNro = g.OrdenNro AND b.Item = g.Item
	Where a.OrdenNro = @Orden

if @Cond=2	--Tela Cruda
	Select a.*, b.*, d.TelaCrudaDescripcion as Descripcion, ContactoNombres=c.clienterazonsocial,
          ContactoDireccion=c.clientedireccion,ContactoTelefonos=c.clientetelefono,ContactoRuc=c.clienteruc 
          ,contactoFaxes=c.clientefax,e.UnidadMedida as UAlma, e.UnidadMedida2 as UCompr,	Factor = case UnidadOperador when 'm' then unidaddimension else 1/unidaddimension end,
	f.PagoCondDescripcion, 	MonedaNomb = case a.moneda when 1 then 'Soles' else 'Dolares' end, g.Fecha as CronoFech, g.Cant as CronCant
	from  OrdenCompra a INNER JOIN
                      OrdenCompraDetalle b ON a.OrdenNro = b.OrdenNro INNER JOIN
                      dbo.[cp_proveedor] c ON C.clientecodigo = A.ContactoNro INNER JOIN
                      dbo.Unidades e ON b.Um = e.UnidadMedidaID INNER JOIN
                      dbo.[Maestro Tela Cruda] d ON b.Codigo = d.TelaCrudaId INNER JOIN
                      dbo.PagoCondicion f ON a.PagoCondId = f.PagoCondId LEFT OUTER JOIN
                      OrdenCompraCronograma g ON b.OrdenNro = g.OrdenNro AND b.Item = g.Item
	Where a.OrdenNro = @Orden

if @Cond=5	--Quimico
	Select a.*, b.*, d.QuimicoDescripcion as Descripcion,ContactoNombres=c.clienterazonsocial,
          ContactoDireccion=c.clientedireccion,ContactoTelefonos=c.clientetelefono,ContactoRuc=c.clienteruc 
          ,contactoFaxes=c.clientefax, 
	e.UnidadMedida as UAlma, e.UnidadMedida2 as UCompr,	Factor = case UnidadOperador when 'm' then unidaddimension else 1/unidaddimension end,
	f.PagoCondDescripcion, 	MonedaNomb = case a.moneda when 1 then 'Soles' else 'Dolares' end, g.Fecha as CronoFech, g.Cant as CronCant
	from  OrdenCompra a INNER JOIN
                      OrdenCompraDetalle b ON a.OrdenNro = b.OrdenNro INNER JOIN
                      dbo.[cp_proveedor] c ON C.clientecodigo = A.ContactoNro INNER JOIN
                      dbo.Unidades e ON b.Um = e.UnidadMedidaID INNER JOIN
                      dbo.[Maestro Quimicos] d ON b.Codigo = d.QuimicoId INNER JOIN
                      dbo.PagoCondicion f ON a.PagoCondId = f.PagoCondId LEFT OUTER JOIN
                      OrdenCompraCronograma g ON b.OrdenNro = g.OrdenNro AND b.Item = g.Item
	Where a.OrdenNro = @Orden



if @@Error<>0 goto ErProc
	return(0)
ErProc:
	return(1)

