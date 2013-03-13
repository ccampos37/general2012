
/*Reporte de Movimientos */ /* sustento de movimientos diversos*/
select * from dbo.te_cabecerarecibos

select * from te_detallerecibos
concepto=detrec_tipodoc_concepto
tipomov=cabrec_ingsal
transfer=cabrec_transferenciaautomatico(1)
--si utiliza el mismo reporte


/*Consistencia de movimientos*/
obejtivos       : Imprimir la cabecera con sus detalles
criterio        :rango fechas
campos a mostrar:
Mismo formato del voucher.

/*Antiguedad de Deudas */ cuentas por pagar
Objetivos  : Consolidacion de deudas por proveedor (Igual aviso de cobranzas )
             Totalizado por cada proveedor
Criterios  : Adicional TC y Moneda de Presentacion

Tablas     : cp_cargo ,
Campos a mostrar : Proveedor,total en moneda de origen,moneda presentacion 
2 presentaciones 
  - Un resumen por proveedor (Resumido)
  - Un resumen por provedor y tipo de documento (normal)

/*Documentos vencidos y por vencer*/ cuentas por pagar
los documentos del dia son vencidos y por vencer
no hay todos vecer o vencidos

Criterio : intervalo de dias de vencimiento o por vencer,proveedores
campos   : cliente,td ,descripcion corta tipo doc , nrdoc,fechaemi,fechvcto,saldo,soles,dolares
tabla    : cp_cargo             

/*Relacion de Documentos*/
Criterio : El rango de fechas es con respectos a la fecha de emision del documentos 
Tabla : cp_cargo
misma presentacion del reporte que esta en ese formulario.
