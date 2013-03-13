use contaprueba

select * from ct_detcomprob2003 where cabcomprobmes=5 
					and asientocodigo in ('071','072','073','074') and subasientocodigo='0005' 

select * from ct_cabcomprob2003 where cabcomprobmes=5 
					and asientocodigo in ('071','072','073','074') and subasientocodigo='0004' 

/*Eliminar los documentos del SubAsiento 0004 de los Asientos 071,072,073,073*/
delete from ct_cabcomprob2003 where cabcomprobmes=5 
					and asientocodigo in ('071','072','073','074') and subasientocodigo='0004' 

delete from ct_detcomprob2003 where cabcomprobmes=4 
					and asientocodigo in ('071','072','073','074') and subasientocodigo='0001' 


/*Restaurar el archivo de exportación a la Base Datos en Producción*/


/*Actualizar las Cabeceras*/
exec marfice.dbo.cc_actualizacab 'Contaprueba','2003','05','071'
exec marfice.dbo.cc_actualizacab 'Contaprueba','2003','05','072'
exec marfice.dbo.cc_actualizacab 'Contaprueba','2003','05','073'
exec marfice.dbo.cc_actualizacab 'Contaprueba','2003','05','074'