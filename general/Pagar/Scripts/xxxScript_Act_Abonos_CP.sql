--update cp_CARGO set cargoapeimppag=0,cargoapeflgcan=0 where cargoapeflgcan='0'
set nocount on
Declare @documentocargo varchar(2)
declare @cargoapeimpape float
Declare @cargonumdoc varchar(15)
Declare @clientecodigo varchar(20)
DECLARE tablas CURSOR FOR 
SELECT documentocargo,cargonumdoc,cargoapeimpape,clientecodigo from cp_CARGO WHERE cargoapeflgcan='0'

	OPEN tablas
	/* Leer cada registro del cursor  */
	FETCH NEXT FROM tablas INTO @documentocargo,@cargonumdoc,@cargoapeimpape,@clientecodigo

	WHILE @@FETCH_STATUS = 0
	BEGIN
        	declare @cadsql varchar(3000)
			
		set @cadsql=''
		set @cadsql='update cp_CARGO 
	           set cargoapeimppag=(select isnull(sum(isnull(abonocanimpsol,0)),0) from cp_abono 
		   	where documentoabono=''' +@documentocargo+ ''' and abonocancli=''' +@clientecodigo+ ''' and abononumdoc=''' +@cargonumdoc+''')
		   where documentocargo=''' +@documentocargo+ ''' and clientecodigo=''' +@clientecodigo+ ''' and cargonumdoc=''' +@cargonumdoc+ ''''
		exec(@cadsql)
						
		set @cadsql=''
		set @cadsql='if round( cast( (select isnull(sum(isnull(abonocanimpsol,0)),0) from cp_abono 
		  where documentoabono=''' +@documentocargo+ ''' and abonocancli=''' +@clientecodigo+ ''' and abononumdoc=''' +@cargonumdoc +''') as varchar(15)),2) >=' + cast(@cargoapeimpape as varchar(15)) + 
			'  begin
			   update cp_CARGO set cargoapeflgcan=''1'' where documentocargo=''' +@documentocargo+ ''' and clientecodigo=''' +@clientecodigo+ ''' and cargonumdoc=''' +@cargonumdoc+ ''' end'
		exec(@cadsql)
		

 	FETCH NEXT FROM tablas INTO @documentocargo,@cargonumdoc,@cargoapeimpape,@clientecodigo
    END
	CLOSE tablas
	DEALLOCATE tablas
set nocount off

--select * from cp_abono
--select * into bk_cp_cargo from cp_cargo
--delete from cp_cargo
--insert cp_cargo select * from bk_cp_cargo
--select * from vt_cargo where cargoapeimpape<=cargoapeimppag and cargoapeflgcan=0

/*Para Actualizar Fecha de Cancelación*/
--select * from cp_abono where abonocanfecan<>abonocanfecpro
--update cp_abono set abonocanfecan=abonocanfecpla

--documentoabono abononumdoc abonocannumpag zonacodigo abonotipoplanilla vendedorcodigo abononumplanilla abonocanfecpla                                         abonocanfecpro                                         abonocanext                                           abonocancli abonocantcli abonocantdqc abonocanndqc abonocanmoneda abonocanimcan                                         abonocanforcan abonocanfecan                                          abonocancarabo abonocancuenta       abonocanbco abonocantipcam                                        abonocanflpres abonocanflreg abonocantmp01                                         abonocantmp02                                         abonocannumord abonocanmoncan abonocanimpcan                                        abonocanimpsol                                        usuariocodigo fechaact                                               
---------------- ----------- -------------- ---------- ----------------- -------------- ---------------- ------------------------------------------------------ ------------------------------------------------------ ----------------------------------------------------- ----------- ------------ ------------ ------------ -------------- ----------------------------------------------------- -------------- ------------------------------------------------------ -------------- -------------------- ----------- ----------------------------------------------------- -------------- ------------- ----------------------------------------------------- ----------------------------------------------------- -------------- -------------- ----------------------------------------------------- ----------------------------------------------------- ------------- ------------------------------------------------------ 
--01             00300000357 3              01         01                001            000140           2002-12-12 00:00:00.000                                2002-12-12 00:00:00.000                                NULL                                                  20382355751 NULL         07           00100000738  02             42.140000000000001                                    P              2002-12-12 00:00:00.000                                A              421101                           3.5                                                   1              NULL          NULL                                                  NULL                                                  NULL           02             42.140000000000001                                    42.140000000000001                                    CAMTEX        2002-12-12 00:00:00.000
--07             00100000738 1              01         01                001            000140           2002-12-12 00:00:00.000                                2002-12-12 00:00:00.000                                NULL                                                  20382355751 NULL         01           00300000357  02             42.140000000000001                                    P              2002-12-12 00:00:00.000                                A              421101                           3.5                                                   1              NULL          NULL                                                  NULL                                                  NULL           02             42.140000000000001                                    42.140000000000001                                    CAMTEX        2002-12-12 00:00:00.000