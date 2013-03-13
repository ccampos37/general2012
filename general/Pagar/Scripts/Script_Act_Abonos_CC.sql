set nocount on
Declare @documentocargo varchar(2)
declare @cargoapeimpape float
Declare @cargonumdoc varchar(15)
DECLARE tablas CURSOR FOR 
--update vt_CARGO set cargoapeimppag=0,cargoapeflgcan='0'
SELECT documentocargo,cargonumdoc,cargoapeimpape from VT_CARGO WHERE cargoapeflgcan='0'
	OPEN tablas
	/* Leer cada registro del cursor  */
	FETCH NEXT FROM tablas INTO @documentocargo,@cargonumdoc,@cargoapeimpape

	WHILE @@FETCH_STATUS = 0
	BEGIN
        	declare @cadsql varchar(3000)
				
		set @cadsql=''
		set @cadsql='update VT_CARGO 
	           set cargoapeimppag=(select isnull(sum(isnull(abonocanimpsol,0)),0) from vt_abono 
		   	where documentoabono=''' +@documentocargo+ ''' and abononumdoc=''' +@cargonumdoc+''')
		   where documentocargo=''' +@documentocargo+ ''' and cargonumdoc=''' +@cargonumdoc+ ''''
		exec(@cadsql)
				
		set @cadsql=''
		set @cadsql='if round( cast( (select isnull(sum(isnull(abonocanimpsol,0)),0) from vt_abono 
		  where documentoabono=''' +@documentocargo+ ''' and abononumdoc=''' +@cargonumdoc +''') as varchar(15)),2) >=' + cast(@cargoapeimpape as varchar(15)) + 
			'  begin
			   update VT_CARGO set cargoapeflgcan=''1'' where documentocargo=''' +@documentocargo+ ''' and cargonumdoc=''' +@cargonumdoc+ ''' end'
		exec(@cadsql)
		

 	FETCH NEXT FROM tablas INTO @documentocargo,@cargonumdoc,@cargoapeimpape
    END
	CLOSE tablas
	DEALLOCATE tablas

set nocount off
--select * from vt_abono
--select * from vt_cargo
