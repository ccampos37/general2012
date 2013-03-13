SELECT A.detcomprobfechaemision,A.cabcomprobnumero,A.documentocodigo, 
    A.detcomprobnumdocumento,A.tipdocref,A.detcomprobnumref,A.detcomprobglosa,A.detcomprobtipocambio, 
    A.detcomprobussdebe-A.detcomprobusshaber as ComprobUSS,A.detcomprobdebe,A.detcomprobhaber,
    SaldoDebe=C.saldodebe00,
    SaldoHaber=C.saldohaber00,
	SaldoIni=(C.saldodebe00-C.saldohaber00),SaldoFin=A.detcomprobdebe-A.detcomprobhaber,
    A.cuentacodigo,
    B.cuentadescripcion,
    A.monedacodigo,
    Cuenta2=left(A.cuentacodigo,2)
    FROM  
		[PRUEBA_contaprueba_SANIL].dbo.[ct_detcomprob2002] A, 
		[PRUEBA_contaprueba_SANIL].dbo.[ct_cuenta] B, 
    	[PRUEBA_contaprueba_SANIL].dbo.[ct_saldos2002] C

    WHERE A.cuentacodigo = B.cuentacodigo AND
       A.cuentacodigo = C.cuentacodigo AND
       A.cuentacodigo like '104115' AND
       A.cabcomprobmes>='01'  and A.cabcomprobmes<='01' 
       ORDER BY A.cuentacodigo
