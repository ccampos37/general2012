
INSERT INTO gr_moneda
    (monedacodigo, monedadescripcion, monedaabreviatura, 
    monedasimbolo, usuariocodigo, fechaact)
SELECT '00', '(Ninguno)', '(Ninguno)', '(Ninguno)','Sistema', '20/08/2002' UNION ALL
SELECT '01', 'MONEDA NACIONAL', 'SOLES', 'S/.','Sistema', '20/08/2002' UNION ALL
SELECT '02', 'MONEDA EXTRANJERA', 'DOLAR', 'USD','Sistema', '20/08/2002' 


INSERT INTO ct_operacion
    (operacioncodigo, operaciondescripcion, usuariocodigo, 
    fechaact)
VALUES ('00', '(Ninguno)', 'Sistema', '20/08/2002')

go

INSERT INTO ct_centrocosto
    (centrocostocodigo, centrocostodescripcion, 
    centrocostodescrcorta, centrocostotipo, usuariocodigo, 
    fechaact)
VALUES ('00', '(Ninguno)', '(Ninguno)', 'N', 'Sistema', '20/08/2002')

go

INSERT INTO ct_tipoanalitico
    (tipoanaliticocodigo,tipoanaliticodescripcion,usuariocodigo,fechaact)
VALUES ('00','(Ninguno)','Sistema','20/08/2002')

go

INSERT INTO ct_entidad
    (entidadcodigo,entidadrazonsocial,entidaddireccion,entidadtelefono,entidadruc,usuariocodigo,fechaact)
VALUES ('00','(Ninguno)','(Ninguno)','(Ninguno)','(Ninguno)','Sistema','20/08/2002')

go

INSERT INTO ct_analitico
    (analiticocodigo,entidadcodigo,tipoanaliticocodigo,usuariocodigo,fechaact)
VALUES ('00','00','00','Sistema','20/08/2002')

go



INSERT INTO ct_estcomprob
   (estcomprobcodigo,estcomprobdescripcion,usuariocodigo,fechaact)
SELECT '01','ERRADO','Sistema','20/08/2002' UNION ALL
SELECT '02','REGISTRADO','Sistema','20/08/2002' UNION ALL
SELECT '03','PROCESADO','Sistema','20/08/2002'

go

INSERT INTO gr_documento
    (documentocodigo, documentodescripcion, documentoregcompras,documentoregventas,
     documentoregletrasxcobrar, documentoregletrasxpagar, documentonotacredito, 
     usuariocodigo,fechaact)
VALUES ('00', '(Ninguno)',0,0,0,0,0,'Sistema','20/08/2002')


--delete from ct_subasiento
--delete from ct_asiento
--delete from ct_operacion
--delete from ct_analitico
--delete from ct_tipoanalitico
--delete from ct_cuenta