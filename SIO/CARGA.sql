--CARGA DE CATALOGO DE PRODUCTOS
INSERT INTO TFPRODUC(
CLAPROVE,CLAFAMIL,DESCRIPC,NOMCORTO,CONTENID,FLETEX,FLESUB,MEDIDA,PAQUETES,TASAIEPS,BARRASPZA,BARRASCAJA,CONSEC,PEDIR
,FECACT,BAJA,MEDMAY,OTRSAB,COSTOPAQ,COSTOCAJ,EXICAJA,  EXIPZA , PESO, fechaintro ,  procedencia, familia ,precosto 
,ofertado, fechaini, fechafin, fechaactivo ,   activo, costototal , clavedelprov, actualizado ,IVA  ,  IEPS  , LINEA ,cajas  ,     encajas 
,  CARGO122,CARGO222,CARGO_EFECTIVO22,      FLETE_EFECTIVO22 ,     MANIOBRAS22 ,          preciolista,           preciomaximo   ,       usuario     
,observaciones ,                                     tasaiva ,    tipoproducto,         manual, interno, enpromo ,dectofin1 , observa1
, dectofin2,  observa2,                       dectofin3,  observa3 ,                      dectofin4,  observa4,dectofin5,  observa5
,  dectofin,   observa  ,                      igualofi ,largo,  ancho,  alto ,   volmtocub 
) 
SELECT  
CLAPROVE,CLAFAMIL,DESCRIPC,NOMCORTO,CONTENID,FLETEX,FLESUB,MEDIDA,PAQUETES,TASAIEPS,BARRASPZA,BARRASCAJA,CONSEC,PEDIR
,FECACT,BAJA,MEDMAY,OTRSAB,COSTOPAQ,COSTOCAJ,EXICAJA,  EXIPZA , PESO, fechaintro ,  procedencia, familia ,precosto 
,ofertado, fechaini, fechafin, fechaactivo ,   activo, costototal , clavedelprov, actualizado ,IVA  ,  IEPS  , LINEA ,cajas  ,     encajas 
,  CARGO122,CARGO222,CARGO_EFECTIVO22,      FLETE_EFECTIVO22 ,     MANIOBRAS22 ,          preciolista,           preciomaximo   ,       usuario     
,observaciones ,                                     tasaiva ,    tipoproducto,         manual, interno, enpromo ,dectofin1 , observa1
, dectofin2,  observa2,                       dectofin3,  observa3 ,                      dectofin4,  observa4,dectofin5,  observa5
,  dectofin,   observa  ,                      igualofi ,largo,  ancho,  alto ,   volmtocub 
FROM TFPRODUCT
UPDATE TFPRODUC SET MANUAL = 0

--CARGA DE CATALOGO DE PROVEEDORES
INSERT INTO CATPROV(
PROVE, NOMPROVE,DIRPRO,COLPRO , DELPRO, CODPRO, CIUPRO,  LOCPRO, TELPRO ,  frecuencia,  activo, tipo, comprador,  rfc,  backorder, visita,procedencia,usuario,PLAZOPAGO
,fechaactivo,razon,  dectofin1, observa1 , dectofin2, observa2 , dectofin3, observa3,dectofin4,observa4, dectofin5,observa5 , volumen,pagoneto, clamaster
)
SELECT 
PROVE, NOMPROVE,DIRPRO,COLPRO , DELPRO, CODPRO, CIUPRO,  LOCPRO, TELPRO ,  frecuencia,  activo, tipo, comprador,  rfc,  backorder, visita,procedencia,usuario,PLAZOPAGO
,fechaactivo,razon,  dectofin1, observa1 , dectofin2, observa2 , dectofin3, observa3,dectofin4,observa4, dectofin5,observa5 , volumen,pagoneto, clamaster
FROM CATPROVT  

-- TABLA DE PRECIOS
DELETE FROM PREPROD
INSERT INTO PREPROD(PRECLAVE) SELECT CONSEC FROM TFPRODUC
UPDATE PREPROD SET PRECIO5 = 0,PRECIO6 = 0,PRECIO2 = PRECAJA,PRECIO3 = PRECAJA,PRECIO4 = PRECAJA, PRECIO1 = PREPAQUE 
FROM PRECIOS WHERE PRECLAVE = PRECIOS.CONSEC
SELECT * FROM PREPROD

--TABLA DE CARGOS
DELETE FROM CARGOS
INSERT INTO CARGOS(CAPROD) SELECT CONSEC FROM TFPRODUC
UPDATE CARGOS SET CARGO1 = PORCARGO,CARGO2 = 0, IEPS = PRECIOS.IEPS, IVA = PRECIOS.IVA, cargo_efectivo = OTROSREC, flete_efectivo = FLETES, maniobras = FLESUB
FROM PRECIOS WHERE CAPROD = PRECIOS.CONSEC
SELECT * FROM CARGOS

--TABLA DE DESCUENTOS
DELETE FROM DESCUENTOS
INSERT INTO DESCUENTOS(DEPROD) SELECT CONSEC FROM TFPRODUC
UPDATE DESCUENTOS SET DECTO1 = DESCTO01, DECTO2 = DESCTO02,  DECTO3 = DESCTO03,  DECTOOFERTA = DESCTO04,  DECTO5 = DESCTO05,
dectoFinanciero = DESCEFEC , dectoefectivo = 0
FROM PRECIOS WHERE DEPROD = PRECIOS.CONSEC
SELECT * FROM DESCUENTOS


--TABLA DE MARGENES
DELETE FROM MARGEN
INSERT INTO MARGEN(PRODUCTO) SELECT CONSEC FROM TFPRODUC
UPDATE MARGEN SET ESCALA1 = GANANPAQ , ESCALA2 =GANANCAJ , ESCALA3 =GANANCAJ , ESCALA4 =GANANCAJ , ESCALA5 =0,  ESCALA6 =0
FROM PRECIOS WHERE PRODUCTO = PRECIOS.CONSEC
SELECT * FROM MARGEN ORDER BY ESCALA1 

DELETE FROM CAMBPRE
SELECT * FROM CATCLIENTE

SELECT *  FROM CARGOS WHERE IVA > 10
SELECT * FROM TFPRODUC WHERE CONSEC = 31025
UPDATE TFPRODUC SET NOMCORTO = DESCRIPC WHERE NOMCORTO IS NULL
SELECT * FROM TFPRODUC ORDER BY NOMCORTO

SELECT * FROM TFPRODUC