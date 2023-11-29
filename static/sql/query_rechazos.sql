SELECT 
DTTAPC "Código aplicación",    
DTBRCH "Sucursal libros",    
DTENBR "Sucursal ingreso",    
DTDATE "Fecha efectiva",    
DTTXCD "Código TRN",    
DTDBCR "Código DB o CR",  
  
case
    when DTDBCR = 'C' then -DTAMT
    else DTAMT
end as "Monto",

DTUPST "Código devolución",    
DTACCT "Número cuenta",   
DTPSDT "Fecha vinculación",    
DTSERL "Número cheque o serie",    
TRIM(DFTRDS) "Descripción TRN",
(SDSDODSP - SDRETACT - SDVLRCC - SDVLREMB) AS "Saldo disponible TR",
SDTIPOSOB "Estado cupo sobregiro",
SDTPESOB "Cupo sobregiro",
DMSTAT "Estado cuenta",
SDESTADO "Estado cuenta TR",
DFRRCD "Código TRN rechazo",
DTPODÑ "Código rastreo",
CNCDTI "Tipo identificación",
CNNOSS "Número identificación",        
CNNAME "Nombre",        
CNCDBI "Segmento",       
CNCDTY "Tipo cliente"
FROM BVDLIBT.DTRNPMMDD
LEFT JOIN VISIONR.DTRCD
ON DTTAPC = DFRAPC AND DTTXCD = DFRNCD
LEFT JOIN SCILIBRAMD.SCIFFSALDO
ON DTTAPC =  SDTIPOCTA AND DTACCT = SDCUENTA      
LEFT JOIN VISIONR.DBAL
ON DTTAPC = DMAPCD AND DTACCT = DMACCT
LEFT JOIN VISIONR.CXREF
ON DMAPCD = CXCDAP AND DMACCT = CXNOAC
AND CXCDAR = 'T' 
LEFT JOIN VISIONR.CNAME
ON CXNAMK = CNNAMK
WHERE DTUPST <> ' '
