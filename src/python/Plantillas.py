import sys
from Plantilla.terceros import *
from Plantilla.novedades import *
from Plantilla.liquidacion import *
from Plantilla.fileupload import *
from Access.Getaccess import *
import json

def controlador():
    
    json_string = sys.argv[1]
    json_object = json.loads(json_string)
    
    if(json_object['form'] == "TERCEROS"):
        
        record_id = json_object['ID']
        d_temporal = json_object['temporal']
        d_id_temporal = json_object['id_temporal']
        d_cliente = json_object['cliente']
        d_id_cliente = json_object['id_cliente']
        d_empleados = json_object['empleados']
        d_concepto = json_object['concepto']
        d_id_concepto = json_object['id_concepto']
        d_fecha = json_object['fecha']    
        d_tipo = "Prestamo"
        d_valor = 0
        d_realizacion_descuento = "Cada periodo"
        d_n_cuotas = 1
        d_modo_pago = "Pago nomina"
        d_estado_des_total = "Sin calcular"
        
        file_path = terceros(d_temporal,d_cliente,d_empleados,d_concepto,d_fecha,d_tipo,d_valor,d_realizacion_descuento,d_n_cuotas,d_modo_pago,d_estado_des_total)
        access_token = obtener_access_token()
        
        if access_token:
            if(access_token != None):
                updatedata(access_token,file_path,record_id)
    
    if(json_object['form'] == "NOVEDADES"):
        
        record_id = json_object['ID']
        d_temporal = json_object['temporal']
        d_id_temporal = json_object['id_temporal']
        d_cliente = json_object['cliente']
        d_id_cliente = json_object['id_cliente']
        d_empleados = json_object['empleados']
        d_concepto = json_object['concepto']
        d_id_concepto = json_object['id_concepto']
        d_periodo = json_object['periodo']
        d_id_periodo = json_object['id_periodo']
        d_valor = 0
        d_unidades = 0    
        
        file_path = novedades(d_temporal,d_cliente,d_empleados,d_periodo,d_concepto,d_valor,d_unidades)
        access_token = obtener_access_token()
        
        if access_token:
            if(access_token != None):
                updatedata(access_token,file_path,record_id)
        
    if(json_object['form'] == "LIQUIDACION"):
        
        record_id = json_object['ID']
        d_temporal = json_object['temporal']
        d_id_temporal = json_object['id_temporal']
        d_cliente = json_object['cliente']
        d_id_cliente = json_object['id_cliente']
        d_empleados = json_object['empleados']
        d_concepto = json_object['concepto']
        d_id_concepto = json_object['id_concepto']
        d_periodo = json_object['periodo']
        d_id_periodo = json_object['id_periodo']
        d_valor = 0
        d_unidades = 0    
        
        file_path = liquidacion(d_temporal,d_cliente,d_empleados,d_periodo,d_concepto,d_valor,d_unidades,d_id_temporal,d_id_cliente,d_id_concepto,d_id_periodo)    
        access_token = obtener_access_token()
        
        if access_token:
            if(access_token != None):
                updatedata(access_token,file_path,record_id)
    
controlador()