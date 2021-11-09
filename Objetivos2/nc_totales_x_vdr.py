import pgconnection
from datetime import date


#Ingresa una conexion abierta usando pgconnection y un string con el query deseado. Retorna una tupla con los datos que trae esa query
def fetch_data(connection, query_string):
    cursor = connection.cursor()
    cursor.execute(query_string)
    data = cursor.fetchall()
    cursor.close()

    return data

#Convierte el formato del nombre y monto del vendedor a diccionario
def data_reformat(vdr, data):
    new_data = {}
    new_data[vdr] = round(float(data[0][0]),2)

    return new_data

#Ejecuta el query del listado de vendedores activos, y los vuelca en una lista
def list_vdrs():
    query = "SELECT ALIAS_1.NOMBRE ALIAS_1_NOMBRE FROM V_VENDEDOR ALIAS_0 LEFT OUTER JOIN V_PERSONA ALIAS_1 ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_1.ID WHERE ALIAS_0.BO_PLACE_ID = '{89C234D8-3F01-11D5-86AD-0080AD403F5F}' AND  ALIAS_0.ACTIVESTATUS = 0 AND   ALIAS_0.TIPOOBJETOESTATICO_ID IS NULL"
    conn = pgconnection.get_connection("BD")
    data = fetch_data(conn, query)
    vendedores = []
    for vendedor in enumerate(data):
        vendedores.append(vendedor[1][0])

    return vendedores

def main_process():
    vdrs = list_vdrs()
    query_1 = "SELECT SUM(ALIAS_0.SUBTOTALIMPORTE) FROM V_TRCREDITOVENTA ALIAS_0 LEFT OUTER JOIN V_FLAG ALIAS_4 ON ALIAS_0.FLAG_ID = ALIAS_4.ID WHERE ALIAS_0.BO_PLACE_ID = '{89C23513-3F01-11D5-86AD-0080AD403F5F}'   AND   0=0  AND  ALIAS_0.TIPOTRANSACCION_ID = '{9B9915C3-4FA6-11D5-B060-004854841C8A}' AND ALIAS_0.ESTADO = 'C'"
    totales_por_vendedor = {}
    conn = pgconnection.get_connection("BD")

    # Cicla el listado de vendedores, agregando su nombre al query, y ajusta automaticamente la fecha
    # para reconocer el mes en curso, hasta el dia de ejecucion.
    for vdr in vdrs:
        # Formato de fecha AAAAMMDD sin espacios ni caracteres intermedios
        date_from = str(date.today().replace(day=1)).replace("-","")
        date_until = str(date.today()).replace("-","")
        query_2 = f" AND ALIAS_0.NOMBREORIGINANTE = '{vdr}'"
        query_3 = f"AND ALIAS_0.FECHAACTUAL <= '{date_until}000000000' AND ALIAS_0.FECHAACTUAL >= '{date_from}000000000'"

        query = query_1 + query_2 + query_3
        formatted = fetch_data(conn, query)
        if formatted[0][0] == None: #Saltea el vendedor si no tiene ventas facturadas.
            continue
        totales_por_vendedor = totales_por_vendedor | data_reformat(vdr, formatted)
    print('nc_totales_x_vdr calculado')
    return totales_por_vendedor

# Siguiente porcion diseñada para probar cambios en el modulo.
# Si se ejecuta este modulo como principal, prueba sobre el vendedor declarado
if __name__ == '__main__':
    print(main_process()['VDR'])

